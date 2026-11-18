import copy
import json
from pathlib import Path

import pytest
from helpers import FakeBridge, connect_ppt, text_of

from deskhand.core.apply_token import TokenTable
from deskhand.core.powershell import BridgeError
from deskhand.ppt.server import BRIDGE_SCRIPT, INSTRUCTIONS
from deskhand.ppt.tools import PptDeps, WritePlan, _run_write

pytestmark = pytest.mark.anyio

CONTRACT = json.loads(
    (Path(__file__).parent.parent / "parity" / "fixtures" / "ppt" / "tools.json").read_text("utf-8")
)

TOOLS = [
    "presentations",
    "read_deck",
    "read_slide",
    "render",
    "create_deck",
    "add_slide",
    "add_text",
    "add_table",
    "add_chart",
    "add_picture",
    "add_shapes",
    "set_notes",
    "update_shape",
    "delete_shape",
    "move_slide",
    "delete_slide",
    "save_deck",
    "close_deck",
]


def test_the_bridge_has_strict_readers_snapshots_and_fallback_parts():
    script = BRIDGE_SCRIPT.read_text("utf-8")
    for name in (
        "Get-ShapeTextStrict",
        "Get-ShapeNameStrict",
        "Get-ShapeParagraphsStrict",
        "Get-ShapeFillStrict",
        "Get-ShapeLineStrict",
        "Get-ShapeTableStrict",
        "Get-ShapeChartStrict",
        "Get-ShapePictureStrict",
        "Get-NotesTextStrict",
        "Get-SlideLayoutStrict",
        "Get-SlideTitleStrict",
        "Get-SlideShapesStrict",
        "Get-ShapeSnapshot",
        "Get-SlideSnapshot",
        "Get-DeckSnapshot",
    ):
        assert f"function {name}" in script
    for fallback in ("fill|?", "line|?", "table|?", "chart|?", "picture|?", "name|?"):
        assert fallback in script
    assert "fill|none" in script and "line|none" in script
    assert "Get-DeckSnapshot $pres $true $indexes $includeNotes" in script
    assert "Get-SlideSnapshot $slide $true $includeNotes" in script


PRESENTATION = {
    "name": "Deck",
    "path": "C:\\tmp\\deck.pptx",
    "slides": 2,
    "widthPx": 1280,
    "heightPx": 720,
    "saved": True,
}
SHAPE = {
    "id": 4,
    "name": "TextBox 1",
    "type": "textBox",
    "x": 96,
    "y": 64,
    "w": 1088,
    "h": 96,
    "text": "hello",
}
SLIDE = {"index": 1, "layout": "blank", "title": "", "shapes": [SHAPE]}
SLIDE_WITH_NOTES = {**SLIDE, "notes": "speaker note"}


def write_handler(**extra):
    def handle(args):
        data = {
            "applied": args.get("apply") is True,
            "fingerprint": "fp-1",
            "fingerprintQuality": "full",
            "slide": SLIDE,
            "presentation": PRESENTATION,
        }
        data.update(extra)
        return data

    return handle


def deck_handler(args):
    return {
        "presentation": PRESENTATION,
        "slides": [SLIDE, {**SLIDE, "index": 2}],
        "fingerprint": "deck-fp",
        "fingerprintQuality": "full",
    }


async def test_list_tools_and_instructions():
    async with connect_ppt(
        FakeBridge({"presentations": lambda args: {"running": False, "presentations": []}})
    ) as t:
        listed = await t.client.list_tools()
        assert [tool.name for tool in listed.tools] == TOOLS
        by_name = {tool.name: tool for tool in listed.tools}
        read_deck = by_name["read_deck"].annotations
        delete_slide = by_name["delete_slide"].annotations
        assert read_deck is not None and read_deck.read_only_hint is True
        assert delete_slide is not None and delete_slide.destructive_hint is True
        assert by_name["render"].annotations is not None
        assert by_name["render"].annotations.destructive_hint is True
        assert "applyToken" in (by_name["add_text"].description or "")
        assert "expectFingerprint" in INSTRUCTIONS


async def test_the_advertised_contract_matches_the_recorded_fixture():
    async with connect_ppt(FakeBridge()) as t:
        listed = await t.client.list_tools()
        actual = [tool.model_dump(mode="json", by_alias=True, exclude_none=True) for tool in listed.tools]
        assert t.client.instructions == CONTRACT["instructions"] == INSTRUCTIONS
    assert [tool["name"] for tool in actual] == [tool["name"] for tool in CONTRACT["tools"]]
    for got, want in zip(actual, CONTRACT["tools"], strict=True):
        assert got == want, want["name"]


async def test_presentations_when_powerpoint_is_not_running():
    bridge = FakeBridge({"presentations": lambda args: {"running": False, "presentations": []}})
    async with connect_ppt(bridge) as t:
        result = await t.call("presentations")
    assert not result.is_error
    assert "not running" in text_of(result)
    assert result.structured_content == {"running": False, "presentations": []}
    assert bridge.calls[0]["timeout"] == 10.0


async def test_read_deck_converts_a_wsl_path_and_passes_notes():
    bridge = FakeBridge({"read_deck": deck_handler})
    async with connect_ppt(bridge) as t:
        result = await t.call(
            "read_deck",
            {"presentation": "/mnt/c/tmp/deck.pptx", "slides": "1-2", "includeNotes": True},
        )
    assert not result.is_error, text_of(result)
    call = bridge.calls[0]
    assert call["op"] == "read_deck"
    assert call["args"]["presentation"] == "C:\\tmp\\deck.pptx"
    assert call["args"]["slides"] == "1-2"
    assert call["args"]["includeNotes"] is True
    assert call["timeout"] == 10.0
    assert "slide 1" in text_of(result) and "slide 2" in text_of(result)


async def test_read_deck_rejects_a_path_powerpoint_cannot_reach():
    async with connect_ppt(FakeBridge()) as t:
        result = await t.call("read_deck", {"presentation": "/home/me/deck.pptx"})
    assert result.is_error
    assert "Windows drive" in text_of(result)


async def test_read_slide_returns_the_fingerprint():
    bridge = FakeBridge(
        {
            "read_slide": lambda args: {
                "slide": SLIDE_WITH_NOTES,
                "fingerprint": "fp-1",
                "fingerprintQuality": "full",
            }
        }
    )
    async with connect_ppt(bridge) as t:
        result = await t.call("read_slide", {"slide": 1, "includeNotes": True})
    assert not result.is_error, text_of(result)
    assert result.structured_content["fingerprint"] == "fp-1"
    assert "speaker note" in text_of(result)


async def test_read_slide_omits_an_unavailable_fingerprint():
    bridge = FakeBridge(
        {
            "read_slide": lambda args: {
                "slide": SLIDE,
                "fingerprintQuality": "unavailable",
            }
        }
    )
    async with connect_ppt(bridge) as t:
        result = await t.call("read_slide", {"slide": 1})
    assert not result.is_error, text_of(result)
    assert "fingerprint" not in result.structured_content
    assert result.structured_content["fingerprintQuality"] == "unavailable"


async def test_read_deck_omits_an_unavailable_fingerprint():
    bridge = FakeBridge(
        {
            "read_deck": lambda args: {
                "presentation": PRESENTATION,
                "slides": [SLIDE],
                "fingerprintQuality": "unavailable",
            }
        }
    )
    async with connect_ppt(bridge) as t:
        result = await t.call("read_deck", {})
    assert not result.is_error, text_of(result)
    assert "fingerprint" not in result.structured_content
    assert result.structured_content["fingerprintQuality"] == "unavailable"


async def test_render_converts_out_dir_and_uses_a_long_timeout():
    bridge = FakeBridge(
        {
            "render": lambda args: {
                "files": [{"path": "C:\\tmp\\out\\slide-1.png", "widthPx": 1600, "heightPx": 900}],
                "elapsedMs": 12,
                "outDir": "C:\\tmp\\out",
            }
        }
    )
    async with connect_ppt(bridge) as t:
        result = await t.call("render", {"outDir": "/mnt/c/tmp/out", "slides": "1"})
    assert not result.is_error, text_of(result)
    assert bridge.calls[0]["args"]["outDir"] == "C:\\tmp\\out"
    assert bridge.calls[0]["timeout"] == 55.0
    assert "rendered 1 file" in text_of(result)


async def test_render_requires_out_dir():
    async with connect_ppt(FakeBridge()) as t:
        result = await t.call("render", {})
    assert result.is_error
    assert "outDir" in text_of(result)


async def test_render_rejects_pdf_sizing_instead_of_ignoring_it():
    async with connect_ppt(FakeBridge()) as t:
        result = await t.call("render", {"outDir": "/mnt/c/tmp/out", "format": "pdf", "width": 1600})
    assert result.is_error
    assert "PNG only" in text_of(result)


async def test_preview_rejects_an_unknown_fit_before_issuing_a_token():
    bridge = FakeBridge({"add_picture": write_handler()})
    async with connect_ppt(bridge) as t:
        result = await t.call(
            "add_picture",
            {"slide": 1, "x": 0, "y": 0, "w": 100, "h": 100, "path": "/mnt/c/a.png", "fit": "fill"},
        )
    assert result.is_error
    assert "fit" in text_of(result)
    assert bridge.calls == []


async def test_preview_rejects_a_bad_column_width_count():
    bridge = FakeBridge({"add_table": write_handler()})
    async with connect_ppt(bridge) as t:
        result = await t.call(
            "add_table",
            {
                "slide": 1,
                "x": 0,
                "y": 0,
                "w": 400,
                "h": 200,
                "header": ["a", "b"],
                "rows": [["1", "2"]],
                "columnWidths": [100],
            },
        )
    assert result.is_error
    assert "columnWidths" in text_of(result)
    assert bridge.calls == []


@pytest.mark.parametrize(
    ("tool", "args"),
    [
        (
            "add_text",
            {
                "slide": 1,
                "x": 0,
                "y": 0,
                "w": 100,
                "h": 100,
                "paragraphs": [{"runs": [{"text": "x", "size": 1}]}],
            },
        ),
        (
            "add_table",
            {"slide": 1, "x": 0, "y": 0, "w": 100, "h": 100, "rows": [["x"]], "fontSize": 1},
        ),
        (
            "update_shape",
            {
                "slide": 1,
                "shapeId": 4,
                "paragraphs": [{"runs": [{"text": "x", "size": 1}]}],
            },
        ),
    ],
)
async def test_a_font_size_below_one_point_is_refused_before_a_token(tool, args):
    bridge = FakeBridge({tool: write_handler()})
    async with connect_ppt(bridge) as t:
        result = await t.call(tool, args)
    assert result.is_error
    assert "at least 1 pt" in text_of(result)
    assert bridge.calls == []


async def test_update_shape_can_carry_paragraphs():
    bridge = FakeBridge({"update_shape": write_handler()})
    args = {
        "slide": 1,
        "shapeId": 4,
        "paragraphs": [{"runs": [{"text": "Hello", "bold": True}]}],
    }
    async with connect_ppt(bridge) as t:
        preview = await t.call("update_shape", args)
        assert not preview.is_error, text_of(preview)
        applied = await t.call(
            "update_shape",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert not applied.is_error, text_of(applied)
    assert bridge.calls[0]["args"]["paragraphs"][0]["runs"][0]["bold"] is True


async def test_update_shape_rejects_an_empty_change():
    async with connect_ppt(FakeBridge()) as t:
        result = await t.call("update_shape", {"slide": 1, "shapeId": 4})
    assert result.is_error
    assert "at least one" in text_of(result)


async def test_add_chart_uses_value_axis_title():
    bridge = FakeBridge({"add_chart": write_handler()})
    args = {
        "slide": 1,
        "x": 0,
        "y": 0,
        "w": 400,
        "h": 300,
        "type": "column",
        "categories": ["Q1"],
        "series": [{"name": "Revenue", "values": [10]}],
        "valueAxisTitle": "USD",
    }
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_chart", args)
    assert not preview.is_error, text_of(preview)
    assert bridge.calls[0]["args"]["valueAxisTitle"] == "USD"


async def test_an_applied_response_without_a_fingerprint_is_not_success():
    def handler(args):
        if args.get("apply"):
            return {"applied": True, "slide": SLIDE}
        return {"applied": False, "fingerprint": "fp-1", "fingerprintQuality": "full", "slide": SLIDE}

    bridge = FakeBridge({"add_text": handler})
    args = {"slide": 1, "x": 0, "y": 0, "w": 100, "h": 100, "paragraphs": [{"text": "hi"}]}
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_text", args)
        result = await t.call(
            "add_text",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert result.is_error
    assert "unknown" in text_of(result)


@pytest.mark.parametrize("quality", [None, "unavailable"])
async def test_a_preview_without_a_usable_fingerprint_quality_does_not_issue_a_token(quality):
    data = {"applied": False, "fingerprint": "fp-1", "slide": SLIDE}
    if quality is not None:
        data["fingerprintQuality"] = quality
    tokens = TokenTable()
    plan = WritePlan(tool="add_text", op="add_text", args={"slide": 1}, summary="summary")
    with pytest.raises(BridgeError) as raised:
        await _run_write(PptDeps(FakeBridge({"add_text": lambda args: data}), tokens), plan)
    assert raised.value.code == "fingerprint_unavailable"
    assert tokens._entries == {}


async def test_a_coarse_preview_is_still_usable():
    data = {
        "applied": False,
        "fingerprint": "fp-1",
        "fingerprintQuality": "coarse",
        "slide": SLIDE,
    }
    tokens = TokenTable()
    plan = WritePlan(tool="add_text", op="add_text", args={"slide": 1}, summary="summary")
    result = await _run_write(PptDeps(FakeBridge({"add_text": lambda args: data}), tokens), plan)
    assert result.structured["fingerprintQuality"] == "coarse"
    assert result.structured["applyToken"].startswith("dh1_")


def test_short_text_counts_utf16_units_and_keeps_a_whole_cluster():
    from deskhand.core.compat import utf16_len
    from deskhand.ppt.format import short_text

    assert utf16_len(short_text("😀" * 200)) <= 91
    assert not short_text("a" * 89 + "🇨🇳" + "tail").endswith("🇨")
    assert short_text("e\u0301" * 100).endswith("e\u0301…")


async def test_read_deck_defaults_to_a_one_line_per_slide_overview():
    slides = [{**SLIDE, "index": index, "title": f"Slide {index}"} for index in range(1, 26)]
    bridge = FakeBridge(
        {
            "read_deck": lambda args: {
                "presentation": PRESENTATION,
                "slides": slides,
                "fingerprint": "deck-fp",
                "fingerprintQuality": "full",
            }
        }
    )
    async with connect_ppt(bridge) as t:
        result = await t.call("read_deck", {})
    assert not result.is_error, text_of(result)
    lines = text_of(result).splitlines()
    assert len(lines) <= 25
    assert "more slide(s)" in text_of(result)
    assert "Pass slides=" in text_of(result)


async def test_render_caps_the_printed_file_list():
    files = [
        {"path": f"C:\\tmp\\out\\slide-{index}.png", "widthPx": 1600, "heightPx": 900}
        for index in range(1, 41)
    ]
    bridge = FakeBridge({"render": lambda args: {"files": files, "elapsedMs": 12, "outDir": "C:\\tmp\\out"}})
    async with connect_ppt(bridge) as t:
        result = await t.call("render", {"outDir": "/mnt/c/tmp/out"})
    assert not result.is_error, text_of(result)
    assert len(text_of(result).splitlines()) <= 8
    assert "and 35 more file(s)" in text_of(result)


async def test_a_write_preview_puts_the_token_and_fingerprint_first():
    bridge = FakeBridge({"add_text": write_handler()})
    async with connect_ppt(bridge) as t:
        preview = await t.call(
            "add_text", {"slide": 1, "x": 0, "y": 0, "w": 100, "h": 100, "paragraphs": [{"text": "hi"}]}
        )
    assert not preview.is_error, text_of(preview)
    lines = text_of(preview).splitlines()
    assert lines[0] == f"applyToken: {preview.structured_content['applyToken']}"
    assert lines[1] == f"fingerprint: {preview.structured_content['fingerprint']}"


@pytest.mark.parametrize(
    ("tool", "args"),
    [("update_shape", {"slide": 1, "shapeId": 10, "x": 10}), ("delete_shape", {"slide": 1, "shapeId": 10})],
)
async def test_a_shape_write_preview_always_shows_the_target_shape(tool, args):
    shapes = [{**SHAPE, "id": index, "text": f"shape {index}"} for index in range(1, 11)]
    slide = {**SLIDE, "shapes": shapes}
    bridge = FakeBridge({tool: lambda call: {**write_handler()(call), "slide": slide}})
    async with connect_ppt(bridge) as t:
        preview = await t.call(tool, args)
    assert not preview.is_error, text_of(preview)
    assert "shape 10 " in text_of(preview)
    assert "and 4 more shape(s)" in text_of(preview)


async def test_a_delete_slide_preview_shows_the_content_that_will_be_lost():
    doomed = {**SLIDE, "title": "Quarterly result", "shapes": [{**SHAPE, "text": "revenue grew"}]}
    bridge = FakeBridge(
        {
            "delete_slide": lambda args: {
                "applied": False,
                "fingerprint": "fp-1",
                "fingerprintQuality": "full",
                "presentation": PRESENTATION,
                "slide": doomed,
            }
        }
    )
    async with connect_ppt(bridge) as t:
        preview = await t.call("delete_slide", {"slide": 1})
    assert not preview.is_error, text_of(preview)
    assert 'title="Quarterly result"' in text_of(preview)
    assert "revenue grew" in text_of(preview)


async def test_create_deck_preview_then_apply():
    bridge = FakeBridge(
        {
            "create_deck": lambda args: {
                "applied": args.get("apply") is True,
                "fingerprint": "new-fp",
                "presentation": {**PRESENTATION, "slides": 0, "saved": False},
            }
        }
    )
    async with connect_ppt(bridge) as t:
        preview = await t.call("create_deck", {"width": 1280, "height": 720})
        assert not preview.is_error, text_of(preview)
        token = preview.structured_content["applyToken"]
        assert preview.structured_content["fingerprint"] == "new-fp"
        applied = await t.call(
            "create_deck",
            {
                "width": 1280,
                "height": 720,
                "dryRun": False,
                "applyToken": token,
                "expectFingerprint": "new-fp",
            },
        )
    assert not applied.is_error, text_of(applied)
    assert applied.structured_content["status"] == "applied"
    assert bridge.calls[0]["args"]["apply"] is False
    assert bridge.calls[1]["args"]["apply"] is True
    assert bridge.calls[1]["args"]["expectFingerprint"] == "new-fp"


async def test_apply_requires_the_fingerprint_from_the_preview():
    bridge = FakeBridge({"add_text": write_handler()})
    async with connect_ppt(bridge) as t:
        preview = await t.call(
            "add_text",
            {"slide": 1, "x": 96, "y": 64, "w": 1088, "h": 96, "paragraphs": [{"text": "hi"}]},
        )
        token = preview.structured_content["applyToken"]
        missing = await t.call(
            "add_text",
            {
                "slide": 1,
                "x": 96,
                "y": 64,
                "w": 1088,
                "h": 96,
                "paragraphs": [{"text": "hi"}],
                "dryRun": False,
                "applyToken": token,
            },
        )
        wrong = await t.call(
            "add_text",
            {
                "slide": 1,
                "x": 96,
                "y": 64,
                "w": 1088,
                "h": 96,
                "paragraphs": [{"text": "hi"}],
                "dryRun": False,
                "applyToken": token,
                "expectFingerprint": "other",
            },
        )
    assert missing.is_error and "expectFingerprint" in text_of(missing)
    assert wrong.is_error and "spent" in text_of(wrong)


async def test_add_text_preview_and_apply():
    bridge = FakeBridge({"add_text": write_handler()})
    args = {
        "slide": 1,
        "x": 96,
        "y": 64,
        "w": 1088,
        "h": 96,
        "paragraphs": [{"runs": [{"text": "Hello", "bold": True}, {"text": " world", "color": "#112233"}]}],
        "padding": 12,
        "anchor": "middle",
    }
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_text", args)
        assert not preview.is_error, text_of(preview)
        assert "Would add a text box to slide 1" in text_of(preview)
        applied = await t.call(
            "add_text",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert not applied.is_error, text_of(applied)
    assert applied.structured_content["status"] == "applied"
    assert applied.structured_content["slide"]["index"] == 1
    assert bridge.calls[1]["args"]["paragraphs"][0]["runs"][1]["text"] == " world"


async def test_add_shapes_preview_and_apply():
    bridge = FakeBridge({"add_shapes": write_handler(shapeIds=[11, 12], appliedCount=2)})
    args = {
        "slide": 1,
        "shapes": [
            {"kind": "text", "x": 10, "y": 20, "w": 100, "h": 40, "paragraphs": [{"text": "hi"}]},
            {"kind": "table", "x": 10, "y": 80, "w": 200, "h": 60, "rows": [["a"]]},
        ],
    }
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_shapes", args)
        assert not preview.is_error, text_of(preview)
        assert preview.structured_content["applied"] is False
        assert preview.structured_content["shapeIds"] == []
        assert preview.structured_content["appliedCount"] == 0
        applied = await t.call(
            "add_shapes",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert not applied.is_error, text_of(applied)
    assert applied.structured_content["shapeIds"] == [11, 12]
    assert applied.structured_content["appliedCount"] == 2
    assert bridge.calls[0]["args"]["apply"] is False
    assert bridge.calls[1]["args"]["apply"] is True
    assert bridge.calls[1]["args"]["expectFingerprint"] == "fp-1"


async def test_add_shapes_partial_details_are_machine_readable():
    def partial(args):
        raise BridgeError(
            "shape 2 failed",
            "applied_unverified",
            {
                "status": "partial",
                "tokenSpent": True,
                "appliedCount": 1,
                "failedIndex": 2,
                "total": 3,
                "shapeIds": [11],
                "failedShapeMayExist": True,
            },
        )

    bridge = FakeBridge({"add_shapes": write_handler()})
    args = {
        "slide": 1,
        "shapes": [
            {"kind": "text", "x": 10, "y": 20, "w": 100, "h": 40, "paragraphs": [{"text": "a"}]},
            {"kind": "text", "x": 10, "y": 80, "w": 100, "h": 40, "paragraphs": [{"text": "b"}]},
            {"kind": "text", "x": 10, "y": 140, "w": 100, "h": 40, "paragraphs": [{"text": "c"}]},
        ],
    }
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_shapes", args)
        bridge.handlers["add_shapes"] = partial
        applied = await t.call(
            "add_shapes",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert applied.is_error
    details = applied.structured_content["error"]["details"]
    assert details["status"] == "partial"
    assert details["appliedCount"] == 1
    assert details["failedIndex"] == 2
    assert details["shapeIds"] == [11]
    assert details["failedShapeMayExist"] is True
    assert "1 of 3" in text_of(applied)
    assert "Nothing was written" not in text_of(applied)


async def test_add_shapes_rejects_empty_and_too_many():
    async with connect_ppt(FakeBridge()) as t:
        empty = await t.call("add_shapes", {"slide": 1, "shapes": []})
        too_many = await t.call(
            "add_shapes",
            {
                "slide": 1,
                "shapes": [{"kind": "text", "x": 0, "y": 0, "w": 10, "h": 10, "paragraphs": [{"text": "x"}]}]
                * 21,
            },
        )
    assert empty.is_error and empty.structured_content["error"]["details"]["failedIndex"] is None
    assert too_many.is_error and too_many.structured_content["error"]["details"]["failedIndex"] is None


async def test_add_shapes_apply_requires_all_ids():
    bridge = FakeBridge({"add_shapes": write_handler(shapeIds=[11], appliedCount=1)})
    args = {
        "slide": 1,
        "shapes": [
            {"kind": "text", "x": 10, "y": 20, "w": 100, "h": 40, "paragraphs": [{"text": "a"}]},
            {"kind": "text", "x": 10, "y": 80, "w": 100, "h": 40, "paragraphs": [{"text": "b"}]},
        ],
    }
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_shapes", args)
        applied = await t.call(
            "add_shapes",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert applied.is_error
    assert applied.structured_content["error"]["code"] == "outcome_unknown"


async def test_a_token_is_bound_to_the_request():
    bridge = FakeBridge({"add_text": write_handler()})
    args = {"slide": 1, "x": 96, "y": 64, "w": 1088, "h": 96, "paragraphs": [{"text": "hi"}]}
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_text", args)
        changed = await t.call(
            "add_text",
            {
                **args,
                "x": 100,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert changed.is_error
    assert "differ" in text_of(changed)


async def test_a_token_is_single_use():
    bridge = FakeBridge({"add_text": write_handler()})
    args = {"slide": 1, "x": 96, "y": 64, "w": 1088, "h": 96, "paragraphs": [{"text": "hi"}]}
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_text", args)
        apply_args = {
            **args,
            "dryRun": False,
            "applyToken": preview.structured_content["applyToken"],
            "expectFingerprint": preview.structured_content["fingerprint"],
        }
        first = await t.call("add_text", apply_args)
        second = await t.call("add_text", apply_args)
    assert not first.is_error, text_of(first)
    assert second.is_error and "used" in text_of(second)


async def test_a_changed_deck_is_reported_as_nothing_written():
    def changed(args):
        if args.get("apply"):
            raise BridgeError("PowerPoint changed_since_preview: moved", "changed_since_preview")
        return write_handler()(args)

    bridge = FakeBridge({"add_text": changed})
    async with connect_ppt(bridge) as t:
        preview = await t.call(
            "add_text",
            {"slide": 1, "x": 96, "y": 64, "w": 1088, "h": 96, "paragraphs": [{"text": "hi"}]},
        )
        result = await t.call(
            "add_text",
            {
                "slide": 1,
                "x": 96,
                "y": 64,
                "w": 1088,
                "h": 96,
                "paragraphs": [{"text": "hi"}],
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert result.is_error
    assert "Nothing was written" in text_of(result)


async def test_an_uncertain_write_is_not_reported_as_safe_to_retry():
    def uncertain(args):
        if args.get("apply"):
            raise BridgeError("PowerPoint applied_unverified: half", "applied_unverified")
        return write_handler()(args)

    bridge = FakeBridge({"add_text": uncertain})
    async with connect_ppt(bridge) as t:
        preview = await t.call(
            "add_text",
            {"slide": 1, "x": 96, "y": 64, "w": 1088, "h": 96, "paragraphs": [{"text": "hi"}]},
        )
        result = await t.call(
            "add_text",
            {
                "slide": 1,
                "x": 96,
                "y": 64,
                "w": 1088,
                "h": 96,
                "paragraphs": [{"text": "hi"}],
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert result.is_error
    assert "unknown" in text_of(result)
    assert "Nothing was written" not in text_of(result)


async def test_a_write_result_that_fails_its_own_schema_is_unknown():
    bridge = FakeBridge({"x": lambda args: {"applied": True, "fingerprint": "fp"}})
    tokens = TokenTable()
    token = tokens.issue("x", {}, "fp")
    plan = WritePlan(
        tool="x",
        op="x",
        args={"dryRun": False, "applyToken": token, "expectFingerprint": "fp"},
        summary="x",
        output_schema={"type": "object", "required": ["must"]},
    )
    with pytest.raises(BridgeError) as raised:
        await _run_write(PptDeps(bridge, tokens), plan)
    assert raised.value.code == "outcome_unknown"
    assert raised.value.details is not None
    assert raised.value.details["status"] == "unknown"


@pytest.mark.parametrize(
    ("tool", "args"),
    [
        ("add_slide", {"layout": "title", "title": "Hi"}),
        (
            "add_table",
            {"slide": 1, "x": 96, "y": 96, "w": 800, "h": 300, "header": ["a", "b"], "rows": [["1", "2"]]},
        ),
        (
            "add_chart",
            {
                "slide": 1,
                "x": 96,
                "y": 96,
                "w": 800,
                "h": 400,
                "type": "column",
                "categories": ["Q1", "Q2"],
                "series": [{"name": "Revenue", "values": [1, 2]}],
            },
        ),
        (
            "add_picture",
            {"slide": 1, "x": 96, "y": 96, "w": 400, "h": 300, "path": "/mnt/c/tmp/a.png"},
        ),
        ("set_notes", {"slide": 1, "text": "note"}),
        ("update_shape", {"slide": 1, "shapeId": 4, "x": 10}),
        ("delete_shape", {"slide": 1, "shapeId": 4}),
        ("move_slide", {"from": 1, "to": 2}),
        ("delete_slide", {"slide": 1}),
        ("save_deck", {"path": "/mnt/c/tmp/out.pptx"}),
        ("close_deck", {}),
    ],
)
async def test_every_write_has_a_preview_and_an_apply(tool, args):
    bridge = FakeBridge({tool: write_handler()})
    async with connect_ppt(bridge) as t:
        preview = await t.call(tool, args)
        assert not preview.is_error, f"{tool}: {text_of(preview)}"
        applied = await t.call(
            tool,
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    assert not applied.is_error, f"{tool}: {text_of(applied)}"
    assert applied.structured_content["status"] == "applied"


async def test_add_picture_requires_exactly_one_source():
    async with connect_ppt(FakeBridge()) as t:
        both = await t.call(
            "add_picture",
            {"slide": 1, "x": 0, "y": 0, "w": 100, "h": 100, "path": "/mnt/c/a.png", "dataBase64": "aGk="},
        )
        neither = await t.call("add_picture", {"slide": 1, "x": 0, "y": 0, "w": 100, "h": 100})
    assert both.is_error and "exactly one" in text_of(both)
    assert neither.is_error and "exactly one" in text_of(neither)


async def test_save_deck_and_close_deck_pass_their_arguments():
    bridge = FakeBridge(
        {
            "save_deck": lambda args: {
                "applied": args.get("apply") is True,
                "fingerprint": "file-fp",
                "fingerprintQuality": "full",
                "presentation": PRESENTATION,
                "path": "C:\\tmp\\out.pptx",
            },
            "close_deck": lambda args: {
                "applied": args.get("apply") is True,
                "fingerprint": "fp-1",
                "fingerprintQuality": "full",
                "closed": "Deck",
                "path": "C:\\tmp\\deck.pptx",
            },
        }
    )
    async with connect_ppt(bridge) as t:
        save = await t.call("save_deck", {"path": "/mnt/c/tmp/out.pptx", "overwrite": True})
        close = await t.call("close_deck", {})
    assert not save.is_error, text_of(save)
    assert bridge.calls[0]["args"]["path"] == "C:\\tmp\\out.pptx"
    assert bridge.calls[0]["args"]["overwrite"] is True
    assert not close.is_error, text_of(close)
    assert bridge.calls[1]["args"]["discard"] is True
    assert "closed: Deck" in text_of(close)


async def test_a_write_does_not_leak_its_token_into_the_bridge_request():
    bridge = FakeBridge({"add_text": write_handler()})
    args = {"slide": 1, "x": 96, "y": 64, "w": 1088, "h": 96, "paragraphs": [{"text": "hi"}]}
    async with connect_ppt(bridge) as t:
        preview = await t.call("add_text", args)
        await t.call(
            "add_text",
            {
                **args,
                "dryRun": False,
                "applyToken": preview.structured_content["applyToken"],
                "expectFingerprint": preview.structured_content["fingerprint"],
            },
        )
    apply_call = bridge.calls[1]["args"]
    assert "applyToken" not in apply_call
    assert "dryRun" not in apply_call
    assert apply_call["expectFingerprint"] == "fp-1"


async def test_read_deck_keeps_the_bridge_result_shape():
    bridge = FakeBridge({"read_deck": deck_handler})
    async with connect_ppt(bridge) as t:
        result = await t.call("read_deck", {})
    assert not result.is_error, text_of(result)
    assert copy.deepcopy(result.structured_content)["slides"][0]["shapes"][0]["id"] == 4
