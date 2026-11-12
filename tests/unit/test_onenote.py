import re

import pytest
from helpers import PAGE, FakeBridge, connect, preview_handler, text_of

from deskhand.core.powershell import BridgeError
from deskhand.onenote.server import INSTRUCTIONS

pytestmark = pytest.mark.anyio

TOOLS = [
    "notebooks",
    "pages",
    "search",
    "read_page",
    "create_page",
    "append_lines",
    "insert_table",
    "insert_image",
    "edit",
    "delete_page",
]
WRITE_TOOLS = ["create_page", "append_lines", "insert_table", "insert_image", "edit", "delete_page"]


async def test_list_tools_exposes_the_ten_tools_with_titles_annotations_and_schemas():
    async with connect() as t:
        tools = (await t.client.list_tools()).tools
        assert sorted(tool.name for tool in tools) == sorted(TOOLS)
        for tool in tools:
            assert tool.title, f"{tool.name} has a title"
            assert tool.input_schema["type"] == "object"
            assert tool.output_schema and tool.output_schema["type"] == "object"
            assert tool.annotations, f"{tool.name} has annotations"
            assert (tool.annotations.read_only_hint is True) == (tool.name not in WRITE_TOOLS), tool.name
        by_name = {tool.name: tool for tool in tools}
        for name in WRITE_TOOLS:
            properties = by_name[name].input_schema["properties"]
            assert "dryRun" in properties and "applyToken" in properties, name
            assert "applyToken" in (by_name[name].description or "")
        for name in ("edit", "delete_page"):
            annotations = by_name[name].annotations
            assert annotations and annotations.destructive_hint is True, name
        assert t.client.instructions == INSTRUCTIONS


async def test_read_tools_validate_against_their_output_schemas_including_powershell_quirks():
    bridge = FakeBridge(
        {
            "notebooks": lambda _: {
                "notebooks": [
                    {
                        "id": "n1",
                        "name": "NB",
                        "path": "p",
                        "sections": {"id": "s1", "name": "Sec", "path": "p2"},
                    }
                ]
            },
            "pages": lambda _: {
                "sectionId": "s1",
                "section": "Sec",
                "pages": {"id": "p1", "name": "Probe", "level": "0", "created": "t"},
            },
            "search": lambda _: {
                "total": 1,
                "returned": 1,
                "pages": {"pageId": "p1", "page": "Probe", "section": "Sec", "notebook": "NB"},
            },
            "read_page": lambda _: {
                "page": {**PAGE, "lines": PAGE["lines"][0], "outlines": PAGE["outlines"][0]}
            },
        }
    )
    async with connect(bridge) as t:
        for name, args in [
            ("notebooks", {}),
            ("pages", {"sectionId": "s1"}),
            ("search", {"query": "pro"}),
            ("read_page", {"pageId": "p1"}),
        ]:
            result = await t.call(name, args)
            assert not result.is_error, f"{name}: {text_of(result)}"
            assert result.structured_content, name
        page = (await t.call("read_page", {"pageId": "p1"})).structured_content["page"]
        assert isinstance(page["lines"], list) and isinstance(page["outlines"], list)
        assert not any("apply" in call["args"] for call in bridge.calls)


async def test_a_dry_run_puts_the_token_on_the_first_line_and_never_applies():
    bridge = FakeBridge({"append": preview_handler()})
    async with connect(bridge) as t:
        result = await t.call("append_lines", {"pageId": "p1", "lines": ["a", "b"]})
        assert not result.is_error, text_of(result)
        first = text_of(result).split("\n")[0]
        assert re.fullmatch(r"applyToken: dh1_[0-9a-f]{32}", first)
        assert result.structured_content["status"] == "dry-run"
        assert result.structured_content["applyToken"] == first.removeprefix("applyToken: ")
        assert bridge.apply_calls() == []
        assert [call["args"]["apply"] for call in bridge.calls] == [False]


async def test_an_apply_without_a_valid_token_is_refused_before_the_bridge_is_reached():
    bridge = FakeBridge({"append": preview_handler()})
    async with connect(bridge) as t:
        args = {"pageId": "p1", "lines": ["a"]}
        token = (await t.call("append_lines", args)).structured_content["applyToken"]
        attempts = [
            {**args, "dryRun": False},
            {**args, "dryRun": False, "applyToken": "dh1_forged"},
            {**args, "lines": ["changed"], "dryRun": False, "applyToken": token},
            {"pageId": "other", "lines": ["a"], "dryRun": False, "applyToken": token},
        ]
        for attempt in attempts:
            assert (await t.call("append_lines", attempt)).is_error is True, attempt
        assert bridge.apply_calls() == [], "the bridge never saw an apply"


async def test_a_valid_token_applies_once_with_the_previewed_fingerprint_and_cannot_be_replayed():
    bridge = FakeBridge({"append": preview_handler("fp-A")})
    async with connect(bridge) as t:
        args = {"pageId": "p1", "lines": ["a"]}
        token = (await t.call("append_lines", args)).structured_content["applyToken"]

        applied = await t.call("append_lines", {**args, "dryRun": False, "applyToken": token})
        assert not applied.is_error, text_of(applied)
        assert applied.structured_content["status"] == "applied"
        sent = bridge.apply_calls()
        assert len(sent) == 1
        assert sent[0]["args"]["expectFingerprint"] == "fp-A"
        assert sorted(sent[0]["args"]) == ["apply", "expectFingerprint", "lines", "pageId"]

        replay = await t.call("append_lines", {**args, "dryRun": False, "applyToken": token})
        assert replay.is_error is True
        assert len(bridge.apply_calls()) == 1, "the replay never reached the bridge"


async def test_a_token_issued_for_one_tool_does_not_apply_another():
    bridge = FakeBridge({"append": preview_handler(), "delete_page": preview_handler()})
    async with connect(bridge) as t:
        token = (await t.call("append_lines", {"pageId": "p1", "lines": ["a"]})).structured_content[
            "applyToken"
        ]
        result = await t.call(
            "delete_page", {"pageId": "p1", "expectName": "Probe", "dryRun": False, "applyToken": token}
        )
        assert result.is_error is True
        assert bridge.apply_calls() == []


@pytest.mark.parametrize(
    ("tool", "args", "op"),
    [
        ("create_page", {"sectionId": "s1", "title": "T", "lines": ["x"]}, "create_page"),
        ("append_lines", {"pageId": "p1", "lines": ["x"]}, "append"),
        ("insert_table", {"pageId": "p1", "rows": [["a", "b"]]}, "insert_table"),
        ("edit", {"pageId": "p1", "set": [{"objectID": "oe-1", "text": "new"}]}, "edit"),
        ("delete_page", {"pageId": "p1", "expectName": "Probe"}, "delete_page"),
    ],
)
async def test_every_write_tool_follows_the_same_dry_run_then_apply_flow(tool, args, op):
    bridge = FakeBridge({op: preview_handler(f"fp-{tool}")})
    async with connect(bridge) as t:
        preview = await t.call(tool, args)
        assert not preview.is_error, f"{tool}: {text_of(preview)}"
        applied = await t.call(
            tool, {**args, "dryRun": False, "applyToken": preview.structured_content["applyToken"]}
        )
        assert not applied.is_error, f"{tool}: {text_of(applied)}"
        assert bridge.apply_calls()[0]["args"]["expectFingerprint"] == f"fp-{tool}"


@pytest.mark.parametrize(
    "args",
    [
        {"pageId": "p1", "append": [{"objectID": "ol-1", "lines": ["x"], "xml": ["<one:OE/>"]}]},
        {"pageId": "p1", "set": [{"objectID": "oe-1", "text": "x", "xml": "<one:OE/>"}]},
    ],
)
async def test_edit_accepts_only_the_whitelisted_fields_so_raw_xml_cannot_be_smuggled_in(args):
    bridge = FakeBridge({"edit": preview_handler()})
    async with connect(bridge) as t:
        assert (await t.call("edit", args)).is_error is True
        assert bridge.calls == [], "rejected by schema validation, before any bridge call"


async def test_bridge_failures_surface_as_tool_errors_with_their_code():
    def append(args):
        if args.get("apply"):
            raise BridgeError("OneNote changed_since_preview: the target changed", "changed_since_preview")
        return preview_handler()(args)

    async with connect(FakeBridge({"append": append})) as t:
        args = {"pageId": "p1", "lines": ["a"]}
        token = (await t.call("append_lines", args)).structured_content["applyToken"]
        result = await t.call("append_lines", {**args, "dryRun": False, "applyToken": token})
        assert result.is_error is True
        assert "changed_since_preview" in text_of(result)


async def test_insert_image_binds_the_token_to_the_file_contents(tmp_path):
    path = tmp_path / "a.png"
    path.write_bytes(b"first")
    bridge = FakeBridge({"insert_image": preview_handler()})
    async with connect(bridge) as t:
        args = {"pageId": "p1", "path": str(path)}
        token = (await t.call("insert_image", args)).structured_content["applyToken"]
        path.write_bytes(b"second")
        result = await t.call("insert_image", {**args, "dryRun": False, "applyToken": token})
        assert result.is_error is True, "the file changed after the preview"
        assert bridge.apply_calls() == []


async def test_an_unknown_tool_is_a_protocol_error_and_a_missing_argument_object_is_validated():
    async with connect() as t:
        with pytest.raises(Exception, match="not found"):
            await t.call("nope")
        result = await t.client.call_tool("append_lines", None)
        assert result.is_error is True
        assert text_of(result).startswith("Input validation error")


async def test_a_paragraph_object_carries_depth_list_and_style_to_the_bridge():
    bridge = FakeBridge({"append": preview_handler()})
    async with connect(bridge) as t:
        args = {
            "pageId": "p1",
            "lines": [
                {"text": "heading", "quickStyle": 2, "style": {"fontFamily": "Calibri", "fontSize": 19}},
                {"text": "child", "depth": 1, "list": "bullet", "style": {"bold": True}},
                {"text": "grandchild", "depth": 2, "list": "number"},
            ],
        }
        preview = await t.call("append_lines", args)
        assert not preview.is_error, text_of(preview)
        sent = bridge.calls[0]["args"]["lines"]
        # depth is always explicit on the wire; everything else is passed through as given.
        assert sent == [{**args["lines"][0], "depth": 0}, *args["lines"][1:]], sent
        assert "1. heading   [style=2, 19pt Calibri]" in text_of(preview), text_of(preview)
        assert "2.   child   [depth=1, bullet, bold]" in text_of(preview), text_of(preview)


async def test_a_plain_string_stays_a_plain_string_so_an_unformatted_call_is_unchanged():
    bridge = FakeBridge({"append": preview_handler()})
    async with connect(bridge) as t:
        preview = await t.call("append_lines", {"pageId": "p1", "lines": ["a", "b"]})
        assert bridge.calls[0]["args"]["lines"] == ["a", "b"]
        assert text_of(preview).count("\n  1. a") == 1, text_of(preview)


@pytest.mark.parametrize(
    ("lines", "message"),
    [
        ([{"text": "a", "depth": 1}], "the first paragraph must be at depth 0"),
        ([{"text": "a"}, {"text": "b", "depth": 2}], "depth 2 skips a level"),
        ([{"text": "a", "depth": -1}], "depth must be a whole number"),
        ([{"text": "a", "depth": 1.5}], "depth must be a whole number"),
        ([{"text": "a", "style": {"fontSize": 0}}], "style.fontSize must be a number above 0"),
        ([{"text": "a", "style": {"color": "blue"}}], "style.color must be #RRGGBB"),
        ([{"text": "a", "spaceBefore": -1}], "spaceBefore must be a number"),
        ([{"text": "a", "quickStyle": True}], "is not valid under any of the given schemas"),
    ],
)
async def test_paragraph_objects_are_checked_before_anything_reaches_windows(lines, message):
    bridge = FakeBridge({"append": preview_handler()})
    async with connect(bridge) as t:
        result = await t.call("append_lines", {"pageId": "p1", "lines": lines})
        assert result.is_error is True
        assert message in text_of(result), text_of(result)
        assert bridge.calls == []


async def test_after_anchors_an_image_table_or_appended_paragraph_to_a_paragraph(tmp_path):
    path = tmp_path / "a.png"
    path.write_bytes(b"png bytes")
    bridge = FakeBridge(
        {"insert_image": preview_handler(), "insert_table": preview_handler(), "edit": preview_handler()}
    )
    async with connect(bridge) as t:
        image = await t.call("insert_image", {"pageId": "p1", "path": str(path), "after": "oe-3"})
        assert not image.is_error, text_of(image)
        assert bridge.calls[0]["args"]["after"] == "oe-3"
        assert "after paragraph oe-3" in text_of(image), text_of(image)

        table = await t.call("insert_table", {"pageId": "p1", "rows": [["a"]], "after": "oe-3"})
        assert not table.is_error, text_of(table)
        assert bridge.calls[1]["args"]["after"] == "oe-3"

        edit = await t.call(
            "edit", {"pageId": "p1", "append": [{"objectID": "ol-1", "lines": ["x"], "after": "oe-3"}]}
        )
        assert not edit.is_error, text_of(edit)
        assert bridge.calls[2]["args"]["append"][0]["after"] == "oe-3"
        assert "add 1 paragraph(s) to ol-1 after oe-3" in text_of(edit), text_of(edit)
