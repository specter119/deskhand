"""Live PowerPoint smoke test. Run by hand from WSL: `uv run python tests/live/ppt.py`.

It drives the real `ppt` server over MCP, creates a deck, adds content, renders it, saves it,
closes it, and reads it back. It never touches an existing deck: every file lives under the
Windows temp directory and is removed at the end.
"""

from __future__ import annotations

import asyncio
import shutil
import statistics
import subprocess
import time
import uuid
from pathlib import Path
from typing import Any

from mcp import Client

from deskhand.ppt.server import create_ppt_server


def run_powershell(command: str) -> str:
    return subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
        capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        timeout=60,
        check=True,
    ).stdout.strip()


def windows_temp() -> tuple[str, str]:
    win = run_powershell("[IO.Path]::GetTempPath()").replace("\r", "").replace("\n", "")
    wsl = subprocess.run(["wslpath", "-u", win], capture_output=True, text=True, check=True).stdout.strip()
    return win, wsl


def text_of(result: Any) -> str:
    return "\n".join(getattr(block, "text", "") for block in result.content)


def process_count() -> int:
    return int(run_powershell("@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count"))


def p50(values: list[float]) -> float:
    return statistics.median(values)


def wait_for_process_count(target: int, timeout: float = 20.0) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if process_count() == target:
            return True
        time.sleep(0.5)
    return process_count() == target


async def preview_apply(client: Client, tool: str, args: dict[str, Any]) -> Any:
    preview = await client.call_tool(tool, args)
    if preview.is_error:
        raise RuntimeError(f"{tool} preview: {text_of(preview)}")
    token = preview.structured_content["applyToken"]
    fingerprint = preview.structured_content["fingerprint"]
    applied = await client.call_tool(
        tool,
        {
            **args,
            "dryRun": False,
            "applyToken": token,
            "expectFingerprint": fingerprint,
        },
    )
    if applied.is_error:
        raise RuntimeError(f"{tool} apply: {text_of(applied)}")
    return applied


async def a1b_live_scenarios(record: Any) -> None:
    if process_count() != 0:
        print("SKIP A1b live scenarios: PowerPoint is already running")
        return

    server = create_ppt_server()
    async with Client(server, mode="legacy") as client:
        started = await client.call_tool("presentations", {})
        record(
            "a1b_start_no_deck",
            not started.is_error,
            text_of(started).splitlines()[0] if text_of(started) else "",
        )
    record("a1b_no_deck_quits", wait_for_process_count(0))
    if process_count() != 0:
        return

    server_a = create_ppt_server()
    async with Client(server_a, mode="legacy") as client_a:
        first = await client_a.call_tool("presentations", {})
        record("a1b_start", not first.is_error)
        if first.is_error:
            return
        created = await preview_apply(client_a, "create_deck", {"width": 1280, "height": 720})
        record("a1b_unsaved_deck", created.structured_content["status"] == "applied")

        server_b = create_ppt_server()
        async with Client(server_b, mode="legacy") as client_b:
            attached = await client_b.call_tool("presentations", {})
            decks = attached.structured_content.get("presentations", []) if not attached.is_error else []
            record("a1b_attach_existing", not attached.is_error and len(decks) >= 1, text_of(attached))
        time.sleep(0.5)
        record("a1b_attach_does_not_quit", process_count() > 0)

    time.sleep(0.5)
    record("a1b_unsaved_deck_survives_shutdown", process_count() > 0)

    server_c = create_ppt_server()
    async with Client(server_c, mode="legacy") as client_c:
        await client_c.call_tool("presentations", {})
        try:
            await preview_apply(client_c, "close_deck", {"discard": True})
        except Exception as error:
            record("a1b_cleanup_close", False, str(error))
    run_powershell(
        "try { [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application').Quit() } catch { }"
    )
    record("a1b_cleanup", wait_for_process_count(0))


async def main() -> int:
    if shutil.which("powershell.exe") is None:
        print("SKIP: powershell.exe is not available")
        return 0
    win_temp, wsl_temp = windows_temp()
    name = f"deskhand-ppt-live-{uuid.uuid4().hex[:8]}"
    win_dir = win_temp.rstrip("\\") + "\\" + name
    wsl_dir = Path(wsl_temp) / name
    pptx = win_dir + "\\live.pptx"
    results: list[tuple[str, bool, str]] = []

    def record(label: str, ok: bool, detail: str = "") -> None:
        results.append((label, ok, detail))
        print(f"{'PASS' if ok else 'FAIL'} {label}" + (f": {detail}" if detail else ""))

    await a1b_live_scenarios(record)
    server = create_ppt_server()
    baseline_ppt = process_count()
    try:
        async with Client(server, mode="legacy") as client:
            listed = await client.list_tools()
            record("tools", len(listed.tools) == 17, f"{len(listed.tools)} tools")

            started = time.monotonic()
            ping = await client.call_tool("presentations", {})
            cold = time.monotonic() - started
            record("powerpoint", not ping.is_error, text_of(ping).splitlines()[0] if text_of(ping) else "")
            if ping.is_error:
                return 1

            created = await preview_apply(client, "create_deck", {"width": 1280, "height": 720})
            record("create_deck", created.structured_content["status"] == "applied")

            slide = await preview_apply(
                client,
                "add_slide",
                {"layout": "title-content", "title": "Live probe", "subtitle": "deskhand"},
            )
            record("add_slide", slide.structured_content["slide"]["index"] == 1)

            hot_presentations: list[float] = []
            for _ in range(5):
                started = time.monotonic()
                await client.call_tool("presentations", {})
                hot_presentations.append(time.monotonic() - started)
            hot_slides: list[float] = []
            for _ in range(5):
                started = time.monotonic()
                await client.call_tool("read_slide", {"slide": 1})
                hot_slides.append(time.monotonic() - started)
            await client.call_tool("read_slide", {"slide": 999})
            started = time.monotonic()
            await client.call_tool("presentations", {})
            after_error = time.monotonic() - started
            record(
                "latency",
                cold <= 10.0
                and p50(hot_presentations) <= 0.3
                and p50(hot_slides) <= 0.5
                and after_error <= 0.5,
                f"cold={cold:.2f}s presentations_p50={p50(hot_presentations) * 1000:.0f}ms "
                f"read_slide_p50={p50(hot_slides) * 1000:.0f}ms after_error={after_error * 1000:.0f}ms",
            )

            text = await preview_apply(
                client,
                "add_text",
                {
                    "slide": 1,
                    "x": 96,
                    "y": 220,
                    "w": 1088,
                    "h": 120,
                    "paragraphs": [{"runs": [{"text": "Hello ", "bold": True}, {"text": "PowerPoint"}]}],
                },
            )
            record("add_text", len(text.structured_content["slide"]["shapes"]) >= 3)

            table = await preview_apply(
                client,
                "add_table",
                {
                    "slide": 1,
                    "x": 96,
                    "y": 360,
                    "w": 600,
                    "h": 160,
                    "header": ["A", "B"],
                    "rows": [["1", "2"]],
                },
            )
            record("add_table", len(table.structured_content["slide"]["shapes"]) >= 4)

            try:
                chart = await preview_apply(
                    client,
                    "add_chart",
                    {
                        "slide": 1,
                        "x": 720,
                        "y": 360,
                        "w": 460,
                        "h": 240,
                        "type": "column",
                        "categories": ["Q1", "Q2"],
                        "series": [
                            {"name": "Revenue", "values": [10, 20], "color": "#4472C4"},
                            {"name": "Cost", "values": [6, 12], "color": "#ED7D31"},
                            {"name": "Margin", "values": [4, 8], "color": "#A5A5A5"},
                            {"name": "Target", "values": [12, 18], "color": "#FFC000"},
                        ],
                        "dataLabels": True,
                        "numberFormat": "0.0",
                    },
                )
                record("add_chart", len(chart.structured_content["slide"]["shapes"]) >= 5)
            except Exception as error:
                record("add_chart", False, str(error))

            picture = await preview_apply(
                client,
                "add_picture",
                {
                    "slide": 1,
                    "x": 96,
                    "y": 540,
                    "w": 120,
                    "h": 80,
                    "dataBase64": (
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
                    ),
                    "fit": "contain",
                },
            )
            picture_shapes = picture.structured_content["slide"]["shapes"]
            picture_id = picture_shapes[-1]["id"]
            record("add_picture", picture_shapes[-1]["type"] == "picture")

            text_shapes = text.structured_content["slide"]["shapes"]
            text_id = next(shape["id"] for shape in text_shapes if shape["type"] == "textBox")
            updated = await preview_apply(
                client,
                "update_shape",
                {"slide": 1, "shapeId": text_id, "x": 120, "y": 240, "text": "Updated live text"},
            )
            updated_shape = next(
                shape for shape in updated.structured_content["slide"]["shapes"] if shape["id"] == text_id
            )
            record("update_shape", updated_shape["x"] == 120 and "Updated" in updated_shape["text"])

            deleted = await preview_apply(client, "delete_shape", {"slide": 1, "shapeId": picture_id})
            record(
                "delete_shape",
                all(shape["id"] != picture_id for shape in deleted.structured_content["slide"]["shapes"]),
            )

            second = await preview_apply(client, "add_slide", {"layout": "blank"})
            record("add_second_slide", second.structured_content["slide"]["index"] == 2)
            moved = await preview_apply(client, "move_slide", {"from": 2, "to": 1})
            record("move_slide", moved.structured_content["presentation"]["slides"] == 2)
            removed = await preview_apply(client, "delete_slide", {"slide": 1})
            record("delete_slide", removed.structured_content["presentation"]["slides"] == 1)

            notes = await preview_apply(client, "set_notes", {"slide": 1, "text": "live note"})
            record("set_notes", notes.structured_content["slide"].get("notes") == "live note")

            rendered = await client.call_tool("render", {"outDir": win_dir, "slides": "1", "overwrite": True})
            record("render", not rendered.is_error, text_of(rendered))
            rendered_pdf = await client.call_tool(
                "render", {"outDir": win_dir, "format": "pdf", "overwrite": True}
            )
            record("render_pdf", not rendered_pdf.is_error, text_of(rendered_pdf))

            saved = await preview_apply(client, "save_deck", {"path": pptx, "overwrite": True})
            record("save_deck", saved.structured_content["status"] == "applied")

            before_pdf = await client.call_tool("presentations", {})
            before_paths = [
                item.get("path", "") for item in before_pdf.structured_content.get("presentations", [])
            ]
            await client.call_tool("render", {"outDir": win_dir, "format": "pdf", "overwrite": True})
            after_pdf = await client.call_tool("presentations", {})
            after_paths = [
                item.get("path", "") for item in after_pdf.structured_content.get("presentations", [])
            ]
            record(
                "pdf_keeps_saved_deck_path",
                not after_pdf.is_error
                and before_paths == after_paths
                and all(path.lower().endswith(".pptx") for path in after_paths),
                f"before={before_paths} after={after_paths}",
            )

            protected = await client.call_tool("save_deck", {"path": pptx})
            record(
                "save_deck_overwrite_guard",
                protected.is_error and "file_exists" in text_of(protected),
                text_of(protected),
            )

            closed = await preview_apply(client, "close_deck", {"discard": True})
            record("close_deck", closed.structured_content["status"] == "applied")

            reopened = await client.call_tool("read_deck", {"presentation": pptx})
            record("read_saved_deck", not reopened.is_error, text_of(reopened).splitlines()[0])
            if not reopened.is_error:
                record(
                    "read_saved_deck_shape",
                    reopened.structured_content["presentation"]["slides"] == 1,
                )
                direct_save = await preview_apply(
                    client, "save_deck", {"presentation": pptx, "path": pptx, "overwrite": True}
                )
                record(
                    "save_read_only_deck_in_place",
                    direct_save.structured_content["status"] == "applied",
                )
                first_shape = reopened.structured_content["slides"][0]["shapes"][0]
                edited = await preview_apply(
                    client,
                    "update_shape",
                    {
                        "presentation": pptx,
                        "slide": 1,
                        "shapeId": first_shape["id"],
                        "x": first_shape["x"] + 10,
                    },
                )
                record("edit_saved_deck", edited.structured_content["status"] == "applied")
                saved_again = await preview_apply(
                    client, "save_deck", {"presentation": pptx, "path": pptx, "overwrite": True}
                )
                record("save_saved_deck", saved_again.structured_content["status"] == "applied")
                final = await client.call_tool("read_deck", {"presentation": pptx, "slides": "1"})
                final_shape = final.structured_content["slides"][0]["shapes"][0]
                record("read_edited_deck", final_shape["x"] == first_shape["x"] + 10)
                await preview_apply(client, "close_deck", {"discard": True})
    finally:
        if wsl_dir.exists():
            shutil.rmtree(wsl_dir, ignore_errors=True)

    deadline = time.monotonic() + 20
    while process_count() > baseline_ppt and time.monotonic() < deadline:
        time.sleep(0.5)
    record("powerpoint_cleanup", process_count() <= baseline_ppt, f"baseline={baseline_ppt}")

    failed = [item for item in results if not item[1]]
    print(f"\n{len(results) - len(failed)}/{len(results)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
