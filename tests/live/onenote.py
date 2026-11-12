"""Live checks against a real OneNote desktop client. Needs WSL with Windows interop.

    DESKHAND_LIVE_NOTEBOOK="<existing notebook name>" uv run python tests/live/onenote.py

Every page it creates is titled "deskhand-live-<pid>-..." in the section DESKHAND_LIVE_SECTION
(default "deskhand-probe", created if missing) and is deleted again, also when a check fails.
"""

from __future__ import annotations

import asyncio
import base64
import json
import os
import re
import shutil
import signal
import statistics
import struct
import sys
import tempfile
import time
import traceback
import zlib
from collections.abc import Awaitable, Callable
from contextlib import AsyncExitStack
from pathlib import Path
from typing import Any

from mcp import Client, StdioServerParameters

from deskhand.core.powershell import BridgeError, PowerShellBridge, work_dir_from_temp
from deskhand.onenote.server import BRIDGE_SCRIPT

NOTEBOOK = os.environ.get("DESKHAND_LIVE_NOTEBOOK")
SECTION = os.environ.get("DESKHAND_LIVE_SECTION", "deskhand-probe")
if not NOTEBOOK:
    sys.exit("Set DESKHAND_LIVE_NOTEBOOK to the name of an existing notebook.")

ENSURE_SECTION = (Path(__file__).parent / "ensure-section.ps1").read_text("utf-8")
TAG = f"deskhand-live-{os.getpid()}-{int(time.time()):x}"
WHERE = {"notebook": NOTEBOOK, "section": SECTION}


def quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


async def run_powershell(script: str) -> str:
    proc = await asyncio.create_subprocess_exec(
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        base64.b64encode(script.encode("utf-16-le")).decode(),
        stdin=asyncio.subprocess.DEVNULL,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    stdout, stderr = await proc.communicate()
    if proc.returncode:
        raise RuntimeError(stderr.decode("utf-8", "replace"))
    return stdout.decode("utf-8", "replace").strip()


async def ensure_section() -> str:
    return await run_powershell(
        f"$Notebook = {quote(NOTEBOOK or '')}; $Section = {quote(SECTION)}\n{ENSURE_SECTION}"
    )


def server_children() -> set[int]:
    """Pids of the deskhand servers this process started."""
    found: set[int] = set()
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        try:
            ppid = int((entry / "stat").read_text().rsplit(")", 1)[1].split()[1])
            if ppid == os.getpid() and b"deskhand" in (entry / "cmdline").read_bytes():
                found.add(int(entry.name))
        except (OSError, IndexError, ValueError):
            continue
    return found


async def windows_workers() -> dict[int, list[int]]:
    """Maps each WSL server pid to the Windows pids of the powershell.exe workers it started."""
    listing = await run_powershell(
        "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" | "
        "ForEach-Object { '{0} {1}' -f $_.ProcessId, $_.CommandLine }"
    )
    workers: dict[int, list[int]] = {}
    for line in listing.splitlines():
        match = re.match(r"(\d+) .*-EncodedCommand (\S+)", line)
        if not match:
            continue
        bootstrap = base64.b64decode(match[2]).decode("utf-16-le")
        owner = re.search(r"worker-(\d+)-[0-9a-f-]+\.ps1", bootstrap)
        if owner:
            workers.setdefault(int(owner[1]), []).append(int(match[1]))
    return workers


async def onenote_running() -> bool:
    return await run_powershell("[bool](Get-Process ONENOTE -ErrorAction SilentlyContinue)") == "True"


async def windows_process_exists(pid: int) -> bool:
    return await run_powershell(f"[bool](Get-Process -Id {pid} -ErrorAction SilentlyContinue)") == "True"


async def lone_background_onenote() -> int | None:
    """The pid of the only ONENOTE.EXE when it is safe to kill, else None.

    Safe means: one process, started by COM (-Embedding) with no window, and no deskhand worker that this run
    did not start, so a session the user has open somewhere is never interrupted.
    """
    listing = await run_powershell(
        "Get-CimInstance Win32_Process -Filter \"Name='ONENOTE.EXE'\" | ForEach-Object { "
        "'{0}|{1}|{2}' -f $_.ProcessId, "
        "[bool](Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }), "
        "$_.CommandLine }"
    )
    rows = [line.split("|", 2) for line in listing.splitlines() if line.strip()]
    if len(rows) != 1 or rows[0][1] != "False" or "-Embedding" not in rows[0][2]:
        return None
    ours = {os.getpid(), *server_children()}
    if set(await windows_workers()) - ours:
        return None
    return int(rows[0][0])


async def update_page_directly(page_id: str, xml_body: str) -> str:
    """Sends UpdatePageContent through COM with the bridge's own classification and returns '<hex> <code>'."""
    temp = await run_powershell("[IO.Path]::GetTempPath()")
    work = work_dir_from_temp(temp, f"live-{os.getpid()}")
    work.wsl.mkdir(parents=True, exist_ok=True)
    try:
        head = BRIDGE_SCRIPT.read_text("utf-8").split("# --- worker loop", 1)[0]
        (work.wsl / "head.ps1").write_text(head, encoding="utf-8")
        ns = "http://schemas.microsoft.com/office/onenote/2013/onenote"
        return await run_powershell(
            f"Invoke-Expression ([IO.File]::ReadAllText({quote(work.win + chr(92) + 'head.ps1')}, [Text.Encoding]::UTF8))\n"
            "$app = New-Object -ComObject OneNote.Application\n"
            f"$xml = '<one:Page xmlns:one=\"{ns}\" ID=\"' + {quote(page_id)} + '\">' + {quote(xml_body)} + '</one:Page>'\n"
            "try { $app.UpdatePageContent($xml, [DateTime]::MinValue, 2, $false); 'no-error' }\n"
            "catch { $h = Get-HResultHex $_; $h + ' ' + (Get-WriteFailureCode $h) }"
        )
    finally:
        shutil.rmtree(work.wsl, ignore_errors=True)


class Result:
    def __init__(self, raw: Any) -> None:
        self.is_error: bool = bool(raw.is_error)
        self.text = "\n".join(getattr(block, "text", "") for block in raw.content)
        self.data: dict[str, Any] = raw.structured_content or {}


class Server:
    def __init__(self, client: Client) -> None:
        self.client = client

    async def call(self, name: str, args: dict[str, Any] | None = None) -> Result:
        return Result(await self.client.call_tool(name, args or {}))

    async def write(self, name: str, args: dict[str, Any]) -> tuple[Result, Result, str]:
        """Dry run, then apply with the token it issued. Returns both results and the token."""
        dry = await self.call(name, args)
        assert not dry.is_error, f"{name} dry run failed: {dry.text}"
        assert dry.data["status"] == "dry-run"
        token = dry.data["applyToken"]
        assert dry.text.split("\n")[0] == f"applyToken: {token}", "token must be the first line"
        applied = await self.call(name, {**args, "dryRun": False, "applyToken": token})
        return dry, applied, token


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    size = 16
    header = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    rows = b"".join(b"\x00" + b"\x80" * (size * 3) for _ in range(size))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


results: list[tuple[str, bool]] = []


async def check(name: str, fn: Callable[[], Awaitable[None]]) -> None:
    try:
        await fn()
        results.append((name, True))
        print(f"ok   {name}")
    except Exception:
        results.append((name, False))
        trace = "\n     ".join(traceback.format_exc().splitlines()[-6:])
        print(f"FAIL {name}\n     {trace}")


async def main() -> int:
    created: list[str] = []
    # When the user has not opened OneNote, COM starts it for the first call, which adds about a second.
    was_running = await onenote_running()
    print(f"OneNote was {'already' if was_running else 'not'} running")
    print(f"section {SECTION}: {(await ensure_section()).split(' ')[0]}")

    async with AsyncExitStack() as stack:
        params = StdioServerParameters(command=sys.executable, args=["-m", "deskhand", "onenote"])
        a = Server(await stack.enter_async_context(Client(params)))
        (pid_a,) = server_children()
        b = Server(await stack.enter_async_context(Client(params)))
        (pid_b,) = server_children() - {pid_a}
        first_worker: dict[str, int] = {}

        async def page_id_of(title: str) -> str | None:
            listing = await a.call("pages", WHERE)
            return next((p["id"] for p in listing.data["pages"] if p["name"] == title), None)

        async def read_page(server: Server, page_id: str) -> dict[str, Any]:
            return (await server.call("read_page", {"pageId": page_id})).data["page"]

        title = f"{TAG}-main"
        state: dict[str, Any] = {}

        async def line_by_text(text: str) -> dict[str, Any] | None:
            page = await read_page(a, state["page_id"])
            return next((line for line in page["lines"] if line["text"] == text), None)

        try:

            async def lists_tools() -> None:
                tools = (await a.client.list_tools()).tools
                assert len(tools) == 10
                for tool in tools:
                    assert tool.output_schema, f"{tool.name} has no output_schema"

            await check("lists ten tools with output schemas", lists_tools)

            async def notebooks() -> None:
                data = (await a.call("notebooks")).data
                notebook = next((n for n in data["notebooks"] if n["name"] == NOTEBOOK), None)
                assert notebook, "notebook missing"
                assert any(s["name"] == SECTION for s in notebook["sections"]), "section missing"

            await check("notebooks shows the live notebook and the probe section", notebooks)

            async def one_worker_started() -> None:
                workers = await windows_workers()
                assert len(workers.get(pid_a, [])) == 1, workers
                first_worker["a"] = workers[pid_a][0]

            await check("the first call starts exactly one resident powershell.exe", one_worker_started)

            async def create() -> None:
                args = {**WHERE, "title": title, "lines": ["first line", "keep line"]}
                dry = await a.call("create_page", args)
                assert not dry.is_error, dry.text
                assert await page_id_of(title) is None, "dry run created the page"

                refused = await a.call("create_page", {**args, "dryRun": False})
                assert refused.is_error, "apply without a token must fail"
                assert await page_id_of(title) is None, "refused apply created the page"

                token = dry.data["applyToken"]
                applied = await a.call("create_page", {**args, "dryRun": False, "applyToken": token})
                assert not applied.is_error, applied.text
                page_id = await page_id_of(title)
                assert page_id, "page not found after apply"
                state["page_id"] = page_id
                created.append(page_id)

                reused = await a.call("create_page", {**args, "dryRun": False, "applyToken": token})
                assert reused.is_error, "a token must be single use"
                pages = (await a.call("pages", WHERE)).data["pages"]
                assert sum(p["name"] == title for p in pages) == 1, "reused token created a second page"

            await check(
                "create_page: dry run writes nothing, apply without token is refused, apply creates", create
            )

            async def reads() -> None:
                page = await read_page(a, state["page_id"])
                assert page["name"] == title
                assert any(line["text"] == "first line" for line in page["lines"])

            await check("read_page returns paragraphs that validate against the output schema", reads)

            async def append() -> None:
                _, applied, _ = await a.write(
                    "append_lines", {"pageId": state["page_id"], "lines": ["second line"]}
                )
                assert not applied.is_error, applied.text
                texts = [line["text"] for line in (await read_page(a, state["page_id"]))["lines"]]
                assert "first line" in texts and "second line" in texts, "|".join(texts)

            await check("append_lines adds a paragraph and leaves the old ones", append)

            async def formatting() -> None:
                _, applied, _ = await a.write(
                    "append_lines",
                    {
                        "pageId": state["page_id"],
                        "lines": [
                            {
                                "text": "formatted heading",
                                "list": "bullet",
                                "quickStyle": "p",
                                "spaceBefore": 5,
                                "style": {
                                    "fontFamily": "Calibri",
                                    "fontSize": 19,
                                    "color": "#2E75B5",
                                    "bold": True,
                                },
                            },
                            {"text": "nested child", "depth": 1, "list": "bullet"},
                            {"text": "nested grandchild", "depth": 2, "list": "number"},
                        ],
                    },
                )
                assert not applied.is_error, applied.text
                page = await read_page(a, state["page_id"])
                by_text = {line["text"]: line for line in page["lines"]}
                heading = by_text["formatted heading"]
                child = by_text["nested child"]
                grandchild = by_text["nested grandchild"]
                assert heading["list"] == "bullet" and heading["depth"] == 0, heading
                assert child["depth"] == 1 and child["list"] == "bullet", child
                assert grandchild["depth"] == 2 and grandchild["list"] == "number", grandchild
                # A name resolves through the page's own QuickStyleDef list. OneNote then maps the index onto
                # the style that best matches the paragraph, so only "a defined index" is stable to assert.
                assert heading["quickStyle"] in {"0", "1"}, heading
                assert "19" in heading.get("style", ""), heading

            await check("append_lines writes depth, list, quick style and font size", formatting)

            async def insert_after() -> None:
                anchor = await line_by_text("second line")
                assert anchor, "anchor paragraph missing"
                with tempfile.TemporaryDirectory(prefix="deskhand-live-") as directory:
                    file = Path(directory) / "probe.png"
                    file.write_bytes(png())
                    _, applied, _ = await a.write(
                        "insert_image",
                        {"pageId": state["page_id"], "path": str(file), "after": anchor["objectID"]},
                    )
                assert not applied.is_error, applied.text
                page = await read_page(a, state["page_id"])
                index = next(i for i, line in enumerate(page["lines"]) if line["text"] == "second line")
                assert page["lines"][index + 1]["hasImage"], [line["text"] for line in page["lines"]]

            await check("insert_image with after lands right after that paragraph", insert_after)

            async def second_dry_run() -> None:
                target = await line_by_text("keep line")
                assert target
                args = {"pageId": state["page_id"], "set": [{"objectID": target["objectID"], "text": "kept"}]}
                one = await a.call("edit", args)
                two = await a.call("edit", args)
                assert not one.is_error and not two.is_error, one.text + two.text
                assert one.data["applyToken"] != two.data["applyToken"]
                applied = await a.call(
                    "edit", {**args, "dryRun": False, "applyToken": one.data["applyToken"]}
                )
                assert not applied.is_error, applied.text
                assert await line_by_text("kept")

            await check(
                "a second dry run of an unchanged page does not invalidate the first token", second_dry_run
            )

            async def stale_token() -> None:
                # The fingerprint covers the outline text that would be written, so the concurrent change
                # must touch another paragraph of the same outline.
                target = await line_by_text("first line")
                neighbour = await line_by_text("kept")
                assert target and neighbour
                page_id = state["page_id"]
                stale_args = {"pageId": page_id, "set": [{"objectID": target["objectID"], "text": "STALE"}]}
                stale = await a.call("edit", stale_args)
                _, other, _ = await a.write(
                    "edit",
                    {"pageId": page_id, "set": [{"objectID": neighbour["objectID"], "text": "neighbour"}]},
                )
                assert not other.is_error, other.text
                refused = await a.call(
                    "edit", {**stale_args, "dryRun": False, "applyToken": stale.data["applyToken"]}
                )
                assert refused.is_error
                assert re.search(r"changed_since_preview", refused.text), refused.text
                assert not await line_by_text("STALE")
                assert await line_by_text("neighbour")

            await check("a token is refused when the outline changed after the dry run", stale_token)

            async def other_arguments() -> None:
                page_id = state["page_id"]
                dry = await a.call("append_lines", {"pageId": page_id, "lines": ["bound"]})
                refused = await a.call(
                    "append_lines",
                    {
                        "pageId": page_id,
                        "lines": ["something else"],
                        "dryRun": False,
                        "applyToken": dry.data["applyToken"],
                    },
                )
                assert refused.is_error
                page = await read_page(a, page_id)
                assert not any(line["text"] == "something else" for line in page["lines"])

            await check("a token for different arguments is refused", other_arguments)

            async def edit_remove() -> None:
                page_id = state["page_id"]
                first = await line_by_text("first line")
                kept = await line_by_text("neighbour")
                assert first and kept
                assert first["outlineID"] == kept["outlineID"], "both lines should share the first outline"
                _, applied, _ = await a.write("edit", {"pageId": page_id, "remove": [kept["objectID"]]})
                assert not applied.is_error, applied.text
                assert not await line_by_text("neighbour")
                assert await line_by_text("first line")
                rejected = await a.call(
                    "edit",
                    {
                        "pageId": page_id,
                        "set": [{"objectID": first["objectID"], "text": "x", "xml": "<one:OE/>"}],
                    },
                )
                assert rejected.is_error, "raw xml must be rejected"

            await check(
                "edit: set replaces text, remove deletes a paragraph, raw xml is rejected", edit_remove
            )

            async def table() -> None:
                _, applied, _ = await a.write(
                    "insert_table",
                    {
                        "pageId": state["page_id"],
                        "headers": ["a", "b"],
                        "rows": [["1", "2"], ["3", "4"]],
                        "hasHeaderRow": True,
                        "bordersVisible": True,
                    },
                )
                assert not applied.is_error, applied.text
                page = await read_page(a, state["page_id"])
                assert len(page["tables"]) == 1
                assert page["tables"][0]["rows"] == 3
                assert page["tables"][0]["columns"] == 2

            await check("insert_table adds a table", table)

            async def image() -> None:
                with tempfile.TemporaryDirectory(prefix="deskhand-live-") as directory:
                    file = Path(directory) / "probe.png"
                    file.write_bytes(png())
                    _, applied, _ = await a.write(
                        "insert_image", {"pageId": state["page_id"], "path": str(file), "alt": "probe"}
                    )
                assert not applied.is_error, applied.text
                # The after-anchor check above already put one image on this page.
                assert len((await read_page(a, state["page_id"]))["images"]) >= 1

            await check("insert_image adds an image bound to the file contents", image)

            async def refused_by_onenote() -> None:
                # Well-formed XML that is not in the page schema: OneNote refuses it, and force=false is accepted.
                verdict = await update_page_directly(state["page_id"], "<one:Bogus/>")
                assert verdict == "0x80042001 rejected", verdict

            await check(
                "a schema violation is refused with 0x80042001 and classified rejected", refused_by_onenote
            )

            async def two_processes() -> None:
                titles = [f"{TAG}-a", f"{TAG}-b"]
                (_, ra, _), (_, rb, _) = await asyncio.gather(
                    a.write("create_page", {**WHERE, "title": titles[0], "lines": ["from a"]}),
                    b.write("create_page", {**WHERE, "title": titles[1], "lines": ["from b"]}),
                )
                assert not ra.is_error, ra.text
                assert not rb.is_error, rb.text
                ids = await asyncio.gather(*(page_id_of(t) for t in titles))
                assert all(ids), "a page is missing"
                created.extend(i for i in ids if i)
                id_a, id_b = ids
                reads = await asyncio.gather(
                    *(a.call("read_page", {"pageId": id_a}) for _ in range(3)),
                    *(b.call("read_page", {"pageId": id_b}) for _ in range(3)),
                    a.call("pages", WHERE),
                    b.call("notebooks"),
                )
                for read in reads:
                    assert not read.is_error, read.text
                appends = await asyncio.gather(
                    a.write("append_lines", {"pageId": id_a, "lines": ["a2"]}),
                    b.write("append_lines", {"pageId": id_b, "lines": ["b2"]}),
                )
                for _, applied, _ in appends:
                    assert not applied.is_error, applied.text
                assert any(line["text"] == "a2" for line in (await read_page(a, id_a or ""))["lines"])
                assert any(line["text"] == "b2" for line in (await read_page(b, id_b or ""))["lines"])

            await check("two server processes work on separate pages at the same time", two_processes)

            async def latency() -> None:
                before = server_children()
                async with Client(params) as client:
                    fresh = Server(client)
                    (pid_fresh,) = server_children() - before

                    async def timed(name: str, args: dict[str, Any] | None = None) -> float:
                        started = time.monotonic()
                        result = await fresh.call(name, args)
                        elapsed = time.monotonic() - started
                        assert not result.is_error, result.text
                        return elapsed

                    cold = await timed("notebooks")
                    hot_notebooks = [await timed("notebooks") for _ in range(10)]
                    page_args = {"pageId": state["page_id"]}
                    paragraphs = len((await read_page(fresh, state["page_id"]))["lines"])
                    hot_reads = [await timed("read_page", page_args) for _ in range(10)]
                    workers_before = (await windows_workers())[pid_fresh]

                    broken = await fresh.call(
                        "read_page", {"pageId": "{00000000-0000-0000-0000-000000000000}{1}{B0}"}
                    )
                    assert broken.is_error
                    after_error = await timed("notebooks")
                    workers_after = (await windows_workers())[pid_fresh]
                    assert workers_before == workers_after and len(workers_after) == 1, "worker restarted"

                n50, r50 = statistics.median(hot_notebooks), statistics.median(hot_reads)
                print(
                    f"     first call {cold * 1000:.0f} ms; hot notebooks p50 {n50 * 1000:.0f} ms "
                    f"(max {max(hot_notebooks) * 1000:.0f}); hot read_page p50 {r50 * 1000:.0f} ms "
                    f"(max {max(hot_reads) * 1000:.0f}, {paragraphs} paragraphs); "
                    f"after an error {after_error * 1000:.0f} ms"
                )
                limit = 2.5 if was_running else 4.0
                assert cold <= limit, f"first call took {cold:.2f} s (limit {limit} s)"
                assert n50 <= 0.2, f"notebooks p50 {n50 * 1000:.0f} ms"
                assert r50 <= 0.3, f"read_page p50 {r50 * 1000:.0f} ms"
                assert after_error <= 0.5, f"call after an error took {after_error * 1000:.0f} ms"

            await check("a fresh server answers hot calls within the latency gates", latency)

            async def errors_keep_the_worker() -> None:
                bridge = PowerShellBridge(BRIDGE_SCRIPT, "onenote")
                try:
                    await bridge("ping", {})
                    pid = bridge.worker_pid
                    assert pid
                    cases = [
                        ("unknown_op", {}, {"unknown_op"}),
                        ("pages", {}, {"invalid_request"}),
                        (
                            "read_page",
                            {"pageId": "{00000000-0000-0000-0000-000000000000}{1}{B0}"},
                            {"not_found", "com_error"},
                        ),
                        ("read_page", {}, {"invalid_request"}),
                    ]
                    for op, args, codes in cases:
                        try:
                            await bridge(op, args)
                        except BridgeError as error:
                            assert error.code in codes, f"{op}: {error.code} {error}"
                        else:
                            raise AssertionError(f"{op} did not fail")
                        await bridge("ping", {})
                        assert bridge.worker_pid == pid, f"worker restarted after {op}"
                finally:
                    await bridge.close()

            await check("bridge errors of every kind leave the same worker running", errors_keep_the_worker)

            async def sigkill_leaves_nothing() -> None:
                before = server_children()
                process = await asyncio.create_subprocess_exec(
                    sys.executable,
                    "-m",
                    "deskhand",
                    "onenote",
                    stdin=asyncio.subprocess.PIPE,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.DEVNULL,
                )
                assert process.stdin and process.stdout
                to_server, from_server = process.stdin, process.stdout
                try:

                    async def send(message: dict[str, Any]) -> None:
                        to_server.write((json.dumps(message) + "\n").encode())
                        await to_server.drain()

                    async def answer(request_id: int) -> None:
                        while True:
                            line = await asyncio.wait_for(from_server.readline(), 60)
                            assert line, "server closed its output"
                            if json.loads(line).get("id") == request_id:
                                return

                    await send(
                        {
                            "jsonrpc": "2.0",
                            "id": 1,
                            "method": "initialize",
                            "params": {
                                "protocolVersion": "2025-06-18",
                                "capabilities": {},
                                "clientInfo": {"name": "live", "version": "0"},
                            },
                        }
                    )
                    await answer(1)
                    await send({"jsonrpc": "2.0", "method": "notifications/initialized"})
                    await send(
                        {
                            "jsonrpc": "2.0",
                            "id": 2,
                            "method": "tools/call",
                            "params": {"name": "notebooks", "arguments": {}},
                        }
                    )
                    await answer(2)
                    (pid,) = server_children() - before
                    (worker,) = (await windows_workers())[pid]
                    assert await windows_process_exists(worker)
                    os.kill(pid, signal.SIGKILL)
                    await process.wait()
                    deadline = time.monotonic() + 90
                    while await windows_process_exists(worker):
                        assert time.monotonic() < deadline, "the Windows worker outlived its killed server"
                        await asyncio.sleep(2)
                finally:
                    if process.returncode is None:
                        process.kill()
                        await process.wait()

            await check("a SIGKILLed server leaves no Windows worker behind", sigkill_leaves_nothing)

            async def workers_never_restarted() -> None:
                workers = await windows_workers()
                assert workers.get(pid_a) == [first_worker["a"]], workers
                assert len(workers.get(pid_b, [])) == 1, workers

            await check("neither server restarted its worker during the run", workers_never_restarted)

            async def delete() -> None:
                page_id = state["page_id"]
                wrong = await a.call("delete_page", {"pageId": page_id, "expectName": "not the title"})
                assert wrong.is_error, "a wrong expectName must fail"
                assert await page_id_of(title)
                _, applied, _ = await a.write("delete_page", {"pageId": page_id, "expectName": title})
                assert not applied.is_error, applied.text
                assert await page_id_of(title) is None

            await check("delete_page checks the name, then removes the page", delete)

            async def survives_onenote_restart() -> None:
                if was_running:
                    print("     skipped: OneNote was open before the run, so it is not killed")
                    return
                bridge = PowerShellBridge(BRIDGE_SCRIPT, "onenote")
                try:
                    await bridge("ping", {})
                    pid = bridge.worker_pid
                    assert pid
                    assert await onenote_running(), "the worker's COM reference should keep OneNote running"
                    victim = await lone_background_onenote()
                    if victim is None:
                        print("     skipped: ONENOTE.EXE is not a lone windowless instance owned by this run")
                        return
                    await run_powershell(f"Stop-Process -Id {victim} -Force")
                    await asyncio.sleep(1)
                    await bridge("ping", {})
                    assert bridge.worker_pid == pid, "the worker restarted instead of reconnecting"
                finally:
                    await bridge.close()

            await check("a worker reconnects after OneNote exits between calls", survives_onenote_restart)
        finally:
            for page_id in created:
                pages = (await a.call("pages", WHERE)).data["pages"]
                name = next((p["name"] for p in pages if p["id"] == page_id), None)
                if not name:
                    continue
                try:
                    _, applied, _ = await a.write("delete_page", {"pageId": page_id, "expectName": name})
                    if applied.is_error:
                        print(f"cleanup of {name} failed: {applied.text}")
                except Exception as error:
                    print(f"cleanup of {name} failed: {error}")

    failed = sum(1 for _, ok in results if not ok)
    print(f"\n{len(results) - failed}/{len(results)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
