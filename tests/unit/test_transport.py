"""PowerShellBridge against a fake resident powershell.exe, so no Windows is needed.

The fake speaks the same stdin/stdout protocol as bridge.ps1's worker loop. What it cannot test is
PowerShell semantics; test_bridge_script.py and the live suite cover those.
"""

from __future__ import annotations

import asyncio
import os
import stat
import sys
import textwrap
from pathlib import Path

import anyio
import pytest

from deskhand.core.powershell import BridgeError, PowerShellBridge, WorkDir, work_dir_from_temp

pytestmark = pytest.mark.anyio

FAKE = textwrap.dedent(
    """\
    #!{python}
    import base64, json, os, re, sys, time

    encoded = sys.argv[sys.argv.index("-EncodedCommand") + 1]
    bootstrap = base64.b64decode(encoded).decode("utf-16-le")
    name = re.search(r'\\\\(worker-[^"]+\\.ps1)"', bootstrap).group(1)
    work = os.environ["FAKE_WORK"]
    start = os.environ.get("FAKE_START", "ok")
    script = open(os.path.join(work, name), encoding="utf-8").read()
    with open(os.path.join(work, "pids"), "a") as pids:
        pids.write(f"{{os.getpid()}}\\n")

    def say(text):
        sys.stdout.write(text + "\\n")
        sys.stdout.flush()

    if start == "flood":
        sys.stderr.write("x" * 1_000_000)
        sys.stderr.flush()
    elif start == "exit":
        sys.stderr.buffer.write(b"cannot start \\x82 oem")
        sys.exit(1)
    elif start == "hang":
        time.sleep(60)
    say("deskhand-ready " + ("2" if start == "version" else "1") + f" {{os.getpid()}}")

    for line in sys.stdin:
        stem = line.strip()
        path = os.path.join(work, f"req-{{stem}}.json")
        request = json.load(open(path, encoding="utf-8"))
        behave = request.pop("behave", "")

        def respond(body):
            with open(os.path.join(work, f"resp-{{stem}}.json"), "w", encoding="utf-8-sig") as out:
                json.dump(body, out)

        if behave == "hang":
            time.sleep(60)
        elif behave == "exit":
            sys.stderr.buffer.write(b"boom \\x82 oem")
            sys.stderr.flush()
            os._exit(3)
        elif behave == "error":
            respond({{"ok": False, "error": "no such page", "code": "not_found"}})
        elif behave == "silent":
            pass
        elif behave == "garbled":
            with open(os.path.join(work, f"resp-{{stem}}.json"), "w", encoding="utf-8") as out:
                out.write("{{not json")
        elif behave == "longline":
            say("x" * 2_000_000)
        else:
            if behave == "noise":
                say("WARNING: something")
                say("deskhand-done someone-else")
            respond({{"ok": True, "data": {{"request": request, "script": script, "pid": os.getpid()}}}})
        say("deskhand-done " + stem)
    """
)


class Env:
    def __init__(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        self.monkeypatch = monkeypatch
        self.fake = tmp_path / "powershell.exe"
        self.fake.write_text(FAKE.format(python=sys.executable))
        self.fake.chmod(self.fake.stat().st_mode | stat.S_IEXEC)
        self.work = tmp_path / "work"
        self.work.mkdir()
        self.script = tmp_path / "bridge.ps1"
        self.script.write_text("# \u811a\u672c \U0001f680\n", encoding="utf-8")
        monkeypatch.setenv("FAKE_WORK", str(self.work))

    def start(self, mode: str) -> None:
        self.monkeypatch.setenv("FAKE_START", mode)

    def bridge(self, **kwargs) -> PowerShellBridge:
        return PowerShellBridge(
            self.script,
            "test",
            powershell=str(self.fake),
            work_dir=WorkDir("C:\\fake", self.work),
            **kwargs,
        )

    def pids(self) -> list[int]:
        path = self.work / "pids"
        return [int(p) for p in path.read_text().split()] if path.exists() else []

    def leftovers(self) -> list[str]:
        return sorted(p.name for p in self.work.iterdir() if p.name != "pids")


@pytest.fixture
def env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Env:
    return Env(tmp_path, monkeypatch)


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    return True


async def _gone(pid: int) -> bool:
    for _ in range(100):
        if not _alive(pid):
            return True
        await asyncio.sleep(0.02)
    return False


async def test_a_call_round_trips_unicode_and_cleans_up(env):
    bridge = env.bridge()
    data = await bridge("append", {"lines": ["\u4e2d\u6587 \U0001f680"], "width": 80.0, "gone": None})
    assert data["request"] == {"op": "append", "lines": ["\u4e2d\u6587 \U0001f680"], "width": 80}
    assert data["script"] == "# \u811a\u672c \U0001f680\n"
    assert env.leftovers() == [], "request, response and worker script files must all be gone"
    await bridge.close()


async def test_calls_reuse_one_worker(env):
    bridge = env.bridge()
    first = await bridge("a", {})
    second = await bridge("b", {})
    assert first["pid"] == second["pid"] == env.pids()[0]
    assert len(env.pids()) == 1
    assert bridge.worker_pid == first["pid"]
    await bridge.close()


async def test_parallel_calls_are_serialized_and_do_not_share_files(env):
    bridge = env.bridge()
    results = await asyncio.gather(*(bridge("echo", {"n": n}) for n in range(8)))
    assert [r["request"]["n"] for r in results] == list(range(8))
    assert len(env.pids()) == 1
    assert env.leftovers() == []
    await bridge.close()


async def test_a_bridge_error_keeps_its_code_and_leaves_the_worker_usable(env):
    bridge = env.bridge()
    first = await bridge("a", {})
    with pytest.raises(BridgeError, match="OneNote not_found: no such page") as raised:
        await bridge("read_page", {"behave": "error"})
    assert raised.value.code == "not_found"
    assert (await bridge("b", {}))["pid"] == first["pid"]
    assert env.leftovers() == []
    await bridge.close()


async def test_a_missing_response_is_reported_and_the_worker_stays(env):
    bridge = env.bridge()
    first = await bridge("a", {})
    with pytest.raises(BridgeError) as raised:
        await bridge("read_page", {"behave": "silent"})
    assert raised.value.code == "no_response"
    assert (await bridge("b", {}))["pid"] == first["pid"]
    await bridge.close()


@pytest.mark.parametrize("behave", ["silent", "garbled"])
@pytest.mark.parametrize(("apply", "code"), [(False, "no_response"), (True, "outcome_unknown")])
async def test_a_write_without_a_readable_response_is_outcome_unknown(env, behave, apply, code):
    bridge = env.bridge()
    first = await bridge("a", {})
    with pytest.raises(BridgeError) as raised:
        await bridge("append", {"behave": behave, "apply": apply})
    assert raised.value.code == code
    assert (await bridge("b", {}))["pid"] == first["pid"]
    await bridge.close()


@pytest.mark.parametrize(("apply", "code"), [(False, "bridge_failed"), (True, "outcome_unknown")])
async def test_an_over_long_line_during_a_write_is_outcome_unknown(env, apply, code):
    bridge = env.bridge()
    first = await bridge("a", {})
    with pytest.raises(BridgeError) as raised:
        await bridge("append", {"behave": "longline", "apply": apply})
    assert raised.value.code == code
    assert (await bridge("b", {}))["pid"] != first["pid"]
    await bridge.close()


async def test_output_that_is_not_the_protocol_is_ignored(env):
    bridge = env.bridge()
    assert (await bridge("a", {"behave": "noise"}))["request"]["op"] == "a"
    await bridge.close()


async def test_stderr_written_before_ready_does_not_block_the_worker(env):
    env.start("flood")
    bridge = env.bridge(startup_timeout=10)
    assert (await bridge("a", {}))["request"]["op"] == "a"
    await bridge.close()


@pytest.mark.parametrize(("apply", "code"), [(False, "bridge_failed"), (True, "outcome_unknown")])
async def test_a_worker_that_exits_mid_call_is_replaced_on_the_next_call(env, apply, code):
    bridge = env.bridge()
    first = await bridge("a", {})
    with pytest.raises(BridgeError) as raised:
        await bridge("edit", {"behave": "exit", "apply": apply})
    assert raised.value.code == code
    if not apply:
        assert "boom" in str(raised.value)
    assert bridge.worker_pid is None
    second = await bridge("b", {})
    assert second["pid"] != first["pid"]
    assert env.leftovers() == []
    await bridge.close()


@pytest.mark.parametrize("apply", [False, True])
async def test_a_worker_that_died_while_idle_is_replaced_without_an_error(env, apply):
    bridge = env.bridge()
    first = await bridge("a", {})
    os.kill(first["pid"], 9)
    assert await _gone(first["pid"])
    second = await bridge("b", {"apply": apply})
    assert second["pid"] != first["pid"]
    await bridge.close()


@pytest.mark.parametrize(("apply", "code"), [(False, "bridge_failed"), (True, "outcome_unknown")])
async def test_a_timeout_kills_the_worker_and_only_a_write_is_outcome_unknown(env, apply, code):
    bridge = env.bridge(timeout=0.5)
    first = await bridge("a", {})
    with pytest.raises(BridgeError) as raised:
        await bridge("edit", {"behave": "hang", "apply": apply})
    assert raised.value.code == code
    assert await _gone(first["pid"])
    assert (await bridge("b", {}))["pid"] != first["pid"]
    assert env.leftovers() == []
    await bridge.close()


async def test_asyncio_cancellation_kills_the_worker_and_propagates(env):
    bridge = env.bridge()
    first = await bridge("a", {})
    task = asyncio.ensure_future(bridge("edit", {"behave": "hang", "apply": True}))
    await asyncio.sleep(0.3)
    task.cancel()
    with pytest.raises(asyncio.CancelledError):
        await task
    assert await _gone(first["pid"])
    assert (await bridge("b", {}))["pid"] != first["pid"]
    assert env.leftovers() == []
    await bridge.close()


async def test_anyio_cancellation_kills_the_worker_before_its_first_await(env):
    # anyio re-raises at every await until its scope exits, which is how the MCP SDK cancels a handler.
    bridge = env.bridge()
    first = await bridge("a", {})
    with anyio.move_on_after(0.4) as scope:
        await bridge("edit", {"behave": "hang", "apply": True})
    assert scope.cancelled_caught
    assert bridge.worker_pid is None
    assert await _gone(first["pid"])
    assert (await bridge("b", {}))["pid"] != first["pid"]
    assert env.leftovers() == []
    await bridge.close()


async def test_a_cancelled_call_waiting_for_the_lock_leaves_the_worker_alone(env):
    bridge = env.bridge()
    first = await bridge("a", {})
    slow = asyncio.ensure_future(bridge("c", {"behave": "noise"}))
    queued = asyncio.ensure_future(bridge("d", {}))
    queued.cancel()
    assert (await slow)["pid"] == first["pid"]
    with pytest.raises(asyncio.CancelledError):
        await queued
    assert (await bridge("e", {}))["pid"] == first["pid"]
    await bridge.close()


async def test_an_idle_worker_is_stopped_and_the_next_call_starts_a_new_one(env):
    bridge = env.bridge(idle_timeout=0.3)
    first = await bridge("a", {})
    assert await _gone(first["pid"])
    assert bridge.worker_pid is None
    assert (await bridge("b", {}))["pid"] != first["pid"]
    await bridge.close()


async def test_a_failed_start_is_reported_with_stderr_and_not_cached(env):
    env.start("exit")
    bridge = env.bridge()
    with pytest.raises(BridgeError, match="exited during start-up: cannot start") as raised:
        await bridge("a", {})
    assert raised.value.code == "bridge_failed"
    assert env.leftovers() == []
    env.start("ok")
    assert (await bridge("a", {}))["request"]["op"] == "a"
    await bridge.close()


async def test_a_worker_that_never_gets_ready_is_killed(env):
    env.start("hang")
    bridge = env.bridge(startup_timeout=0.5)
    with pytest.raises(BridgeError, match="not ready within") as raised:
        await bridge("a", {})
    assert raised.value.code == "bridge_failed"
    assert await _gone(env.pids()[0])
    assert env.leftovers() == []


async def test_a_worker_with_another_protocol_version_is_refused(env):
    env.start("version")
    bridge = env.bridge()
    with pytest.raises(BridgeError, match="protocol 2") as raised:
        await bridge("a", {})
    assert raised.value.code == "bridge_failed"
    assert await _gone(env.pids()[0])


async def test_close_stops_the_worker_and_does_nothing_when_none_ran(env):
    bridge = env.bridge()
    first = await bridge("a", {})
    await bridge.close()
    assert await _gone(first["pid"])
    assert bridge.worker_pid is None

    never_started = PowerShellBridge(
        env.script, "test", powershell="/nonexistent/powershell.exe", work_dir=WorkDir("C:\\fake", env.work)
    )
    await never_started.close()


def test_the_windows_temp_path_maps_to_a_wsl_path():
    work = work_dir_from_temp("C:\\Users\\me\\AppData\\Local\\Temp\\\r\n", "onenote")
    assert work.win == "C:\\Users\\me\\AppData\\Local\\Temp\\deskhand\\onenote"
    assert work.wsl == Path("/mnt/c/Users/me/AppData/Local/Temp/deskhand/onenote")
    with pytest.raises(BridgeError):
        work_dir_from_temp("\\\\server\\share", "onenote")
