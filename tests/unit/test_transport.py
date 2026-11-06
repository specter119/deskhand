"""PowerShellBridge against a fake powershell.exe, so no Windows is needed."""

from __future__ import annotations

import asyncio
import os
import stat
import sys
import textwrap
from pathlib import Path

import pytest

from deskhand.core.powershell import BridgeError, PowerShellBridge, WorkDir, work_dir_from_temp

pytestmark = pytest.mark.anyio

FAKE = textwrap.dedent(
    """\
    #!{python}
    import base64, json, os, re, sys, time

    encoded = sys.argv[sys.argv.index("-EncodedCommand") + 1]
    bootstrap = base64.b64decode(encoded).decode("utf-16-le")
    stem = re.search(r"^\\$f='(.*)'$", bootstrap, re.M).group(1)
    work = os.environ["FAKE_WORK"]
    mode = os.environ.get("FAKE_MODE", "ok")
    open(os.path.join(work, "pid"), "w").write(str(os.getpid()))
    request = json.load(open(os.path.join(work, f"req-{{stem}}.json"), encoding="utf-8"))

    def respond(body):
        with open(os.path.join(work, f"resp-{{stem}}.json"), "w", encoding="utf-8-sig") as out:
            json.dump(body, out)

    if mode == "hang":
        time.sleep(60)
    elif mode == "fail":
        sys.stderr.buffer.write(b"boom \\x82 oem")
        sys.exit(1)
    elif mode == "error":
        respond({{"ok": False, "error": "no such page", "code": "not_found"}})
    elif mode == "silent":
        pass
    else:
        script = request.pop("__script")
        respond({{"ok": True, "data": {{"request": request, "script": script, "stdin": sys.stdin.read()}}}})
    """
)


@pytest.fixture
def bridge_factory(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    fake = tmp_path / "powershell.exe"
    fake.write_text(FAKE.format(python=sys.executable))
    fake.chmod(fake.stat().st_mode | stat.S_IEXEC)
    work = tmp_path / "work"
    work.mkdir()
    script = tmp_path / "bridge.ps1"
    script.write_text("# 脚本 🚀\n", encoding="utf-8")
    monkeypatch.setenv("FAKE_WORK", str(work))

    def make(mode: str = "ok", **kwargs) -> PowerShellBridge:
        monkeypatch.setenv("FAKE_MODE", mode)
        return PowerShellBridge(
            script, "test", powershell=str(fake), work_dir=WorkDir("C:\\fake", work), **kwargs
        )

    make.work = work  # type: ignore[attr-defined]
    return make


def _leftovers(work: Path) -> list[str]:
    return sorted(p.name for p in work.iterdir() if p.name != "pid")


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    return True


async def test_a_call_round_trips_unicode_and_cleans_up(bridge_factory):
    data = await bridge_factory()("append", {"lines": ["中文 🚀"], "width": 80.0, "gone": None})
    assert data["request"] == {"op": "append", "lines": ["中文 🚀"], "width": 80}
    assert data["script"] == "# 脚本 🚀\n"
    assert data["stdin"] == "", "the child must not inherit the server's stdin"
    assert _leftovers(bridge_factory.work) == []


async def test_a_bridge_error_keeps_its_code(bridge_factory):
    with pytest.raises(BridgeError, match="OneNote not_found: no such page") as raised:
        await bridge_factory("error")("read_page", {})
    assert raised.value.code == "not_found"
    assert _leftovers(bridge_factory.work) == []


async def test_a_failing_process_reports_its_stderr_even_when_it_is_not_utf8(bridge_factory):
    with pytest.raises(BridgeError, match="OneNote bridge failed: boom") as raised:
        await bridge_factory("fail")("read_page", {})
    assert raised.value.code == "bridge_failed"


async def test_a_missing_response_is_reported(bridge_factory):
    with pytest.raises(BridgeError) as raised:
        await bridge_factory("silent")("read_page", {})
    assert raised.value.code == "no_response"
    assert _leftovers(bridge_factory.work) == []


@pytest.mark.parametrize(("apply", "code"), [(False, "bridge_failed"), (True, "outcome_unknown")])
async def test_a_timeout_kills_the_child_and_only_a_write_is_outcome_unknown(bridge_factory, apply, code):
    with pytest.raises(BridgeError) as raised:
        await bridge_factory("hang", timeout=0.5)("edit", {"apply": apply})
    assert raised.value.code == code
    assert not _alive(int((bridge_factory.work / "pid").read_text()))
    assert _leftovers(bridge_factory.work) == []


async def test_cancellation_kills_the_child_and_propagates_instead_of_becoming_a_bridge_error(bridge_factory):
    task = asyncio.ensure_future(bridge_factory("hang")("edit", {"apply": True}))
    pid_file = bridge_factory.work / "pid"
    for _ in range(100):
        if pid_file.exists() and pid_file.read_text():
            break
        await asyncio.sleep(0.05)
    task.cancel()
    with pytest.raises(asyncio.CancelledError):
        await task
    assert not _alive(int(pid_file.read_text()))
    assert _leftovers(bridge_factory.work) == []


async def test_parallel_calls_do_not_share_files(bridge_factory):
    bridge = bridge_factory()
    results = await asyncio.gather(*(bridge("echo", {"n": n}) for n in range(8)))
    assert [r["request"]["n"] for r in results] == list(range(8))


def test_the_windows_temp_path_maps_to_a_wsl_path():
    work = work_dir_from_temp("C:\\Users\\me\\AppData\\Local\\Temp\\\r\n", "onenote")
    assert work.win == "C:\\Users\\me\\AppData\\Local\\Temp\\deskhand\\onenote"
    assert work.wsl == Path("/mnt/c/Users/me/AppData/Local/Temp/deskhand/onenote")
    with pytest.raises(BridgeError):
        work_dir_from_temp("\\\\server\\share", "onenote")
