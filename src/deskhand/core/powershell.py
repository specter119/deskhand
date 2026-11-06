"""Keeps one resident powershell.exe per server and sends it one bridge op at a time, from WSL.

The transport constraints (UTF-8 request file, ~200-character -EncodedCommand bootstrap, no -File, no
environment variables for paths, a private stdin pipe) are explained in onenote/AGENTS.md and must stay as
they are.
"""

from __future__ import annotations

import asyncio
import base64
import contextlib
import json
import logging
import os
import re
import shutil
import time
import uuid
from asyncio.subprocess import DEVNULL, PIPE
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from deskhand.core.compat import compact

POWERSHELL_FALLBACK = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
# Clients give up on a tool call after their own timeout (Pi: 60 s by default, and it does not wait for a reply
# after that). Startup, the temp-dir probe and the call add up to a cold first call, so the sum has to stay
# under that: a hot call waits at most CALL_TIMEOUT_S, a cold one about that plus a few seconds of startup.
CALL_TIMEOUT_S = 40.0
STARTUP_TIMEOUT_S = 15.0
IDLE_TIMEOUT_S = 1800.0
PROBE_TIMEOUT_S = 15.0
EXIT_GRACE_S = 5.0
CLOSE_GRACE_S = 2.0
STDERR_SETTLE_S = 0.5
STDERR_TAIL_BYTES = 64 * 1024
STDOUT_LINE_LIMIT = 1 << 20
PROTOCOL_VERSION = "1"

log = logging.getLogger("deskhand.powershell")

BridgeFn = Callable[[str, dict[str, Any]], Awaitable[dict[str, Any]]]
"""One bridge call: an op name plus its arguments, resolved to the op's `data` object."""


class BridgeError(Exception):
    """A failure reported by the bridge, or by the transport around it. `code` is machine-readable."""

    def __init__(self, message: str, code: str) -> None:
        super().__init__(message)
        self.code = code


@dataclass(frozen=True)
class WorkDir:
    win: str
    """Windows path, e.g. C:\\Users\\me\\AppData\\Local\\Temp\\deskhand\\onenote"""
    wsl: Path
    """The same directory as a WSL path."""


@dataclass(frozen=True)
class _Setup:
    powershell: str
    work: WorkDir
    script: str


class _WorkerGone(Exception):
    """The worker's stdout closed: the process exited or was killed."""


def new_call_stem() -> str:
    """Every call uses its own request/response pair.

    The pid keeps two server processes apart (each client starts its own), and the uuid keeps calls
    inside one process apart.
    """
    return f"{os.getpid()}-{uuid.uuid4()}"


def find_powershell(path: str | None = None) -> str:
    found = shutil.which("powershell.exe", path=path)
    if found:
        return found
    if os.access(POWERSHELL_FALLBACK, os.X_OK):
        return POWERSHELL_FALLBACK
    raise BridgeError(
        "powershell.exe was not found: this server must run in WSL with Windows interop enabled",
        "no_windows_host",
    )


def _ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _decode(data: bytes) -> str:
    # Windows PowerShell writes errors in the OEM code page, so undecodable bytes are expected.
    return data.decode("utf-8", "replace")


async def _resolve_work_dir(powershell: str, subdir: str) -> WorkDir:
    try:
        proc = await asyncio.create_subprocess_exec(
            powershell,
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "[IO.Path]::GetTempPath()",
            stdin=DEVNULL,
            stdout=PIPE,
            stderr=PIPE,
        )
        try:
            stdout, _ = await asyncio.wait_for(proc.communicate(), PROBE_TIMEOUT_S)
        except BaseException:
            _kill_now(proc)
            raise
        if proc.returncode != 0:
            raise OSError(f"exit code {proc.returncode}")
    except (OSError, TimeoutError) as error:
        raise BridgeError(
            f"powershell.exe could not be run from WSL ({error or type(error).__name__})",
            "no_windows_host",
        ) from error

    work = work_dir_from_temp(_decode(stdout), subdir)
    work.wsl.mkdir(parents=True, exist_ok=True)
    return work


def work_dir_from_temp(win_temp: str, subdir: str) -> WorkDir:
    win_temp = re.sub(r"[\\/]+$", "", win_temp.replace("\r", "").replace("\n", "").strip())
    win = f"{win_temp}\\deskhand\\{subdir}"
    drive = re.match(r"^([A-Za-z]):[\\/](.*)$", win)
    if not drive:
        raise BridgeError(
            f"Unsupported Windows temp path (expected a drive-letter path): {win}", "no_windows_host"
        )
    return WorkDir(win=win, wsl=Path(f"/mnt/{drive[1].lower()}/{drive[2].replace(chr(92), '/')}"))


_reapers: set[asyncio.Task[Any]] = set()


def _kill_now(proc: asyncio.subprocess.Process) -> None:
    """Kill the child without awaiting anything.

    The MCP SDK cancels handlers with an anyio scope, which re-raises on every `await` until the scope
    exits, so state changes must finish before the first await of a cleanup path. asyncio reaps the
    child in the background.
    """
    if proc.returncode is None:
        with contextlib.suppress(ProcessLookupError):
            proc.kill()
    with contextlib.suppress(Exception):
        task = asyncio.get_running_loop().create_task(proc.wait())
        _reapers.add(task)
        task.add_done_callback(_reapers.discard)


class _Worker:
    """One powershell.exe running bridge.ps1, plus the stderr it has written so far."""

    def __init__(self, proc: asyncio.subprocess.Process) -> None:
        self.proc = proc
        self.windows_pid = 0
        self.last_used = time.monotonic()
        self._stderr = bytearray()
        self._stderr_total = 0
        self.drain = asyncio.get_running_loop().create_task(self._drain_stderr())

    async def _drain_stderr(self) -> None:
        # An undrained stderr pipe fills up and blocks the worker, so this must always be reading.
        assert self.proc.stderr is not None
        while chunk := await self.proc.stderr.read(4096):
            self._stderr += chunk
            self._stderr_total += len(chunk)
            if len(self._stderr) > STDERR_TAIL_BYTES:
                del self._stderr[:-STDERR_TAIL_BYTES]

    @property
    def stderr_mark(self) -> int:
        return self._stderr_total

    def stderr_since(self, mark: int) -> str:
        fresh = min(self._stderr_total - mark, len(self._stderr))
        return _decode(bytes(self._stderr[len(self._stderr) - fresh :])).strip() if fresh > 0 else ""

    def alive(self) -> bool:
        return self.proc.returncode is None and not (self.proc.stdout and self.proc.stdout.at_eof())

    def discard(self) -> None:
        """Kill the worker. Synchronous on purpose, see _kill_now."""
        self.drain.cancel()
        _kill_now(self.proc)

    async def settle_stderr(self) -> None:
        """After the worker has exited, give the reader a moment to take the last of its stderr."""
        await asyncio.wait({self.drain}, timeout=STDERR_SETTLE_S)

    async def stop(self, grace: float) -> None:
        """Ask the worker to leave by closing its stdin, and kill it if it does not."""
        try:
            if self.proc.stdin and not self.proc.stdin.is_closing():
                self.proc.stdin.close()
            await asyncio.wait_for(self.proc.wait(), grace)
        except TimeoutError:
            pass
        finally:
            self.discard()


class PowerShellBridge:
    """Calls one op of a PowerShell bridge script. Nothing touches Windows until the first call.

    The first call starts a resident powershell.exe that loads the script once; every later call
    costs one stdin line and two small files. Calls run one at a time because OneNote's COM server
    is single-threaded anyway.
    """

    def __init__(
        self,
        script_path: Path,
        work_subdir: str,
        *,
        timeout: float = CALL_TIMEOUT_S,
        startup_timeout: float = STARTUP_TIMEOUT_S,
        idle_timeout: float = IDLE_TIMEOUT_S,
        powershell: str | None = None,
        work_dir: WorkDir | None = None,
    ) -> None:
        self._script_path = script_path
        self._work_subdir = work_subdir
        self._timeout = timeout
        self._startup_timeout = startup_timeout
        self._idle_timeout = idle_timeout
        self._powershell = powershell
        self._work_dir = work_dir
        self._setup: _Setup | None = None
        self._lock = asyncio.Lock()
        self._worker: _Worker | None = None
        self._idle_handle: asyncio.TimerHandle | None = None
        self._stopping: set[asyncio.Task[Any]] = set()

    @property
    def worker_pid(self) -> int | None:
        """The Windows process id of the resident worker, or None while none is running."""
        return self._worker.windows_pid if self._worker else None

    async def _ensure_setup(self) -> _Setup:
        # Only a successful setup is cached, and it is not shared as a future: one cancelled caller must
        # not take every other waiting caller down with it.
        if self._setup is not None:
            return self._setup
        powershell = self._powershell or find_powershell()
        work = self._work_dir or await _resolve_work_dir(powershell, self._work_subdir)
        script = self._script_path.read_text(encoding="utf-8")
        self._setup = _Setup(powershell, work, script)
        return self._setup

    async def __call__(self, op: str, args: dict[str, Any]) -> dict[str, Any]:
        setup = await self._ensure_setup()
        stem = new_call_stem()
        request_path = setup.work.wsl / f"req-{stem}.json"
        response_path = setup.work.wsl / f"resp-{stem}.json"
        started = time.monotonic()

        try:
            request_path.write_text(
                json.dumps(compact({"op": op, **args}), ensure_ascii=True), encoding="utf-8"
            )
            applying = args.get("apply") is True
            async with self._lock:
                cold = await self._exchange(setup, stem, applying=applying)
            data = self._read_response(response_path, applying=applying)
            log.info("%s %.0f ms%s", op, (time.monotonic() - started) * 1000, " (cold start)" if cold else "")
            return data
        finally:
            request_path.unlink(missing_ok=True)
            response_path.unlink(missing_ok=True)

    async def _send(self, setup: _Setup, stem: str) -> tuple[_Worker, int, bool]:
        """Hand the stem to a live worker, starting one if needed. Returns (worker, stderr mark, cold)."""
        cold = False
        failure: _Worker | None = None
        mark = 0
        for _ in range(2):
            if self._worker is not None and not self._worker.alive():
                self._drop(self._worker)
            if self._worker is None:
                self._worker = await self._spawn(setup)
                cold = True
            worker = self._worker
            self._cancel_idle()
            mark = worker.stderr_mark
            try:
                assert worker.proc.stdin is not None
                worker.proc.stdin.write(f"{stem}\n".encode("ascii"))
                await worker.proc.stdin.drain()
            except (BrokenPipeError, ConnectionResetError):
                # The worker never read the stem, so sending it to another worker cannot run it twice.
                self._drop(worker)
                failure = worker
            except BaseException:
                self._drop(worker)
                raise
            else:
                return worker, mark, cold

        assert failure is not None
        await failure.settle_stderr()
        raise BridgeError(
            f"OneNote bridge failed: {failure.stderr_since(mark) or 'powershell.exe is not accepting requests'}",
            "bridge_failed",
        )

    async def _exchange(self, setup: _Setup, stem: str, *, applying: bool) -> bool:
        """Send one stem and wait for its done line. Returns whether a worker had to be started."""
        worker, mark, cold = await self._send(setup, stem)

        try:
            async with asyncio.timeout(self._timeout):
                await self._wait_done(worker, stem)
        except TimeoutError:
            self._drop(worker)
            if applying:
                raise BridgeError(
                    "The call timed out while a write was in flight, so whether OneNote applied it is "
                    "unknown. OneNote may be showing a sign-in or update dialog, or be hung. "
                    "Read the page again and preview before retrying.",
                    "outcome_unknown",
                ) from None
            raise BridgeError(
                f"OneNote bridge failed: powershell.exe did not finish within {self._timeout:g} s "
                "(OneNote may be showing a sign-in or update dialog, or be hung)",
                "bridge_failed",
            ) from None
        except _WorkerGone:
            self._drop(worker, cancel_drain=False)
            await worker.settle_stderr()
            if applying:
                raise BridgeError(
                    "powershell.exe exited while a write was in flight, so whether OneNote applied it is "
                    "unknown. Read the page again and preview before retrying.",
                    "outcome_unknown",
                ) from None
            detail = worker.stderr_since(mark)
            raise BridgeError(
                f"OneNote bridge failed: {detail or 'powershell.exe exited during the call'}", "bridge_failed"
            ) from None
        except ValueError:
            self._drop(worker)
            if applying:
                raise BridgeError(
                    "powershell.exe wrote an over-long line while a write was in flight, so whether OneNote "
                    "applied it is unknown. Read the page again and preview before retrying.",
                    "outcome_unknown",
                ) from None
            raise BridgeError(
                "OneNote bridge failed: powershell.exe wrote an over-long line", "bridge_failed"
            ) from None
        except BaseException:
            # Includes cancellation by the client: the worker may be mid-call, so it cannot be reused.
            self._drop(worker)
            raise

        worker.last_used = time.monotonic()
        self._schedule_idle()
        return cold

    @staticmethod
    async def _wait_done(worker: _Worker, stem: str) -> None:
        assert worker.proc.stdout is not None
        expected = f"deskhand-done {stem}"
        while True:
            line = await worker.proc.stdout.readline()
            if not line:
                raise _WorkerGone
            text = _decode(line).strip()
            if text == expected:
                return
            log.debug("ignored worker output: %s", text[:200])

    async def _spawn(self, setup: _Setup) -> _Worker:
        script_path = setup.work.wsl / f"worker-{os.getpid()}-{uuid.uuid4()}.ps1"
        bootstrap = (
            "$ErrorActionPreference='Stop';$ProgressPreference='SilentlyContinue';"
            f"$d={_ps_quote(setup.work.win)};"
            f'Invoke-Expression ([IO.File]::ReadAllText("$d\\{script_path.name}",[Text.Encoding]::UTF8))'
        )
        encoded = base64.b64encode(bootstrap.encode("utf-16-le")).decode("ascii")
        started = time.monotonic()
        worker: _Worker | None = None
        try:
            script_path.write_text(setup.script, encoding="utf-8")
            try:
                proc = await asyncio.create_subprocess_exec(
                    setup.powershell,
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded,
                    # A private pipe: the worker must never inherit the MCP connection's stdin.
                    stdin=PIPE,
                    stdout=PIPE,
                    stderr=PIPE,
                    limit=STDOUT_LINE_LIMIT,
                )
            except OSError as error:
                raise BridgeError(
                    f"powershell.exe could not be run from WSL ({error})", "no_windows_host"
                ) from error
            worker = _Worker(proc)
            try:
                async with asyncio.timeout(self._startup_timeout):
                    await self._wait_ready(worker)
            except TimeoutError:
                worker.discard()
                raise BridgeError(
                    f"OneNote bridge failed: powershell.exe was not ready within {self._startup_timeout:g} s"
                    + self._tail(worker),
                    "bridge_failed",
                ) from None
            except _WorkerGone:
                _kill_now(worker.proc)
                await worker.settle_stderr()
                worker.discard()
                raise BridgeError(
                    "OneNote bridge failed: powershell.exe exited during start-up" + self._tail(worker),
                    "bridge_failed",
                ) from None
            except BaseException:
                worker.discard()
                raise
        finally:
            script_path.unlink(missing_ok=True)

        log.info(
            "worker started: windows pid %d in %.0f ms",
            worker.windows_pid,
            (time.monotonic() - started) * 1000,
        )
        return worker

    @staticmethod
    def _tail(worker: _Worker) -> str:
        detail = worker.stderr_since(0)
        return f": {detail}" if detail else ""

    @staticmethod
    async def _wait_ready(worker: _Worker) -> None:
        assert worker.proc.stdout is not None
        while True:
            line = await worker.proc.stdout.readline()
            if not line:
                raise _WorkerGone
            parts = _decode(line).split()
            if len(parts) == 3 and parts[0] == "deskhand-ready":
                if parts[1] != PROTOCOL_VERSION:
                    raise BridgeError(
                        f"OneNote bridge failed: bridge protocol {parts[1]} does not match {PROTOCOL_VERSION}",
                        "bridge_failed",
                    )
                worker.windows_pid = int(parts[2]) if parts[2].isdigit() else 0
                return

    def _drop(self, worker: _Worker, *, cancel_drain: bool = True) -> None:
        """Forget a worker and kill it, all before any await."""
        if self._worker is worker:
            self._worker = None
        if cancel_drain:
            worker.discard()
        else:
            _kill_now(worker.proc)

    def _cancel_idle(self) -> None:
        if self._idle_handle is not None:
            self._idle_handle.cancel()
            self._idle_handle = None

    def _schedule_idle(self) -> None:
        self._cancel_idle()
        if self._worker is not None and self._idle_timeout > 0:
            self._idle_handle = asyncio.get_running_loop().call_later(
                self._idle_timeout, self._on_idle, self._worker
            )

    def _on_idle(self, worker: _Worker) -> None:
        self._idle_handle = None
        # A call in progress reschedules this when it finishes.
        if self._worker is not worker or self._lock.locked():
            return
        self._worker = None
        self._track(asyncio.get_running_loop().create_task(worker.stop(EXIT_GRACE_S)))

    def _track(self, task: asyncio.Task[Any]) -> None:
        self._stopping.add(task)
        task.add_done_callback(self._stopping.discard)

    async def close(self) -> None:
        """Stop the worker if one is running. Never starts one."""
        self._cancel_idle()
        worker, self._worker = self._worker, None
        pending = list(self._stopping)
        if worker is not None:
            try:
                await asyncio.wait_for(worker.stop(CLOSE_GRACE_S), CLOSE_GRACE_S + 1)
            except BaseException:
                worker.discard()
                raise
        if pending:
            await asyncio.wait(pending, timeout=CLOSE_GRACE_S)

    @staticmethod
    def _read_response(response_path: Path, *, applying: bool) -> dict[str, Any]:
        try:
            parsed = json.loads(response_path.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError):
            # A write that left no readable answer may still have landed, so it must not read as a plain failure.
            if applying:
                raise BridgeError(
                    "OneNote bridge produced no readable response for a write, so whether OneNote applied it "
                    "is unknown. Read the page again and preview before retrying.",
                    "outcome_unknown",
                ) from None
            raise BridgeError(
                "OneNote bridge produced no response; the bridge script may have failed before it "
                "could report an error",
                "no_response",
            ) from None
        if not parsed.get("ok"):
            code = parsed.get("code") or "error"
            raise BridgeError(f"OneNote {code}: {parsed.get('error')}", code)
        return parsed["data"]
