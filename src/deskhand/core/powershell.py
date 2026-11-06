"""Runs one bridge op per powershell.exe process, from WSL.

The transport constraints (UTF-8 request file, ~200-character -EncodedCommand bootstrap, no -File, no
environment variables for paths) are explained in onenote/AGENTS.md and must stay as they are.
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
CALL_TIMEOUT_S = 180.0
PROBE_TIMEOUT_S = 30.0

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


def new_call_stem() -> str:
    """Every call uses its own request/response pair.

    The pid keeps two server processes apart (each client starts its own), and the uuid keeps parallel
    calls inside one process apart.
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
            await _kill(proc)
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


async def _kill(proc: asyncio.subprocess.Process) -> None:
    """Stop the child and reap it, even when this task is being cancelled."""
    if proc.returncode is None:
        with contextlib.suppress(ProcessLookupError):
            proc.kill()
    await asyncio.shield(proc.wait())


class PowerShellBridge:
    """Calls one op of a PowerShell bridge script. Nothing touches Windows until the first call."""

    def __init__(
        self,
        script_path: Path,
        work_subdir: str,
        *,
        timeout: float = CALL_TIMEOUT_S,
        powershell: str | None = None,
        work_dir: WorkDir | None = None,
    ) -> None:
        self._script_path = script_path
        self._work_subdir = work_subdir
        self._timeout = timeout
        self._powershell = powershell
        self._work_dir = work_dir
        self._setup: _Setup | None = None

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

        body = json.dumps(compact({"op": op, "__script": setup.script, **args}), ensure_ascii=True)
        request_path.write_text(body, encoding="utf-8")
        bootstrap = "\n".join(
            [
                "$ErrorActionPreference='Stop'",
                f"$d={_ps_quote(setup.work.win)}",
                f"$f={_ps_quote(stem)}",
                '$i=[IO.File]::ReadAllText("$d\\req-$f.json",[Text.Encoding]::UTF8)|ConvertFrom-Json',
                "Invoke-Expression $i.__script",
                "",
            ]
        )
        encoded = base64.b64encode(bootstrap.encode("utf-16-le")).decode("ascii")

        try:
            await self._run(setup.powershell, encoded, applying=args.get("apply") is True)
            data = self._read_response(response_path)
            log.info("%s %.2fs", op, time.monotonic() - started)
            return data
        finally:
            request_path.unlink(missing_ok=True)
            response_path.unlink(missing_ok=True)

    async def _run(self, powershell: str, encoded: str, *, applying: bool) -> None:
        proc = await asyncio.create_subprocess_exec(
            powershell,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-EncodedCommand",
            encoded,
            # Without this the child inherits the MCP connection's stdin and can swallow JSON-RPC.
            stdin=DEVNULL,
            stdout=PIPE,
            stderr=PIPE,
        )
        try:
            async with asyncio.timeout(self._timeout):
                stdout, stderr = await proc.communicate()
        except TimeoutError:
            await _kill(proc)
            if applying:
                raise BridgeError(
                    "The call timed out while a write was in flight, so whether OneNote applied it is "
                    "unknown. Read the page again and preview before retrying.",
                    "outcome_unknown",
                ) from None
            raise BridgeError(
                f"OneNote bridge failed: powershell.exe did not finish within {self._timeout:g} s",
                "bridge_failed",
            ) from None
        except BaseException:
            # Cancellation by the client is not a bridge failure: stop the child and let it propagate.
            await _kill(proc)
            raise

        if proc.returncode != 0:
            detail = (_decode(stderr) or _decode(stdout)).strip()
            raise BridgeError(
                f"OneNote bridge failed: {detail or 'powershell.exe returned a non-zero exit code'}",
                "bridge_failed",
            )

    @staticmethod
    def _read_response(response_path: Path) -> dict[str, Any]:
        try:
            raw = response_path.read_text(encoding="utf-8-sig")
        except OSError:
            raise BridgeError(
                "OneNote bridge produced no response; the bridge script may have failed before it "
                "could report an error",
                "no_response",
            ) from None
        parsed = json.loads(raw)
        if not parsed.get("ok"):
            code = parsed.get("code") or "error"
            raise BridgeError(f"OneNote {code}: {parsed.get('error')}", code)
        return parsed["data"]
