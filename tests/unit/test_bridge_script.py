"""Static checks on bridge.ps1 that the fake worker in test_transport.py cannot make."""

from __future__ import annotations

import re
import shutil
import subprocess

import pytest

from deskhand.onenote.server import BRIDGE_SCRIPT


def test_the_script_never_exits_the_worker():
    # exit between writing a response and printing deskhand-done would make a finished write look lost.
    code = [
        line
        for line in BRIDGE_SCRIPT.read_text(encoding="utf-8").splitlines()
        if not line.lstrip().startswith("#")
    ]
    offenders = [line for line in code if re.search(r"\bexit\b|\[Environment\]::Exit", line, re.I)]
    assert offenders == []


def test_the_script_is_ascii():
    # It is read back with ReadAllText(..., UTF8) today, but ASCII keeps it safe from any decoding mistake.
    BRIDGE_SCRIPT.read_bytes().decode("ascii")


@pytest.mark.skipif(shutil.which("powershell.exe") is None, reason="needs Windows PowerShell 5.1")
def test_the_script_parses_in_windows_powershell():
    path = subprocess.run(
        ["wslpath", "-w", str(BRIDGE_SCRIPT)], capture_output=True, text=True, check=True
    ).stdout.strip()
    command = (
        "$e=$null;$t=$null;"
        f"[void][System.Management.Automation.Language.Parser]::ParseFile('{path}',[ref]$t,[ref]$e);"
        "foreach($x in $e){'{0}:{1} {2}' -f $x.Extent.StartLineNumber,$x.Extent.StartColumnNumber,$x.Message};"
        "'errors=' + @($e).Count"
    )
    out = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
        capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        timeout=60,
    ).stdout
    assert out.strip().splitlines()[-1] == "errors=0", out
