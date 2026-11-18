"""Static checks on bridge.ps1 that the fake worker in test_transport.py cannot make."""

from __future__ import annotations

import re
import shutil
import subprocess

import pytest

from deskhand.onenote.server import BRIDGE_SCRIPT
from deskhand.ppt.server import BRIDGE_SCRIPT as PPT_BRIDGE_SCRIPT

SCRIPTS = [BRIDGE_SCRIPT, PPT_BRIDGE_SCRIPT]


def _ppt_function_body(text: str, name: str) -> str:
    match = re.search(rf"^function {re.escape(name)}(?=[\s({{])", text, re.M)
    assert match is not None, name
    start = match.start()
    end = text.index("\n}\n", start) + 3
    return text[start:end]


@pytest.mark.parametrize("script", SCRIPTS, ids=["onenote", "ppt"])
def test_the_script_never_exits_the_worker(script):
    # exit between writing a response and printing deskhand-done would make a finished write look lost.
    code = [
        line for line in script.read_text(encoding="utf-8").splitlines() if not line.lstrip().startswith("#")
    ]
    offenders = [line for line in code if re.search(r"\bexit\b|\[Environment\]::Exit", line, re.I)]
    assert offenders == []


@pytest.mark.parametrize("script", SCRIPTS, ids=["onenote", "ppt"])
def test_the_script_is_ascii(script):
    # It is read back with ReadAllText(..., UTF8) today, but ASCII keeps it safe from any decoding mistake.
    script.read_bytes().decode("ascii")


def test_the_ppt_worker_does_not_save_or_close_user_decks_on_shutdown():
    text = PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8")
    body = _ppt_function_body(text, "Reset-Ppt")
    assert "Saved" not in body
    assert ".Close(" not in body
    assert "ReleaseComObject" in _ppt_function_body(text, "Reset-PptReleaseRefs")
    assert "ReleaseComObject($script:PptApp" not in text


def test_the_ppt_reset_state_is_explicit():
    text = PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8")
    for name in (
        "RESET_PENDING",
        "RESET_ATTEMPTS",
        "RESET_TERMINAL",
        "PPT_PID",
        "PPT_STARTED_AT",
        "PPT_MAPPING_FAILED",
    ):
        assert f"$script:{name} =" in text


def test_the_ppt_process_probe_has_three_states():
    body = _ppt_function_body(PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8"), "Test-PptProcessAliveWithProbe")
    assert "'unknown'" in body
    assert "'dead'" in body
    assert "'alive'" in body
    assert "PPT_PID" in body
    assert "PPT_STARTED_AT" in body
    assert "StartTime" in body


def test_the_ppt_new_instance_requires_zero_processes_and_cross_checks_window():
    text = PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8")
    new_body = _ppt_function_body(text, "New-PptInstance")
    assert "$before.Count -ne 0" in new_body
    assert "Set-PptNewMapping" in new_body
    map_body = _ppt_function_body(text, "Set-PptNewMapping")
    assert "AddSeconds(3)" in map_body
    assert "Milliseconds 100" in map_body
    assert "StartTime" in map_body
    assert "Get-PptWindowPid" in map_body


def test_the_ppt_attach_uses_getactiveobject_only_when_not_started():
    text = PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8")
    attach_body = _ppt_function_body(text, "Attach-PptInstance")
    assert "[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')" in attach_body
    resolve_body = _ppt_function_body(text, "Resolve-PptNormal")
    assert "(-not $script:STARTED_PPT) -and $processes.Count -eq 1" in resolve_body
    assert "$script:STARTED_PPT -and $processes.Count -gt 0" in resolve_body


def test_the_ppt_get_retries_pid_mapping_without_new_object():
    text = PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8")
    resolve_body = _ppt_function_body(text, "Resolve-PptNormal")
    assert "[int]$script:PPT_PID -eq 0" in resolve_body
    assert "Set-PptMappingFromWindows" in resolve_body
    get_body = _ppt_function_body(text, "Get-Ppt")
    assert "Reset-Ppt -Recovery" in get_body
    assert "continue" in get_body
    assert "New-Object" not in get_body


def test_the_ppt_recovery_only_quits_instances_it_started():
    body = _ppt_function_body(PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8"), "Reset-Ppt")
    assert "if (-not $script:STARTED_PPT)" in body
    assert body.index("if (-not $script:STARTED_PPT)") < body.index("Quit()")
    assert "presentation count failed" in body
    assert "Add-PptResetAttempt" in body
    assert "Clear-PptResetPending" in body
    assert "Clear-PptResetState" not in body


def test_the_ppt_reset_attempts_become_terminal():
    body = _ppt_function_body(PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8"), "Add-PptResetAttempt")
    assert "RESET_ATTEMPTS" in body
    assert "-ge 2" in body
    assert "RESET_TERMINAL = $true" in body


def test_the_ppt_add_path_uses_structured_shape_errors():
    text = PPT_BRIDGE_SCRIPT.read_text(encoding="utf-8")
    assert "'add_shapes' {" in text
    for name in (
        "ConvertTo-Color",
        "Set-TextFrameOptions",
        "Set-Paragraphs",
        "Add-TextShape",
        "Add-TableShape",
        "Test-ShapeSpecs",
        "Add-Shape",
        "Add-PictureShape",
    ):
        start = text.index(f"function {name}")
        end = text.index("\n}\n", start) + 3
        body = text[start:end]
        assert not re.search(r"(?<!-)\bFail\s", body), name


@pytest.mark.skipif(shutil.which("powershell.exe") is None, reason="needs Windows PowerShell 5.1")
@pytest.mark.parametrize("script", SCRIPTS, ids=["onenote", "ppt"])
def test_the_script_parses_in_windows_powershell(script):
    path = subprocess.run(
        ["wslpath", "-w", str(script)], capture_output=True, text=True, check=True
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
