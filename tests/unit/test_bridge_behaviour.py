"""Runs bridge.ps1's request handling against a fake OneNote object in Windows PowerShell 5.1.

The fake raises COM exceptions with chosen HRESULTs, so these tests pin how a failure around a write is
classified and what the connection to OneNote does when it is busy or gone. Everything runs in one
powershell.exe, because starting it is the expensive part.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import uuid
from typing import Any

import pytest

from deskhand.core.powershell import work_dir_from_temp
from deskhand.onenote.server import BRIDGE_SCRIPT

pytestmark = pytest.mark.skipif(shutil.which("powershell.exe") is None, reason="needs Windows PowerShell 5.1")

HARNESS = r"""param([string]$Head, [string]$Work)
$ErrorActionPreference = 'Stop'
Invoke-Expression ([IO.File]::ReadAllText($Head, [Text.Encoding]::UTF8))
$script:WORK = $Work
$PROBE_BUSY_S = 1
$X = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'

function Com([string]$hex, [string]$msg = '') {
	$bits = [uint32]::Parse($hex.Substring(2), [Globalization.NumberStyles]::AllowHexSpecifier)
	$hr = [BitConverter]::ToInt32([BitConverter]::GetBytes($bits), 0)
	if ([string]::IsNullOrEmpty($msg)) { $msg = 'fake ' + $hex }
	return [Runtime.InteropServices.COMException]::new($msg, $hr)
}

function New-Fake($b) {
	$o = [pscustomobject]@{ Calls = (New-Object System.Collections.ArrayList); Probes = 0; Updated = $false; B = $b }
	Add-Member -InputObject $o -MemberType ScriptMethod -Name GetHierarchy -Value {
		if ($args[0] -eq '') {
			$this.Probes++
			if ($this.B.Probe) { & $this.B.Probe $this }
			$args[2].Value = ('<one:Notebooks ' + $X + '/>')
		}
		else {
			$args[2].Value = ('<one:Section ' + $X + ' ID="S1" name="Sec"/>')
		}
	}
	Add-Member -InputObject $o -MemberType ScriptMethod -Name GetPageContent -Value {
		if ($this.B.Read) { & $this.B.Read $this }
		$args[1].Value = ('<one:Page ' + $X + ' ID="' + $args[0] + '" name="Probe page" lastModifiedTime="2026-01-01T00:00:00.000Z"><one:Title><one:OE><one:T><![CDATA[Probe page]]></one:T></one:OE></one:Title></one:Page>')
	}
	Add-Member -InputObject $o -MemberType ScriptMethod -Name UpdatePageContent -Value {
		[void]$this.Calls.Add('update/' + $args.Count + '/' + $args[3].GetType().Name + '/' + $args[3])
		if ($this.B.Update) { & $this.B.Update $this }
		$this.Updated = $true
	}
	Add-Member -InputObject $o -MemberType ScriptMethod -Name CreateNewPage -Value {
		[void]$this.Calls.Add('create')
		if ($this.B.Create) { & $this.B.Create $this }
		$args[1].Value = 'NEWPAGE'
	}
	Add-Member -InputObject $o -MemberType ScriptMethod -Name DeleteHierarchy -Value {
		[void]$this.Calls.Add('delete/' + $args.Count)
		if ($this.B.Delete) { & $this.B.Delete $this }
	}
	return $o
}

$script:NewCalls = 0
$script:Factory = $null
function New-OneNoteApp { $script:NewCalls++; & $script:Factory $script:NewCalls }

$results = New-Object System.Collections.ArrayList
function Run($name, $req, $app, $factory, $fakes) {
	Reset-OneNote
	$script:NewCalls = 0
	$script:Factory = $factory
	$script:OneNoteApp = $app
	$stem = 't' + $results.Count
	$json = $req | ConvertTo-Json -Depth 10 -Compress
	[IO.File]::WriteAllText((Join-Path $Work ('req-' + $stem + '.json')), $json, [Text.UTF8Encoding]::new($false))
	$null = Invoke-Request $stem *> $null
	$resp = [IO.File]::ReadAllText((Join-Path $Work ('resp-' + $stem + '.json'))) | ConvertFrom-Json
	$calls = @()
	foreach ($f in $fakes) { $calls += @($f.Calls) }
	[void]$results.Add([pscustomobject]@{
			name = $name; ok = $resp.ok; code = $resp.code; error = $resp.error
			calls = @($calls); probes = $fakes[0].Probes; newCalls = $script:NewCalls
			kept = ($null -ne $app -and [object]::ReferenceEquals($script:OneNoteApp, $app))
		})
}

function Append-Req { @{ op = 'append'; pageId = 'P1'; lines = @('hello'); apply = $true; expectFingerprint = (Get-Sha256Hex 'page:P1') } }
function Delete-Req { @{ op = 'delete_page'; pageId = 'P1'; expectName = 'Probe page'; apply = $true; expectFingerprint = (Get-Sha256Hex 'delete:P1|Probe page') } }
function Create-Req { @{ op = 'create_page'; sectionId = 'S1'; title = 'T'; lines = @('x'); apply = $true; expectFingerprint = (Get-Sha256Hex 'create:S1') } }
function Read-Req { @{ op = 'read_page'; pageId = 'P1' } }
function Case($name, $req, $behaviour) {
	$f = New-Fake $behaviour
	Run $name $req $f $null @($f)
}

try {
	Case 'append_ok' (Append-Req) @{}
	Case 'append_rejected' (Append-Req) @{ Update = { throw (Com '0x80042001') } }
	Case 'append_stale' (Append-Req) @{ Update = { throw (Com '0x80042010') } }
	Case 'append_busy' (Append-Req) @{ Update = { throw (Com '0x80010001') } }
	Case 'append_recording' (Append-Req) @{ Update = { throw (Com '0x80042024') } }
	Case 'append_gone' (Append-Req) @{ Update = { throw (Com '0x800706BA') } }
	Case 'append_efail' (Append-Req) @{ Update = { throw (Com '0x80004005') } }
	Case 'append_onenote_timeout' (Append-Req) @{ Update = { throw (Com '0x80042023') } }
	Case 'append_object_missing' (Append-Req) @{ Update = { throw (Com '0x8004200E') } }
	Case 'append_unknown_in_range' (Append-Req) @{ Update = { throw (Com '0x8004207F') } }
	Case 'append_unknown_long' (Append-Req) @{ Update = { throw (Com '0x80004005' ('x' * 4000)) } }
	Case 'append_readback_fails' (Append-Req) @{ Read = { param($f) if ($f.Updated) { throw (Com '0x80042014') } } }
	Case 'delete_ok' (Delete-Req) @{}
	Case 'delete_refused' (Delete-Req) @{ Delete = { throw (Com '0x80042014') } }
	Case 'delete_gone' (Delete-Req) @{ Delete = { throw (Com '0x800706BE') } }
	Case 'create_ok' (Create-Req) @{}
	Case 'create_refused' (Create-Req) @{ Create = { throw (Com '0x80042001') } }
	Case 'create_gone' (Create-Req) @{ Create = { throw (Com '0x800706BA') } }
	Case 'create_title_fails' (Create-Req) @{ Update = { throw (Com '0x80042001') } }

	Case 'probe_busy_then_ok' (Read-Req) @{ Probe = { param($f) if ($f.Probes -le 2) { throw (Com '0x80010001') } } }
	Case 'probe_always_busy' (Read-Req) @{ Probe = { throw (Com '0x8001010A') } }

	foreach ($hr in '0x80004005', '0x800706BA') {
		$script:DeadHr = $hr
		$dead = New-Fake @{ Probe = { throw (Com $script:DeadHr) } }
		$script:Fresh = New-Fake @{}
		Run ('probe_dead_' + $hr) (Read-Req) $dead { $script:Fresh } @($dead, $script:Fresh)
	}

	$script:Fresh = New-Fake @{}
	Run 'create_fails_once' (Read-Req) $null { param($n) if ($n -eq 1) { throw (Com '0x800706BF') } else { $script:Fresh } } @($script:Fresh)
	Run 'create_fails_twice' (Read-Req) $null { throw (Com '0x800706BF') } @($script:Fresh)

	$bindingHex = 'unset'
	try { [void][int]'not a number' } catch { $bindingHex = Get-HResultHex $_ }
	[void]$results.Add([pscustomobject]@{ name = 'binding_error_hex'; hex = $bindingHex; code = (Get-WriteFailureCode $bindingHex) })
	[Console]::Out.WriteLine('RESULTS ' + (ConvertTo-Json -InputObject @($results) -Depth 6 -Compress))
}
catch {
	[Console]::Out.WriteLine('HARNESS-ERROR ' + $_.Exception.Message + ' | ' + $_.InvocationInfo.PositionMessage)
}
"""

UPDATE_OK = "update/4/Boolean/False"


def _powershell(*args: str, timeout: float = 120) -> str:
    return subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", *args],
        capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        timeout=timeout,
    ).stdout


@pytest.fixture(scope="module")
def outcomes() -> dict[str, dict[str, Any]]:
    temp = _powershell("-Command", "[IO.Path]::GetTempPath()")
    work = work_dir_from_temp(temp, f"test-behaviour-{uuid.uuid4().hex[:8]}")
    work.wsl.mkdir(parents=True)
    try:
        source = BRIDGE_SCRIPT.read_text(encoding="utf-8")
        (work.wsl / "head.ps1").write_text(source.split("# --- worker loop", 1)[0], encoding="utf-8")
        (work.wsl / "harness.ps1").write_text(HARNESS, encoding="utf-8")
        out = _powershell(
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            f"{work.win}\\harness.ps1",
            "-Head",
            f"{work.win}\\head.ps1",
            "-Work",
            work.win,
        )
    finally:
        shutil.rmtree(work.wsl, ignore_errors=True)
    lines = [line.strip() for line in out.splitlines() if line.startswith(("RESULTS ", "HARNESS-ERROR"))]
    assert lines, out
    assert lines[-1].startswith("RESULTS "), lines[-1]
    return {item["name"]: item for item in json.loads(lines[-1].removeprefix("RESULTS "))}


@pytest.mark.parametrize(
    ("name", "ok", "code"),
    [
        ("append_ok", True, None),
        ("append_rejected", False, "rejected"),
        ("append_stale", False, "stale"),
        ("append_busy", False, "busy"),
        ("append_recording", False, "busy"),
        ("append_gone", False, "outcome_unknown"),
        ("append_efail", False, "outcome_unknown"),
        ("append_onenote_timeout", False, "outcome_unknown"),
        ("append_object_missing", False, "rejected"),
        ("append_unknown_in_range", False, "outcome_unknown"),
        ("delete_ok", True, None),
        ("delete_refused", False, "rejected"),
        ("delete_gone", False, "outcome_unknown"),
        ("create_ok", True, None),
        ("create_refused", False, "rejected"),
        ("create_gone", False, "outcome_unknown"),
    ],
)
def test_a_failure_around_a_write_is_classified_by_what_is_known(outcomes, name, ok, code):
    outcome = outcomes[name]
    assert outcome["ok"] is ok, outcome
    assert outcome["code"] == code, outcome


def test_the_error_text_of_each_classification_is_pinned(outcomes):
    # A wording change is a contract change: it belongs in a commit that says so, not in a drive-by edit.
    assert outcomes["append_stale"]["error"] == (
        "OneNote refused the update as stale (0x80042010); re-read the page and retry"
    )
    assert outcomes["append_busy"]["error"] == (
        "OneNote is busy (0x80010001) and did not update the page; try again shortly"
    )
    assert outcomes["append_rejected"]["error"] == (
        "OneNote refused to update the page (0x80042001): fake 0x80042001"
    )
    assert outcomes["append_gone"]["error"] == (
        "Read the page again and preview before retrying: whether OneNote managed to update the page "
        "is unknown (0x800706BA): fake 0x800706BA"
    )


def test_an_unrecognised_code_in_the_onenote_range_is_not_called_refused(outcomes):
    # The write may have landed; calling it refused would invite a retry that duplicates it.
    outcome = outcomes["append_unknown_in_range"]
    assert outcome["code"] == "outcome_unknown", outcome
    assert outcome["error"].startswith("Read the page again and preview before retrying"), outcome


def test_a_long_diagnostic_does_not_push_the_guidance_out_of_the_reply(outcomes):
    outcome = outcomes["append_unknown_long"]
    assert outcome["code"] == "outcome_unknown", outcome
    assert outcome["error"].startswith("Read the page again and preview before retrying"), outcome
    assert "0x80004005" in outcome["error"], outcome
    assert len(outcome["error"]) < 500, outcome
    assert "x" * 400 not in outcome["error"], outcome


def test_a_write_that_returned_is_never_called_refused(outcomes):
    # The write is in OneNote already: a failure to read it back, or a later step failing, must not invite a retry.
    for name in ("append_readback_fails", "create_title_fails"):
        outcome = outcomes[name]
        assert outcome["code"] == "applied_unverified", outcome
        assert outcome["error"].startswith("Part or all of this write was applied"), outcome
    assert outcomes["create_title_fails"]["calls"] == ["create", UPDATE_OK], "no step runs after the failure"


def test_update_page_content_is_never_forced(outcomes):
    assert outcomes["append_ok"]["calls"] == [UPDATE_OK]
    assert outcomes["create_ok"]["calls"] == ["create", UPDATE_OK, UPDATE_OK]
    assert outcomes["delete_ok"]["calls"] == ["delete/1"]


def test_a_refused_write_changes_nothing_else(outcomes):
    assert outcomes["create_refused"]["calls"] == ["create"]
    assert outcomes["delete_refused"]["calls"] == ["delete/1"]


def test_a_busy_onenote_is_waited_for_and_kept(outcomes):
    waited = outcomes["probe_busy_then_ok"]
    assert (waited["ok"], waited["probes"], waited["newCalls"]) == (True, 3, 0), waited

    stuck = outcomes["probe_always_busy"]
    assert (stuck["ok"], stuck["code"], stuck["newCalls"], stuck["kept"]) == (False, "busy", 0, True), stuck


@pytest.mark.parametrize("hresult", ["0x80004005", "0x800706BA"])
def test_a_dead_connection_is_replaced_once(outcomes, hresult):
    replaced = outcomes[f"probe_dead_{hresult}"]
    assert (replaced["ok"], replaced["newCalls"], replaced["kept"]) == (True, 1, False), replaced


def test_creating_the_com_object_is_retried_once(outcomes):
    assert (outcomes["create_fails_once"]["ok"], outcomes["create_fails_once"]["newCalls"]) == (True, 2)
    twice = outcomes["create_fails_twice"]
    assert (twice["ok"], twice["code"], twice["newCalls"]) == (False, "com_error", 2), twice


def test_a_powershell_level_failure_is_not_a_com_result(outcomes):
    assert outcomes["binding_error_hex"]["hex"] == ""
    assert outcomes["binding_error_hex"]["code"] == "com_error"
