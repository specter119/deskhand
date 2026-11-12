"""The Python server must show clients exactly what the TypeScript server showed.

tests/parity/fixtures/ts-transcripts.json was recorded from the TypeScript implementation (branch
`ts`) by replaying tests/parity/scenario.json against a scripted bridge. This test replays the same
scenario here. Error texts are compared only by `isError`: their wording came from the TypeScript SDK.

fixtures/tools.json is the advertised tool list. It matched the TypeScript record until paragraph
formatting and the `after` anchor were added on purpose; regenerate it with
`uv run python tests/parity/dump_tools.py` whenever the contract is meant to change.
"""

from __future__ import annotations

import base64
import hashlib
import json
import re
from pathlib import Path
from typing import Any

import pytest
from helpers import FakeBridge, connect

from deskhand.core.apply_token import TokenTable
from deskhand.core.powershell import BridgeError
from deskhand.onenote.server import INSTRUCTIONS

PARITY = Path(__file__).parent.parent / "parity"
SCENARIO = json.loads((PARITY / "scenario.json").read_text("utf-8"))
EXPECTED_TOOLS = json.loads((PARITY / "fixtures" / "tools.json").read_text("utf-8"))
EXPECTED_TRANSCRIPTS = json.loads((PARITY / "fixtures" / "ts-transcripts.json").read_text("utf-8"))

pytestmark = pytest.mark.anyio


def _normalize(value: Any, tmp: Path) -> Any:
    raw = json.dumps(value, ensure_ascii=False).replace(json.dumps(str(tmp))[1:-1], "<TMP>")
    return json.loads(re.sub(r"dh1_[0-9a-f]{32}", "<TOKEN>", raw))


def _substitute(value: Any, tmp: Path) -> Any:
    if isinstance(value, str):
        return value.replace("$TMP", str(tmp))
    if isinstance(value, list):
        return [_substitute(item, tmp) for item in value]
    if isinstance(value, dict):
        return {key: _substitute(member, tmp) for key, member in value.items()}
    return value


def _redact(args: dict[str, Any]) -> dict[str, Any]:
    copy = dict(args)
    if isinstance(copy.get("base64"), str):
        copy["base64"] = "sha256:" + hashlib.sha256(copy["base64"].encode()).hexdigest()
    return copy


def _handlers(step: dict[str, Any]) -> dict[str, Any]:
    def make(entry: dict[str, Any]):
        def handle(args: dict[str, Any]) -> Any:
            if "error" in entry:
                raise BridgeError(entry["error"]["message"], entry["error"]["code"])
            if "any" in entry:
                return entry["any"]
            return entry["apply"] if args.get("apply") is True else entry["dry"]

        return handle

    return {op: make(entry) for op, entry in step["bridge"].items()}


async def test_tool_list_and_instructions_match_the_typescript_server():
    async with connect() as t:
        listed = await t.client.list_tools()
        actual = [tool.model_dump(mode="json", by_alias=True, exclude_none=True) for tool in listed.tools]
        assert t.client.instructions == EXPECTED_TOOLS["instructions"] == INSTRUCTIONS

    assert [tool["name"] for tool in actual] == [tool["name"] for tool in EXPECTED_TOOLS["tools"]]
    for got, want in zip(actual, EXPECTED_TOOLS["tools"], strict=True):
        got = {key: got[key] for key in want}
        assert got == want, want["name"]


async def test_every_scenario_step_produces_the_recorded_transcript(tmp_path: Path):
    for name, content in SCENARIO["files"].items():
        (tmp_path / name).write_bytes(base64.b64decode(content))

    tokens = TokenTable()
    issued: dict[str, str | None] = {}
    assert len(SCENARIO["steps"]) == len(EXPECTED_TRANSCRIPTS)

    for step, want in zip(SCENARIO["steps"], EXPECTED_TRANSCRIPTS, strict=True):
        bridge = FakeBridge(_handlers(step))
        args = _substitute(step["args"], tmp_path)
        if "applyFrom" in step:
            args.update(dryRun=False, applyToken=issued[step["applyFrom"]])

        async with connect(bridge, tokens) as t:
            result = await t.call(step["tool"], args)

        got: dict[str, Any] = {"name": step["name"], "tool": step["tool"], "isError": bool(result.is_error)}
        if not result.is_error:
            got["text"] = "\n".join(block.text for block in result.content)  # type: ignore[union-attr]
            got["structuredContent"] = result.structured_content
            issued[step["name"]] = (result.structured_content or {}).get("applyToken")
        got["bridgeCalls"] = [{"op": call["op"], "args": _redact(call["args"])} for call in bridge.calls]
        got = _normalize(got, tmp_path)

        if want["isError"]:
            want = {key: want[key] for key in ("name", "tool", "isError", "bridgeCalls")}
        assert got == want, step["name"]
