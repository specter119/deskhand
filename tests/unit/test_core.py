import os
from typing import Any

import pytest
from helpers import text_of
from mcp import Client
from mcp.types import ToolAnnotations

from deskhand.core.apply_token import TokenError, TokenTable, hash_request, stable_stringify
from deskhand.core.powershell import BridgeError, new_call_stem
from deskhand.core.server import ToolOutput, ToolSpec, build_server

pytestmark = pytest.mark.anyio


def test_call_stems_never_collide_across_calls_and_carry_the_pid():
    stems = {new_call_stem() for _ in range(2000)}
    assert len(stems) == 2000
    assert all(stem.startswith(f"{os.getpid()}-") for stem in stems)


def test_stable_stringify_ignores_key_order_and_missing_members():
    assert stable_stringify({"b": 1, "a": [{"d": 1, "c": None, "e": 2}]}) == stable_stringify(
        {"a": [{"e": 2, "d": 1}], "b": 1}
    )
    assert hash_request("edit", {"a": 1}) != hash_request("append_lines", {"a": 1})


def test_a_token_is_single_use_bound_to_its_request_and_expires():
    now = 1_000.0
    table = TokenTable(10_000, lambda: now)
    request = {"pageId": "p", "lines": ["x"]}
    token = table.issue("append_lines", request, "fp")

    def reason(token: str | None, tool: str, request: dict[str, Any]) -> str:
        with pytest.raises(TokenError) as raised:
            table.consume(token, tool, request)
        return raised.value.reason

    assert reason(None, "append_lines", request) == "missing"
    assert reason("dh1_forged", "append_lines", request) == "unknown"
    assert reason(token, "append_lines", {**request, "lines": ["y"]}) == "mismatch"
    assert reason(token, "edit", request) == "mismatch"

    assert table.consume(token, "append_lines", request) == "fp"
    assert reason(token, "append_lines", request) == "used"

    late = table.issue("append_lines", request, "fp")
    now += 10_001
    assert reason(late, "append_lines", request) == "expired"


def test_the_table_stays_bounded():
    table = TokenTable()
    first = table.issue("t", {"n": 0}, "fp")
    for n in range(1, 400):
        table.issue("t", {"n": n}, "fp")
    with pytest.raises(TokenError) as raised:
        table.consume(first, "t", {"n": 0})
    assert raised.value.reason == "unknown"


@pytest.mark.parametrize("status", ["applied", "failed"])
async def test_output_validation_after_a_write_says_the_call_had_returned(status: str):
    async def run(_: dict[str, object]) -> ToolOutput:
        return ToolOutput(text="done", structured={"status": status})

    spec = ToolSpec(
        name="w",
        title="W",
        description="w",
        input_schema={"type": "object"},
        output_schema={"type": "object", "required": ["id"], "properties": {"id": {"type": "string"}}},
        annotations=ToolAnnotations(read_only_hint=False),
        run=run,
    )
    async with Client(build_server("t", "1", "i", [spec]), mode="legacy") as client:
        result = await client.call_tool("w", {})
    assert result.is_error is True
    text = text_of(result)
    assert "Output validation error" in text
    assert ("read the page before retrying" in text) == (status == "applied")


async def test_a_bridge_error_can_carry_structured_details():
    async def run(_: dict[str, object]) -> ToolOutput:
        raise BridgeError("PowerPoint partial: shape 3 failed", "applied_unverified", {"appliedCount": 2})

    spec = ToolSpec(
        name="w",
        title="W",
        description="w",
        input_schema={"type": "object"},
        output_schema={"type": "object"},
        annotations=ToolAnnotations(read_only_hint=False),
        run=run,
    )
    async with Client(build_server("t", "1", "i", [spec]), mode="legacy") as client:
        result = await client.call_tool("w", {})
    assert result.is_error is True
    assert result.structured_content == {
        "error": {"code": "applied_unverified", "details": {"appliedCount": 2}}
    }
