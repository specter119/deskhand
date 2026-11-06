import os

import pytest

from deskhand.core.apply_token import TokenError, TokenTable, hash_request, stable_stringify
from deskhand.core.powershell import new_call_stem


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

    def reason(*args: object) -> str:
        with pytest.raises(TokenError) as raised:
            table.consume(*args)  # type: ignore[arg-type]
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
