"""Small helpers that keep text and request bodies identical to what a JavaScript runtime produces.

OneNote replies come from PowerShell as JSON, and the text shown to people and models was fixed by the first
implementation. These differences would otherwise change it silently: truthiness of empty containers,
float formatting (`36.0` versus `36`), round-half-up versus banker's rounding, and `null` members.
"""

from __future__ import annotations

import math
from typing import Any

MAX_TEXT_CHARS = 40_000


def js_truthy(value: Any) -> bool:
    """JavaScript truthiness: empty lists and dicts are true, 0, "" and null are false."""
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, int | float):
        return value != 0 and not (isinstance(value, float) and math.isnan(value))
    if isinstance(value, str):
        return value != ""
    return True


def num(value: Any) -> str:
    """Interpolate a JSON value the way a JavaScript template literal does."""
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return "null"
    if isinstance(value, float):
        if value.is_integer() and abs(value) < 1e21:
            return str(int(value))
        return repr(value)
    return str(value)


def js_round(value: float) -> int:
    """Math.round: halves go up, unlike Python's round()."""
    return math.floor(value + 0.5)


def js_number(value: Any) -> float:
    """Number(value) for the values OneNote sends: null and missing become 0."""
    if value is None:
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return math.nan


def utf16_len(text: str) -> int:
    return len(text.encode("utf-16-le", "surrogatepass")) // 2


def utf16_slice(text: str, limit: int) -> str:
    """text.slice(0, limit) where limit counts UTF-16 code units, as a JavaScript string does."""
    units = text.encode("utf-16-le", "surrogatepass")
    if len(units) // 2 <= limit:
        return text
    return units[: limit * 2].decode("utf-16-le", "replace")


def truncate(text: str) -> str:
    if utf16_len(text) <= MAX_TEXT_CHARS:
        return text
    return f"{utf16_slice(text, MAX_TEXT_CHARS)}\n... truncated at {MAX_TEXT_CHARS} characters"


def as_array(value: Any) -> list[Any]:
    """PowerShell 5.1 collapses one-element arrays to the bare element; make every list a list again."""
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def as_string(value: Any) -> str:
    return value if isinstance(value, str) else ""


def compact(value: Any) -> Any:
    """Drop null members and turn whole floats into ints.

    JSON.stringify omits `undefined` members and writes `80.0` as `80`. Every request that is hashed,
    sent to the bridge or recorded goes through here so that all three see the same thing.
    """
    if isinstance(value, dict):
        return {key: compact(member) for key, member in value.items() if member is not None}
    if isinstance(value, list):
        return [compact(item) for item in value]
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e21:
        return int(value)
    return value
