"""Writes tests/parity/fixtures/tools.json, the tool list a client sees.

    uv run python tests/parity/dump_tools.py

Run it only when the advertised contract is meant to change; test_parity compares against the file.
"""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "unit"))

from helpers import connect  # noqa: E402  # pyright: ignore[reportMissingImports]

from deskhand.onenote.server import INSTRUCTIONS  # noqa: E402


async def main() -> None:
    async with connect() as t:
        listed = await t.client.list_tools()
        tools = [tool.model_dump(mode="json", by_alias=True, exclude_none=True) for tool in listed.tools]
    document = {"instructions": INSTRUCTIONS, "tools": tools}
    (HERE / "fixtures" / "tools.json").write_text(
        json.dumps(document, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


asyncio.run(main())
