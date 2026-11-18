"""Writes tests/parity/fixtures/<app>/tools.json, the tool list a client sees.

    uv run python tests/parity/dump_tools.py onenote
    uv run python tests/parity/dump_tools.py ppt

Run it only when the advertised contract is meant to change; test_parity compares against the file.
"""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "unit"))

from helpers import connect, connect_ppt  # noqa: E402  # pyright: ignore[reportMissingImports]  # ty: ignore[import-not-found]

from deskhand.onenote.server import INSTRUCTIONS  # noqa: E402
from deskhand.ppt.server import INSTRUCTIONS as PPT_INSTRUCTIONS  # noqa: E402


async def main(app: str) -> None:
    if app == "onenote":
        connector, instructions = connect, INSTRUCTIONS
    elif app == "ppt":
        connector, instructions = connect_ppt, PPT_INSTRUCTIONS
    else:
        raise SystemExit(f"unknown app: {app}")
    async with connector() as t:
        listed = await t.client.list_tools()
        tools = [tool.model_dump(mode="json", by_alias=True, exclude_none=True) for tool in listed.tools]
    document = {"instructions": instructions, "tools": tools}
    target = HERE / "fixtures" / app / "tools.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(document, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    asyncio.run(main(sys.argv[1] if len(sys.argv) > 1 else "onenote"))
