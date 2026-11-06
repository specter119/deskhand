from __future__ import annotations

import copy
from collections.abc import AsyncIterator, Callable
from contextlib import asynccontextmanager
from dataclasses import dataclass
from typing import Any

from mcp import Client
from mcp.types import CallToolResult

from deskhand.core.apply_token import TokenTable
from deskhand.onenote.server import create_onenote_server

PAGE: dict[str, Any] = {
    "id": "page-1",
    "name": "Probe",
    "level": "0",
    "lang": "en-US",
    "created": "2026-01-01T00:00:00.000Z",
    "lastModified": "2026-01-01T00:00:00.000Z",
    "lines": [
        {
            "index": 0,
            "objectID": "oe-1",
            "text": "hello",
            "inTitle": False,
            "inTable": False,
            "hasImage": False,
            "hasTable": False,
            "hasInk": False,
            "list": "",
            "depth": 0,
            "outlineID": "ol-1",
            "quickStyle": "",
            "alignment": "",
        }
    ],
    "outlines": [
        {
            "objectID": "ol-1",
            "x": None,
            "y": None,
            "z": None,
            "width": None,
            "height": None,
            "inTitle": False,
            "lineCount": 1,
        }
    ],
    "images": [],
    "tables": [],
}

Handler = Callable[[dict[str, Any]], Any]


class FakeBridge:
    """A bridge that records every call and answers from `handlers[op]`."""

    def __init__(self, handlers: dict[str, Handler] | None = None) -> None:
        self.handlers = handlers or {}
        self.calls: list[dict[str, Any]] = []

    async def __call__(self, op: str, args: dict[str, Any]) -> dict[str, Any]:
        self.calls.append({"op": op, "args": copy.deepcopy(args)})
        handler = self.handlers.get(op)
        if handler is None:
            raise AssertionError(f"no fake handler for {op}")
        return handler(args)

    def apply_calls(self) -> list[dict[str, Any]]:
        return [call for call in self.calls if call["args"].get("apply") is True]


def preview_handler(fingerprint: str = "fp-1") -> Handler:
    def handle(args: dict[str, Any]) -> dict[str, Any]:
        return {"applied": args.get("apply") is True, "fingerprint": fingerprint, "page": PAGE, "changes": []}

    return handle


@dataclass
class Connection:
    client: Client
    bridge: FakeBridge
    tokens: TokenTable

    async def call(self, name: str, args: dict[str, Any] | None = None) -> CallToolResult:
        return await self.client.call_tool(name, args or {})


@asynccontextmanager
async def connect(
    bridge: FakeBridge | None = None, tokens: TokenTable | None = None
) -> AsyncIterator[Connection]:
    bridge = bridge or FakeBridge()
    tokens = tokens or TokenTable()
    server = create_onenote_server(bridge=bridge, tokens=tokens)
    async with Client(server, mode="legacy") as client:
        yield Connection(client, bridge, tokens)


def text_of(result: CallToolResult) -> str:
    return "\n".join(getattr(block, "text", "") for block in result.content)
