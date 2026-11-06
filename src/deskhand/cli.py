"""Entry point: `deskhand <app>` serves one app over MCP on stdio."""

from __future__ import annotations

import contextlib
import logging
import os
import sys
from collections.abc import Callable
from typing import TYPE_CHECKING

from deskhand import __version__

if TYPE_CHECKING:
    from mcp.server import Server


def _onenote() -> Server:
    from deskhand.onenote.server import create_onenote_server

    return create_onenote_server()


APPS: dict[str, Callable[[], Server]] = {"onenote": _onenote}


def usage() -> str:
    return "\n".join(
        [
            "Usage: deskhand <app>",
            "",
            "Apps:",
            *(f"  {name}" for name in APPS),
            "",
            "Each app speaks MCP over stdio.",
        ]
    )


async def serve(app: str) -> None:
    from mcp.server.stdio import stdio_server

    server = APPS[app]()
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    app, rest = (args[0] if args else None), args[1:]

    if app in ("--version", "-v"):
        print(__version__)
        return 0
    if not app or app in ("--help", "-h"):
        print(usage())
        return 0
    if app not in APPS or rest:
        print(f"Unexpected argument: {rest[0]}" if app in APPS else f"Unknown app: {app}", file=sys.stderr)
        print(usage(), file=sys.stderr)
        return 2

    # stdout carries JSON-RPC only; everything else goes to stderr.
    logging.basicConfig(
        stream=sys.stderr,
        level=os.environ.get("DESKHAND_LOG", "WARNING").upper(),
        format=f"[deskhand {app}] %(name)s %(message)s",
    )
    import anyio

    with contextlib.suppress(KeyboardInterrupt):
        anyio.run(serve, app)
    return 0
