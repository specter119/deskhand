from __future__ import annotations

from pathlib import Path

from mcp.server import Server

from deskhand import __version__
from deskhand.core.apply_token import TokenTable
from deskhand.core.powershell import BridgeFn, PowerShellBridge
from deskhand.core.server import build_server
from deskhand.ppt.tools import PptDeps, build_tools

BRIDGE_SCRIPT = Path(__file__).with_name("bridge.ps1")

INSTRUCTIONS = "\n".join(
    [
        "These tools drive the Windows PowerPoint desktop client through its COM automation interface, from WSL.",
        "They read a deck's real geometry and text, add and change shapes on a 1280x720 px canvas by default,",
        "render slides with PowerPoint itself, and save the result as .pptx.",
        "Every length in the API is px at 96 dpi; PowerPoint's points are converted in the bridge.",
        "Writes are two-step: the first call previews and returns an applyToken and a fingerprint; the second",
        "call repeats the arguments with dryRun=false, that applyToken, and that expectFingerprint. A changed",
        "deck is refused rather than overwritten. Nothing is saved until save_deck is called.",
        "The server only closes presentations it opened or created; it never quits a PowerPoint the user had",
        "already running.",
    ]
)


def create_ppt_server(*, bridge: BridgeFn | None = None, tokens: TokenTable | None = None) -> Server:
    on_close = None
    if bridge is None:
        owned = PowerShellBridge(
            BRIDGE_SCRIPT,
            "ppt",
            app_name="PowerPoint",
            target_name="deck",
            close_grace=15.0,
        )
        bridge, on_close = owned, owned.close
    deps = PptDeps(bridge=bridge, tokens=tokens or TokenTable())
    return build_server("ppt", __version__, INSTRUCTIONS, build_tools(deps), on_close=on_close)
