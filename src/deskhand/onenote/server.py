from __future__ import annotations

from pathlib import Path

from mcp.server import Server

from deskhand import __version__
from deskhand.core.apply_token import TokenTable
from deskhand.core.powershell import BridgeFn, PowerShellBridge
from deskhand.core.server import build_server
from deskhand.onenote.tools import OneNoteDeps, build_tools

BRIDGE_SCRIPT = Path(__file__).with_name("bridge.ps1")

INSTRUCTIONS = "\n".join(
    [
        "These tools drive the Windows OneNote desktop client through its COM automation interface, from WSL.",
        "They work with the user's own signed-in notebooks, so writes land in real notes that sync to SharePoint/OneDrive.",
        "",
        "Identifying a page: prefer pageId, taken from pages or search. When only names are known, pass notebook + section + page together;",
        "a section name that exists in several notebooks is rejected as ambiguous.",
        "",
        "Writing is always two steps. Call a write tool without dryRun: nothing is written, the reply starts with 'applyToken: <token>' and shows the exact",
        "change OneNote would receive. To apply, repeat the same call with dryRun=false and that applyToken. A token is single use, expires after 10 minutes,",
        "belongs to one server process, and is refused when the arguments differ or the target changed in OneNote after the dry run.",
        "If an apply is interrupted its outcome is unknown: read the page again and preview before retrying.",
        "",
        "Several changes: batch takes a list of ops, each one of append_lines, insert_table, insert_image or edit with that tool's",
        "arguments; it previews all of them, returns one applyToken and applies them in order. Prefer it when several independent changes to",
        "one page are planned, so the caller confirms once. Two ops that rewrite the same outline are refused: put several changes inside one",
        "outline in one edit call. create_page and delete_page are not batched. If a change fails, the error says what landed and what was",
        "not attempted.",
        "",
        "Editing text: OneNote replaces an entire outline with whatever one:OEChildren the caller sends, so a paragraph omitted from that list is DELETED.",
        "That is why edit exists and why it resends each affected outline in full, echoing back OneNote's own XML for everything it is not changing.",
        "Use edit for changes inside an outline and append_lines for a new outline at the end of a page. Raw outline XML is never accepted.",
        "",
        "Formatting: a line is a plain string, or an object with text, depth, list, quickStyle, alignment, spaceBefore,",
        "spaceAfter and style (fontFamily, fontSize, colour, bold, italic, underline, strikethrough, highlight). Depth",
        "nests a paragraph under the nearest earlier one at the previous level, which is how a bullet tree is written.",
        "A quickStyle name is resolved against the page's own QuickStyleDef list, so pass the index read_page reported",
        "when the page does not define that name. insert_image, insert_table and edit append take after: the objectID",
        "of a paragraph to insert after, at any depth.",
        "",
        "Concurrency: lastModifiedTime does not advance when content is written through COM and the hierarchy API can report a different value than the page",
        "itself, so it is informational only. Safety comes from the token's content fingerprint, and delete_page additionally requires the exact page name.",
    ]
)


def create_onenote_server(bridge: BridgeFn | None = None, tokens: TokenTable | None = None) -> Server:
    """Cheap and free of side effects: nothing touches Windows until a tool runs."""
    on_close = None
    if bridge is None:
        owned = PowerShellBridge(BRIDGE_SCRIPT, "onenote")
        bridge, on_close = owned, owned.close
    deps = OneNoteDeps(bridge=bridge, tokens=tokens or TokenTable())
    return build_server("onenote", __version__, INSTRUCTIONS, build_tools(deps), on_close=on_close)
