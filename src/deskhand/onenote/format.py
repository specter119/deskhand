"""Text shown to people and models. The wording and layout are part of the tool contract."""

from __future__ import annotations

import json
from typing import Any

from deskhand.core.compat import (
    as_array,
    as_string,
    js_number,
    js_round,
    js_truthy,
    num,
    truncate,
    utf16_len,
    utf16_slice,
)

Page = dict[str, Any]

# A preview is read, not parsed: a replaced section can delete 40 paragraphs and add 30, and printing
# every objectID, every full text and every resulting paragraph buries the one thing to check.
MAX_REMOVED_LINES = 5
MAX_RESULTING_LINES = 5
TEXT_PREVIEW_CHARS = 100


def _short(text: str, limit: int = TEXT_PREVIEW_CHARS) -> str:
    if utf16_len(text) <= limit:
        return text
    return utf16_slice(text, limit) + "…"


def normalize_page(raw: Any) -> Page:
    """PowerShell 5.1 collapses one-element arrays in places; make every list a list again."""
    page: Page = dict(raw) if isinstance(raw, dict) else {}
    for key in ("lines", "outlines", "images", "tables"):
        page[key] = as_array(page.get(key))
    return page


def _location(outline_id: Any) -> str:
    return as_string(outline_id) or "(page level)"


def format_page(page: Page) -> str:
    lines: list[str] = [
        f"page: {as_string(page.get('name'))}",
        f"id: {as_string(page.get('id'))}",
        f"lastModified: {as_string(page.get('lastModified'))} (informational)",
    ]

    for table in as_array(page.get("tables")):
        lines.append(
            f"table {as_string(table.get('objectID'))}: {num(table.get('rows'))}x{num(table.get('columns'))}"
            f" cells, in outline {_location(table.get('outlineID'))}"
        )

    for outline in as_array(page.get("outlines")):
        x = outline.get("x")
        geometry = f" at ({num(x)},{num(outline.get('y'))})" if x is not None else ""
        title = " (title)" if js_truthy(outline.get("inTitle")) else ""
        lines.append(
            f"outline {as_string(outline.get('objectID'))}{title}{geometry}, "
            f"{num(outline.get('lineCount'))} paragraph(s)"
        )

    for image in as_array(page.get("images")):
        size = ""
        if js_truthy(image.get("width")):
            size = f" {js_round(js_number(image.get('width')))}x{js_round(js_number(image.get('height')))}"
        alt = f' alt="{as_string(image.get("alt"))}"' if as_string(image.get("alt")) else ""
        bits = "" if js_truthy(image.get("embedded")) else " - NO BITS"
        lines.append(f"image{size}{alt} in outline {_location(image.get('outlineID'))}{bits}")

    lines.append("content:")
    page_lines = as_array(page.get("lines"))
    for line in page_lines:
        marks: list[str] = []
        if js_truthy(line.get("inTitle")):
            marks.append("title")
        if js_truthy(line.get("inTable")):
            marks.append("table-cell")
        if js_truthy(line.get("hasTable")):
            marks.append("table")
        if js_truthy(line.get("hasImage")):
            marks.append("image")
        if js_truthy(line.get("hasInk")):
            marks.append("ink")
        if js_truthy(line.get("list")):
            marks.append(num(line.get("list")))
        depth = line.get("depth")
        if js_truthy(depth):
            marks.append(f"depth={num(depth)}")

        indent = "  " * int(min(js_number(depth), 6))
        text = as_string(line.get("text"))
        body = text.replace("\n", " / ")
        object_id = line.get("objectID")
        suffix = " ".join(
            part
            for part in (f"[{','.join(marks)}]" if marks else "", f"oe={object_id}" if object_id else "")
            if part
        )
        lines.append(
            f"  [{num(line.get('index'))}] {indent}{body or '(no text)'}{f'   {suffix}' if suffix else ''}"
        )
    if not page_lines:
        lines.append("  (empty page)")

    return truncate("\n".join(lines))


def format_page_list(pages: Any) -> str:
    items = as_array(pages)
    if not items:
        return "no pages"
    # No lastModified here on purpose: OneNote's hierarchy API reports a stale one.
    return "\n".join(
        f"- {as_string(page.get('name'))}  {as_string(page.get('id'))}  created={as_string(page.get('created'))}"
        for page in items
    )


def format_edit_changes(data: dict[str, Any]) -> str:
    changes = as_array(data.get("changes"))
    if not changes:
        return "no change"
    out: list[str] = []
    for change in changes:
        out.append(f"outline {as_string(change.get('outlineId'))}")
        for item in as_array(change.get("changed")):
            before = _short(as_string(item.get("before")))
            after_text = _short(as_string(item.get("after")))
            out.append(f'  ~ "{before}"  ->  "{after_text}"')
        removed = as_array(change.get("removed"))
        for item in removed[:MAX_REMOVED_LINES]:
            out.append(f'  - delete "{_short(as_string(item.get("text")))}"')
        if len(removed) > MAX_REMOVED_LINES:
            out.append(f"  … and {len(removed) - MAX_REMOVED_LINES} more deleted")
        after = as_array(change.get("after"))
        if len(after) > MAX_RESULTING_LINES:
            out.append(f"  resulting paragraphs: {len(after)} paragraph(s)")
        else:
            listed = ", ".join(json.dumps(entry, ensure_ascii=False) for entry in after)
            out.append(f"  resulting paragraphs: {listed}")
    return "\n".join(out)


def describe_lines(lines: list[Any]) -> str:
    """One numbered row per paragraph; a plain string renders exactly as it always has."""
    rows = []
    for index, line in enumerate(lines):
        if not isinstance(line, dict):
            rows.append(f"  {index + 1}. {line}")
            continue
        marks = []
        depth = js_number(line.get("depth"))
        if depth:
            marks.append(f"depth={int(depth)}")
        if line.get("list"):
            marks.append(as_string(line.get("list")))
        quick_style = line.get("quickStyle")
        if quick_style is not None:
            label = num(quick_style) if isinstance(quick_style, (int, float)) else as_string(quick_style)
            marks.append(f"style={label}")
        style = line.get("style")
        if isinstance(style, dict):
            bits = []
            if style.get("fontSize"):
                bits.append(f"{num(style['fontSize'])}pt")
            if style.get("fontFamily"):
                bits.append(as_string(style["fontFamily"]))
            if style.get("color"):
                bits.append(as_string(style["color"]))
            bits.extend(
                flag for flag in ("bold", "italic", "underline", "strikethrough") if style.get(flag) is True
            )
            if style.get("highlight"):
                bits.append(f"highlight {as_string(style['highlight'])}")
            if bits:
                marks.append(" ".join(bits))
        suffix = f"   [{', '.join(marks)}]" if marks else ""
        indent = "  " * int(depth)
        rows.append(f"  {index + 1}. {indent}{as_string(line.get('text'))}{suffix}")
    return "\n".join(rows)
