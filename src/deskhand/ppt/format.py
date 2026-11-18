"""Text shown to people and models for the PowerPoint tools."""

from __future__ import annotations

import unicodedata
from typing import Any

from deskhand.core.compat import as_array, as_string, num, truncate, utf16_len

Deck = dict[str, Any]

TEXT_PREVIEW_CHARS = 90
MAX_SHAPES = 5
MAX_EXPANDED_SLIDES = 5
MAX_OVERVIEW_SLIDES = 20
MAX_RENDER_FILES = 5


def short_text(text: str, limit: int = TEXT_PREVIEW_CHARS) -> str:
    if utf16_len(text) <= limit:
        return text
    out: list[str] = []
    used = 0
    index = 0
    while index < len(text):
        end = index + 1
        while end < len(text):
            char = text[end]
            if char == "\u200d" and end + 1 < len(text):
                end += 2
                continue
            if unicodedata.combining(char) or char in ("\ufe0f", "\u20e3") or 0x1F1E6 <= ord(char) <= 0x1F1FF:
                end += 1
                continue
            break
        cluster = text[index:end]
        width = utf16_len(cluster)
        if used + width > limit:
            break
        out.append(cluster)
        used += width
        index = end
    return "".join(out) + "…"


def _one_line(text: str) -> str:
    return text.replace("\r", " ").replace("\n", " / ")


def normalize_deck(raw: Any) -> Deck:
    deck: Deck = dict(raw) if isinstance(raw, dict) else {}
    deck["slides"] = as_array(deck.get("slides"))
    for slide in deck["slides"]:
        if isinstance(slide, dict):
            slide["shapes"] = as_array(slide.get("shapes"))
    return deck


def format_presentation(presentation: Any) -> str:
    item = presentation if isinstance(presentation, dict) else {}
    return (
        f"{as_string(item.get('name'))}  {as_string(item.get('path')) or '(unsaved)'}  "
        f"{num(item.get('widthPx'))}x{num(item.get('heightPx'))} px  "
        f"{num(item.get('slides'))} slide(s)  {'saved' if item.get('saved') else 'unsaved'}"
    )


def format_presentations(data: dict[str, Any]) -> str:
    if not data.get("running"):
        return "PowerPoint is not running; no presentation is open."
    items = as_array(data.get("presentations"))
    if not items:
        return "PowerPoint is running; no presentation is open."
    return "\n".join(f"- {format_presentation(item)}" for item in items)


def format_shape(shape: Any) -> str:
    item = shape if isinstance(shape, dict) else {}
    bits = [
        f"shape {num(item.get('id'))} {as_string(item.get('type'))} "
        f"at ({num(item.get('x'))},{num(item.get('y'))}) {num(item.get('w'))}x{num(item.get('h'))}"
    ]
    text = as_string(item.get("text"))
    if text:
        bits.append(f'"{short_text(_one_line(text))}"')
    table = item.get("table")
    if isinstance(table, dict):
        bits.append(f"table {num(table.get('rows'))}x{num(table.get('cols'))}")
    chart = item.get("chart")
    if isinstance(chart, dict):
        names = ", ".join(as_string(name) for name in as_array(chart.get("seriesNames")))
        bits.append(f"chart {as_string(chart.get('chartType'))}{f' ({names})' if names else ''}")
    picture = item.get("picture")
    if isinstance(picture, dict):
        path = as_string(picture.get("path"))
        bits.append(f"picture{f' {path}' if path else ' (embedded)'}")
    return "  ".join(bits)


def format_slide(slide: Any, *, include_notes: bool = False, focus_shape_id: int | None = None) -> str:
    item = slide if isinstance(slide, dict) else {}
    title = as_string(item.get("title"))
    lines = [
        f"slide {num(item.get('index'))}: layout={as_string(item.get('layout')) or '(unknown)'}"
        + (f' title="{short_text(_one_line(title))}"' if title else "")
    ]
    shapes = as_array(item.get("shapes"))
    shown = [
        shape
        for index, shape in enumerate(shapes)
        if index < MAX_SHAPES
        or (focus_shape_id is not None and isinstance(shape, dict) and shape.get("id") == focus_shape_id)
    ]
    for shape in shown:
        lines.append(f"  {format_shape(shape)}")
    if len(shapes) > len(shown):
        lines.append(f"  … and {len(shapes) - len(shown)} more shape(s)")
    if not shapes:
        lines.append("  (no shapes)")
    if include_notes:
        notes = as_string(item.get("notes"))
        if notes:
            lines.append(f"  notes: {short_text(_one_line(notes))}")
    return "\n".join(lines)


def format_deck(data: dict[str, Any], *, include_notes: bool = False, expand_shapes: bool = False) -> str:
    deck = normalize_deck(data)
    presentation = deck.get("presentation")
    lines = [f"presentation: {format_presentation(presentation)}"]
    slides = deck["slides"]
    if not slides:
        lines.append("(no slides)")
    if expand_shapes:
        for slide in slides[:MAX_EXPANDED_SLIDES]:
            lines.append(format_slide(slide, include_notes=include_notes))
        if len(slides) > MAX_EXPANDED_SLIDES:
            lines.append(
                f"… and {len(slides) - MAX_EXPANDED_SLIDES} more slide(s); use read_slide for the rest"
            )
    else:
        for slide in slides[:MAX_OVERVIEW_SLIDES]:
            item = slide if isinstance(slide, dict) else {}
            title = as_string(item.get("title"))
            lines.append(
                f"slide {num(item.get('index'))}: layout={as_string(item.get('layout')) or '(unknown)'}"
                + (f' title="{short_text(_one_line(title))}"' if title else "")
                + f" shapes={len(as_array(item.get('shapes')))}"
            )
        if len(slides) > MAX_OVERVIEW_SLIDES:
            lines.append(f"… and {len(slides) - MAX_OVERVIEW_SLIDES} more slide(s)")
        if slides:
            lines.append('Pass slides="1-3" to expand shapes for those slides.')
    return truncate("\n".join(lines))


def format_slide_result(data: dict[str, Any], *, include_notes: bool = False) -> str:
    return truncate(format_slide(data.get("slide"), include_notes=include_notes))


def format_render(data: dict[str, Any]) -> str:
    files = as_array(data.get("files"))
    lines = [f"rendered {len(files)} file(s) in {num(data.get('elapsedMs'))} ms"]
    for item in files[:MAX_RENDER_FILES]:
        lines.append(
            f"  {as_string(item.get('path'))}  {num(item.get('widthPx'))}x{num(item.get('heightPx'))} px"
        )
    if len(files) > MAX_RENDER_FILES:
        lines.append(
            f"  … and {len(files) - MAX_RENDER_FILES} more file(s) in {as_string(data.get('outDir'))}"
        )
    return "\n".join(lines)


def format_write(tool: str, data: dict[str, Any], *, applied: bool, focus_shape_id: int | None = None) -> str:
    """The text after a write preview or apply, from the state the bridge returned."""
    lines: list[str] = []
    if applied:
        lines.append(f"Applied {tool}.")
    else:
        lines.append(f"Dry run of {tool}, nothing was written.")
    slide = data.get("slide")
    if isinstance(slide, dict):
        lines.append(format_slide(slide, focus_shape_id=focus_shape_id))
    presentation = data.get("presentation")
    if isinstance(presentation, dict):
        lines.append(f"presentation: {format_presentation(presentation)}")
    path = as_string(data.get("path"))
    if path:
        lines.append(f"path: {path}")
    closed = as_string(data.get("closed"))
    if closed:
        lines.append(f"closed: {closed}")
    return truncate("\n".join(lines))
