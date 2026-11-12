"""The ten OneNote tools."""

from __future__ import annotations

import asyncio
import base64
import json
import os
import re
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from mcp.types import ToolAnnotations

from deskhand.core import schema as s
from deskhand.core.apply_token import TokenTable
from deskhand.core.compat import (
    as_array,
    as_string,
    compact,
    js_round,
    js_truthy,
    num,
    truncate,
    utf16_slice,
)
from deskhand.core.powershell import BridgeFn
from deskhand.core.server import ToolOutput, ToolSpec
from deskhand.onenote.format import (
    describe_lines,
    format_edit_changes,
    format_page,
    format_page_list,
    normalize_page,
)
from deskhand.onenote.schemas import (
    LINE,
    NOTEBOOK,
    PAGE,
    PAGE_META,
    TARGET_PARAMS,
    WRITE_PARAMS,
    WRITE_RESULT,
)

# Measured through the MCP server against a local notebook (incompressible PNG, dry run / apply):
# 8 MB 0.5 s / 0.8 s, 16 MB 0.8 s / 1.5 s, 30 MB 1.5 s / 2.9 s. The cap leaves the apply far under the 40 s call timeout.
MAX_IMAGE_BYTES = 32 * 1024 * 1024

READ_ONLY = ToolAnnotations(read_only_hint=True, destructive_hint=False, open_world_hint=False)
ADDITIVE = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False
)
# Resends whole outlines, so a wrong request would delete paragraphs: treat as destructive.
REWRITING = ToolAnnotations(
    read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False
)

WRITE_FLOW = (
    "Two steps: call it without dryRun first (nothing is written; the reply starts with a line "
    "'applyToken: <token>' and shows the exact change), "
    "then call it again with the same arguments plus dryRun=false and that applyToken. "
    "A token is single use, expires after 10 minutes, "
    "and is refused if the arguments differ or the target changed in OneNote after the dry run."
)

Params = dict[str, Any]


@dataclass(frozen=True)
class OneNoteDeps:
    bridge: BridgeFn
    tokens: TokenTable


async def _call(deps: OneNoteDeps, op: str, args: Params) -> Params:
    # Members that are None never reach the bridge, the token hash or a recording.
    return await deps.bridge(op, compact(args))


def _pick_target(params: Params) -> Params:
    return {
        key: params[key]
        for key in ("pageId", "notebook", "section", "page")
        if js_truthy(params.get(key)) and isinstance(params[key], str)
    }


def _or(value: Any, default: str) -> str:
    return default if value is None else num(value)


def _target_label(target: Params) -> str:
    return as_string(target.get("pageId")) or (
        f"{as_string(target.get('notebook')) or '?'} / {as_string(target.get('section')) or '?'}"
        f" / {as_string(target.get('page')) or '?'}"
    )


def _removal_lines(remove: list[str]) -> list[str]:
    """A long delete list is a wall of objectIDs; a reader checks the texts, which the preview shows."""
    if len(remove) <= 3:
        return [f"  delete {object_id}" for object_id in remove]
    return [f"  delete {len(remove)} paragraph(s)"]


def _placement(outline: Any, after: Any) -> str:
    """Where a block lands: a new outline, inside an existing one, or after one paragraph of it."""
    if js_truthy(after):
        return (
            f"inside outline {outline} after paragraph {after}"
            if js_truthy(outline)
            else f"after paragraph {after}"
        )
    if js_truthy(outline):
        return f"inside outline {outline}"
    return "as a new outline at the end of the page"


def _image_format_for(path: str, requested: str | None) -> str:
    """OneNote's image format pattern is auto|png|emf|jpg, so jpeg is spelled jpg and anything else gets auto."""
    if requested:
        lowered = requested.lower()
        return "jpg" if lowered == "jpeg" else lowered
    return {".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".emf": "emf"}.get(
        Path(path).suffix.lower(), "auto"
    )


@dataclass(frozen=True)
class WritePlan:
    tool: str
    op: str
    request: Params
    """Exactly what is sent to the bridge, minus apply and expectFingerprint. The token is bound to it."""
    intent: str
    describe: Callable[[Params, bool], str]


async def _run_write(deps: OneNoteDeps, plan: WritePlan, params: Params) -> ToolOutput:
    request = compact(plan.request)

    if params.get("dryRun") is not False:
        data = await _call(deps, plan.op, {**request, "apply": False})
        fingerprint = as_string(data.get("fingerprint"))
        if not fingerprint:
            raise RuntimeError("The OneNote bridge returned no fingerprint for this preview")
        token = deps.tokens.issue(plan.tool, request, fingerprint)
        preview = plan.describe(data, False)
        structured: Params = {"status": "dry-run", "applyToken": token, "preview": preview}
        if js_truthy(data.get("page")):
            structured["page"] = normalize_page(data["page"])
        if js_truthy(data.get("changes")):
            structured["changes"] = as_array(data["changes"])
        return ToolOutput(
            # The token goes first: long results are cut in the middle by some clients.
            text="\n".join(
                [
                    f"applyToken: {token}",
                    "Dry run, nothing was written.",
                    plan.intent,
                    "",
                    preview,
                    "",
                    f"To apply, call {plan.tool} again with the same arguments plus dryRun=false and "
                    f'applyToken="{token}". The token is single use and expires in 10 minutes.',
                ]
            ),
            structured=structured,
        )

    fingerprint = deps.tokens.consume(params.get("applyToken"), plan.tool, request)
    data = await _call(deps, plan.op, {**request, "apply": True, "expectFingerprint": fingerprint})
    page = normalize_page(data["page"]) if js_truthy(data.get("page")) else None
    structured = {"status": "applied"}
    if page is not None:
        structured["page"] = page
    if js_truthy(data.get("changes")):
        structured["changes"] = as_array(data["changes"])
    described = plan.describe(data, True)
    return ToolOutput(
        text=truncate(f"{described}\n\n{format_page(page)}" if page is not None else described),
        structured=structured,
    )


def _page_of(data: Params) -> Params:
    page = data.get("page")
    return page if isinstance(page, dict) else {}


def _strings(value: Any) -> list[str]:
    return [item for item in as_array(value) if isinstance(item, str)]


_PARAGRAPH_KEYS = (
    "text",
    "depth",
    "list",
    "quickStyle",
    "alignment",
    "spaceBefore",
    "spaceAfter",
    "style",
)
_COLOR = re.compile(r"^(#[0-9A-Fa-f]{6}|automatic)$")


def _paragraphs(value: Any, where: str = "lines") -> list[Any]:
    """Normalises lines, which are plain strings or paragraph objects, and checks what the schema cannot.

    Depth describes a tree as a flat list: a paragraph at depth d hangs under the nearest earlier one at
    depth d-1, so a level can only be entered from the level above it.
    """
    out: list[Any] = []
    depth = 0
    for index, item in enumerate(as_array(value)):
        if isinstance(item, str):
            # A plain string stays a plain string on the wire, so an unformatted call is unchanged.
            depth = 0
            out.append(item)
            continue
        if not isinstance(item, dict):
            raise ValueError(f"{where}[{index}] must be a string or a paragraph object")
        paragraph: dict[str, Any] = {key: item[key] for key in _PARAGRAPH_KEYS if item.get(key) is not None}
        text = paragraph.get("text")
        if not isinstance(text, str):
            raise ValueError(f"{where}[{index}] needs text")
        level = paragraph.get("depth", 0)
        if isinstance(level, bool) or not isinstance(level, (int, float)) or level != int(level):
            raise ValueError(f"{where}[{index}] depth must be a whole number of 0 or more")
        level = int(level)
        if level < 0:
            raise ValueError(f"{where}[{index}] depth must be a whole number of 0 or more")
        if not out and level != 0:
            raise ValueError(f"{where}[{index}] the first paragraph must be at depth 0")
        if level > depth + 1:
            raise ValueError(
                f"{where}[{index}] depth {level} skips a level: the previous paragraph is at depth {depth}"
            )
        paragraph["depth"] = level
        depth = level
        _check_paragraph(paragraph, where, index)
        out.append(paragraph)
    return out


def _check_paragraph(paragraph: dict[str, Any], where: str, index: int) -> None:
    for key in ("spaceBefore", "spaceAfter"):
        value = paragraph.get(key)
        if value is None:
            continue
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
            raise ValueError(f"{where}[{index}] {key} must be a number of 0 or more")
    style = paragraph.get("style")
    if not isinstance(style, dict):
        return
    size = style.get("fontSize")
    if size is not None and (isinstance(size, bool) or not isinstance(size, (int, float)) or size <= 0):
        raise ValueError(f"{where}[{index}] style.fontSize must be a number above 0")
    for key in ("color", "highlight"):
        value = style.get(key)
        if value is not None and (not isinstance(value, str) or not _COLOR.match(value)):
            raise ValueError(f"{where}[{index}] style.{key} must be #RRGGBB or automatic")


def build_tools(deps: OneNoteDeps) -> list[ToolSpec]:
    async def notebooks(args: Params) -> ToolOutput:
        data = await _call(deps, "notebooks", {})
        listed = [{**nb, "sections": as_array(nb.get("sections"))} for nb in as_array(data.get("notebooks"))]
        blocks = []
        for notebook in listed:
            head = f"{as_string(notebook.get('name'))}  {as_string(notebook.get('id'))}"
            body = "\n".join(
                f"  - {as_string(sec.get('name'))}  {as_string(sec.get('id'))}"
                for sec in notebook["sections"]
            )
            blocks.append(f"{head}\n{body}" if body else head)
        return ToolOutput(truncate("\n\n".join(blocks) or "no notebooks"), {"notebooks": listed})

    async def pages(args: Params) -> ToolOutput:
        data = await _call(deps, "pages", dict(args))
        listed = as_array(data.get("pages"))
        return ToolOutput(
            truncate(f"section: {as_string(data.get('section'))}\n{format_page_list(listed)}"),
            {**data, "pages": listed},
        )

    async def search(args: Params) -> ToolOutput:
        data = await _call(deps, "search", {"query": args["query"], "limit": args.get("limit")})
        hits = as_array(data.get("pages"))
        out = (
            "\n".join(
                f"- {as_string(h.get('notebook'))} / {as_string(h.get('section'))} / {as_string(h.get('page'))}"
                f"  {as_string(h.get('pageId'))}"
                for h in hits
            )
            if hits
            else "no pages matched"
        )
        note = (
            f"\n({num(data.get('total'))} matches, showing {num(data.get('returned'))})"
            if data.get("total") != data.get("returned")
            else ""
        )
        return ToolOutput(truncate(out + note), {**data, "pages": hits})

    async def read_page(args: Params) -> ToolOutput:
        data = await _call(
            deps, "read_page", {**_pick_target(args), "includeXml": args.get("includeXml") is True}
        )
        page = normalize_page(data.get("page"))
        raw = as_string(page.get("rawXml"))
        text = format_page(page)
        if raw:
            text += utf16_slice(f"\n\nraw xml:\n{raw}", 40_000)
        return ToolOutput(text, {"page": page})

    async def create_page(args: Params) -> ToolOutput:
        lines = _paragraphs(args.get("lines"))
        where = (
            f"section {args['sectionId']}"
            if js_truthy(args.get("sectionId"))
            else f"{_or(args.get('notebook'), '?')} / {_or(args.get('section'), '?')}"
        )
        intent = "\n".join(
            [
                f'Create page "{args["title"]}" in {where}',
                f"with {len(lines)} body paragraph(s):\n{describe_lines(lines)}"
                if lines
                else "with no body text",
            ]
        )

        def describe(data: Params, applied: bool) -> str:
            if applied:
                return f"Created page {as_string(_page_of(data).get('id'))}"
            return f'Section "{as_string(data.get("section"))}" exists; the page would be created there.'

        plan = WritePlan(
            "create_page",
            "create_page",
            {
                "sectionId": args.get("sectionId"),
                "notebook": args.get("notebook"),
                "section": args.get("section"),
                "title": args["title"],
                "lines": lines,
            },
            intent,
            describe,
        )
        return await _run_write(deps, plan, args)

    async def append_lines(args: Params) -> ToolOutput:
        lines = _paragraphs(args.get("lines"))
        if not lines:
            raise ValueError("append_lines needs at least one line")
        target = _pick_target(args)
        intent = (
            f"Append {len(lines)} paragraph(s) to {_target_label(target)} as a new outline:\n"
            f"{describe_lines(lines)}"
        )

        def describe(data: Params, applied: bool) -> str:
            page = _page_of(data)
            paragraphs = len(as_array(page.get("lines")))
            outlines = len(as_array(page.get("outlines")))
            if applied:
                return f"Appended. The page now has {paragraphs} paragraph(s) across {outlines} outline(s)."
            return (
                f"Would add one outline with {len(lines)} paragraph(s). "
                f"The page currently has {paragraphs} paragraph(s) across {outlines} outline(s)."
            )

        plan = WritePlan("append_lines", "append", {**target, "lines": lines}, intent, describe)
        return await _run_write(deps, plan, args)

    async def insert_table(args: Params) -> ToolOutput:
        headers = _strings(args.get("headers"))
        rows = [
            [cell if isinstance(cell, str) else _or(cell, "") for cell in as_array(row)]
            for row in as_array(args.get("rows"))
        ]
        if not rows and not headers:
            raise ValueError("insert_table needs headers or rows")
        target = _pick_target(args)
        column_count = max([len(headers), *(len(row) for row in rows)])
        row_count = (1 if headers else 0) + len(rows)
        outline = args.get("outline")
        after = args.get("after")
        placement = _placement(outline, after)
        intent = "\n".join(
            part
            for part in [
                f"Insert a {row_count}x{column_count} table into {_target_label(target)}",
                placement,
                f"header: {' | '.join(headers)}" if headers else "",
                *(f"row: {' | '.join(row)}" for row in rows),
            ]
            if part
        )

        def describe(data: Params, applied: bool) -> str:
            tables = len(as_array(_page_of(data).get("tables")))
            if applied:
                return f"Inserted the table. The page now has {tables} table(s)."
            return (
                f"Would insert a {row_count}x{column_count} table. The page currently has {tables} table(s)."
            )

        plan = WritePlan(
            "insert_table",
            "insert_table",
            {
                **target,
                "headers": headers,
                "rows": rows,
                "bordersVisible": args.get("bordersVisible") is True,
                "hasHeaderRow": args.get("hasHeaderRow") is True,
                "columnWidths": args.get("columnWidths"),
                "outline": outline,
                "after": after,
            },
            intent,
            describe,
        )
        return await _run_write(deps, plan, args)

    async def insert_image(args: Params) -> ToolOutput:
        path: str = args["path"]
        try:
            size = os.stat(path).st_size
        except OSError as error:
            raise ValueError(f"Cannot read {path}: {error.strerror or error}") from error
        if size > MAX_IMAGE_BYTES:
            raise ValueError(
                f"{path} is {size} bytes, above the {MAX_IMAGE_BYTES}-byte limit this tool imposes"
            )
        image_format = _image_format_for(path, args.get("format"))
        target = _pick_target(args)
        outline = args.get("outline")
        after = args.get("after")
        sized = js_truthy(args.get("width")) and js_truthy(args.get("height"))
        intent = "\n".join(
            [
                f"Insert {path} ({js_round(size / 1024)} KB, format {image_format}) into {_target_label(target)}",
                _placement(outline, after),
                f"rendered at {num(args['width'])}x{num(args['height'])}"
                if sized
                else "at OneNote's chosen size",
            ]
        )
        try:
            content = await asyncio.to_thread(Path(path).read_bytes)
        except OSError as error:
            raise ValueError(f"Cannot read {path}: {error.strerror or error}") from error

        def describe(data: Params, applied: bool) -> str:
            images = as_array(_page_of(data).get("images"))
            embedded = sum(1 for image in images if image.get("embedded") is True)
            if applied:
                return f"Inserted the image. OneNote reports {len(images)} image(s), {embedded} with embedded bits."
            return f"Would insert one image. The page currently has {len(images)} image(s), {embedded} with embedded bits."

        plan = WritePlan(
            "insert_image",
            "insert_image",
            {
                **target,
                "base64": base64.b64encode(content).decode("ascii"),
                "format": image_format,
                "width": args.get("width"),
                "height": args.get("height"),
                "alt": args.get("alt"),
                "outline": outline,
                "after": after,
            },
            intent,
            describe,
        )
        return await _run_write(deps, plan, args)

    async def edit(args: Params) -> ToolOutput:
        set_ = [{"objectID": item["objectID"], "text": item["text"]} for item in as_array(args.get("set"))]
        remove = _strings(args.get("remove"))
        append = [
            {
                "objectID": item["objectID"],
                "lines": _paragraphs(item.get("lines"), "append.lines"),
                "after": item.get("after"),
            }
            for item in as_array(args.get("append"))
        ]
        if not set_ and not remove and not append:
            raise ValueError("edit needs at least one of set, remove or append")
        target = _pick_target(args)
        intent = "\n".join(
            [
                f"Edit {_target_label(target)}:",
                *(
                    f"  replace {i['objectID']} with {json.dumps(i['text'], ensure_ascii=False)}"
                    for i in set_
                ),
                *_removal_lines(remove),
                *(
                    f"  add {len(i['lines'])} paragraph(s) to {i['objectID']}"
                    + (f" after {i['after']}" if js_truthy(i.get("after")) else "")
                    for i in append
                ),
            ]
        )

        def describe(data: Params, applied: bool) -> str:
            paragraphs = len(as_array(_page_of(data).get("lines")))
            added = [line for item in append for line in item["lines"]]
            parts = [f"{'Changed' if applied else 'Would change'}: {format_edit_changes(data)}"]
            # A single appended line is already visible in the intent; a rewritten section is not.
            if len(added) > 3:
                heading = "Added paragraphs" if applied else "New paragraphs"
                parts.append(f"{heading}:\n{describe_lines(added)}")
            parts.append(
                f"{'OneNote now reports' if applied else 'The page currently has'} {paragraphs} paragraph(s)."
            )
            return "\n\n".join(parts)

        plan = WritePlan(
            "edit", "edit", {**target, "set": set_, "remove": remove, "append": append}, intent, describe
        )
        return await _run_write(deps, plan, args)

    async def delete_page(args: Params) -> ToolOutput:
        target = _pick_target(args)
        intent = (
            f'Delete OneNote page {_target_label(target)}, which must be named "{args["expectName"]}". '
            "It goes to the recycle bin."
        )

        def describe(data: Params, applied: bool) -> str:
            name, page_id = as_string(data.get("name")), as_string(data.get("pageId"))
            if applied:
                return f'Deleted page "{name}" ({page_id}). It is in the notebook recycle bin.'
            return f'Page "{name}" ({page_id}) would be deleted.'

        plan = WritePlan(
            "delete_page", "delete_page", {**target, "expectName": args["expectName"]}, intent, describe
        )
        return await _run_write(deps, plan, args)

    def spec(
        name: str,
        title: str,
        description: str,
        properties: dict[str, Any],
        output: dict[str, Any],
        annotations: ToolAnnotations,
        run: Callable[[Params], Awaitable[ToolOutput]],
    ) -> ToolSpec:
        return ToolSpec(name, title, description, s.obj(properties), output, annotations, run)

    section_params: dict[str, Any] = {
        "sectionId": s.optional(s.string("Section ID from notebooks")),
        "notebook": s.optional(s.string("Notebook name")),
        "section": s.optional(s.string("Section name, required when sectionId is omitted")),
    }

    return [
        spec(
            "notebooks",
            "OneNote notebooks",
            "List the Microsoft OneNote notebooks and sections available in the user's Windows OneNote desktop client.",
            {},
            s.obj({"notebooks": s.array(NOTEBOOK)}, additional=True),
            READ_ONLY,
            notebooks,
        ),
        spec(
            "pages",
            "OneNote pages",
            "List the pages in one OneNote section. Pass sectionId, or notebook plus section names. "
            "Returns names, ids and creation times; for a page's live modification time use read_page.",
            section_params,
            s.obj(
                {"sectionId": s.string(), "section": s.string(), "pages": s.array(PAGE_META)}, additional=True
            ),
            READ_ONLY,
            pages,
        ),
        spec(
            "search",
            "Search OneNote",
            "Find OneNote pages by name across every notebook. Matching is a case-insensitive substring match "
            "on the page name; page bodies are not searched.",
            {
                "query": s.string("Substring to look for in page names"),
                "limit": s.optional(s.number("Maximum matches to return (default 25)")),
            },
            s.obj(
                {
                    "total": s.number(),
                    "returned": s.number(),
                    "pages": s.array(
                        s.obj(
                            {
                                "pageId": s.string(),
                                "page": s.string(),
                                "section": s.string(),
                                "notebook": s.string(),
                            },
                            additional=True,
                        )
                    ),
                },
                additional=True,
            ),
            READ_ONLY,
            search,
        ),
        spec(
            "read_page",
            "Read OneNote page",
            "Read one OneNote page: its paragraphs as text, each with the objectID needed to reference it in "
            "edit, plus its outlines, tables and images. The reported lastModified is informational and must "
            "not be used as a concurrency check.",
            {
                **TARGET_PARAMS,
                "includeXml": s.optional(
                    s.boolean(
                        "Also return the raw OneNote XML. Large, and it includes the author's account metadata."
                    )
                ),
            },
            s.obj({"page": PAGE}, additional=True),
            READ_ONLY,
            read_page,
        ),
        spec(
            "create_page",
            "Create OneNote page",
            "Create a new page in a OneNote section with a title and optional body paragraphs, each its own "
            "paragraph in one outline. A line is a plain string, or an object that adds depth, a list, a "
            "quick style and character formatting to it. The dry run validates the section but cannot show "
            f"the page, which does not exist yet. {WRITE_FLOW}",
            {
                **section_params,
                "title": s.string("Page title"),
                "lines": s.optional(s.array(LINE, "Body paragraphs, in order")),
                **WRITE_PARAMS,
            },
            WRITE_RESULT,
            ADDITIVE,
            create_page,
        ),
        spec(
            "append_lines",
            "Append to OneNote page",
            "Append paragraphs to a page as a new outline at the end. Existing content is not touched. A line "
            "is a plain string, or an object that adds depth, a list, a quick style and character formatting "
            "to it; depth nests a paragraph under the nearest earlier one at the previous level. Use this for "
            f"adding to the end of a page; use edit to change text inside an existing outline. {WRITE_FLOW}",
            {
                **TARGET_PARAMS,
                "lines": s.array(LINE, "Paragraphs to append, in order"),
                **WRITE_PARAMS,
            },
            WRITE_RESULT,
            ADDITIVE,
            append_lines,
        ),
        spec(
            "insert_table",
            "Insert OneNote table",
            "Insert a table into a page: give headers and rows as arrays of cell text. Without outline it "
            "becomes a new outline at the end of the page; with outline it is appended inside that existing "
            "outline; with after it goes right after that paragraph instead. Cells are single paragraphs and "
            f"cannot contain nested tables. {WRITE_FLOW}",
            {
                **TARGET_PARAMS,
                "headers": s.optional(s.array(s.string("One header cell"), "Optional header row")),
                "rows": s.array(s.array(s.string("One cell")), "Rows of cell text"),
                "bordersVisible": s.optional(s.boolean("Draw cell borders; OneNote's default is false")),
                "hasHeaderRow": s.optional(s.boolean("Render the first row as a header row")),
                "columnWidths": s.optional(
                    s.array(s.number("Width of one column"), "Optional per-column widths")
                ),
                "outline": s.optional(s.string("objectID of an existing outline to append the table to")),
                "after": s.optional(
                    s.string(
                        "objectID of a paragraph to insert after, at any depth; the outline is derived from "
                        "it when outline is omitted"
                    )
                ),
                **WRITE_PARAMS,
            },
            WRITE_RESULT,
            ADDITIVE,
            insert_table,
        ),
        spec(
            "insert_image",
            "Insert OneNote image",
            "Insert an image file into a page, read from a path this server can reach (a WSL path, or "
            "/mnt/c/... for a file on the Windows side). Without outline it becomes a new outline at the end "
            "of the page; with outline it is appended inside that existing outline; with after it goes right "
            "after that paragraph instead. Without width and height OneNote picks the size. The image is "
            f"embedded as given: crop it before calling if only part of it belongs on the page. {WRITE_FLOW}",
            {
                **TARGET_PARAMS,
                "path": s.string("Image file path readable from this server"),
                "format": s.optional(
                    s.string("png, jpg, emf or auto; derived from the extension when omitted")
                ),
                "width": s.optional(s.number("Rendered width; needs height too")),
                "height": s.optional(s.number("Rendered height; needs width too")),
                "alt": s.optional(s.string("Alternative text stored on the image")),
                "outline": s.optional(s.string("objectID of an existing outline to append the image to")),
                "after": s.optional(
                    s.string(
                        "objectID of a paragraph to insert after, at any depth; the outline is derived from "
                        "it when outline is omitted"
                    )
                ),
                **WRITE_PARAMS,
            },
            WRITE_RESULT,
            ADDITIVE,
            insert_image,
        ),
        spec(
            "edit",
            "Edit OneNote page text",
            "Change paragraphs inside existing outlines of a page: replace their text with set, delete them "
            "with remove, or add paragraphs to an outline with append, optionally after a given paragraph. "
            "Paragraph objectIDs come from read_page, and outline objectIDs may be used as an append target. "
            "OneNote replaces an outline "
            "with exactly what the caller sends, so this tool resends each affected outline in full, echoing "
            "back OneNote's own XML for everything it is not changing - that is why paragraph XML is never "
            f"accepted from callers. {WRITE_FLOW}",
            {
                **TARGET_PARAMS,
                "set": s.optional(
                    s.array(
                        s.obj(
                            {
                                "objectID": s.string("objectID of the paragraph to replace"),
                                "text": s.string("New text for that paragraph"),
                            },
                            additional=False,
                        ),
                        "Paragraphs whose text is replaced",
                    )
                ),
                "remove": s.optional(
                    s.array(s.string("objectID of a paragraph to delete"), "Paragraphs to delete")
                ),
                "append": s.optional(
                    s.array(
                        s.obj(
                            {
                                "objectID": s.string(
                                    "objectID of the outline (or of any paragraph in it) to add to"
                                ),
                                "lines": s.array(LINE, "Paragraphs to add inside that outline"),
                                "after": s.optional(
                                    s.string("objectID of a paragraph in that outline to insert after")
                                ),
                            },
                            additional=False,
                        ),
                        "Paragraphs to add inside an existing outline",
                    )
                ),
                **WRITE_PARAMS,
            },
            WRITE_RESULT,
            REWRITING,
            edit,
        ),
        spec(
            "delete_page",
            "Delete OneNote page",
            "Delete a OneNote page. It goes to the notebook's recycle bin, so it can be restored from OneNote. "
            "expectName is required and must match the page the target resolves to, so a stale or wrong id "
            f"cannot delete a different page. {WRITE_FLOW}",
            {
                **TARGET_PARAMS,
                "expectName": s.string("Exact page name the target must have; required safety check"),
                **WRITE_PARAMS,
            },
            WRITE_RESULT,
            REWRITING,
            delete_page,
        ),
    ]
