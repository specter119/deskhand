"""The PowerPoint tools."""

from __future__ import annotations

import base64
import re
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from typing import Any

from mcp.types import ToolAnnotations

from deskhand.core.apply_token import TokenError, TokenTable
from deskhand.core.compat import as_array, as_string, compact
from deskhand.core.powershell import BridgeError, BridgeFn
from deskhand.core.server import ToolOutput, ToolSpec, output_problem
from deskhand.ppt import format as fmt
from deskhand.ppt import schemas as p

READ_TIMEOUT = 10.0
WRITE_TIMEOUT = 30.0
RENDER_TIMEOUT = 55.0
ADD_SHAPES_TIMEOUT = 55.0
MAX_ADD_SHAPES = 20
SUPPORTED_SHAPE_KINDS = ("text", "table", "picture")

READ_ONLY = ToolAnnotations(read_only_hint=True, destructive_hint=False, open_world_hint=False)
ADDITIVE = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False
)
REWRITING = ToolAnnotations(
    read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False
)
EXPORTING = ToolAnnotations(
    read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False
)

WRITE_FLOW = (
    "Two steps: call it without dryRun first (nothing is written; the reply starts with an applyToken "
    "and a fingerprint), then call it again with the same arguments plus dryRun=false, that applyToken, "
    "and that expectFingerprint. A token is single use, expires after 10 minutes, and is bound to the "
    "exact arguments and to the deck state the preview saw."
)

_WINDOWS_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_WSL_DRIVE = re.compile(r"^/mnt/([A-Za-z])(?:/(.*))?$")
_COLOR = re.compile(r"^#[0-9A-Fa-f]{6}$")


def to_windows_path(path: str) -> str:
    """PowerPoint's COM server cannot read a WSL path, so only a Windows drive path is accepted."""
    text = path.strip()
    if _WINDOWS_DRIVE.match(text):
        return text.replace("/", "\\")
    match = _WSL_DRIVE.match(text)
    if match:
        rest = (match.group(2) or "").replace("/", "\\")
        return f"{match.group(1).upper()}:\\{rest}" if rest else f"{match.group(1).upper()}:\\"
    raise BridgeError(f"Path must be on a Windows drive (/mnt/c/...): {path}", "invalid_request")


def _presentation_ref(value: Any) -> str:
    text = as_string(value)
    if text.startswith("/"):
        return to_windows_path(text)
    if _WINDOWS_DRIVE.match(text):
        return text.replace("/", "\\")
    return text


def _invalid(message: str) -> BridgeError:
    return BridgeError(message, "invalid_request")


def _invalid_shape(message: str, index: int | None = None) -> BridgeError:
    return BridgeError(
        message,
        "invalid_request",
        {
            "status": "invalid",
            "tokenSpent": False,
            "appliedCount": 0,
            "failedIndex": index,
            "shapeIds": [],
            "failedShapeMayExist": False,
        },
    )


def _color(value: Any, name: str) -> None:
    if value is None:
        return
    if not isinstance(value, str) or not _COLOR.match(value):
        raise _invalid(f"{name} must be #RRGGBB: {value}")


def _enum(value: Any, allowed: set[str], name: str, default: str | None = None) -> str:
    text = as_string(value)
    if text == "":
        if default is None:
            raise _invalid(f"{name} is required")
        return default
    if text not in allowed:
        raise _invalid(f"{name} must be one of {', '.join(sorted(allowed))}: {text}")
    return text


def _positive(value: Any, name: str, default: float | None = None) -> float:
    if value is None:
        if default is None:
            raise _invalid(f"{name} is required")
        return default
    number = float(value)
    if number <= 0:
        raise _invalid(f"{name} must be greater than 0: {number:g}")
    return number


def _font_size(value: Any, name: str) -> None:
    if value is None:
        return
    number = _positive(value, name)
    if number * 0.75 < 1.0 - 1e-9:
        raise _invalid(f"{name} must be at least 1 pt (1.34 px): {number:g}")


def _non_negative(value: Any, name: str) -> None:
    if value is None:
        return
    number = float(value)
    if number < 0:
        raise _invalid(f"{name} must be 0 or greater: {number:g}")


def _paragraphs(value: Any) -> list[Any]:
    paragraphs = as_array(value)
    if not paragraphs:
        raise _invalid("paragraphs must contain at least one paragraph")
    for paragraph in paragraphs:
        if not isinstance(paragraph, dict):
            raise _invalid("each paragraph must be an object")
        runs = as_array(paragraph.get("runs"))
        if not runs and not as_string(paragraph.get("text")):
            raise _invalid("each paragraph needs runs or text")
        for run in runs:
            if not isinstance(run, dict):
                raise _invalid("each run must be an object")
            _color(run.get("color"), "run color")
            if run.get("size") is not None:
                _font_size(run.get("size"), "run size")
        _enum(paragraph.get("align"), {"left", "center", "right", "justify"}, "align", "left")
        for key in ("spaceBefore", "spaceAfter"):
            _non_negative(paragraph.get(key), key)
        if paragraph.get("lineSpacing") is not None:
            _positive(paragraph["lineSpacing"], "lineSpacing")
    return paragraphs


def _base64(value: str) -> None:
    try:
        base64.b64decode(value, validate=True)
    except Exception as error:
        raise _invalid(f"dataBase64 is not valid base64: {error}") from None


def _base64_prefix(value: str) -> None:
    compact = "".join(value.split())
    if not compact or len(compact) % 4 != 0 or not re.fullmatch(r"[A-Za-z0-9+/]*={0,2}", compact):
        raise _invalid("dataBase64 is not valid base64")
    head = compact[:12]
    head += "=" * ((4 - len(head) % 4) % 4)
    try:
        raw = base64.b64decode(head, validate=True)
    except Exception:
        raise _invalid("dataBase64 is not valid base64") from None
    if not raw.startswith((b"\x89PNG", b"\xff\xd8\xff", b"GIF8")):
        raise _invalid("dataBase64 must be a PNG, JPEG or GIF image")


@dataclass(frozen=True)
class PptDeps:
    bridge: BridgeFn
    tokens: TokenTable


@dataclass(frozen=True)
class WritePlan:
    tool: str
    op: str
    args: dict[str, Any]
    summary: str
    timeout: float = WRITE_TIMEOUT
    output_schema: dict[str, Any] | None = None
    describe: Callable[[dict[str, Any], bool], str] | None = None
    extras: Callable[[dict[str, Any], bool], dict[str, Any]] | None = None
    requires_fingerprint_quality: bool = True


async def _call(deps: PptDeps, op: str, args: dict[str, Any], timeout: float | None = None) -> dict[str, Any]:
    payload = compact(args)
    if timeout is None:
        return await deps.bridge(op, payload)
    return await deps.bridge(op, payload, timeout=timeout)


def _write_request(args: dict[str, Any]) -> dict[str, Any]:
    return compact(
        {
            key: value
            for key, value in args.items()
            if key not in ("dryRun", "applyToken", "expectFingerprint")
        }
    )


def _state_fields(data: dict[str, Any]) -> dict[str, Any]:
    return {key: data[key] for key in ("presentation", "slide", "path", "closed") if key in data}


_NOT_WRITTEN = {
    "invalid_request",
    "apply_requires_fingerprint",
    "changed_since_preview",
    "read_only",
    "file_exists",
    "no_presentation",
    "slide_not_found",
    "fingerprint_unavailable",
    "reset_pending",
    "com_error",
    "shape_not_found",
    "file_locked",
    "ambiguous",
    "unsupported",
    "unknown_op",
    "no_powerpoint",
}


def _not_applied_details(token_spent: bool) -> dict[str, Any]:
    return {
        "status": "not_applied",
        "tokenSpent": token_spent,
        "appliedCount": 0,
        "failedIndex": None,
        "shapeIds": [],
        "failedShapeMayExist": False,
    }


def _stale_details() -> dict[str, Any]:
    return {
        "status": "stale",
        "tokenSpent": True,
        "appliedCount": 0,
        "failedIndex": None,
        "shapeIds": [],
        "failedShapeMayExist": False,
    }


def _invalid_details(token_spent: bool, details: dict[str, Any] | None) -> dict[str, Any]:
    merged = dict(details or {})
    merged.update(
        {
            "status": "invalid",
            "tokenSpent": token_spent,
            "appliedCount": 0,
            "shapeIds": [],
            "failedShapeMayExist": False,
        }
    )
    merged.setdefault("failedIndex", None)
    return merged


def _write_failure(error: BridgeError, *, applied: bool) -> BridgeError:
    code = error.code
    details = error.details
    if code == "changed_since_preview":
        if applied:
            return BridgeError(
                f"{error}. Nothing was written; the slide changed since preview.", code, _stale_details()
            )
        return BridgeError(
            f"{error}. Nothing was written; preview again before retrying.",
            "invalid_request",
            _invalid_details(False, details),
        )
    if code == "invalid_request":
        return BridgeError(f"{error}. Nothing was written.", code, _invalid_details(applied, details))
    if code == "apply_requires_fingerprint":
        return BridgeError(
            f"{error}. Nothing was written.", "invalid_request", _invalid_details(False, details)
        )
    if code == "reset_pending":
        if applied:
            return BridgeError(
                f"{error}. Nothing was written; retry after cleanup.", code, _not_applied_details(True)
            )
        return BridgeError(
            f"{error}. Nothing was written; retry after cleanup.",
            "invalid_request",
            _invalid_details(False, details),
        )
    if code == "fingerprint_unavailable":
        if applied:
            return BridgeError(
                f"{error}. Nothing was written; fingerprint unavailable.", code, _not_applied_details(True)
            )
        return BridgeError(f"{error}. Nothing was written.", code, _invalid_details(False, details))
    if code == "applied_unverified":
        if isinstance(details, dict) and details.get("status") == "partial":
            merged = dict(details)
            merged["tokenSpent"] = True
            count = merged.get("appliedCount")
            total = merged.get("total")
            prefix = (
                f"{count} of {total} shapes were applied; read the slide before retrying."
                if count is not None and total is not None
                else "Part of this write was applied; read the slide before retrying."
            )
            return BridgeError(f"{prefix} {error}", code, merged)
        return BridgeError(
            f"{error}. Whether PowerPoint applied it is unknown. Read the deck again before retrying.",
            code,
            _unknown_details(),
        )
    if code in _NOT_WRITTEN:
        if applied:
            return BridgeError(f"{error}. Nothing was written.", code, _not_applied_details(True))
        return BridgeError(
            f"{error}. Nothing was written.", "invalid_request", _invalid_details(False, details)
        )
    if applied:
        return BridgeError(
            f"{error}. Whether PowerPoint applied it is unknown. Read the deck again before retrying.",
            code,
            _unknown_details(),
        )
    return BridgeError(f"{error}. Nothing was written.", "invalid_request", _invalid_details(False, details))


def _unknown_details() -> dict[str, Any]:
    return {
        "status": "unknown",
        "tokenSpent": True,
        "appliedCount": None,
        "failedIndex": None,
        "shapeIds": None,
        "failedShapeMayExist": None,
    }


def _checked_output(plan: WritePlan, structured: dict[str, Any], *, applied: bool) -> None:
    schema = plan.output_schema or p.WRITE_RESULT
    problem = output_problem(schema, structured)
    if problem is None:
        return
    if applied:
        raise BridgeError(
            f"Output validation error: {problem}. The write call had already returned; read the deck "
            "before retrying.",
            "outcome_unknown",
            _unknown_details(),
        )
    raise BridgeError(
        f"Output validation error: {problem}. Nothing was written.",
        "invalid_request",
        {"status": "invalid", "tokenSpent": False},
    )


async def _run_write(deps: PptDeps, plan: WritePlan) -> ToolOutput:
    request = _write_request(plan.args)
    if plan.args.get("dryRun") is not False:
        try:
            data = await _call(deps, plan.op, {**request, "apply": False}, plan.timeout)
            quality = data.get("fingerprintQuality")
            if plan.requires_fingerprint_quality and quality not in ("full", "coarse"):
                raise BridgeError(
                    "The PowerPoint bridge could not read the deck completely, so this write has no reliable "
                    "preview. Nothing was written.",
                    "fingerprint_unavailable",
                    _invalid_details(False, None),
                )
            fingerprint = as_string(data.get("fingerprint"))
            if not fingerprint:
                raise BridgeError(
                    "The PowerPoint bridge returned no fingerprint for this preview. Nothing was written.",
                    "invalid_request",
                    _invalid_details(False, None),
                )
            token = deps.tokens.issue(plan.tool, request, fingerprint)
            focus_shape_id = plan.args.get("shapeId")
            preview = (
                plan.describe(data, False)
                if plan.describe is not None
                else fmt.format_write(plan.tool, data, applied=False, focus_shape_id=focus_shape_id)
            )
            text = "\n".join(
                [
                    f"applyToken: {token}",
                    f"fingerprint: {fingerprint}",
                    plan.summary,
                    preview,
                    "",
                    f"To apply, call {plan.tool} again with the same arguments plus dryRun=false, "
                    f'applyToken="{token}", and expectFingerprint="{fingerprint}".',
                ]
            )
            structured = {
                "status": "dry-run",
                "fingerprint": fingerprint,
                "applyToken": token,
                "preview": preview,
                **_state_fields(data),
                **(plan.extras(data, False) if plan.extras is not None else {}),
            }
            if data.get("fingerprintQuality"):
                structured["fingerprintQuality"] = data["fingerprintQuality"]
            _checked_output(plan, structured, applied=False)
            return ToolOutput(text=text, structured=structured)
        except BridgeError as error:
            raise _write_failure(error, applied=False) from None
        except Exception as error:
            raise BridgeError(
                f"{error}. Nothing was written.", "invalid_request", _invalid_details(False, None)
            ) from None

    token = as_string(plan.args.get("applyToken"))
    expected = as_string(plan.args.get("expectFingerprint"))
    if not expected:
        raise BridgeError(
            "apply requires expectFingerprint from the dry run",
            "apply_requires_fingerprint",
            _invalid_details(False, None),
        )
    try:
        stored = deps.tokens.consume(token, plan.tool, request)
    except TokenError as error:
        raise BridgeError(f"{error}. Preview again before retrying.", error.reason) from None
    if stored != expected:
        raise BridgeError(
            "expectFingerprint does not match the dry run, so the applyToken is spent. "
            "Read the deck again and preview before retrying.",
            "changed_since_preview",
        )
    try:
        data = await _call(
            deps,
            plan.op,
            {**request, "apply": True, "expectFingerprint": stored},
            plan.timeout,
        )
    except BridgeError as error:
        raise _write_failure(error, applied=True) from None
    try:
        fingerprint = as_string(data.get("fingerprint"))
        if not fingerprint:
            raise BridgeError(
                "The PowerPoint bridge returned no fingerprint after the write, so whether it applied "
                "the whole change is unknown. Read the deck again before retrying.",
                "outcome_unknown",
                _unknown_details(),
            )
        structured = {
            "status": "applied",
            "fingerprint": fingerprint,
            **_state_fields(data),
            **(plan.extras(data, True) if plan.extras is not None else {}),
        }
        if data.get("fingerprintQuality"):
            structured["fingerprintQuality"] = data["fingerprintQuality"]
        _checked_output(plan, structured, applied=True)
        body = (
            plan.describe(data, True)
            if plan.describe is not None
            else fmt.format_write(plan.tool, data, applied=True, focus_shape_id=plan.args.get("shapeId"))
        )
        return ToolOutput(
            text="\n".join([plan.summary, body]),
            structured=structured,
        )
    except BridgeError:
        raise
    except Exception as error:
        raise BridgeError(
            f"{error}. Whether PowerPoint applied it is unknown. Read the deck again before retrying.",
            "outcome_unknown",
            _unknown_details(),
        ) from None


def _slide(args: dict[str, Any]) -> int:
    value = int(args["slide"])
    if value < 1:
        raise _invalid(f"slide must be 1 or greater: {value}")
    return value


def _preview_text(value: Any, limit: int = 90) -> str:
    return fmt.short_text(as_string(value), limit)


def _text_spec(args: dict[str, Any], where: str = "text") -> dict[str, Any]:
    paragraphs = _paragraphs(args.get("paragraphs"))
    for key in ("w", "h"):
        _positive(args.get(key), f"{where}.{key}")
    _non_negative(args.get("padding"), f"{where}.padding")
    anchor = _enum(args.get("anchor"), {"top", "middle", "bottom"}, f"{where}.anchor", "top")
    _color(args.get("fill"), f"{where}.fill")
    _color(args.get("line"), f"{where}.line")
    return {
        "x": args.get("x"),
        "y": args.get("y"),
        "w": args.get("w"),
        "h": args.get("h"),
        "paragraphs": paragraphs,
        "padding": args.get("padding"),
        "anchor": anchor,
        "fill": args.get("fill"),
        "line": args.get("line"),
    }


def _table_spec(args: dict[str, Any], where: str = "table") -> dict[str, Any]:
    rows = as_array(args.get("rows"))
    header = as_array(args.get("header"))
    if not rows and not header:
        raise _invalid(f"{where} needs a header or at least one row")
    if any(not isinstance(row, list) for row in rows):
        raise _invalid(f"each row in {where} must be an array")
    columns = max([len(header), *(len(row) for row in rows)])
    if columns < 1:
        raise _invalid(f"{where} needs at least one column")
    widths = as_array(args.get("columnWidths"))
    if widths and len(widths) != columns:
        raise _invalid(f"{where}.columnWidths must have one entry per column")
    for width in widths:
        _positive(width, f"{where}.columnWidths entry")
    _color(args.get("headerFill"), f"{where}.headerFill")
    for key in ("fontSize", "headerFontSize"):
        _font_size(args.get(key), f"{where}.{key}")
    return {
        "x": args.get("x"),
        "y": args.get("y"),
        "w": args.get("w"),
        "h": args.get("h"),
        "header": header or None,
        "rows": rows,
        "columnWidths": widths or None,
        "headerFill": args.get("headerFill"),
        "fontSize": args.get("fontSize"),
        "headerFontSize": args.get("headerFontSize"),
    }


def _picture_spec(args: dict[str, Any], where: str = "picture") -> dict[str, Any]:
    path = as_string(args.get("path"))
    data = as_string(args.get("dataBase64"))
    if bool(path) == bool(data):
        raise BridgeError(f"{where} needs exactly one of path or dataBase64", "invalid_request")
    fit = _enum(args.get("fit"), {"stretch", "contain", "cover"}, f"{where}.fit", "stretch")
    if data:
        _base64_prefix(data)
    for key in ("w", "h"):
        _positive(args.get(key), f"{where}.{key}")
    return {
        "x": args.get("x"),
        "y": args.get("y"),
        "w": args.get("w"),
        "h": args.get("h"),
        "path": to_windows_path(path) if path else None,
        "dataBase64": data or None,
        "fit": fit,
    }


_SPEC_PLANNERS: dict[str, Callable[[dict[str, Any], str], dict[str, Any]]] = {
    "text": _text_spec,
    "table": _table_spec,
    "picture": _picture_spec,
}


def build_tools(deps: PptDeps) -> list[ToolSpec]:
    def spec(
        name: str,
        title: str,
        description: str,
        input_schema: dict[str, Any],
        output: dict[str, Any],
        annotations: ToolAnnotations,
        run: Callable[[dict[str, Any]], Awaitable[ToolOutput]],
    ) -> ToolSpec:
        return ToolSpec(name, title, description, input_schema, output, annotations, run)

    async def presentations(args: dict[str, Any]) -> ToolOutput:
        data = await _call(deps, "presentations", {}, READ_TIMEOUT)
        return ToolOutput(fmt.format_presentations(data), data)

    async def read_deck(args: dict[str, Any]) -> ToolOutput:
        include_notes = args.get("includeNotes") is True
        slides = as_string(args.get("slides"))
        data = await _call(
            deps,
            "read_deck",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slides": slides,
                "includeNotes": include_notes,
            },
            READ_TIMEOUT,
        )
        return ToolOutput(
            fmt.format_deck(data, include_notes=include_notes, expand_shapes=bool(slides)), data
        )

    async def read_slide(args: dict[str, Any]) -> ToolOutput:
        include_notes = args.get("includeNotes") is True
        data = await _call(
            deps,
            "read_slide",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                "includeNotes": include_notes,
            },
            READ_TIMEOUT,
        )
        return ToolOutput(fmt.format_slide_result(data, include_notes=include_notes), data)

    async def render(args: dict[str, Any]) -> ToolOutput:
        output_format = _enum(args.get("format"), {"png", "pdf"}, "format", "png")
        width = args.get("width")
        height = args.get("height")
        if width is not None:
            _positive(width, "width")
        if height is not None:
            _positive(height, "height")
        if output_format == "pdf" and (width is not None or height is not None):
            raise _invalid("width and height apply to PNG only; a PDF keeps the deck's page size")
        if output_format == "pdf" and as_string(args.get("slides")):
            raise _invalid("slides applies to PNG only; a PDF contains the whole deck")
        data = await _call(
            deps,
            "render",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slides": args.get("slides"),
                "outDir": to_windows_path(as_string(args.get("outDir"))),
                "width": width,
                "height": height,
                "format": output_format,
                "overwrite": args.get("overwrite"),
            },
            RENDER_TIMEOUT,
        )
        return ToolOutput(fmt.format_render(data), data)

    async def create_deck(args: dict[str, Any]) -> ToolOutput:
        path = as_string(args.get("path"))
        width = int(_positive(args.get("width"), "width", 1280))
        height = int(_positive(args.get("height"), "height", 720))
        theme = args.get("theme")
        if theme is not None:
            if not isinstance(theme, dict):
                raise _invalid("theme must be an object")
            for key in ("background", "text", "accent"):
                _color(theme.get(key), f"theme {key}")
        plan = WritePlan(
            "create_deck",
            "create_deck",
            {
                "path": to_windows_path(path) if path else None,
                "width": width,
                "height": height,
                "theme": theme,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would create a {width}x{height} px presentation"
            + (f" at {path}" if path else ".")
            + (
                " Theme: "
                + ", ".join(
                    f"{key}={theme[key]}"
                    for key in ("background", "text", "accent", "font")
                    if theme.get(key)
                )
                if theme
                else ""
            ),
            requires_fingerprint_quality=False,
        )
        return await _run_write(deps, plan)

    async def add_slide(args: dict[str, Any]) -> ToolOutput:
        layout = _enum(
            args.get("layout"),
            {"blank", "title", "title-content", "two-column", "section"},
            "layout",
            "blank",
        )
        position = args.get("position")
        if position is not None:
            _positive(position, "position")
        _color(args.get("background"), "background")
        title = as_string(args.get("title"))
        subtitle = as_string(args.get("subtitle"))
        where = f"at position {int(position)}" if position is not None else "at the end"
        plan = WritePlan(
            "add_slide",
            "add_slide",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "layout": layout,
                "position": position,
                "title": title or None,
                "subtitle": subtitle or None,
                "background": args.get("background"),
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would add a {layout} slide {where}."
            + (f' Title: "{_preview_text(title)}".' if title else "")
            + (f' Subtitle: "{_preview_text(subtitle)}".' if subtitle else ""),
        )
        return await _run_write(deps, plan)

    async def add_text(args: dict[str, Any]) -> ToolOutput:
        spec = _text_spec(args)
        paragraphs = spec["paragraphs"]
        first = as_string(paragraphs[0].get("text"))
        if not first:
            first = "".join(as_string(run.get("text")) for run in as_array(paragraphs[0].get("runs")))
        plan = WritePlan(
            "add_text",
            "add_text",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                **spec,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would add a text box to slide {_slide(args)} with {len(paragraphs)} paragraph(s): "
            f'"{_preview_text(first)}".',
        )
        return await _run_write(deps, plan)

    async def add_table(args: dict[str, Any]) -> ToolOutput:
        spec = _table_spec(args)
        rows = spec["rows"]
        header = spec["header"] or []
        columns = max([len(header), *(len(row) for row in rows)])
        first_row = header or (rows[0] if rows else [])
        plan = WritePlan(
            "add_table",
            "add_table",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                **spec,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would add a {len(rows) + (1 if args.get('header') else 0)}x{columns} table to slide "
            f"{_slide(args)}: {', '.join(_preview_text(cell, 30) for cell in first_row)}.",
        )
        return await _run_write(deps, plan)

    async def add_chart(args: dict[str, Any]) -> ToolOutput:
        series = as_array(args.get("series"))
        categories = as_array(args.get("categories"))
        chart_type = _enum(args.get("type"), {"column", "bar", "line", "pie"}, "type")
        if not categories:
            raise _invalid("categories must contain at least one value")
        if not series:
            raise _invalid("series must contain at least one series")
        for item in series:
            if not isinstance(item, dict):
                raise _invalid("each series must be an object")
            values = as_array(item.get("values"))
            if len(values) != len(categories):
                raise _invalid("every series needs one value per category")
            _color(item.get("color"), "series color")
        for key in ("w", "h"):
            _positive(args.get(key), key)
        names = ", ".join(_preview_text(item.get("name"), 30) for item in series)
        plan = WritePlan(
            "add_chart",
            "add_chart",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                "x": args.get("x"),
                "y": args.get("y"),
                "w": args.get("w"),
                "h": args.get("h"),
                "type": chart_type,
                "categories": categories,
                "series": series,
                "title": args.get("title"),
                "legend": args.get("legend"),
                "dataLabels": args.get("dataLabels"),
                "numberFormat": args.get("numberFormat"),
                "valueAxisTitle": args.get("valueAxisTitle"),
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would add a {chart_type} chart to slide {_slide(args)} with {len(series)} series "
            f"({names}) and {len(categories)} categories: "
            f"{', '.join(_preview_text(category, 20) for category in categories)}.",
        )
        return await _run_write(deps, plan)

    async def add_picture(args: dict[str, Any]) -> ToolOutput:
        spec = _picture_spec(args)
        path = as_string(spec.get("path"))
        data = as_string(spec.get("dataBase64"))
        fit = as_string(spec.get("fit"))
        source = path if path else f"base64 ({len(data)} characters)"
        plan = WritePlan(
            "add_picture",
            "add_picture",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                **spec,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would add a picture to slide {_slide(args)} from {source} with fit={fit}.",
        )
        return await _run_write(deps, plan)

    async def add_shapes(args: dict[str, Any]) -> ToolOutput:
        raw = as_array(args.get("shapes"))
        if not raw:
            raise _invalid_shape("shapes must contain at least one shape")
        if len(raw) > MAX_ADD_SHAPES:
            raise _invalid_shape(f"add_shapes takes at most {MAX_ADD_SHAPES} shapes, got {len(raw)}")
        normalized: list[dict[str, Any]] = []
        for index, item in enumerate(raw, start=1):
            if not isinstance(item, dict):
                raise _invalid_shape(f"shapes[{index}] must be an object", index)
            kind = as_string(item.get("kind"))
            planner = _SPEC_PLANNERS.get(kind)
            if planner is None:
                raise _invalid_shape(
                    f"shapes[{index}].kind must be one of {', '.join(SUPPORTED_SHAPE_KINDS)}: {kind}",
                    index,
                )
            try:
                normalized.append({"kind": kind, **planner(item, f"shapes[{index}]")})
            except BridgeError as error:
                raise _invalid_shape(str(error), index) from None

        def describe(data: dict[str, Any], applied: bool) -> str:
            if applied:
                ids = [int(value) for value in as_array(data.get("shapeIds"))]
                return "Applied shape id(s): " + ", ".join(str(value) for value in ids) + "."
            lines = [f"{len(normalized)} shape(s), added in this order:"]
            for index, spec in enumerate(normalized, start=1):
                kind = spec["kind"]
                at = f"({spec['x']}, {spec['y']}) {spec['w']}x{spec['h']}"
                if kind == "text":
                    paragraphs = spec["paragraphs"]
                    first = as_string(paragraphs[0].get("text")) or "".join(
                        as_string(run.get("text")) for run in as_array(paragraphs[0].get("runs"))
                    )
                    lines.append(f'  {index}. text at {at}: "{_preview_text(first, 60)}"')
                elif kind == "table":
                    rows = spec["rows"]
                    header = spec["header"] or []
                    columns = max([len(header), *(len(row) for row in rows)])
                    lines.append(f"  {index}. table {len(rows) + (1 if header else 0)}x{columns} at {at}")
                else:
                    source = as_string(spec.get("path")) or "base64"
                    lines.append(f"  {index}. picture at {at} fit={spec['fit']} from {source}")
            return "\n".join(lines)

        def extras(data: dict[str, Any], applied: bool) -> dict[str, Any]:
            if not applied:
                return {"applied": False, "shapeIds": [], "appliedCount": 0}
            ids = [int(value) for value in as_array(data.get("shapeIds"))]
            if len(ids) != len(normalized):
                raise BridgeError(
                    "The PowerPoint bridge returned a different number of shape ids than requested; "
                    "read the slide before retrying.",
                    "outcome_unknown",
                    _unknown_details(),
                )
            return {"applied": True, "shapeIds": ids, "appliedCount": len(ids)}

        plan = WritePlan(
            "add_shapes",
            "add_shapes",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                "shapes": normalized,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would add {len(normalized)} shape(s) to slide {_slide(args)}.",
            timeout=ADD_SHAPES_TIMEOUT,
            output_schema=p.ADD_SHAPES_RESULT,
            describe=describe,
            extras=extras,
        )
        return await _run_write(deps, plan)

    async def set_notes(args: dict[str, Any]) -> ToolOutput:
        text = as_string(args.get("text"))
        plan = WritePlan(
            "set_notes",
            "set_notes",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                "text": text,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would replace the notes on slide {_slide(args)} with {len(text)} character(s): "
            f'"{_preview_text(text)}".',
        )
        return await _run_write(deps, plan)

    async def update_shape(args: dict[str, Any]) -> ToolOutput:
        shape_id = int(args["shapeId"])
        if shape_id < 1:
            raise _invalid(f"shapeId must be 1 or greater: {shape_id}")
        paragraphs = _paragraphs(args["paragraphs"]) if args.get("paragraphs") is not None else None
        if not any(
            args.get(key) is not None for key in ("x", "y", "w", "h", "text", "paragraphs", "fill", "line")
        ):
            raise _invalid("update_shape needs at least one of x, y, w, h, text, paragraphs, fill or line")
        for key in ("w", "h"):
            if args.get(key) is not None:
                _positive(args[key], key)
        _color(args.get("fill"), "fill")
        _color(args.get("line"), "line")
        detail = as_string(args.get("text"))
        if paragraphs:
            detail = as_string(paragraphs[0].get("text")) or "".join(
                as_string(run.get("text")) for run in as_array(paragraphs[0].get("runs"))
            )
        plan = WritePlan(
            "update_shape",
            "update_shape",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                "shapeId": shape_id,
                "x": args.get("x"),
                "y": args.get("y"),
                "w": args.get("w"),
                "h": args.get("h"),
                "text": args.get("text"),
                "paragraphs": paragraphs,
                "fill": args.get("fill"),
                "line": args.get("line"),
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would update shape {shape_id} on slide {_slide(args)}"
            + (f': "{_preview_text(detail)}"' if detail else "")
            + ".",
        )
        return await _run_write(deps, plan)

    async def delete_shape(args: dict[str, Any]) -> ToolOutput:
        shape_id = int(args["shapeId"])
        if shape_id < 1:
            raise _invalid(f"shapeId must be 1 or greater: {shape_id}")
        plan = WritePlan(
            "delete_shape",
            "delete_shape",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                "shapeId": shape_id,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would delete shape {shape_id} from slide {_slide(args)}.",
        )
        return await _run_write(deps, plan)

    async def move_slide(args: dict[str, Any]) -> ToolOutput:
        source = int(args["from"])
        target = int(args["to"])
        if source < 1 or target < 1:
            raise _invalid("from and to must be 1 or greater")
        plan = WritePlan(
            "move_slide",
            "move_slide",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "from": source,
                "to": target,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would move slide {source} to position {target}.",
        )
        return await _run_write(deps, plan)

    async def delete_slide(args: dict[str, Any]) -> ToolOutput:
        plan = WritePlan(
            "delete_slide",
            "delete_slide",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "slide": _slide(args),
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would delete slide {_slide(args)}.",
        )
        return await _run_write(deps, plan)

    async def save_deck(args: dict[str, Any]) -> ToolOutput:
        raw_path = as_string(args.get("path"))
        if not raw_path:
            raise _invalid("path is required")
        path = to_windows_path(raw_path)
        plan = WritePlan(
            "save_deck",
            "save_deck",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "path": path,
                "overwrite": args.get("overwrite"),
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            f"Would save the presentation to {path}"
            + (" (overwriting)." if args.get("overwrite") is True else "."),
        )
        return await _run_write(deps, plan)

    async def close_deck(args: dict[str, Any]) -> ToolOutput:
        discard = args.get("discard") is not False
        plan = WritePlan(
            "close_deck",
            "close_deck",
            {
                "presentation": _presentation_ref(args.get("presentation")),
                "discard": discard,
                **{key: args.get(key) for key in ("dryRun", "applyToken", "expectFingerprint")},
            },
            "Would close the presentation " + ("without saving." if discard else "after saving."),
        )
        return await _run_write(deps, plan)

    return [
        spec(
            "presentations",
            "Open presentations",
            "List the presentations PowerPoint currently has open, with their canvas size and slide count.",
            p.PRESENTATIONS_INPUT,
            p.PRESENTATIONS_RESULT,
            READ_ONLY,
            presentations,
        ),
        spec(
            "read_deck",
            "Read a deck",
            "Read a deck's slides, shapes, text, geometry in px at 96 dpi, tables, charts, pictures and "
            'notes. Omit slides for a one-line-per-slide overview; pass slides (for example "1-3") to '
            "expand shapes. A path is opened read-only; a write tool reopens it editable when it has no "
            "unsaved changes.",
            p.READ_DECK_INPUT,
            p.DECK,
            READ_ONLY,
            read_deck,
        ),
        spec(
            "read_slide",
            "Read a slide",
            "Read one slide and its fingerprint. The fingerprint is what a write preview must still match.",
            p.READ_SLIDE_INPUT,
            p.READ_SLIDE_RESULT,
            READ_ONLY,
            read_slide,
        ),
        spec(
            "render",
            "Render slides",
            "Export slides as PNG files or the deck as a PDF, using PowerPoint's own renderer. "
            "outDir must be on a Windows drive. width/height apply to PNG; a PDF keeps the deck's page "
            "size. overwrite replaces existing files. The call ceiling is 55 s, so render large decks in "
            "slide batches.",
            p.RENDER_INPUT,
            p.RENDER_RESULT,
            EXPORTING,
            render,
        ),
        spec(
            "create_deck",
            "Create a deck",
            "Create a new presentation on a 1280x720 px canvas by default. theme accepts background, "
            "text, accent and font. " + WRITE_FLOW,
            p.CREATE_DECK_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            create_deck,
        ),
        spec(
            "add_slide",
            "Add a slide",
            "Add a slide with layout (blank, title, title-content, two-column, section), optional title, "
            "subtitle, position and background. " + WRITE_FLOW,
            p.ADD_SLIDE_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            add_slide,
        ),
        spec(
            "add_text",
            "Add a text box",
            "Add a text box with paragraphs and runs; positions and sizes are px at 96 dpi. Optional "
            "padding, anchor, fill and line. " + WRITE_FLOW,
            p.ADD_TEXT_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            add_text,
        ),
        spec(
            "add_table",
            "Add a table",
            "Add a table with an optional header row, columnWidths, headerFill, fontSize and "
            "headerFontSize. " + WRITE_FLOW,
            p.ADD_TABLE_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            add_table,
        ),
        spec(
            "add_chart",
            "Add a chart",
            "Add a column, bar, line or pie chart with categories, series and colors; optional title, "
            "legend, dataLabels, numberFormat and valueAxisTitle. " + WRITE_FLOW,
            p.ADD_CHART_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            add_chart,
        ),
        spec(
            "add_picture",
            "Add a picture",
            "Add a picture from a Windows-drive path or base64 bytes; fit is stretch, contain or cover. "
            + WRITE_FLOW,
            p.ADD_PICTURE_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            add_picture,
        ),
        spec(
            "add_shapes",
            "Add several shapes",
            f"Add up to {MAX_ADD_SHAPES} text, table or picture shapes to one slide in one preview and "
            "one apply; shapes are added in order. " + WRITE_FLOW,
            p.ADD_SHAPES_INPUT,
            p.ADD_SHAPES_RESULT,
            ADDITIVE,
            add_shapes,
        ),
        spec(
            "set_notes",
            "Set speaker notes",
            "Replace the speaker notes of one slide. " + WRITE_FLOW,
            p.SET_NOTES_INPUT,
            p.WRITE_RESULT,
            REWRITING,
            set_notes,
        ),
        spec(
            "update_shape",
            "Update a shape",
            "Move, resize, retext or recolor one shape by its id from read_deck. text is a plain-text "
            "shortcut; paragraphs keeps runs. " + WRITE_FLOW,
            p.UPDATE_SHAPE_INPUT,
            p.WRITE_RESULT,
            REWRITING,
            update_shape,
        ),
        spec(
            "delete_shape",
            "Delete a shape",
            "Delete one shape by its id from read_deck. " + WRITE_FLOW,
            p.DELETE_SHAPE_INPUT,
            p.WRITE_RESULT,
            REWRITING,
            delete_shape,
        ),
        spec(
            "move_slide",
            "Move a slide",
            "Move a slide to another 1-based position. " + WRITE_FLOW,
            p.MOVE_SLIDE_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            move_slide,
        ),
        spec(
            "delete_slide",
            "Delete a slide",
            "Delete one slide. " + WRITE_FLOW,
            p.DELETE_SLIDE_INPUT,
            p.WRITE_RESULT,
            REWRITING,
            delete_slide,
        ),
        spec(
            "save_deck",
            "Save a deck",
            "Save the presentation as .pptx on a Windows drive; overwrite replaces an existing file. "
            + WRITE_FLOW,
            p.SAVE_DECK_INPUT,
            p.WRITE_RESULT,
            ADDITIVE,
            save_deck,
        ),
        spec(
            "close_deck",
            "Close a deck",
            "Close the presentation, discarding unsaved changes by default. " + WRITE_FLOW,
            p.CLOSE_DECK_INPUT,
            p.WRITE_RESULT,
            REWRITING,
            close_deck,
        ),
    ]
