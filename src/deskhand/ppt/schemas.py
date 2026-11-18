"""Input and output schemas of the PowerPoint tools, written out by hand."""

from __future__ import annotations

from typing import Any

from deskhand.core.schema import (
    array,
    boolean,
    literal,
    number,
    obj,
    optional,
    string,
    union,
)

Schema = dict[str, Any]


def _loose(properties: dict[str, Any]) -> Schema:
    return obj(properties, additional=True)


def _input(properties: dict[str, Any]) -> Schema:
    return obj(properties, additional=False)


COLOR = string("#RRGGBB")
PRESENTATION_REF = optional(
    string("Open presentation name or path. Omit when exactly one presentation is open.")
)
SLIDE_INDEX = number("1-based slide index.")

WRITE_PARAMS: dict[str, Any] = {
    "dryRun": optional(
        boolean(
            "Default true: preview only. Pass false, together with applyToken and expectFingerprint, to apply."
        )
    ),
    "applyToken": optional(
        string("The applyToken from the dry run of this same call. Required when dryRun is false.")
    ),
    "expectFingerprint": optional(string("The fingerprint from the dry run. Required when dryRun is false.")),
}

FINGERPRINT_QUALITY = union(literal("full"), literal("coarse"), literal("unavailable"))


PRESENTATION = _loose(
    {
        "name": string(),
        "path": string(),
        "slides": number(),
        "widthPx": number(),
        "heightPx": number(),
        "saved": boolean(),
    }
)

SHAPE = _loose(
    {
        "id": number(),
        "name": string(),
        "type": string(),
        "x": number(),
        "y": number(),
        "w": number(),
        "h": number(),
        "text": string(),
        "paragraphs": optional(array(_loose({}))),
        "table": optional(_loose({})),
        "chart": optional(_loose({})),
        "picture": optional(_loose({})),
    }
)

SLIDE = _loose(
    {
        "index": number(),
        "layout": string(),
        "title": string(),
        "shapes": array(SHAPE),
        "notes": optional(string()),
    }
)

DECK = _loose(
    {
        "presentation": PRESENTATION,
        "slides": array(SLIDE),
        "fingerprint": optional(string()),
        "fingerprintQuality": FINGERPRINT_QUALITY,
    }
)

READ_SLIDE_RESULT = _loose(
    {
        "slide": SLIDE,
        "fingerprint": optional(string()),
        "fingerprintQuality": FINGERPRINT_QUALITY,
    }
)

WRITE_RESULT = _loose(
    {
        "status": string(),
        "fingerprint": string(),
        "fingerprintQuality": optional(FINGERPRINT_QUALITY),
        "applyToken": optional(string()),
        "preview": optional(string()),
        "presentation": optional(PRESENTATION),
        "slide": optional(SLIDE),
        "path": optional(string()),
        "closed": optional(string()),
    }
)

RENDER_FILE = _loose({"path": string(), "widthPx": number(), "heightPx": number()})
RENDER_RESULT = _loose({"files": array(RENDER_FILE), "elapsedMs": number(), "outDir": string()})
PRESENTATIONS_RESULT = _loose({"running": boolean(), "presentations": array(PRESENTATION)})
PING_RESULT = _loose({"ready": boolean(), "elapsedMs": number()})


RUN = _loose(
    {
        "text": string(),
        "bold": optional(boolean()),
        "italic": optional(boolean()),
        "color": optional(COLOR),
        "size": optional(number("Font size in px at 96 dpi; at least 1 pt (1.34 px).")),
        "font": optional(string()),
    }
)
PARAGRAPH = _loose(
    {
        "runs": optional(array(RUN)),
        "text": optional(string("Shorthand for a single run when runs is omitted.")),
        "align": optional(string("left, center, right or justify.")),
        "spaceBefore": optional(number("px at 96 dpi.")),
        "spaceAfter": optional(number("px at 96 dpi.")),
        "lineSpacing": optional(number("px at 96 dpi.")),
    }
)
SERIES = _loose(
    {
        "name": string(),
        "values": array(number()),
        "color": optional(COLOR),
    }
)
THEME = _loose(
    {
        "background": optional(COLOR),
        "text": optional(COLOR),
        "accent": optional(COLOR),
        "font": optional(string()),
    }
)


PRESENTATIONS_INPUT = _input({})
READ_DECK_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "slides": optional(string('1-based slide selection, for example "1-5" or "1,3,5".')),
        "includeNotes": optional(boolean("Default false.")),
    }
)
READ_SLIDE_INPUT = _input(
    {"presentation": PRESENTATION_REF, "slide": SLIDE_INDEX, "includeNotes": optional(boolean())}
)
RENDER_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "slides": optional(string("1-based slide selection, PNG only.")),
        "outDir": string("Directory on a Windows drive, for example /mnt/c/Users/me/Temp/deck."),
        "width": optional(number("Output width in px for PNG; default 1600.")),
        "height": optional(number("Output height in px for PNG; default 900.")),
        "format": optional(string("png (default) or pdf.")),
        "overwrite": optional(boolean("Default false.")),
    }
)

CREATE_DECK_INPUT = _input(
    {
        "path": optional(string("Optional .pptx path on a Windows drive.")),
        "width": optional(number("Canvas width in px; default 1280.")),
        "height": optional(number("Canvas height in px; default 720.")),
        "theme": optional(THEME),
        **WRITE_PARAMS,
    }
)
ADD_SLIDE_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "layout": optional(string("blank (default), title, title-content, two-column or section.")),
        "position": optional(number("1-based insertion position; default at the end.")),
        "title": optional(string()),
        "subtitle": optional(string()),
        "background": optional(COLOR),
        **WRITE_PARAMS,
    }
)
TEXT_FIELDS = {
    "x": number("px at 96 dpi."),
    "y": number("px at 96 dpi."),
    "w": number("px at 96 dpi."),
    "h": number("px at 96 dpi."),
    "paragraphs": array(PARAGRAPH),
    "padding": optional(number("px at 96 dpi, applied to all four sides.")),
    "anchor": optional(string("top (default), middle or bottom.")),
    "fill": optional(COLOR),
    "line": optional(COLOR),
}
TABLE_FIELDS = {
    "x": number("px at 96 dpi."),
    "y": number("px at 96 dpi."),
    "w": number("px at 96 dpi."),
    "h": number("px at 96 dpi."),
    "header": optional(array(string())),
    "rows": array(array(string())),
    "columnWidths": optional(array(number("px at 96 dpi."))),
    "headerFill": optional(COLOR),
    "fontSize": optional(number("px at 96 dpi; at least 1 pt (1.34 px).")),
    "headerFontSize": optional(number("px at 96 dpi; at least 1 pt (1.34 px).")),
}
PICTURE_FIELDS = {
    "x": number("px at 96 dpi."),
    "y": number("px at 96 dpi."),
    "w": number("px at 96 dpi."),
    "h": number("px at 96 dpi."),
    "path": optional(string("Image path on a Windows drive.")),
    "dataBase64": optional(string("Image bytes as base64, when path is omitted.")),
    "fit": optional(string("stretch (default), contain or cover.")),
}
ADD_TEXT_INPUT = _input(
    {"presentation": PRESENTATION_REF, "slide": SLIDE_INDEX, **TEXT_FIELDS, **WRITE_PARAMS}
)
ADD_TABLE_INPUT = _input(
    {"presentation": PRESENTATION_REF, "slide": SLIDE_INDEX, **TABLE_FIELDS, **WRITE_PARAMS}
)
ADD_CHART_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "slide": SLIDE_INDEX,
        "x": number(),
        "y": number(),
        "w": number(),
        "h": number(),
        "type": string("column, bar, line or pie."),
        "categories": array(string()),
        "series": array(SERIES),
        "title": optional(string()),
        "legend": optional(boolean("Default true.")),
        "dataLabels": optional(boolean("Default false.")),
        "numberFormat": optional(string("For example 0.0% or #,##0.")),
        "valueAxisTitle": optional(string()),
        **WRITE_PARAMS,
    }
)
ADD_PICTURE_INPUT = _input(
    {"presentation": PRESENTATION_REF, "slide": SLIDE_INDEX, **PICTURE_FIELDS, **WRITE_PARAMS}
)
TEXT_SPEC = _input({"kind": literal("text"), **TEXT_FIELDS})
TABLE_SPEC = _input({"kind": literal("table"), **TABLE_FIELDS})
PICTURE_SPEC = _input({"kind": literal("picture"), **PICTURE_FIELDS})
SHAPE_SPEC = union(TEXT_SPEC, TABLE_SPEC, PICTURE_SPEC)
ADD_SHAPES_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "slide": SLIDE_INDEX,
        "shapes": array(SHAPE_SPEC, "One entry per shape, added in order."),
        **WRITE_PARAMS,
    }
)
ADD_SHAPES_RESULT = _input(
    {
        "status": union(literal("dry-run"), literal("applied")),
        "applied": boolean(),
        "applyToken": optional(string()),
        "preview": optional(string()),
        "fingerprint": string(),
        "fingerprintQuality": optional(FINGERPRINT_QUALITY),
        "shapeIds": array(number()),
        "appliedCount": number(),
        "presentation": optional(PRESENTATION),
        "slide": optional(SLIDE),
    }
)
SET_NOTES_INPUT = _input(
    {"presentation": PRESENTATION_REF, "slide": SLIDE_INDEX, "text": string(), **WRITE_PARAMS}
)
UPDATE_SHAPE_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "slide": SLIDE_INDEX,
        "shapeId": number("Shape id from read_deck."),
        "x": optional(number()),
        "y": optional(number()),
        "w": optional(number()),
        "h": optional(number()),
        "text": optional(string()),
        "paragraphs": optional(array(PARAGRAPH)),
        "fill": optional(COLOR),
        "line": optional(COLOR),
        **WRITE_PARAMS,
    }
)
DELETE_SHAPE_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "slide": SLIDE_INDEX,
        "shapeId": number("Shape id from read_deck."),
        **WRITE_PARAMS,
    }
)
MOVE_SLIDE_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "from": SLIDE_INDEX,
        "to": SLIDE_INDEX,
        **WRITE_PARAMS,
    }
)
DELETE_SLIDE_INPUT = _input({"presentation": PRESENTATION_REF, "slide": SLIDE_INDEX, **WRITE_PARAMS})
SAVE_DECK_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "path": string(".pptx path on a Windows drive."),
        "overwrite": optional(boolean("Default false.")),
        **WRITE_PARAMS,
    }
)
CLOSE_DECK_INPUT = _input(
    {
        "presentation": PRESENTATION_REF,
        "discard": optional(boolean("Default true: close without saving.")),
        **WRITE_PARAMS,
    }
)
