"""Input and output schemas of the OneNote tools, written out by hand."""

from __future__ import annotations

from typing import Any

from deskhand.core.schema import (
    anything,
    array,
    boolean,
    literal,
    null,
    number,
    obj,
    optional,
    string,
    union,
)

Schema = dict[str, Any]


def _loose(properties: dict[str, Any]) -> Schema:
    return obj(properties, additional=True)


_num_or_null = union(number(), null())

SECTION = _loose({"id": string(), "name": string(), "path": string()})
NOTEBOOK = _loose({"id": string(), "name": string(), "path": string(), "sections": array(SECTION)})
PAGE_META = _loose({"id": string(), "name": string(), "level": string(), "created": string()})

_PAGE_LINE = _loose(
    {
        "index": number(),
        "objectID": string(),
        "text": string(),
        "inTitle": boolean(),
        "inTable": boolean(),
        "hasImage": boolean(),
        "hasTable": boolean(),
        "hasInk": boolean(),
        "list": string(),
        "depth": number(),
        "outlineID": string(),
        "quickStyle": string(),
        "alignment": string(),
    }
)
_PAGE_OUTLINE = _loose(
    {
        "objectID": string(),
        "x": _num_or_null,
        "y": _num_or_null,
        "z": _num_or_null,
        "width": _num_or_null,
        "height": _num_or_null,
        "inTitle": boolean(),
        "lineCount": number(),
    }
)
_PAGE_IMAGE = _loose(
    {
        "objectID": string(),
        "format": string(),
        "alt": string(),
        "x": _num_or_null,
        "y": _num_or_null,
        "width": _num_or_null,
        "height": _num_or_null,
        "outlineID": string(),
        "embedded": boolean(),
        "callbackID": string(),
    }
)
_PAGE_TABLE = _loose({"objectID": string(), "rows": number(), "columns": number(), "outlineID": string()})

PAGE = _loose(
    {
        "id": string(),
        "name": string(),
        "level": string(),
        "lang": string(),
        "created": string(),
        "lastModified": string(),
        "lines": array(_PAGE_LINE),
        "outlines": array(_PAGE_OUTLINE),
        "images": array(_PAGE_IMAGE),
        "tables": array(_PAGE_TABLE),
        "rawXml": optional(string()),
    }
)

# A dry run is a success too, so one shape covers both outcomes of a write.
WRITE_RESULT = _loose(
    {
        "status": union(literal("dry-run"), literal("applied")),
        "applyToken": optional(string()),
        "preview": optional(string()),
        "page": optional(PAGE),
        "changes": optional(array(anything())),
    }
)

TARGET_PARAMS: dict[str, Any] = {
    "pageId": optional(string("Page ID from pages or search. Preferred.")),
    "notebook": optional(string("Notebook name, used with section and page when pageId is unknown")),
    "section": optional(string("Section name")),
    "page": optional(string("Page name")),
}

WRITE_PARAMS: dict[str, Any] = {
    "dryRun": optional(
        boolean("Default true: preview only. Pass false, together with applyToken, to apply.")
    ),
    "applyToken": optional(
        string("The applyToken from the dry run of this same call. Required when dryRun is false.")
    ),
}
