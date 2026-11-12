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
        "style": optional(string("OneNote's cascaded text style for the paragraph, when it has one")),
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

# A paragraph is either a plain string or this object. Depth is the only way to express nesting, and
# a caller that does not need formatting keeps writing plain strings.
PARAGRAPH_STYLE = obj(
    {
        "fontFamily": optional(string("Font family, for example Calibri")),
        "fontSize": optional(number("Font size in points")),
        "color": optional(string("Text colour as #RRGGBB, or automatic")),
        "bold": optional(boolean("Bold text")),
        "italic": optional(boolean("Italic text")),
        "underline": optional(boolean("Underlined text")),
        "strikethrough": optional(boolean("Struck-through text")),
        "highlight": optional(string("Highlight colour as #RRGGBB")),
    },
    additional=False,
)

PARAGRAPH = obj(
    {
        "text": string("The paragraph's text"),
        "depth": optional(
            number(
                "0 is the outline's top level; a paragraph at depth d nests under the nearest earlier "
                "paragraph at depth d-1"
            )
        ),
        "list": optional(union(literal("bullet"), literal("number"), literal("none"))),
        "quickStyle": optional(
            union(
                number(),
                string(
                    "Style name the page defines, or an index from read_page. OneNote maps the index onto a "
                    "style the page defines and may normalise it, so read_page is the source of truth"
                ),
            )
        ),
        "alignment": optional(union(literal("left"), literal("center"), literal("right"))),
        "spaceBefore": optional(number("Space above the paragraph, in points")),
        "spaceAfter": optional(number("Space below the paragraph, in points")),
        "style": optional(PARAGRAPH_STYLE),
    },
    additional=False,
)

LINE = union(string("One paragraph"), PARAGRAPH)

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
