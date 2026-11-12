# Paragraph formatting and placement in the OneNote write tools

Status: closed 2026-11-12. Observed against deskhand 0.3.0 (xs2013 / OneNote 2013 schema) while
writing a new page that had to match an existing weekly series in the same section; the two gaps below
are now covered by the API.

## What was missing

1. **No paragraph formatting.** `append_lines` and `create_page` took `lines: string[]`; `edit` took
   `text: string`. There was no way to set font family or size, bold/italic/underline, colour,
   alignment, `spaceBefore`, or `quickStyleIndex`. Every new paragraph landed on the default `p`
   quick style (Calibri 11 pt), at depth 0, with no list.
2. **No paragraph-level anchor.** `insert_image` and `insert_table` accepted `outline` only, so a
   block could only be appended after the last paragraph of an outline or as a new outline at the end
   of the page. Inserting after paragraph 3 of an existing outline was impossible.

Not gaps: cropping is the caller's job (`insert_image` embeds the file it is given, so a pre-cropped
PNG is the intended input), and table cell text has no per-cell formatting.

## What a caller can do now

`create_page`, `append_lines`, and `edit`'s `append[].lines` take `lines: (string | Paragraph)[]`:

```json
{
  "text": "…",
  "depth": 2,
  "list": "bullet",
  "quickStyle": "p",
  "alignment": "left",
  "spaceBefore": 5.0,
  "spaceAfter": 0.0,
  "style": {
    "fontFamily": "Calibri",
    "fontSize": 19,
    "color": "#2E75B5",
    "bold": false,
    "italic": false,
    "underline": false,
    "strikethrough": false,
    "highlight": "#FFFF00"
  }
}
```

- A bare string keeps the old behaviour and still produces exactly the XML it produced before.
- `depth` describes the tree as a flat list: a paragraph at depth d nests under the nearest earlier
  paragraph at depth d-1, and a level can only be entered from the level above it.
- `list` is `bullet`, `number` or `none`.
- `quickStyle` is an index or a name the page defines (`read_page` reports both). OneNote maps the
  index onto the style that best matches the paragraph and may normalise it, so a name is a hint and
  `read_page` is the source of truth. Explicit `style` is what makes a heading look like a heading.
- `style` is an object, not a CSS string: the server validates it and writes it where OneNote writes
  it itself (font family, size and colour on the `one:OE` style attribute; bold, italic, underline,
  strike and highlight in a span inside the CDATA).

`insert_image`, `insert_table` and `edit`'s `append[]` take `after`: the objectID of a paragraph, at
any depth. The block becomes that paragraph's next sibling, and the outline is derived from the anchor
when `outline` is omitted. `read_page` now also reports each paragraph's `style` when it has one.

## Acceptance checks, as run

1. A 3-level bullet list at 19 pt reads back with `depth` 0/1/2, `list` `bullet`/`bullet`/`number`
   and `style` containing `19.0pt` — `tests/live/onenote.py`, "append_lines writes depth, list, quick
   style and font size".
2. An image inserted with `after` lands immediately after that paragraph, not at the end of the page —
   same file, "insert_image with after lands right after that paragraph".
3. A bare-string `append_lines` call sends the same request and produces the same XML as before —
   `tests/unit/test_onenote.py`, "a plain string stays a plain string…", plus the parity transcripts.

## Evidence

- A page written by another author in the same section (read back through `read_page`):
  40 paragraphs, depths `{0:2, 1:6, 2:23, 3:1, 4:6, 5:2}`, 33 `bullet` paragraphs, per-`OE`
  `style="font-family:...;font-size:19.0pt"`, and
  `<one:List><one:Bullet bullet="2" fontSize="11.0"/></one:List>` on body bullets. Images sit inline
  at depth 2 between bullets.
- A page written through `append_lines` + `insert_image` before this change: 38 paragraphs, all depth
  0, no `list`, no `style` attribute, no `quickStyleIndex` beyond the default; the images were
  appended after the last paragraph.
- OneNote rewrites `quickStyleIndex` on write: a paragraph whose formatting does not match a defined
  style comes back on the default index. Verified against the live client, which is why the API
  documents the index as a hint.
- `numberFormat` on `one:Number` is a pattern (`"##."`), not a name, and `text` is the rendered
  number OneNote starts from (`"1."`). Anything else is refused with 0x80042001.

## Reload after a change

`uv tool install --reinstall <repo>` replaces the package on disk, but the running MCP server keeps
the old code in its resident worker. The client must reconnect the server (`/mcp reconnect onenote`)
or restart the session before the new schema is visible.
