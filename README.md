# deskhand

MCP servers that drive live Windows desktop applications from WSL, so an agent can read and edit
what you already have open. Each application is its own server; they share only a small core.

| Server    | Status  | What it does                                                                 |
| --------- | ------- | ---------------------------------------------------------------------------- |
| `onenote` | working | Read, search and edit pages of the OneNote desktop client over COM, no Graph |
| `ppt`     | working | Read, build and render decks in the PowerPoint desktop client over COM       |

Requirements: WSL 2 with Windows interop and [uv](https://docs.astral.sh/uv/) (or Python 3.11 or newer).
The OneNote server needs the OneNote desktop client (not "OneNote for Windows 10") signed in; the PPT
server needs the PowerPoint desktop client. Nothing needs to be installed on the Windows side.

This is the Python implementation (`master`). The earlier TypeScript implementation of the OneNote
server, with the same tools and the same behaviour, is frozen on the
[`ts`](https://github.com/specter119/deskhand/tree/ts) branch.

## Install

```sh
uv tool install git+https://github.com/specter119/deskhand
deskhand --version
```

Upgrade with `uv tool install --reinstall git+https://github.com/specter119/deskhand`.

## Use

Register the server with an MCP client as a stdio command:

```json
{
  "mcpServers": {
    "onenote": { "command": "deskhand", "args": ["onenote"] },
    "ppt": { "command": "deskhand", "args": ["ppt"] }
  }
}
```

The server logs to stderr only; stdout carries JSON-RPC.

The first tool call of a server process starts one `powershell.exe` (about 2 s). That process stays
alive and answers later calls in under 100 ms; it exits when the client closes the server, or after 30
idle minutes, and the next call starts a new one. While it lives it also holds a connection to the
desktop application, so if that application was not open it keeps running in the background until the
worker exits.

A call that does not finish within its ceiling is abandoned and its worker is killed: 40 s by default,
10 s for PPT reads, 30 s for PPT writes and 55 s for `render`. A cold PPT server plus a 55 s render
can approach a client's default 60 s tool timeout, so give that client a longer `timeout` or render in
slide batches.

### OneNote tools

| Tool                                                       | Kind                                    |
| ---------------------------------------------------------- | --------------------------------------- |
| `notebooks`, `pages`, `search`, `read_page`                | read-only                               |
| `create_page`, `append_lines`, `insert_table`, `insert_image` | additive                                |
| `edit`                                                     | rewrites paragraphs in place            |
| `delete_page`                                              | moves a page to the notebook recycle bin |
| `batch`                                                    | several of the above, one token          |

A paragraph is a plain string, or an object with `text`, `depth`, `list`, `quickStyle`, `alignment`,
`spaceBefore`, `spaceAfter` and `style` (`fontFamily`, `fontSize`, `color`, `bold`, `italic`,
`underline`, `strikethrough`, `highlight`). `depth` nests a paragraph under the nearest earlier one at
the previous level, which is how a bullet tree is written. `insert_image`, `insert_table` and `edit`'s
append take `after`: the objectID of a paragraph to insert after, at any depth.

### Writes are two steps

Every write tool is a dry run unless told otherwise. A call without `dryRun=false` writes nothing.
Its reply starts with `applyToken: <token>` and shows the exact change. To apply, repeat the call
with the same arguments plus `dryRun=false` and that token.

- A token is single use and expires after 10 minutes.
- It is bound to the exact arguments, so a different request is refused.
- For `edit` and inserts into an existing outline it is also bound to the outline as previewed. The
  apply fails with `changed_since_preview` if that outline no longer matches when the write starts.
  The check runs immediately before the write but is not atomic with it, so a concurrent change in
  that window can be overwritten.
- `edit` never accepts raw OneNote XML, and `delete_page` needs the exact page name.

`batch` takes a list of ops, each one of `append_lines`, `insert_table`, `insert_image` or `edit` with
that tool's arguments, previews all of them and returns one token, then applies them in order. It is
how several changes reach one confirmation. Two ops that rewrite the same outline are refused during
the dry run, because the second fingerprint would be stale once the first is applied; several changes
inside one outline belong in one `edit` call. `create_page` and `delete_page` are not batched. A batch
is not a transaction: if a change fails, the error says what landed and what was not attempted.

The token proves that a preview was produced, not that a person read it. If you want a person in the
loop, have your client ask before any call with `dryRun=false`.

### PPT tools

| Tool                                                                                                              | Kind                       |
| ----------------------------------------------------------------------------------------------------------------- | -------------------------- |
| `presentations`, `read_deck`, `read_slide`                                                                        | read-only                  |
| `render`                                                                                                          | exports PNG/PDF files      |
| `create_deck`, `add_slide`, `add_text`, `add_table`, `add_chart`, `add_picture`, `set_notes`, `move_slide`         | additive                   |
| `update_shape`, `delete_shape`, `delete_slide`, `close_deck`                                                       | rewrites or closes         |
| `save_deck`                                                                                                       | writes a `.pptx` file      |

Positions and sizes are px at 96 dpi on a 1280x720 canvas by default. `read_deck` prints one line per
slide unless `slides` is given, in which case it expands those slides' shapes. Every PPT write follows
the same two-step flow as OneNote: the first call writes nothing and returns an `applyToken` plus a
fingerprint; repeat it with `dryRun=false`, that token and that `expectFingerprint` to apply. The
fingerprint covers the target slide's geometry, text, fill, line, table, chart and notes, so a change
made after the preview is refused with `changed_since_preview`. A path is opened read-only by a read
tool; a write tool reopens it editable when the open copy has no unsaved changes. `render` exports
PNG/PDF without touching the deck, so it does not use a token; with `overwrite=true` it replaces
existing output files and is advertised as destructive.

## Develop

```sh
uv sync
uv run pytest                 # unit, parity and transport tests; no Windows needed
uv run ruff check . && uv run ruff format --check . && uv run pyright
uv run python tests/live/onenote.py
uv run python tests/live/ppt.py
```

The live test needs a running OneNote desktop client. Put `notebook = "My Notebook"` (and optionally
`section`, default `deskhand-probe`) in `~/.config/deskhand/live.toml`, or set
`DESKHAND_LIVE_NOTEBOOK`; the variables win over the file. It writes and deletes its own pages there
and leaves the section in place.

The PPT live test needs the PowerPoint desktop client. It creates its own temporary deck, exercises
reads, writes, PNG/PDF rendering and saving, then deletes it and checks that the PowerPoint process
count is back to its starting value.

Set `DESKHAND_LOG=info` to log every bridge call and its duration to stderr. See `AGENTS.md` for the
layout, `src/deskhand/onenote/AGENTS.md` for the OneNote behaviour the bridge depends on, and
`src/deskhand/ppt/AGENTS.md` for the PowerPoint one.

## License

Apache-2.0
