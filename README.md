# deskhand

MCP servers that drive live Windows desktop applications from WSL, so an agent can read and edit
what you already have open. Each application is its own server; they share only a small core.

| Server    | Status  | What it does                                                                 |
| --------- | ------- | ---------------------------------------------------------------------------- |
| `onenote` | working | Read, search and edit pages of the OneNote desktop client over COM, no Graph |

Requirements: WSL 2 with Windows interop, [uv](https://docs.astral.sh/uv/) (or Python 3.11 or newer), and
the OneNote desktop client (not "OneNote for Windows 10") signed in on the Windows side. Nothing needs
to be installed on the Windows side.

This is the Python implementation (`master`). The earlier TypeScript implementation, with the same tools
and the same behaviour, is frozen on the [`ts`](https://github.com/specter119/deskhand/tree/ts) branch.

## Install

```sh
uv tool install git+https://github.com/specter119/deskhand
deskhand --version
```

Upgrade with `uv tool install --reinstall git+https://github.com/specter119/deskhand`.

## Use

Register the server with an MCP client as a stdio command:

```json
{ "mcpServers": { "onenote": { "command": "deskhand", "args": ["onenote"], "timeout": 120 } } }
```

The server logs to stderr only; stdout carries JSON-RPC.

The first tool call of a server process starts one `powershell.exe` (about 2 s). That process stays
alive and answers later calls in under 100 ms; it exits when the client closes the server, or after 30
idle minutes, and the next call starts a new one. While it lives it also holds a connection to
OneNote, so if OneNote was not open it keeps running in the background until the worker exits.

### OneNote tools

| Tool                                                       | Kind                                    |
| ---------------------------------------------------------- | --------------------------------------- |
| `notebooks`, `pages`, `search`, `read_page`                | read-only                               |
| `create_page`, `append_lines`, `insert_table`, `insert_image` | additive                                |
| `edit`                                                     | rewrites paragraphs in place            |
| `delete_page`                                              | moves a page to the notebook recycle bin |

### Writes are two steps

Every write tool is a dry run unless told otherwise. A call without `dryRun=false` writes nothing.
Its reply starts with `applyToken: <token>` and shows the exact change. To apply, repeat the call
with the same arguments plus `dryRun=false` and that token.

- A token is single use and expires after 10 minutes.
- It is bound to the exact arguments, so a different request is refused.
- For `edit` and inserts into an existing outline it is also bound to the outline as previewed. If
  that outline changed in OneNote after the dry run, the apply fails with `changed_since_preview`.
- `edit` never accepts raw OneNote XML, and `delete_page` needs the exact page name.

The token proves that a preview was produced, not that a person read it. If you want a person in the
loop, have your client ask before any call with `dryRun=false`.

## Develop

```sh
uv sync
uv run pytest                 # unit, parity and transport tests; no Windows needed
uv run ruff check . && uv run ruff format --check . && uv run pyright
DESKHAND_LIVE_NOTEBOOK="My Notebook" uv run python tests/live/onenote.py
```

The live test needs a running OneNote desktop client. It creates a section called `deskhand-probe`
in the notebook you name, writes and deletes its own pages there, and leaves the section in place.

Set `DESKHAND_LOG=info` to log every bridge call and its duration to stderr. See `AGENTS.md` for the
layout and `src/deskhand/onenote/AGENTS.md` for the OneNote behaviour the bridge depends on.

## License

Apache-2.0
