# deskhand

MCP servers that drive live Windows desktop applications from WSL. One repository, one server per
application, a small shared core. Only `onenote` exists today.

## Layout

- `src/deskhand/cli.py`: `deskhand <app>` entry point; serves one app over stdio. Logs go to stderr only.
- `src/deskhand/core/`: what every app shares and nothing app-specific.
  - `powershell.py`: keeps one resident `powershell.exe` per server and sends it one op at a time through request/response files.
  - `apply_token.py`: single-use, request-bound tokens for the two-step write flow.
  - `server.py`: `ToolSpec` and `build_server`, which validate arguments and results (the MCP SDK does not).
  - `schema.py`, `compat.py`: JSON Schema builders, and the JavaScript-compatible text and number helpers.
- `src/deskhand/onenote/`: the OneNote server. Read `src/deskhand/onenote/AGENTS.md` before touching
  `bridge.ps1`.
- `tests/unit/`: run with `uv run pytest`; fake bridge, in-memory MCP client, fake resident `powershell.exe`, no Windows.
- `tests/parity/`: the scenario and the transcripts recorded from the TypeScript implementation (branch `ts`).
  `test_parity.py` replays them; a wording change in a tool reply fails it on purpose. Regenerate the
  fixtures only when the contract is meant to change. `fixtures/tools.json` is the advertised tool
  list, written by `tests/parity/dump_tools.py`; it matched the TypeScript record until paragraph
  formatting and the `after` anchor were added on purpose.
- `tests/live/`: real OneNote, run by hand with `DESKHAND_LIVE_NOTEBOOK`.
- Branch `ts` holds the frozen TypeScript implementation, with its own tests.

## Rules

- Add a new application as `src/deskhand/<app>/` with its own server; do not widen `core` for one app.
  Promote code into `core` only when a second app needs it.
- A write is never a single call: dry run first, apply with the token. Keep both layers described in
  `src/deskhand/onenote/AGENTS.md`.
- `bridge.ps1` is shipped as UTF-8 and sent to Windows PowerShell 5.1 as text, so it must parse there.
  Check it after editing, and run the live test. It runs inside a resident worker, so it must never `exit`.
- Run `uv run ruff check .`, `uv run ruff format --check .`, `uv run pyright` and `uv run pytest` before
  committing. Do not commit `uv.lock`: it records the local package index URL.
- Install the local copy with `uv tool install <repo>` (not `--editable`) and `--reinstall` after a change;
  an editable install breaks `deskhand` whenever the working copy is on another branch.
- Python code that formats numbers or text for users goes through `core/compat.py`: JavaScript and Python
  differ on `36.0`, rounding of halves and truthiness of `[]`, and clients saw the JavaScript results first.
- Keep personal data out of the repository: no notebook names, account names, paths under a home
  directory or work-related identifiers in code, tests, docs or commit messages.
