# deskhand

MCP servers that drive live Windows desktop applications from WSL. One repository, one server per
application, a small shared core. Only `onenote` exists today.

## Layout

- `src/cli.ts`: `deskhand <app>` entry point; serves one app over stdio. Logs go to stderr only.
- `src/core/`: what every app shares and nothing app-specific.
  - `powershell.ts`: runs one op per `powershell.exe` process through request/response files.
  - `apply-token.ts`: single-use, request-bound tokens for the two-step write flow.
  - `server.ts`: `defineTool`, which wires a TypeBox schema pair and a handler into the MCP SDK.
- `src/onenote/`: the OneNote server. Read `src/onenote/AGENTS.md` before touching `bridge.ps1`.
- `tests/unit/`: run with `npm test`; fake bridge, in-memory MCP client, no Windows.
- `tests/live/`: real OneNote, run by hand with `DESKHAND_LIVE_NOTEBOOK`.
- `dist/`: committed build output. a global install from the GitHub tarball runs no build.

## Rules

- Add a new application as `src/<app>/` with its own server; do not widen `src/core` for one app.
  Promote code into `core` only when a second app needs it.
- A write is never a single call: dry run first, apply with the token. Keep both layers described in
  `src/onenote/AGENTS.md`.
- `bridge.ps1` is shipped as UTF-8 and sent to Windows PowerShell 5.1 as text, so it must parse there.
  Check it after editing, and run the live test.
- Run `npm run typecheck`, `npm run lint`, `npm test` and `npm run build` before committing, and
  commit the rebuilt `dist/`.
- Keep personal data out of the repository: no notebook names, account names, paths under a home
  directory or work-related identifiers in code, tests, docs or commit messages.
