# OneNote server

Development notes for the `onenote` MCP server (`src/deskhand/onenote`): what it may and may not do, and the
OneNote behaviour it depends on. The repository-level `AGENTS.md` covers the rest of deskhand.

Drives the **Microsoft OneNote desktop application on the Windows host** from a process running
inside WSL, over the OneNote COM automation interface.

Graph was rejected on purpose: it needs an app registration and a write scope (the available one
has none), and it cannot see the user's already-signed-in local notebooks. COM needs no API scope,
no admin consent and no network, and it writes through the same client the user is already using,
so the change syncs like any other edit.

It is also **silent**: the COM proxy starts OneNote without a visible window and never takes focus,
so it does not disturb whatever the user is doing on the Windows desktop.

## Invariants

The rest of this file explains _why_, but these are the rules to keep:

1. Never send an `objectID` on new content; let OneNote assign it.
2. Never hand-build outline XML. Edit a whole outline through `edit`, which echoes back OneNote's own
   XML for everything it is not changing. `edit` refuses callers that send raw `xml`.
3. Never put a `lastModified` comparison back into a write path.
4. Keep the transport exactly as it is: no `-File`, no environment variables for paths, no
   namespace manager, no `[string]` on parameters that receive `$null`.
5. Every write is default-deny, in two layers. The bridge writes only with `apply=true` **and** the
   `expectFingerprint` of a preview (`apply_requires_fingerprint`, `changed_since_preview`). The
   server hands out that fingerprint only inside a single-use, 10-minute `applyToken` that is bound
   to the exact request (`core/apply_token.py`). Do not add a path that skips either layer.

## Files

| File                     | Role                                                                          |
| ------------------------ | ----------------------------------------------------------------------------- |
| `tools.py`               | The ten tools, the dry-run/apply flow, image file reading and base64 encoding |
| `schemas.py`             | Input and output schemas (loose: OneNote may add fields)                      |
| `format.py`              | Text rendering and normalisation of bridge results (PowerShell array quirks)  |
| `bridge.ps1`             | The whole COM layer, run by one resident `powershell.exe` per server process  |
| `../core/powershell.py`  | Transport: worker lifecycle, request/response files, timeouts, `outcome_unknown` |

**Why a resident worker:** a call used to start a fresh `powershell.exe`, which cost about 0.5 s for the bare
process, about 0.3 s for `Invoke-Expression` of this 900-line file and the first `ConvertFrom-Json`, and the
process exit. A whole call measured about **1.6 s** at p50. The COM work itself is small: a fresh
`New-Object -ComObject OneNote.Application` costs 15-22 ms and `GetHierarchy` 50-80 ms. So each server process
now starts **one** `powershell.exe` on its first call, loads this file once and keeps it, together with the OneNote COM object (see below). Measured with the live
suite on the development machine: the first call takes about **2.1 s** (temp-directory probe, worker start of
about 1.4 s, then the call), and later calls take about **70-100 ms** (`notebooks` p50 94 ms, `read_page` of an
11-paragraph page p50 67 ms). Two clients are two server processes and so two workers; OneNote's COM server
serialises them. A Windows-side Python or a TCP/named-pipe tunnel was rejected: the WSL interop stdin/stdout of
`powershell.exe` is already the tunnel, and per-process start-up was never the language's fault.
`DESKHAND_LOG=info` logs each call's duration and whether it started the worker.

### Worker protocol

`core/powershell.py` writes `bridge.ps1` to `worker-<server pid>-<uuid>.ps1` in the work directory and starts
`powershell.exe -EncodedCommand` with a bootstrap that `Invoke-Expression`s that file. The file ends with the
worker loop, so the bootstrap never returns. The Python side deletes the file once the worker is ready.

| Direction | Content |
| --------- | ------- |
| worker stdout | `deskhand-ready 1 <windows pid>` once, then `deskhand-done <stem>` after each call |
| worker stdin  | one call stem per line (`[0-9A-Za-z-]+`); EOF ends the worker |
| files         | `req-<stem>.json` in, `resp-<stem>.json` out, UTF-8, in the work directory |

Rules that keep it correct:

- **A worker is reused only after the `deskhand-done` line of the current stem.** Every other outcome
  (timeout, cancellation, EOF, a protocol error) kills the worker and forgets it before the first `await`: the MCP
  SDK cancels handlers with an anyio scope that re-raises at every `await`, so cleanup cannot wait.
- A stem that could not be written (broken pipe) is sent again to a fresh worker, apply included, because the
  worker never saw it. A worker that exits after the stem was sent gives `bridge_failed`, or `outcome_unknown`
  for an apply. Reads are not retried.
- Calls run one at a time (an `asyncio.Lock`). The 40 s timeout counts from sending the stem, not from waiting
  for the lock. It is below the 60 s default tool timeout of Pi (the strictest client in use), and startup (15 s) plus
  the call stays under 60 s on a cold first call. Do not raise it without raising the client timeout.
- A worker that has been idle for 30 minutes is closed by closing its stdin; the next call starts a new one.
  Closing the server (the MCP lifespan) does the same, killing it after 2 s.
- **`bridge.ps1` must never `exit`.** An exit between writing a response and printing `deskhand-done` would make a
  finished write look lost. Stop a call with `Fail` (writes the error response, then throws) or `return`.
  `tests/unit/test_bridge_script.py` enforces this and checks that the file parses in Windows PowerShell 5.1.
- `Fail` is control flow. The catch in `Invoke-Request` tells it apart from a real error by the
  `$script:RESPONDED` flag, never by the message: an exception thrown through a .NET callback arrives wrapped.
  The first response written is final.
- PowerShell variables are dynamically scoped: an unassigned name resolves in the worker's globals. `Invoke-Request`
  therefore starts with `$onenote = $null`, and the worker loop's own variables are prefixed `$__dh`.
  `break`/`continue` also propagate up the call stack to the nearest enclosing loop, so do not use them outside
  a loop in the same function.
- **The OneNote COM object is kept for the life of the worker** (`Get-OneNote`). When the user has not opened
  OneNote, COM starts `ONENOTE.EXE` for the first object and ends it as soon as the last one is released
  (measured: five create/release cycles, five restarts, 0.9-1.0 s each; with OneNote already open the same call
  takes 15-22 ms). Releasing per call therefore cost about 0.9 s per call and sometimes raced the shutdown with
  `0x800706BF` ("remote procedure call failed") on the next `New-Object`. Each call probes the kept object with
  `GetHierarchy('', 0)` and builds a new one when the probe throws, which covers OneNote being closed or
  crashing between calls. The probe runs before any write is sent. Side effect: while a worker lives, OneNote
  keeps running in the background, up to 30 idle minutes after the last call.
- stdout carries only the protocol. The worker loop runs each call with `*> $null`, so a stray warning cannot
  splice into a protocol line; Python also ignores any line that is not `deskhand-done <current stem>`.

## Tool surface

Tools are plain MCP tools named `notebooks`, `pages`, `search`, `read_page`, `create_page`,
`append_lines`, `insert_table`, `insert_image`, `edit` and `delete_page`. Each returns text for the
transcript and `structuredContent` for scripts, validated against an output schema. Clients such as
Pi prefix the names (`mcp__onenote__edit`).

| Tool                                                          | Kind                                      |
| ------------------------------------------------------------- | ----------------------------------------- |
| `notebooks`, `pages`, `search`, `read_page`                   | read-only                                 |
| `create_page`, `append_lines`, `insert_table`, `insert_image` | additive                                  |
| `edit`                                                        | rewrites paragraphs in place, destructive |
| `delete_page`                                                 | destructive, goes to the recycle bin      |

Reads are free. A write called without `dryRun=false` is a real bridge call that reports the exact
change set OneNote would receive. Its reply starts with `applyToken: <token>` (long replies get cut
in the middle by some clients, so the token goes first). The apply repeats the call with
`dryRun=false` and that token. `create_page` can only check that the section exists, and
`delete_page` shows the page it would remove.

The fingerprint of `edit` and of inserts into an existing outline covers the outline text that
would be written, so a concurrent change to a different outline does not invalidate a token, and
neither does a change that leaves the written result identical. The fingerprint of `append_lines`
and of inserts as a new outline only ties the token to the page: those writes are additive and
cannot clobber anything.

A client that wants a human in the loop should gate the apply call (`dryRun === false`); the
`applyToken` already proves a preview was produced, not that anyone looked at it. Pi does this in
the dotfiles repository (`pi/extensions/deskhand`).

## Transport constraints

These are not stylistic. Each one cost real debugging time and must stay as it is.

- **The script is never run with `-File`.** Windows PowerShell 5.1 decodes a BOM-less file as ANSI,
  which destroys every CJK literal. `core/powershell.py` writes `bridge.ps1` as UTF-8 to `worker-<id>.ps1` and
  hands it to a ~200-character `-EncodedCommand` bootstrap that reads it back with an explicit encoding.
- **Paths never ride in environment variables.** WSL interop does not forward the Linux environment
  to Windows executables and ignores `WSLENV` in that direction, so the bootstrap carries the work
  directory instead.
- **The bridge cannot be sent through `-EncodedCommand` itself.** A Windows command line caps out
  near 32 K characters, and the script is longer than that once encoded.
- **Payloads travel through UTF-8 files, not arguments.** Page XML reaches 260 KB and image base64
  reached 5.4 MB in testing. Each call uses its own request/response pair, so parallel tool calls
  cannot clobber each other's files.
- **XPath uses `local-name()` instead of a namespace manager.** `XmlNameTable` **and**
  `XmlNamespaceManager` implement `IEnumerable`, so passing or returning one makes PowerShell
  silently unroll it into namespace prefixes and the namespace lookup fails with a confusing
  "cannot find an overload for SelectNodes".
- **Request fields are read through `Get-Field`, not property access.** A hashtable's property
  access falls through to the type's members, so an absent field named `remove` returns
  `Hashtable.Remove`'s overload list instead of `$null`. This produced an error message containing
  a method signature.
- **The worker never inherits the server's stdin.** It gets its own pipe (`stdin=PIPE`); sharing the MCP
  connection's input would let it swallow JSON-RPC.
- **Cancelling and timing out are different.** A client cancel kills the worker and propagates the
  cancellation. Only the 40 s timeout becomes `bridge_failed`, or `outcome_unknown` for an apply. Either way
  the next call starts a new worker.
- **Process lifetime (measured).** Killing the WSL-side `powershell.exe` stub, which is what a timeout or
  cancel does, also ends the Windows process. When the _server_ is SIGKILLed, the worker's stdin reaches EOF
  and the Windows process leaves on its own: at once when idle, and after the call in flight when busy. So an
  in-flight write can still land after the server is gone. `tests/live/onenote.py` checks that no worker is left.
- **Function parameters that must accept `$null` are left untyped.** Declaring `[string]$newText`
  makes PowerShell convert a `$null` argument into `""`, which silently turned "leave this paragraph
  alone" into "replace its text with nothing".

## Writing into OneNote

All of the following is verified against the live client.

### objectID

New elements must be written **without** an `objectID`. OneNote rejects an id it did not issue with
`0x8004200E` (`hrPageObjectDoesNotExist`), and assigns ids itself for anything sent without one.
`ID` on `one:Outline` is not in the schema at all — the attribute is `objectID`.

### UpdatePageContent

The real signature is:

```
UpdatePageContent(string xml, DateTime dateExpectedLastModified, XMLSchema xsSchema, bool force)
```

- `xsSchema` must be **`2`** (xs2013) to match the namespace this bridge writes. `0` and `1` fail
  with `0x80042001`.
- The fourth argument is `force`, **not** an error output (confirmed from the OneNote 15.0 type library).
  It is always `$false`: with `$true` OneNote overwrites page data it does not understand. A refusal arrives
  as an exception, never through a return value.
- `dateExpectedLastModified` is always `MinValue` ("do not check") — see the next section.

### Failures around a write

Writes are the three COM calls `UpdatePageContent`, `CreateNewPage` and `DeleteHierarchy`. A refusal is only
called a refusal when OneNote is known to have refused; anything else may have landed.

| Failure                                                          | Code               |
| ---------------------------------------------------------------- | ------------------ |
| no COM call was made (PowerShell conversion or binding error)    | `com_error`        |
| `0x80010001`, `0x8001010A` (turned away while busy), `0x80042024` (recording)  | `busy`             |
| `0x80042010`                                                     | `stale`            |
| any other `0x800420xx` except `0x80042023`                       | `rejected`         |
| connection lost (`0x800706BA/BE/BF`, ...), `0x80042023` (OneNote's own timeout), anything unrecognised | `outcome_unknown`  |
| the write call returned, then something else failed              | `applied_unverified` |

`applied_unverified` is decided by `$script:MUTATED`, set right after a write call returns: `Fail` and the catch in
`Invoke-Request` then rewrite whatever code they were given. This covers `create_page` (three writes: the page, the
title, the body) and a read-back that fails after a successful write. On the Python side a write (`apply`) that
leaves no readable response, or an over-long line, is also `outcome_unknown`.
HRESULTs are found by walking the exception chain past `MethodInvocationException` (`Get-HResultHex`) and compared
as upper-case strings. `tests/unit/test_bridge_behaviour.py` drives all of this against a fake OneNote object in
Windows PowerShell 5.1.

The kept OneNote object is probed before each call. A busy answer (`0x80010001`, `0x8001010A`) is waited out
for up to 5 s and then reported as `busy`, without discarding the object; any other probe failure drops the object
and creates a new one, retrying once after 1 s because starting OneNote can race its previous shutdown.

### The edit model

`one:OEChildren` is **authoritative for its children**: the list that is sent becomes the complete
child list, so a paragraph omitted from it is **deleted**. Meanwhile page elements that are not
mentioned at all (the title, other outlines) are preserved. Verified by experiment:

| Action                                                              | Result                                                                       |
| ------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Resend an outline with all paragraphs by objectID, one text changed | text changed in place, objectID kept, other outlines and the title untouched |
| Omit a paragraph in that resend                                     | the omitted paragraph is deleted                                             |
| Send a new paragraph with no objectID                               | added, OneNote assigns its id                                                |
| Omit `Position`/`Size` for an existing outline                      | geometry preserved                                                           |
| Send an empty `one:OEChildren`                                      | rejected: "cannot be empty according to the DTD/Schema"                      |

Therefore **the safe unit of edit is one whole outline, resent in full**, which is what
`edit` does: it groups the requested changes by owning outline, echoes back OneNote's own
XML (`OuterXml`) for every paragraph it is not changing — so styles, lists, tables and images
survive — and substitutes only the `<one:T>` of the paragraphs being edited. Paragraph XML must
never be hand-built by a caller.

### Rich text

`one:T` holds an **HTML-ish fragment inside CDATA**, with entities still escaped even inside CDATA.
Consequences both ways:

- **Reading** needs tag stripping _and_ entity decoding, plus a manual numeric-reference pass:
  .NET's `HtmlDecode` silently skips astral-plane references such as `&#128640;`, so emoji would
  otherwise come back as literal text.
- **Writing** must escape `&`, `<` and `>` first (`New-TextCdata`). Otherwise a literal
  `<angles>` is parsed as a tag and disappears. OneNote does the same itself: `->` is stored as
  `-&gt;`.

### Tables

`one:Table` is a child of **`one:OE`**, never of `one:Outline`. `Table` is `Columns?` then `Row+`,
`one:Column` requires both `index` and `width` (so `Columns` is emitted only when widths are given),
and every `one:Cell` **must** contain an `OEChildren`. There is no `columnSpan`/`rowSpan` in the
schema at all.

### Images

`one:Data` with base64 is the supported import path — no callback object is involved. `format` must
match `auto|png|emf|jpg`, so JPEG is spelled **`jpg`**. Reading back never returns `Data`: OneNote
replaces it with a `one:CallbackID`, which is the evidence that the bits are really stored (the
model reports this as `embedded`). `format` is not echoed back. Without `width`/`height` OneNote
chooses the size. Verified working up to **5.4 MB of base64**.

## `lastModifiedTime` is not a concurrency token

This dictates the write design, and it cost two discarded implementations:

- It **does not advance when content is written through COM.** Three appends in a row left it at the
  page's creation second.
- `GetHierarchy` can report a **different value** than `GetPageContent` for the same page.

So comparing it refuses legitimate writes while detecting nothing. Both guards built on it were
removed during development: one refused a correct append with `stale` even though nothing had
changed, and feeding back the value just read protects nothing. Instead:

- appends are additive and cannot clobber a concurrent edit, so their fingerprint only names the page;
- `delete_page` guards on the **exact page name** (`expectName` is required);
- `lastModified` is reported as informational only.

## Not implemented

- Rewriting a paragraph **in place has since been solved** (`edit`); what remains missing is
  deleting a whole outline, and editing inside table cells by position (cells are reachable as
  ordinary paragraphs through their own objectIDs).
- Ink, `one:HTMLBlock` (xs2013 rich HTML fragments) and page-level `one:Image` positioning.
- The tools are OneNote-only. Anything outside COM's reach (Copilot pane, ink) would need a UIA layer.

## Testing

- `uv run pytest` runs the unit tests against a fake bridge through an in-memory MCP client, the transport
  against a fake resident `powershell.exe` that speaks the worker protocol, static checks on `bridge.ps1`, and replays `tests/parity/scenario.json` against the transcripts recorded
  from the TypeScript implementation. They need neither Windows nor OneNote.
- `DESKHAND_LIVE_NOTEBOOK="<notebook>" uv run python tests/live/onenote.py` starts two real server processes against
  the OneNote desktop client. It creates the section `deskhand-probe` if it is missing and leaves it
  in place, and it deletes every page it created. Run it from WSL after any change to `bridge.ps1`
  or the transport; it covers the lifecycle, token refusal, `changed_since_preview`, tables, images,
  two processes working at once, a worker reconnecting after OneNote exits (only when the run started OneNote itself), one worker per server that survives every kind of bridge error, no worker left
  after a SIGKILL of the server, and the latency gates (first call <= 2.5 s, or <= 4 s when the run had to start OneNote; hot `notebooks` p50 <= 200 ms, hot
  `read_page` p50 <= 300 ms; it prints the measured numbers).
