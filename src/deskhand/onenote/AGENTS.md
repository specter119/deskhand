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
| `bridge.ps1`             | The whole COM layer, run as a one-shot `powershell.exe` process per call      |
| `../core/powershell.py`  | Transport: request/response files, bootstrap, timeouts, `outcome_unknown`     |

**Why one process per call instead of a resident worker:** a warm
`New-Object -ComObject OneNote.Application` costs **6-11 ms** (cold: 226 ms), so PowerShell process
startup dominates and a resident worker with its own lifecycle buys almost nothing. Measured, not
assumed.

A whole call measured about **1.6 s** at p50 (`notebooks`, 8 calls, warm; the first call 2.1 s). A bare
`powershell.exe -NoProfile` start is about 0.5 s, so the rest is parsing the 900-line script and the COM
work. `DESKHAND_LOG=info` logs each call's duration to stderr. If this becomes a problem, the answer is a
resident Windows-side worker, not tuning the Linux side.

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
  which destroys every CJK literal. `core/powershell.py` reads `bridge.ps1` as UTF-8 and hands it to a
  ~200-character `-EncodedCommand` bootstrap _through the request file_.
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
- **The child must not inherit the server's stdin.** `powershell.exe` is started with `stdin=DEVNULL`;
  otherwise it shares the MCP connection's input and can swallow JSON-RPC.
- **Cancelling and timing out are different.** A client cancel kills the child and propagates the
  cancellation. Only the 180 s timeout becomes `bridge_failed`, or `outcome_unknown` for an apply.
- **Process lifetime (measured).** Killing the WSL-side `powershell.exe` stub, which is what a timeout or
  cancel does, also ends the Windows process. SIGKILL of the _server_ does not: the Windows process then
  runs to completion on its own, so an in-flight write can still land after the server is gone.
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
UpdatePageContent(string xml, DateTime dateExpectedLastModified, XMLSchema xsSchema, out string error)
```

- `xsSchema` must be **`2`** (xs2013) to match the namespace this bridge writes. `0` and `1` fail
  with `0x80042001`.
- Naming all four arguments keeps the error channel. PowerShell accepts fewer, but then OneNote's
  own refusals are invisible.
- `dateExpectedLastModified` is always `MinValue` ("do not check") — see the next section.

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
  against a fake `powershell.exe`, and replays `tests/parity/scenario.json` against the transcripts recorded
  from the TypeScript implementation. They need neither Windows nor OneNote.
- `DESKHAND_LIVE_NOTEBOOK="<notebook>" uv run python tests/live/onenote.py` starts two real server processes against
  the OneNote desktop client. It creates the section `deskhand-probe` if it is missing and leaves it
  in place, and it deletes every page it created. Run it from WSL after any change to `bridge.ps1`
  or the transport; it covers the lifecycle, token refusal, `changed_since_preview`, tables, images
  and two processes working at once.
