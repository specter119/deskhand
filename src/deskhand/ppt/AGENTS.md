# PowerPoint server

`bridge.ps1` drives the PowerPoint desktop client through COM. `tools.py` builds each request, validates
arguments, and owns the two-step write gate; `format.py` owns every reply string a client reads.

## bridge.ps1

- It is shipped as UTF-8, sent to Windows PowerShell 5.1 as text, and runs for the whole life of the
  server process. It must parse on 5.1, stay ASCII, and never call `exit`: an `exit` between writing a
  response and printing `deskhand-done` makes the caller treat a finished call as a dead worker.
  `tests/unit/test_bridge_script.py` checks all three rules for every `bridge.ps1`.
- One resident `powershell.exe` handles one op at a time. PowerPoint COM is STA; do not add concurrency.
- The API speaks px at 96 dpi. Convert only at the COM boundary with `PxToPt` (`px * 0.75`) and
  `PtToPx` (`pt * 4 / 3`, rounded to 0.1 px). Colors are `#RRGGBB` in the API and `0xBBGGRR` in COM;
  `ConvertTo-Color` is the only place that converts.
- `Get-Ppt` checks the cached COM object is still alive and rebuilds it when it is not. A dead
  PowerPoint must not turn every later call into the same error.
- Process ownership: the server quits PowerPoint only when it started it and no presentation is open.
  It never closes a presentation it did not open unless the caller names it. `Reset-Ppt` closes only
  `$script:OURS` and restores nothing else.
- Charts need a presentation window. `Ensure-ChartWindow` creates a window and minimizes it; do not
  replace chart data through `ChartData.Workbook` (it starts Excel and fails without a window). Write
  `Series.Values` and `Series.XValues` directly.
- `Chart` has no `HasDataLabels` in this PowerPoint version. Turn labels on per series
  (`Series.HasDataLabels`) and set `Point.DataLabel.ShowValue`/`NumberFormat` per point.
- The API speaks px, but PowerPoint's minimum font size is 1 pt (1.34 px). `tools.py` refuses smaller
  run/table sizes before a token is issued; do not let `Font.Size` fail after a partial write.
- PDF export uses `SaveCopyAs(path, 32)` so the open deck keeps its own path. PNG uses
  `Slide.Export`. `overwrite` is the caller's decision; the bridge refuses an existing path otherwise.
- Dialogs are off: `DisplayAlerts = 1` (ppAlertsNone) and `AutomationSecurity = 3`.
- A path is opened read-only by a read op. A write op asks `Get-Presentation` for editable; if the open
  copy is read-only and unsaved, it refuses with `read_only`, and otherwise closes and reopens it
  editable without a window. Wait for the closed copy to leave `Presentations` before reopening, and
  verify the reopened copy is editable; PowerPoint can otherwise hand back the read-only object.
  `save_deck` may keep a genuinely read-only copy to save to a different path, but refuses an in-place
  save with `read_only` instead of reporting `applied_unverified`.

## Write gate

- Every write is a dry run unless `dryRun` is exactly `false`. A preview issues an `applyToken` and
  returns the fingerprint of the state it saw; the token is single-use, bound to the exact request
  body, and expires after 10 minutes.
- The bridge recomputes the fingerprint and calls `Assert-Fingerprint` immediately before the first
  mutation. `Get-SlideFingerprint` covers layout, each shape's id/name/type/geometry/text/fill/line,
  tables, charts, linked picture sources and notes; `Get-DeckFingerprint` adds every slide and its
  order.
  `save_deck` binds both the target file (length + mtime, or absent) and the in-memory deck.
- Set `$script:MUTATED = $true` before the first call that can change anything, including `Save`,
  `SaveAs` and `AddPicture`. After that point a failure is `applied_unverified`, never "nothing was
  written". Before it, a refusal is `invalid_request`, `file_exists`, `changed_since_preview` and so on.
- `tools.py` must not turn a missing post-write fingerprint into success. If an apply reply has no
  fingerprint, the outcome is unknown.

## Timeouts

- Reads use 10 s, writes 30 s, `render` 55 s. The default 40 s core ceiling is too short for a cold
  PowerPoint plus a whole-deck export, and too long to make a stuck read wait. `render` says its
  ceiling in its description; a larger deck is rendered in slide batches.

## Replies

- Preview text starts with `applyToken:` on line 1 and `fingerprint:` on line 2. The rest is bounded:
  five shapes per slide, a target shape is always shown, a deck overview is one line per slide, and a
  render lists five paths. Full data stays in `structuredContent`.
- Truncation counts UTF-16 units and backs off to a whole grapheme cluster, matching the OneNote
  formatter rather than Python's code-point slicing.
