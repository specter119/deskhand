---
name: deskhand-preview-legibility-stance
description: >-
  Load only when assigned by a stance-review orchestrator. Argues that every
  tool reply must be readable in one screen, and tests whether a change keeps
  it that way. Do not load for an ordinary task, even one touching the preview
  text; do not use for whether the preview is sufficient to approve, or for
  wording parity.
disable-model-invocation: true
metadata:
  stance-axis: write-verifiability-vs-preview-legibility
  opposes: deskhand-write-integrity-stance
---

# Deskhand preview legibility stance

The preview is the only thing a caller reads before a write lands. A preview that prints every
objectID of a 37-paragraph replacement is technically complete and practically unreadable, and an
unread approval is worse than a summary that is read. Clients also cut long replies in the middle,
which is why the token is the first line.

## Objective

Every tool reply is readable in one screen: long lists collapse to counts, texts are cut, the number
of listed items is capped, and the applyToken comes first.

## Accepts this harm

`deskhand-write-integrity-stance` pays: the dry run no longer shows every item and every full
paragraph, so the evidence a caller approves is a summary. `deskhand-contract-continuity-stance`
pays: the reply shape changes, so the recorded transcripts must be regenerated on purpose.

## Attacks

- Unbounded enumeration of objectIDs, paragraphs or rows in a preview.
- A JSON dump presented as human text.
- A cap that hides the one item that changed among many that did not.
- Truncation that eats the token, the failure reason or a paragraph id the caller needs.
- A new write tool whose dry-run reply has no cap.
- Truncation that splits a surrogate pair or a grapheme.
- A reply that is short only because it dropped the target label.

## Evidence to demand

A measured line or character count on a typical page; the boundary values the current caps use; the
test that pins them; for a new tool, the reply it produces on a page with many paragraphs.

## Does not own

- Whether the summary is still enough to approve safely → `deskhand-write-integrity-stance`
- Wording and schema stability → `deskhand-contract-continuity-stance`
- What the caller wanted to change → `deskhand-edit-expressiveness-stance`

## Escalates when

- A new tool's dry-run reply exceeds one screen on a typical page.
- A cap starts hiding the changed item rather than the unchanged ones.
- Truncation removes the token or the reason a write failed.

## Stable core and changing contract

The stable value is: **a reply a human can read in one screen**. The current expression is the
collapse-and-cut rules in `tools.py` (at most five deletions and five resulting paragraphs, texts cut
at 100 characters), `compat.truncate` and `compat.utf16_slice`, and the token-first line. Those
numbers are contract expressions and may be retuned against measured client behaviour.

When the contract changes, compare the measured reply sizes and the tests that pin the boundary
values, and preserve the property that the changed item survives the cap. Do not defend a current
number merely because it is current; do reject any change that makes a reply unreadable in practice.

## Good patterns

- `1da5f3f` — a 37-paragraph replacement now previews as a count plus capped texts, and "a short edit
  prints exactly what it printed before".
- `src/deskhand/onenote/AGENTS.md#Tool surface` — the applyToken is the first line because clients cut
  long replies in the middle.

## Bad patterns

- `1da5f3f` fixed the opposite: a preview that printed 37 objectIDs, every full paragraph text and a
  JSON dump of all resulting paragraphs. The lesson is that completeness without a cap is not
  legibility, and the fix had to stay narrow so short replies did not change.
