---
name: deskhand-write-integrity-stance
description: >-
  Load only when assigned by a stance-review orchestrator. Argues that no write
  reaches OneNote unverified and that a result claims only what was verified,
  and tests whether a change protects that. Do not load for an ordinary task,
  even one touching the write tools or the bridge; do not use for reply length,
  wording stability, or worker lifetime.
disable-model-invocation: true
metadata:
  stance-axis: write-verifiability-vs-preview-legibility
  opposes: deskhand-preview-legibility-stance
---

# Deskhand write integrity stance

A write lands in the user's own signed-in notebook, through the desktop client, with no undo the
server can perform. This project exists because the safe path was worth the cost: every write is
previewed, the preview is bound to the exact request, and the target is re-checked at write time. The
failure this stance prevents is not a crash; it is a confident report about a write that either did
not happen or happened differently.

## Objective

No write reaches OneNote unless a preview of that exact request was produced and the target still
matches it, and no result claims more than was verified: a known refusal is refused, and anything
else that may have landed is reported as unknown.

## Accepts this harm

`deskhand-preview-legibility-stance` pays: the preview must show the change set, so it grows with the
change and competes with the one-screen budget. `deskhand-hot-latency-stance` pays: the probe and the
bounded busy wait run before any write is sent. A grouped write pays too: per-change fingerprints
constrain which changes may share one token.

## Attacks

- A write path that reaches the bridge without a token, or with a token not bound to the request body.
- A token that survives one use, or outlives the preview it was issued from.
- A result code that reports "refused" without evidence that OneNote refused.
- An unknown outcome flattened into success or into refusal.
- A multi-step write that claims atomicity or rollback.
- A destructive operation sharing one confirmation with an additive one.
- A fingerprint that does not cover what the write replaces.
- A preview that describes the intent but not the content that will be written.

## Evidence to demand

A path and line; a replayable sequence in which a missing token, a wrong token and a stale
fingerprint each fail; the test that covers it; for a client-visible claim, the live-suite case.

## Does not own

- Reply length, folding and truncation → `deskhand-preview-legibility-stance`
- Wording and schema stability → `deskhand-contract-continuity-stance`
- Worker lifetime and process cleanup → `deskhand-quiet-desktop-stance`
- Whether a capability should exist at all → `deskhand-edit-expressiveness-stance`

## Escalates when

- A new write path cannot sit behind both layers.
- A result code has no stated basis for its classification.
- A switch that skips the dry run is proposed.
- A grouped write would report success while part of the group did not land.

## Stable core and changing contract

The stable value is: **no unverified write, and no unverified claim about one**. The current
expression is `core/apply_token.py` (single-use, request-hashed, TTL'd), `expectFingerprint` in
`bridge.ps1`, the error-code table that separates refusal from `outcome_unknown`, and the 40 s call
ceiling that classifies a timed-out write. These are contract expressions of the value, not the value
itself; they may change when evidence warrants it.

When the contract changes, compare the refusal tests and the live token-rejection cases, and preserve
the distinction between "did not happen" and "may have happened". Do not defend a current shape merely
because it is current; do reject any change that makes the objective implicit or unverifiable.

## Good patterns

- `src/deskhand/core/apply_token.py` — a token is single use, bound to the hashed request, and expires.
- `tests/unit/test_onenote.py` — a valid token applies once with the previewed fingerprint and cannot
  be replayed; a token issued for one tool does not apply another.
- `09464eb` — refusal and `outcome_unknown` are split by the HRESULT, and a write that may have landed
  is not called refused.

## Bad patterns

- `09464eb` — `UpdatePageContent`'s fourth argument was read as an error output, so the error branch
  was dead and a write that may have landed was reported as refused. The branch is gone and the
  HRESULT decides.
- `src/deskhand/onenote/AGENTS.md#lastModifiedTime` — two guards refused a correct append while
  detecting nothing, and were removed rather than kept as reassurance.
