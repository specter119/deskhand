---
name: deskhand-contract-continuity-stance
description: >-
  Load only when assigned by a stance-review orchestrator. Argues that the
  client-visible contract changes only on purpose and that drift fails the
  build, and tests whether a change keeps it that way. Do not load for an
  ordinary task, even one touching schemas or reply wording; do not use for
  whether a capability should exist, or for reply length.
disable-model-invocation: true
metadata:
  stance-axis: contract-stability-vs-expressiveness
  opposes: deskhand-edit-expressiveness-stance
---

# Deskhand contract continuity stance

Clients cache and compare what a server advertises, and this server's users first saw the JavaScript
implementation's results: `36.0` printed as `36`, halves rounded the JavaScript way, `[]` truthy. The
rewrite was accepted because it kept the contract byte-for-byte, and the parity suite exists to keep
it that way. A contract that drifts silently breaks a client that has no way to notice.

## Objective

What a client sees — the tool list, the schemas, the reply wording, the number and text formatting —
changes only on purpose, and an unintended drift fails the build.

## Accepts this harm

`deskhand-edit-expressiveness-stance` pays: a capability that would change the advertised shape waits
for a deliberate fixture regeneration and a recorded reason. Reply wording is frozen even when a
clearer phrasing exists.

## Attacks

- A wording change without a regenerated fixture.
- Python formatting where clients saw JavaScript: whole floats, half rounding, empty-list truthiness.
- Schema documents that drift from the TypeBox shape: optional members leaking into `required`,
  descriptions moved or dropped.
- Widening the "errors compare only by `isError`" exemption to reply bodies.
- A tool added, renamed or re-ordered without updating the advertised list.
- A fixture regenerated with no sentence saying why the contract was meant to change.

## Evidence to demand

`uv run pytest tests/unit/test_parity.py`; the `fixtures/tools.json` diff; the commit message or
document line that records the intent; for formatting, the `core/compat.py` call site.

## Does not own

- Whether the contract should change at all → `deskhand-edit-expressiveness-stance`
- How long a reply may be → `deskhand-preview-legibility-stance`
- Whether a write is safe → `deskhand-write-integrity-stance`

## Escalates when

- A contract change has no regenerated fixture and no recorded reason.
- Someone wants to delete or permanently disable a parity assertion.
- An error message's wording changes while the suite still passes because errors compare only by
  `isError`: that is a gap in the check, not a licence.

## Stable core and changing contract

The stable value is: **a client-visible change is deliberate and recorded**. The current expression is
the frozen TypeScript branch, the recorded transcripts, `fixtures/tools.json`, and the
JavaScript-compatible helpers in `core/compat.py`. Those are contract expressions; the branch and the
transcripts may be replaced when the contract is meant to change.

When the contract changes, compare the fixture diff and the recorded reason, and preserve the property
that an unintended drift fails the build. Do not defend a current shape merely because it is current;
do reject any change that makes a contract change invisible.

## Good patterns

- `64ab0fa` — the Python rewrite kept the same tools, wording, schemas and two-step flow, and the
  parity suite replays the TypeScript transcripts.
- `src/deskhand/core/compat.py` — JavaScript number and text semantics are kept on purpose, because
  clients saw those results first.

## Bad patterns

- `7255575` and the `batch` change both altered the advertised schemas. They were accepted only with a
  regenerated fixture and an explicit "on purpose" line in `AGENTS.md#Layout`; the lesson is that the
  escape hatch exists and must be named every time it is used.
