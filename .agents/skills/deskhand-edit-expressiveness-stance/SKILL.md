---
name: deskhand-edit-expressiveness-stance
description: >-
  Load only when assigned by a stance-review orchestrator. Argues that a caller
  must be able to express the change it actually needs, and tests whether a
  change adds that ability instead of approximating it. Do not load for an
  ordinary task, even one adding a tool or a parameter; do not use for contract
  stability, reply length, or write safety.
disable-model-invocation: true
metadata:
  stance-axis: contract-stability-vs-expressiveness
  opposes: deskhand-contract-continuity-stance
---

# Deskhand edit expressiveness stance

A write API that cannot say what the caller means forces the caller to approximate: rebuild an
outline by hand, append at the end of a page because placement is unavailable, or accept default
formatting because there is no way to set it. The approximation is where content gets lost. This
project has already paid that cost once, when a page had to match an existing series and the API could
not set a single font.

## Objective

A caller can express the change it actually needs — paragraph formatting, placement after a given
paragraph, several changes inside one outline — instead of approximating it with the shapes the
contract already has.

## Accepts this harm

`deskhand-contract-continuity-stance` pays: every new capability changes the advertised schemas and
reply wording, so the recorded fixtures must be regenerated and a client may hold a stale schema.
`deskhand-preview-legibility-stance` pays: a richer change produces a longer preview to cut down.

## Attacks

- A capability withheld because the contract is frozen, leaving the caller to approximate.
- A write that can only append at the end of a page when placement was the point.
- Formatting reachable only by hand-building OneNote XML.
- A needed change expressible only as several calls, each needing its own confirmation.
- A new parameter that is accepted and then ignored by the bridge.
- A schema addition that the tool description never mentions, so the caller cannot find it.

## Evidence to demand

The scenario that could not be expressed before; the test that now covers it; the fixture diff and
the description text that advertises it; for a gap, the record that closed it.

## Does not own

- Whether the contract may change → `deskhand-contract-continuity-stance`
- How much of the change a preview shows → `deskhand-preview-legibility-stance`
- Whether the resulting write is safe → `deskhand-write-integrity-stance`

## Escalates when

- A needed change has no expression and the available workaround can corrupt content.
- A capability would change the contract with no recorded reason and no fixture regeneration.
- The only way to express a change is to send OneNote XML directly.

## Stable core and changing contract

The stable value is: **the caller can express the change it needs**. The current expression is the
paragraph style object, the `after` and `outline` anchors, and the grouped `batch` tool. Those are
contract expressions and may be reshaped when a real need appears.

When the contract changes, compare the gap record and the tests that cover the new expression, and
preserve the property that a caller never has to approximate a change it can describe. Do not defend a
current shape merely because it is current; do reject any addition that the bridge silently ignores.

## Good patterns

- `7255575` — paragraph formatting and the `after` anchor, added because a real page had to match an
  existing one.
- `docs/formatting-gaps.md` — the gap was recorded against a real client before the API changed, and
  closed only when the API covered it.

## Bad patterns

- Before `7255575`, `append_lines` and `create_page` took `lines: string[]` and `edit` took
  `text: string`, so no formatting was expressible at all. `docs/formatting-gaps.md#What was missing`
  records the cost: every new paragraph landed on the default style.
