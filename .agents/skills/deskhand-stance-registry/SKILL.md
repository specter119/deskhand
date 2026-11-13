---
name: deskhand-stance-registry
description: >-
  Use to find which stance owns a question in deskhand, which stances oppose
  each other, and which conflicts are open or already decided. Read by the
  orchestrator before assigning a stance review or recording a resolution. Do
  not use as a review stance itself, and do not hand it to a reviewer.
---

# Deskhand stance registry

The stances below are first-principles objectives, not checklists. They are deliberately opposed: each
one demands something another one pays for. A pass that finds no conflict on a change touching a listed
opposition should be re-examined, but a clean pass is a possible outcome and not evidence of a faulty
review. Do not invent a conflict to satisfy the map.

This file is read by the orchestrator. Do not hand it to a reviewer: it names the oppositions the
orchestrator expects, and a reviewer that has read it will find them.

The stake map that justified this set is `docs/stance-map.md`.

## Stances

| Stance                                | Axis                                | Objective (one line)                                              | Accepts                                                        |
| ------------------------------------- | ----------------------------------- | ----------------------------------------------------------------- | -------------------------------------------------------------- |
| `deskhand-write-integrity-stance`     | write verifiability                 | no unverified write, no unverified claim                          | longer previews, a probe on the write path, per-change fingerprints |
| `deskhand-preview-legibility-stance`  | preview legibility                  | a reply readable in one screen                                    | a summary instead of the full change set, regenerated fixtures  |
| `deskhand-contract-continuity-stance` | contract stability                  | the advertised contract moves only on purpose                     | capabilities wait for a deliberate regeneration                 |
| `deskhand-edit-expressiveness-stance` | expressiveness                      | the caller can express the change it needs                        | contract changes, longer previews                               |
| `deskhand-hot-latency-stance`         | hot-call latency                    | a hot call answers in ~100 ms                                     | a worker and OneNote staying alive on the desktop               |
| `deskhand-quiet-desktop-stance`       | desktop footprint                   | invisible on the desktop, no new authorization surface            | a 30-minute idle close and a ~2 s cold start                    |

## Oppositions

Each row is a pair that must be reviewed together. A pass that reviews only one side of a row is not a
stance review.

| Pair                                                        | Contended resource                            | Where these objectives are known to compete                    |
| ----------------------------------------------------------- | --------------------------------------------- | -------------------------------------------------------------- |
| `deskhand-write-integrity-stance` vs `deskhand-preview-legibility-stance` | paragraphs and characters shown in a dry-run reply | the edit preview after a section rewrite (`1da5f3f`) |
| `deskhand-preview-legibility-stance` vs `deskhand-contract-continuity-stance` | how much of a reply's text is frozen: the guidance, the caps, the wording | the cap added in `1da5f3f`, narrowed so short replies did not change |
| `deskhand-contract-continuity-stance` vs `deskhand-edit-expressiveness-stance` | the advertised schemas, wording and tool list | paragraph formatting and the `after` anchor (`7255575`) |
| `deskhand-hot-latency-stance` vs `deskhand-quiet-desktop-stance` | seconds a process lives on the user's desktop | the resident worker and the long-lived COM object (`8c5b822`) |

## Contract surfaces

Decisions from a stance conflict land in these places. A resolution that cannot name its file here is
incomplete.

- `src/deskhand/core/apply_token.py` — token lifetime, binding and single use
- `src/deskhand/onenote/bridge.ps1` — fingerprint scope and error classification
- `src/deskhand/onenote/tools.py` — preview composition, caps and the tool surface
- `src/deskhand/onenote/schemas.py` — the advertised input and output schemas
- `tests/parity/fixtures/tools.json` — the advertised contract a client sees
- `src/deskhand/onenote/AGENTS.md` — the mechanism the stances constrain
- `docs/stance-map.md` — the stake map and the deferred axes

## Resolved without escalation

Designs that served both sides of a pair. Recorded so the next pass does not relitigate them.

| Conflict                | Resolution                                                                 | Serves | Evidence                                  |
| ----------------------- | -------------------------------------------------------------------------- | ------ | ----------------------------------------- |
| integrity vs legibility | the applyToken is the first line, so truncation cannot remove it           | both   | `1da5f3f`, `src/deskhand/onenote/AGENTS.md#Tool surface` |
| integrity vs legibility | caps were added without changing short replies, so parity survived the fix  | both   | `1da5f3f` ("A short edit prints exactly what it printed before") |
| legibility vs continuity | reorder the reply first, then pin it: the guidance moves ahead of the diagnostic, the diagnostic is capped, and only then does the contract stance freeze the resulting text | both | pass 1, findings F3/F4 |
| integrity vs continuity | a tightened classification changes error codes and text, so the contract stance pins the new behaviour after the fix, not before it | both | pass 1, findings F1/F2/F3 |

## Open essential conflicts

| Conflict | Pair | Status | Decision record |
| -------- | ---- | ------ | --------------- |
| —        | —    | none open | —            |

## Decisions that sacrificed a stance

| Decision  | Sacrificed stance                     | What was given up                                                        | Why                                                        |
| --------- | ------------------------------------- | ------------------------------------------------------------------------ | ---------------------------------------------------------- |
| `1da5f3f` | `deskhand-write-integrity-stance`     | the full change set in the preview: objectIDs, full texts, all results    | a confirmation nobody can read is not a confirmation       |
| `7255575` | `deskhand-contract-continuity-stance` | the advertised schemas stayed byte-identical to the TypeScript record     | a real page had to match an existing series                |
| `8c5b822` | `deskhand-quiet-desktop-stance`       | no process on the desktop between calls                                   | 0.9 s per call and a `0x800706BF` race were worse          |

## Deferred axes

Recorded here so a later pass does not forge them from memory.

| Axis | Activation condition | Notes |
| ---- | -------------------- | ----- |
| change granularity vs confirmation budget (`S1` vs `S12`) | the `batch` change lands | control is the batch diff: one token for a group, per-change fingerprints, same-outline pairs refused |
| client timeout vs slow writes (`S7`) | a second client with a different default timeout, or a request to support an op over 40 s | currently a one-way decision (`6d4d223`) |
| one shared core vs per-app servers (`S11`) | a second `src/deskhand/<app>/` exists | `quiet-desktop` would then compete across applications, not only for OneNote |

## Evidence

Pass 1 changed a design decision: the error reply is reordered and capped before its text is pinned.
It also added one opposition to the map. The control the set was installed against — `09464eb`, where
an exception around a write used to mean the write was refused — was caught by F1 and F2, both from
`deskhand-write-integrity-stance`.

## Passes

| Pass | Target | Pairs reviewed | Outcome |
| ---- | ------ | -------------- | ------- |
| 1    | `09464eb` | integrity vs legibility, continuity vs expressiveness, and a map miss (legibility vs continuity) | 4 findings, all merged by sequencing; no escalation |

Pass 1 detail. Four independent reviewers, one stance each, no shared output. Findings:

- F1 `write-integrity`, seen before: `Get-WriteFailureCode` matches `0x800420*` and reports every
  unlisted code in that range as `rejected`, contradicting the comment above it that only a known
  refusal counts. Reachability of an unknown code in that range is unverified.
- F2 `write-integrity`, new: `_read_response` reads `parsed.get("ok")` and `parsed["data"]` outside the
  JSON guard, so a malformed envelope escapes the `outcome_unknown` classification.
- F3 `contract-continuity`, seen before: `09464eb` rewrote error reply bodies and added no fixture; the
  new tests assert the code and a prefix, and `test_parity.py` compares errors only by `isError`.
- F4 `preview-legibility`, new: the `outcome_unknown` reply concatenates an unbounded exception message
  before "Read the page again and preview before retrying", so a long diagnostic can push the
  actionable sentence past a client's truncation point.

Map miss. F3 and F4 contend over the same reply text, and the registry did not list
`preview-legibility` against `contract-continuity`. The pair is now listed; `1da5f3f` backs it, because
that cap was deliberately narrowed so short replies did not change.

Contest and classification. F3 contests F4: pinning the text before changing it would make the change
look like drift. F4 contests F3: freezing a reply whose actionable sentence can be truncated freezes a
defect. Both survive because the objectives are ordered, not opposed: change the text deliberately
(F4), then pin it (F3). The same ordering resolves F1 and F2 against F3. Nothing was incompatible, so
nothing was escalated. The one unresolved item is F1's reachability: whether OneNote can return an
unlisted `0x800420xx` after a partial write is unmeasured.

Counterexamples harvested. "The stance found it" is not evidence that the stance is accurate: F1 and
F3 are `seen before` findings, so they may only show that the Attacks lists were read. F2 and F4 are
`new`, and F4 is the one that produced a conflict with another stance.

Outcome. F1, F2 and F4 are fixed together: `$HR_REFUSED` lists the four refusals OneNote is known to
make and everything else in the range is `outcome_unknown`; an unreadable or wrongly shaped response
is `outcome_unknown` for a write and `no_response` for a read; the guidance precedes a diagnostic
capped at 300 characters, and the diagnostic now comes from the unwrapped COM exception instead of
PowerShell's `Exception calling ...` wrapper. F3 is resolved by pinning the exact error text in
`tests/unit/test_bridge_behaviour.py` and `tests/unit/test_transport.py`, not in the parity fixture:
that fixture's error wording comes from the TypeScript SDK, and `test_parity.py` compares errors only
by `isError` on purpose. The pin was verified by mutating one word of a message and watching the test
fail, then reverting.

## Operational notes

- Reviewer briefs must name the read set (`git show --stat <target>` files only) and a tool-call cap.
  Without both, a reviewer wandered through the whole bridge and timed out twice.
- `pi-glm-5-3` completed the discovery shape (29 turns) but timed out on both reviewer tasks (83 and
  88 turns, no report). `pi-gpt-6-astra` completed 4 of 4 bounded reviewer tasks. Default reviewer
  workers to the latter until the former is shown to finish a bounded evidence task.
- A clean pass is legitimate, but a one-line pass with no cited evidence is not reviewable. Ask for the
  command that showed the target does not move the objective.
