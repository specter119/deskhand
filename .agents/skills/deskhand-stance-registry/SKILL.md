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
| integrity vs legibility (table removal) | the label leads with the removable `oe=` id, then the table id and size, then a first-cell excerpt collapsed to one line and capped at 40 characters | both | pass 2, P2's objection to W2 |

## Open essential conflicts

| Conflict | Pair | Status | Decision record |
| -------- | ---- | ------ | --------------- |
| a concurrent edit between the fingerprint check and `UpdatePageContent` is overwritten, so `applied` does not imply no lost update | `deskhand-write-integrity-stance` vs `deskhand-contract-continuity-stance` | measured-not-adopted: the window is reproduced live; `lastModifiedTime` has one-second granularity and can lag, so it cannot close it | pass 3; `src/deskhand/onenote/AGENTS.md#concurrency` |

## Decisions that sacrificed a stance

| Decision  | Sacrificed stance                     | What was given up                                                        | Why                                                        |
| --------- | ------------------------------------- | ------------------------------------------------------------------------ | ---------------------------------------------------------- |
| `1da5f3f` | `deskhand-write-integrity-stance`     | the full change set in the preview: objectIDs, full texts, all results    | a confirmation nobody can read is not a confirmation       |
| `7255575` | `deskhand-contract-continuity-stance` | the advertised schemas stayed byte-identical to the TypeScript record     | a real page had to match an existing series                |
| `8c5b822` | `deskhand-quiet-desktop-stance`       | no process on the desktop between calls                                   | 0.9 s per call and a `0x800706BF` race were worse          |
| this change (2026-11-16) | `deskhand-write-integrity-stance` | "applied" no longer implies that no concurrent edit was lost | the check and the COM write cannot be one atomic step; a CAS was measured and rejected (granularity and lag) |

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

Pass 2 binding (2026-11-16). Target: the uncommitted jj change `fc1ac784` (parent `1a6c178a`), which
lets `edit` remove a table by the `one:OE` that holds it. Scope: `src/deskhand/onenote/bridge.ps1`,
`schemas.py`, `tools.py`, `AGENTS.md`; `tests/unit/test_bridge_behaviour.py`,
`tests/unit/test_onenote.py`; `tests/parity/fixtures/tools.json`. Pairs in scope: write-integrity vs
preview-legibility, contract-continuity vs edit-expressiveness, preview-legibility vs
contract-continuity. Prior decisions constraining the pass: `7255575` (the contract stance was
sacrificed so a real page could be matched), pass 1's ordering (change the text deliberately, then pin
it), and the one-screen preview cap from `1da5f3f`. Reviewers: one per stance, independent, no shared
output: write-integrity, preview-legibility, contract-continuity, edit-expressiveness.

## Passes

| Pass | Target | Pairs reviewed | Outcome |
| ---- | ------ | -------------- | ------- |
| 1    | `09464eb` | integrity vs legibility, continuity vs expressiveness, and a map miss (legibility vs continuity) | 4 findings, all merged by sequencing; no escalation |
| 2    | `fc1ac784` | integrity vs legibility, continuity vs expressiveness, legibility vs continuity | 18 findings; 11 fixed, 3 accepted mechanism decisions, 4 deferred axes; no escalation |
| 3    | `d10533cd` + uncommitted batch wording | integrity vs continuity, integrity vs expressiveness | 2 findings on the batch error text fixed; 1 mechanism limit documented and left open |

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

Pass 2 detail. Four independent reviewers, one stance each, no shared output. Findings:

- P1 `preview-legibility`, seen before: `read_page`'s text reply printed the table's own `objectID` but
  not `tables[].oeID`, so a text-only client was pointed at the id `edit` refuses.
- P2 `preview-legibility`, seen before: every removed table or image previewed as `[table]`/`[image]`,
  and a long removal list said `delete N paragraph(s)` although tables are not paragraphs.
- W1 `write-integrity`, seen before: the plan fingerprint covers the rewritten outline, not the
  removed child, so a table edited between preview and apply is deleted. This is the documented
  mechanism decision in `src/deskhand/onenote/AGENTS.md#concurrency`.
- W2 `write-integrity`, seen before: the destructive preview showed no identity or content for a table.
- W3 `write-integrity`, new: no read-back comparison follows an apply; `applied` reports that the
  write call returned, as documented.
- W4 `write-integrity`, seen before: `batch` can mix a destructive `edit` with additive operations.
- W5 `write-integrity`, new: a non-text child other than table/image (an ink OE) previewed as an empty
  string.
- C1 `contract-continuity`, seen before: the schema and fixture changed on purpose, but the root
  ledger and the change description did not say so.
- C2 `contract-continuity`, seen before: the "outline would be empty" refusal called every case a
  table and was not pinned verbatim.
- C3 `contract-continuity`, seen before: the nested-removal refusal called a paragraph a table and was
  pinned only by a substring.
- C4 `contract-continuity`, seen before: `[image]` had no pin.
- C5 `contract-continuity`, new: no bridge test read a table page, so the `oeID` key could drift.
- C6 `contract-continuity`, new: the `remove` description pointed at `read_page` while the text reply
  omitted the id.
- E1 `edit-expressiveness`, new: a table inserted as a new outline cannot be removed, because it is
  that outline's only child.
- E2 `edit-expressiveness`, new: `oeID` is optional in the schema, so a client validating the older
  contract does not see the only removal handle.
- E3 `edit-expressiveness`, new: paragraphs inside table cells can be read and appended to but not
  set or removed.
- E4 `edit-expressiveness`, seen before: the tool descriptions did not mention table removal or the
  new-outline caveat.
- E5 `edit-expressiveness`, new: a table cannot be replaced by a paragraph at the same position in one
  call.

Contest and classification. P2 contests W2 over the excerpt: W2 wanted first-cell content in a
destructive preview, P2 wanted one line. P2's objection held that the table's own `objectID` is not
the removable id and that an unbounded excerpt can span lines; the merged label is
`[table oe=<oe id> <table id> <rows>x<columns> "<first cell>"]`, with whitespace collapsed and 40
characters capped. C1-C6, P1, P2, W2, W5 and E4 are fixed in the same change. W1, W3 and W4 are
pre-existing mechanism decisions recorded in `src/deskhand/onenote/AGENTS.md` (the fingerprint covers
the written outline; `applied` means the write call returned; `batch` takes `edit`), unchanged here.
E1 and E3 are the deferred axes already listed under "Not implemented" (deleting a whole outline,
editing inside table cells); E1 is now stated in the `insert_table` description. E2 is not adopted:
deriving the OE from the table's own `objectID` would accept an id the contract says is refused, and
the text reply now carries `oe=` so the handle is discoverable. E5 predates the change and needs an
insert-at-the-removed-position edit, which this pass does not add. Nothing was incompatible within
the target's scope, so nothing was escalated.

Counterexamples harvested. W5's empty label survived the first implementation (`[table]`/`[image]`
only), and the `oeID` field survived as the only removal handle until C6 and P1 compared the text and
structured channels. W1, W3 and W4 are `seen before` findings that the mechanism docs already accept,
so they show the pass can restate a known limit, not that it found a new one. P2's objection is the
one that changed the implementation: it rejected a label built on the table's own id.

Pass 3 detail. Two independent reviewers, both briefed to attack the proposed fixes rather than the
target, plus a statechart check of the write gate. The statechart found one rule broken:
`AppliedIsFresh`, the claim that an apply whose fingerprint check passed cannot overwrite a
concurrent edit. The fingerprint is recomputed immediately before `UpdatePageContent`, but the two
are separate calls, so a write that lands between them is overwritten. A live replay that paused
after the check reproduced it: the apply reported `applied` and the concurrent append was gone. The
precise claim is therefore "applied does not guarantee no lost update", not "the final content
differs from the preview". `dateExpectedLastModified` is hard-coded to `MinValue`, so no
compare-and-swap is in play.

- I1 `write-integrity`, new: `batch` classified `applied_unverified` as an ordinary failure and could
  print "failed before anything was written" after a write call had returned.
- I2 `write-integrity`, new: an apply that failed on `busy`, `rejected`, `stale` or an unclassified
  code consumed the token without saying so, and the `used` reply did not say to read the page first.
- C7 `contract-continuity`, new: the tool text and `INSTRUCTIONS` promised that a changed target is
  refused, which the race above falsifies.

Contest and classification. I1 and I2 are merged by wording: uncertain outcomes say the write call
returned and could not be verified; only a closed list of known no-write codes says "failed"; an
unclassified code says the token is spent and to read the page. C7 is incompatible with the current
mechanism rather than with another stance: no wording can make the check atomic. A live spike
measured `dateExpectedLastModified`: a cross-second stale value is refused with `0x80042010` and the
content is unchanged, but same-second values are accepted and `GetPageContent` can lag by about two
seconds, so it cannot close the window. Not adopted. It is recorded as an open conflict and stated
in the user-visible text, `README.md`, `bridge.ps1` and `src/deskhand/onenote/AGENTS.md`. Nothing is
claimed as fixed beyond that.

Counterexamples harvested. The statechart rule that broke was an unstated assumption in the tool
text, so the model found what four prior stance passes had accepted. The dissolve review also showed
that `outcome_unknown` and `applied_unverified` must stay distinct: the first has no evidence the
write started, the second has a returned write call. Merging them would let one reply claim more than
is known.

## Operational notes

- Reviewer briefs must name the read set (`git show --stat <target>` files only) and a tool-call cap.
  Without both, a reviewer wandered through the whole bridge and timed out twice.
- `pi-glm-5-3` completed the discovery shape (29 turns) but timed out on both reviewer tasks (83 and
  88 turns, no report). `pi-gpt-6-astra` completed 4 of 4 bounded reviewer tasks. Default reviewer
  workers to the latter until the former is shown to finish a bounded evidence task.
- A clean pass is legitimate, but a one-line pass with no cited evidence is not reviewable. Ask for the
  command that showed the target does not move the objective.
