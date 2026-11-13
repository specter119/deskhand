# Stake map

Written before any stance file, and the gate the stance set had to pass. Every row cites the evidence
that proves the project would sacrifice something to keep the objective. The stances themselves live in
`.agents/skills/deskhand-*-stance/`; the registry that lists them is
`.agents/skills/deskhand-stance-registry/SKILL.md`.

## 1. Stakes

| Id  | Objective (one line)                                                                   | Evidence                                                                                       | What it would demand                                                              |
| --- | -------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------- |
| S1  | No write reaches OneNote unverified, and a result claims only what was verified         | `src/deskhand/onenote/AGENTS.md#Invariants` (5); `2489f6b`; `09464eb`                            | two layers on every write path; single-use request-bound token; unknown outcomes reported as unknown |
| S2  | Every tool reply is readable in one screen                                             | `1da5f3f`; `src/deskhand/onenote/AGENTS.md#Tool surface`                                        | collapse long lists to counts, cut texts, cap items, put the token first           |
| S3  | The client-visible contract changes only on purpose, and drift fails the build          | `64ab0fa`; `AGENTS.md#Layout`; `src/deskhand/core/compat.py` and `core/schema.py` docstrings     | replay recorded transcripts; TypeBox-shaped schemas; JavaScript number and text semantics |
| S4  | A caller can express the change it actually needs                                      | `7255575`; `docs/formatting-gaps.md`                                                            | paragraph formatting, placement after a paragraph, several changes in one outline  |
| S5  | A hot call answers in ~100 ms                                                          | `ed89c84`; `8c5b822`; `src/deskhand/onenote/AGENTS.md#Testing`                                  | one resident worker, one COM object for its life, one op at a time                 |
| S6  | The server is invisible on the Windows desktop                                         | `src/deskhand/onenote/AGENTS.md` (opening: silent, never takes focus); `tests/live/`            | no visible window, no focus, no process left behind, no authorization or network surface |
| S7  | The server answers inside the client's default timeout                                 | `6d4d223`; `src/deskhand/onenote/AGENTS.md#Worker protocol`                                     | a hard 40 s call ceiling and a startup budget                                     |
| S8  | No personal data in the repository                                                     | `AGENTS.md#Rules`                                                                               | no notebook names, account names, home paths or work identifiers                   |
| S9  | Claims are measured, not assumed                                                       | `src/deskhand/onenote/AGENTS.md#Testing`; `docs/formatting-gaps.md`                              | the live suite prints numbers; a gap is closed only against a real client          |
| S10 | The transport stays one shape                                                          | `src/deskhand/onenote/AGENTS.md#Invariants` (4)                                                 | no `-File`, no environment variables for paths, no namespace manager               |
| S11 | One shared core, one server per application                                            | `AGENTS.md#Rules`                                                                               | promote into `core` only when a second app needs it                                |
| S12 | A multi-change task finishes without a confirmation per change                         | the `batch` change in the working tree (uncommitted)                                            | one preview and one token for a group of changes                                   |

## 2. Oppositions

| Pair      | Contended resource                          | A demands                                            | B resists                                     | B demands                              | A resists                                  | Recorded decision                    |
| --------- | ------------------------------------------- | ---------------------------------------------------- | --------------------------------------------- | -------------------------------------- | ------------------------------------------ | ------------------------------------ |
| S1 vs S2  | the dry-run reply surface: paragraphs and characters shown | every item of the exact change set     | a confirmation nobody can read                | collapse and cut                       | an approval based on a summary             | `1da5f3f`                            |
| S3 vs S4  | the advertised contract: schemas, wording, tool list | no drift without a regenerated fixture       | a capability withheld to freeze the contract  | express the change the caller needs    | a contract change without a recorded reason | `7255575`, `AGENTS.md#Layout`        |
| S5 vs S6  | seconds a process lives on the user's desktop | keep the worker and COM object alive for ~100 ms   | a process left on the desktop                 | close on idle, leave nothing behind    | the 1.6 s cold start per call              | `8c5b822`, `src/deskhand/onenote/AGENTS.md#Files` |
| S1 vs S12 | changes per token, and what a failed token costs | each change independently retryable            | a confirmation per change                     | one token for a group                  | a token spent on partial failure           | `hypothetical` (batch uncommitted)   |
| S2 vs S3  | how much of a reply's text is frozen | the guidance stays visible and the diagnostic is capped | a frozen reply whose actionable sentence can be truncated | the reply text is pinned by a test | a contract change that looks like drift | `1da5f3f` (cap narrowed so short replies did not change) |

## 3. Unopposed stakes

| Id  | Why unpaired                                                                                  | Check that covers it                                                   |
| --- | --------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| S7  | No recorded direction asks to wait for a slow write; the 40 s ceiling is a one-way decision     | `6d4d223`, the live latency gates, `#Worker protocol`'s "do not raise it alone" |
| S8  | A standing constraint with no opponent                                                        | review discipline; `_redact` in `tests/unit/test_parity.py`             |
| S9  | A process discipline, not a sacrifice                                                         | the live suite itself                                                  |
| S10 | Already invariant 4; writing a stance for it is the "frozen implementation" failure mode      | `tests/unit/test_bridge_script.py`                                     |
| S11 | No second application exists yet                                                              | `AGENTS.md#Rules`; it becomes a live axis with the second app          |

## 4. Missing stakes

| Theme                                        | Where it appears                | Why no stake holds it                                              |
| -------------------------------------------- | ------------------------------- | ------------------------------------------------------------------ |
| Python rewrite and the frozen TypeScript branch | `64ab0fa`, `86aaf13`         | a one-way door already passed; nothing trades against it now       |
| Graph versus COM                             | `src/deskhand/onenote/AGENTS.md` (opening) | decided at the boundary; the surviving objective is S6             |
| `uv.lock` and the non-editable install       | `AGENTS.md#Rules`               | a maintainer environment procedure, not a project objective        |

## 5. Candidate stance set

| Stance                                     | Objective                                   | In pairs    |
| ------------------------------------------ | ------------------------------------------- | ----------- |
| `deskhand-write-integrity-stance`          | no unverified write, no unverified claim    | S1 vs S2, S1 vs S12 |
| `deskhand-preview-legibility-stance`       | a reply readable in one screen              | S1 vs S2    |
| `deskhand-contract-continuity-stance`      | the advertised contract moves only on purpose | S3 vs S4  |
| `deskhand-edit-expressiveness-stance`      | the caller can express the change it needs  | S3 vs S4    |
| `deskhand-hot-latency-stance`              | a hot call answers in ~100 ms               | S5 vs S6    |
| `deskhand-quiet-desktop-stance`            | invisible on the desktop, no new authorization surface | S5 vs S6 |

## 6. Deviations

- **S1 vs S12 is deferred.** The `batch` change is in the working tree, not in history, so the pair is
  hypothetical. Forge it once that change lands: it is the axis the maintainer actually asked for
  ("several changes, one confirmation"), and its control is the batch diff itself.
- **S3 vs S4 is kept even though `tests/unit/test_parity.py` enforces part of S3.** The fixtures decide
  *how* a contract change is made, not *whether* it should be. That judgment is what the stance owns.
- **S7, S11 and the Graph axis are deferred with activation conditions**, recorded in the registry.
- **S2 vs S3 was a map miss.** The first pass found `preview-legibility` and `contract-continuity`
  contending over the same reply text, and this map did not list the pair. It is listed now, backed by
  `1da5f3f`. The miss is recorded rather than quietly patched, because a map that grows without a
  record cannot be audited.
