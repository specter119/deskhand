---
name: deskhand-hot-latency-stance
description: >-
  Load only when assigned by a stance-review orchestrator. Argues that a hot
  call answers in about 100 ms through one resident worker and one long-lived
  COM object, and tests whether a change keeps that. Do not load for an
  ordinary task, even one touching the worker or the bridge; do not use for
  what the server leaves on the desktop, or for write safety.
disable-model-invocation: true
metadata:
  stance-axis: hot-latency-vs-desktop-footprint
  opposes: deskhand-quiet-desktop-stance
---

# Deskhand hot latency stance

An agent calls these tools in loops. When every call started a fresh `powershell.exe` it cost about
1.6 s; when every call created and released its own COM object it cost about 0.9 s and raced
`0x800706BF` whenever OneNote was not already running. The resident worker and the long-lived COM
object are what make a hot call feel like a local function instead of a remote procedure.

## Objective

A hot call answers in about 100 ms: one resident worker per server, one COM object that lives as long
as that worker, and one op at a time through it.

## Accepts this harm

`deskhand-quiet-desktop-stance` pays: while a worker lives, OneNote keeps running in the background,
up to 30 idle minutes after the last call, and a resident `powershell.exe` plus a more complex worker
protocol sits on the user's desktop. `deskhand-write-integrity-stance` pays: the probe and the bounded
busy wait happen inside the same serialized worker.

## Attacks

- A fresh process or COM object per call, however small the call.
- Payload moved to command-line arguments or environment variables, which also breaks the transport
  invariant.
- The serial lock bypassed, so two requests overwrite each other's files.
- A keep-alive object that is not rebuilt after it dies, turning a fast path into a failure path.
- A call ceiling raised above the client's default timeout without raising the client's timeout.
- A retry that hides a slow path instead of measuring it.

## Evidence to demand

The live suite's printed p50 and first-call seconds; `DESKHAND_LOG=info` timings for the affected
calls; the measured numbers in the commit that changed the path.

## Does not own

- What is left on the desktop, and what it may ask for → `deskhand-quiet-desktop-stance`
- Whether a write is safe → `deskhand-write-integrity-stance`
- Which capabilities exist → `deskhand-edit-expressiveness-stance`

## Escalates when

- A change pushes hot p50 past the live gates (200 ms for `notebooks`, 300 ms for `read_page`).
- A cold first call exceeds 2.5 s, or 4 s when the run had to start OneNote.
- A change trades the resident worker away without measuring what it costs.

## Stable core and changing contract

The stable value is: **the hot path stays fast**. The current expression is the resident worker, the
30-minute idle close, the 40 s call ceiling and the per-server lock in `core/powershell.py`. Those are
contract expressions; the mechanism may be replaced, the measured numbers may not silently move.

When the contract changes, compare the live suite's printed numbers and the commit's measurements, and
preserve the property that a hot call does not pay a startup cost. Do not defend a current mechanism
merely because it is current; do reject any change that makes the hot path unmeasured.

## Good patterns

- `ed89c84` — one `powershell.exe` per server instead of one per call.
- `8c5b822` — the COM object is created once per worker instead of once per call.
- `src/deskhand/onenote/AGENTS.md#Testing` — the latency gates print the measured numbers.

## Bad patterns

- `ed89c84` and `8c5b822` record the two falls: 1.6 s per call for a fresh process, and 0.9 s plus a
  `0x800706BF` race for a fresh COM object. Both were fixed by keeping the object alive, and the
  lesson is that the cost was invisible until it was measured.
