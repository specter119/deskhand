---
name: deskhand-quiet-desktop-stance
description: >-
  Load only when assigned by a stance-review orchestrator. Argues that the
  server stays invisible on the Windows desktop and adds no authorization or
  network surface, and tests whether a change keeps that. Do not load for an
  ordinary task, even one touching the bridge or a dependency; do not use for
  call latency, or for write safety.
disable-model-invocation: true
metadata:
  stance-axis: hot-latency-vs-desktop-footprint
  opposes: deskhand-hot-latency-stance
---

# Deskhand quiet desktop stance

The user is working in OneNote while this server drives it. A window that appears, a focus steal, a
process left running after the client closed, or a new consent prompt all turn a background helper
into an interruption. Graph was rejected on purpose for exactly this reason: it needs an app
registration, a write scope and a network path, where COM needs none of them.

## Objective

The server is invisible on the Windows desktop: no visible window, no focus, no process left behind,
and no authorization or network surface added.

## Accepts this harm

`deskhand-hot-latency-stance` pays: the worker closes after 30 idle minutes and a cold first call costs
about 2 s, instead of an unbounded keep-alive for a faster hot path. A capability that would need a
consent prompt is not built.

## Attacks

- A path that shows a window or takes focus.
- A new dependency that needs network access, an app registration, or admin consent.
- A worker left behind after the server is killed.
- An idle keep-alive extended with no recorded decision.
- A background process that outlives the client and keeps holding OneNote open.
- User data crossing the repository boundary.

## Evidence to demand

The live suite's "no worker left after a SIGKILL" check; the consent and network requirements of a new
dependency, stated explicitly; the measured idle-close seconds.

## Does not own

- How fast a call is → `deskhand-hot-latency-stance`
- OneNote's own behaviour and limits → `src/deskhand/onenote/AGENTS.md`
- Whether a write is safe → `deskhand-write-integrity-stance`

## Escalates when

- A change introduces a network or authorization requirement.
- Process lifetime would extend past the recorded 30 idle minutes.
- A new dependency cannot be shown to be silent and offline.

## Stable core and changing contract

The stable value is: **the desktop is not disturbed and no new authorization surface appears**. The
current expression is COM automation, the silent proxy launch, the 30-minute idle close and the probe
before a write. Those are contract expressions and may be replaced by another silent local mechanism.

When the contract changes, compare the live suite's cleanup check and the dependency's consent and
network requirements, and preserve the property that closing the client leaves nothing behind. Do not
defend a current mechanism merely because it is current; do reject any change that makes the server
visible or makes it ask for permission.

## Good patterns

- `src/deskhand/onenote/AGENTS.md` (opening) — the COM proxy starts OneNote without a visible window
  and never takes focus.
- `AGENTS.md` and `src/deskhand/onenote/AGENTS.md` — Graph was rejected on purpose: no app
  registration, no write scope, no network.
- `tests/unit/` — the whole unit suite runs with a fake bridge and no Windows.

## Bad patterns

- `ed89c84` — `bridge.ps1` used to `exit` on a failure, which would have ended the worker mid-call and
  left the desktop with a dead worker. It now writes the response and throws, and the worker survives.
