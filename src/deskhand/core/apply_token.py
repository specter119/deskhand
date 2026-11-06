"""Single-use tokens that tie an apply to the dry run before it."""

from __future__ import annotations

import hashlib
import json
import secrets
import time
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Literal

from deskhand.core.compat import compact

TokenFailure = Literal["missing", "unknown", "used", "expired", "mismatch"]

DEFAULT_TTL_MS = 10 * 60 * 1000
MAX_ENTRIES = 256


class TokenError(Exception):
    def __init__(self, message: str, reason: TokenFailure) -> None:
        super().__init__(message)
        self.reason = reason


def stable_stringify(value: Any) -> str:
    """JSON with sorted keys and no null members, so equal requests always serialize equally."""
    return json.dumps(compact(value), sort_keys=True, separators=(",", ":"), ensure_ascii=True)


def hash_request(tool: str, request: dict[str, Any]) -> str:
    digest = hashlib.sha256()
    digest.update(tool.encode())
    digest.update(b"\n")
    digest.update(stable_stringify(request).encode())
    return digest.hexdigest()


@dataclass
class _Entry:
    tool: str
    request_hash: str
    fingerprint: str
    expires_at: float
    used: bool = False


class TokenTable:
    """Binds an apply to the dry run that preceded it.

    A token ties together the tool, the exact request that is sent to the bridge (not the caller's raw
    arguments) and the fingerprint of the target that the dry run saw. It is single use, expires, and lives
    only in this process: a restart invalidates every outstanding token, which is intended.

    The methods are synchronous on purpose: checking a token and marking it used must not be interleaved
    with another call.
    """

    def __init__(
        self,
        ttl_ms: float = DEFAULT_TTL_MS,
        now: Callable[[], float] = lambda: time.time() * 1000,
    ) -> None:
        self._ttl_ms = ttl_ms
        self._now = now
        self._entries: dict[str, _Entry] = {}

    def issue(self, tool: str, request: dict[str, Any], fingerprint: str) -> str:
        self._prune()
        while len(self._entries) >= MAX_ENTRIES:
            del self._entries[next(iter(self._entries))]
        token = f"dh1_{secrets.token_hex(16)}"
        self._entries[token] = _Entry(
            tool=tool,
            request_hash=hash_request(tool, request),
            fingerprint=fingerprint,
            expires_at=self._now() + self._ttl_ms,
        )
        return token

    def consume(self, token: str | None, tool: str, request: dict[str, Any]) -> str:
        """Validate the token for exactly this request and spend it. Returns the fingerprint to enforce."""
        if not token:
            raise TokenError(
                f"{tool} with dryRun=false needs the applyToken from a dry run of the same call. "
                "Run it without dryRun first.",
                "missing",
            )
        entry = self._entries.get(token)
        if entry is None:
            raise TokenError(
                "Unknown applyToken: it was not issued by this server process or has expired. "
                "Run the dry run again.",
                "unknown",
            )
        if entry.used:
            raise TokenError("This applyToken was already used. Run the dry run again.", "used")
        if self._now() > entry.expires_at:
            del self._entries[token]
            raise TokenError("This applyToken has expired. Run the dry run again.", "expired")
        if entry.tool != tool or entry.request_hash != hash_request(tool, request):
            raise TokenError(
                "The arguments differ from the dry run this applyToken belongs to. "
                "Repeat the same arguments, or run a new dry run.",
                "mismatch",
            )
        entry.used = True
        return entry.fingerprint

    def _prune(self) -> None:
        now = self._now()
        for token in [t for t, e in self._entries.items() if e.used or now > e.expires_at]:
            del self._entries[token]
