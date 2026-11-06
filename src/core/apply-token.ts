import { createHash, randomBytes } from "node:crypto";

export type TokenFailure = "missing" | "unknown" | "used" | "expired" | "mismatch";

export class TokenError extends Error {
  constructor(
    message: string,
    readonly reason: TokenFailure,
  ) {
    super(message);
    this.name = "TokenError";
  }
}

/** JSON with sorted keys and no undefined members, so equal requests always serialize equally. */
export function stableStringify(value: unknown): string {
  if (value === undefined) return "null";
  if (value === null || typeof value !== "object") return JSON.stringify(value) ?? "null";
  if (Array.isArray(value)) return `[${value.map((item) => stableStringify(item)).join(",")}]`;
  const entries = Object.entries(value as Record<string, unknown>)
    .filter(([, member]) => member !== undefined)
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return `{${entries.map(([key, member]) => `${JSON.stringify(key)}:${stableStringify(member)}`).join(",")}}`;
}

export function hashRequest(tool: string, request: Record<string, unknown>): string {
  return createHash("sha256").update(tool).update("\n").update(stableStringify(request)).digest("hex");
}

interface Entry {
  tool: string;
  requestHash: string;
  fingerprint: string;
  expiresAt: number;
  used: boolean;
}

const DEFAULT_TTL_MS = 10 * 60 * 1000;
const MAX_ENTRIES = 256;

/**
 * Binds an apply to the dry run that preceded it. A token ties together the tool, the exact request
 * that is sent to the bridge (not the caller's raw arguments) and the fingerprint of the target
 * that the dry run saw. It is single use, expires, and lives only in this process: a restart
 * invalidates every outstanding token, which is intended.
 */
export class TokenTable {
  private readonly entries = new Map<string, Entry>();

  constructor(
    private readonly ttlMs = DEFAULT_TTL_MS,
    private readonly now: () => number = Date.now,
  ) {}

  issue(tool: string, request: Record<string, unknown>, fingerprint: string): string {
    this.prune();
    while (this.entries.size >= MAX_ENTRIES) {
      const oldest = this.entries.keys().next();
      if (oldest.done) break;
      this.entries.delete(oldest.value);
    }
    const token = `dh1_${randomBytes(16).toString("hex")}`;
    this.entries.set(token, {
      tool,
      requestHash: hashRequest(tool, request),
      fingerprint,
      expiresAt: this.now() + this.ttlMs,
      used: false,
    });
    return token;
  }

  /** Validates the token for exactly this request and spends it. Returns the fingerprint to enforce. */
  consume(token: string | undefined, tool: string, request: Record<string, unknown>): string {
    if (!token) {
      throw new TokenError(
        `${tool} with dryRun=false needs the applyToken from a dry run of the same call. Run it without dryRun first.`,
        "missing",
      );
    }
    const entry = this.entries.get(token);
    if (!entry) {
      throw new TokenError(
        "Unknown applyToken: it was not issued by this server process or has expired. Run the dry run again.",
        "unknown",
      );
    }
    if (entry.used) throw new TokenError("This applyToken was already used. Run the dry run again.", "used");
    if (this.now() > entry.expiresAt) {
      this.entries.delete(token);
      throw new TokenError("This applyToken has expired. Run the dry run again.", "expired");
    }
    if (entry.tool !== tool || entry.requestHash !== hashRequest(tool, request)) {
      throw new TokenError(
        "The arguments differ from the dry run this applyToken belongs to. Repeat the same arguments, or run a new dry run.",
        "mismatch",
      );
    }
    entry.used = true;
    return entry.fingerprint;
  }

  private prune(): void {
    const now = this.now();
    for (const [token, entry] of this.entries) {
      if (entry.used || now > entry.expiresAt) this.entries.delete(token);
    }
  }
}

/** Module level on purpose: a server factory may run more than once per process. */
export const defaultTokens = new TokenTable();
