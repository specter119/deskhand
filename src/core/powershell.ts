import { execFile } from "node:child_process";
import { randomUUID } from "node:crypto";
import { existsSync } from "node:fs";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { delimiter, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const POWERSHELL_FALLBACK = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe";
const CALL_TIMEOUT_MS = 180_000;

/** A failure reported by the bridge, or by the transport around it. `code` is machine-readable. */
export class BridgeError extends Error {
  constructor(
    message: string,
    readonly code: string,
  ) {
    super(message);
    this.name = "BridgeError";
  }
}

/** One bridge call: an op name plus its arguments, resolved to the op's `data` object. */
export type BridgeFn = (
  op: string,
  args: Record<string, unknown>,
  signal?: AbortSignal,
) => Promise<Record<string, unknown>>;

export interface PowerShellBridgeOptions {
  /** Absolute path of the bridge script that is sent to powershell.exe. */
  scriptPath: string;
  /** Directory name under %TEMP%\\deskhand that holds this bridge's request/response files. */
  workSubdir: string;
  timeoutMs?: number;
}

/**
 * Every call uses its own request/response pair. The pid keeps two server processes apart (each
 * client starts its own), and the uuid keeps parallel calls inside one process apart.
 */
export function newCallStem(): string {
  return `${process.pid}-${randomUUID()}`;
}

export function findPowerShell(env: NodeJS.ProcessEnv = process.env): string {
  for (const dir of (env.PATH ?? "").split(delimiter)) {
    const candidate = dir ? join(dir, "powershell.exe") : "";
    if (candidate && existsSync(candidate)) return candidate;
  }
  if (existsSync(POWERSHELL_FALLBACK)) return POWERSHELL_FALLBACK;
  throw new BridgeError(
    "powershell.exe was not found: this server must run in WSL with Windows interop enabled",
    "no_windows_host",
  );
}

function psQuote(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

interface WorkDir {
  /** Windows path, e.g. C:\\Users\\me\\AppData\\Local\\Temp\\deskhand\\onenote */
  win: string;
  /** The same directory as a WSL path. */
  wsl: string;
}

async function resolveWorkDir(powershell: string, subdir: string): Promise<WorkDir> {
  let stdout: string;
  try {
    ({ stdout } = await execFileAsync(
      powershell,
      ["-NoProfile", "-NonInteractive", "-Command", "[IO.Path]::GetTempPath()"],
      { encoding: "utf8", timeout: 30_000 },
    ));
  } catch (error) {
    throw new BridgeError(
      `powershell.exe could not be run from WSL (${error instanceof Error ? error.message : String(error)})`,
      "no_windows_host",
    );
  }
  const winTemp = stdout
    .replace(/[\r\n]/g, "")
    .trim()
    .replace(/[\\/]+$/, "");
  const win = `${winTemp}\\deskhand\\${subdir}`;
  const drive = /^([A-Za-z]):[\\/](.*)$/.exec(win);
  if (!drive)
    throw new BridgeError(
      `Unsupported Windows temp path (expected a drive-letter path): ${win}`,
      "no_windows_host",
    );
  const wsl = `/mnt/${drive[1].toLowerCase()}/${drive[2].split("\\").join("/")}`;
  await mkdir(wsl, { recursive: true });
  return { win, wsl };
}

/**
 * Runs one op per powershell.exe process. Nothing touches Windows until the first call.
 *
 * The transport constraints (UTF-8 request file, ~200-character -EncodedCommand bootstrap, no -File,
 * no environment variables for paths) are explained in src/onenote/AGENTS.md and must stay as they are.
 */
export function createPowerShellBridge(options: PowerShellBridgeOptions): BridgeFn {
  const timeout = options.timeoutMs ?? CALL_TIMEOUT_MS;
  let setup: Promise<{ powershell: string; work: WorkDir; script: string }> | undefined;

  const ensureSetup = () => {
    setup ??= (async () => {
      const powershell = findPowerShell();
      const [work, script] = await Promise.all([
        resolveWorkDir(powershell, options.workSubdir),
        readFile(options.scriptPath, "utf8"),
      ]);
      return { powershell, work, script };
    })();
    // A failed setup must be retried by the next call, not cached forever.
    setup.catch(() => {
      setup = undefined;
    });
    return setup;
  };

  return async (op, args, signal) => {
    const { powershell, work, script } = await ensureSetup();
    const stem = newCallStem();
    const requestPath = join(work.wsl, `req-${stem}.json`);
    const responsePath = join(work.wsl, `resp-${stem}.json`);

    await writeFile(requestPath, JSON.stringify({ op, __script: script, ...args }), "utf8");

    const bootstrap = [
      "$ErrorActionPreference='Stop'",
      `$d=${psQuote(work.win)}`,
      `$f=${psQuote(stem)}`,
      '$i=[IO.File]::ReadAllText("$d\\req-$f.json",[Text.Encoding]::UTF8)|ConvertFrom-Json',
      "Invoke-Expression $i.__script",
      "",
    ].join("\n");

    try {
      try {
        await execFileAsync(
          powershell,
          [
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-EncodedCommand",
            Buffer.from(bootstrap, "utf16le").toString("base64"),
          ],
          { maxBuffer: 128 * 1024 * 1024, timeout, signal },
        );
      } catch (error) {
        const detail = error as {
          name?: string;
          code?: string;
          killed?: boolean;
          stderr?: string;
          stdout?: string;
          message?: string;
        };
        const interrupted =
          detail.name === "AbortError" || detail.code === "ABORT_ERR" || detail.killed === true;
        if (interrupted && args.apply === true) {
          throw new BridgeError(
            "The call was cancelled or timed out while a write was in flight, so whether OneNote applied it is unknown. Read the page again and preview before retrying.",
            "outcome_unknown",
          );
        }
        const text = (detail.stderr || detail.stdout || detail.message || "").trim();
        throw new BridgeError(
          `OneNote bridge failed: ${text || "powershell.exe returned a non-zero exit code"}`,
          "bridge_failed",
        );
      }

      let raw: string;
      try {
        raw = await readFile(responsePath, "utf8");
      } catch {
        throw new BridgeError(
          "OneNote bridge produced no response; the bridge script may have failed before it could report an error",
          "no_response",
        );
      }
      const parsed = JSON.parse(raw.replace(/^\uFEFF/, "")) as
        | { ok: true; data: Record<string, unknown> }
        | { ok: false; error: string; code?: string };
      if (!parsed.ok)
        throw new BridgeError(`OneNote ${parsed.code ?? "error"}: ${parsed.error}`, parsed.code ?? "error");
      return parsed.data;
    } finally {
      await Promise.all([rm(requestPath, { force: true }), rm(responsePath, { force: true })]);
    }
  };
}
