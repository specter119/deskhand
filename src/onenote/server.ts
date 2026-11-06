import { readFileSync } from "node:fs";
import { McpServer } from "@modelcontextprotocol/server";
import { defaultTokens } from "../core/apply-token.js";
import { type BridgeFn, createPowerShellBridge } from "../core/powershell.js";
import { registerOneNoteTools } from "./tools.js";

const BRIDGE_SCRIPT = new URL("../../src/onenote/bridge.ps1", import.meta.url);
const PACKAGE_JSON = new URL("../../package.json", import.meta.url);

export const INSTRUCTIONS = [
  "These tools drive the Windows OneNote desktop client through its COM automation interface, from WSL.",
  "They work with the user's own signed-in notebooks, so writes land in real notes that sync to SharePoint/OneDrive.",
  "",
  "Identifying a page: prefer pageId, taken from pages or search. When only names are known, pass notebook + section + page together;",
  "a section name that exists in several notebooks is rejected as ambiguous.",
  "",
  "Writing is always two steps. Call a write tool without dryRun: nothing is written, the reply starts with 'applyToken: <token>' and shows the exact",
  "change OneNote would receive. To apply, repeat the same call with dryRun=false and that applyToken. A token is single use, expires after 10 minutes,",
  "belongs to one server process, and is refused when the arguments differ or the target changed in OneNote after the dry run.",
  "If an apply is interrupted its outcome is unknown: read the page again and preview before retrying.",
  "",
  "Editing text: OneNote replaces an entire outline with whatever one:OEChildren the caller sends, so a paragraph omitted from that list is DELETED.",
  "That is why edit exists and why it resends each affected outline in full, echoing back OneNote's own XML for everything it is not changing.",
  "Use edit for changes inside an outline and append_lines for a new outline at the end of a page. Raw outline XML is never accepted.",
  "",
  "Concurrency: lastModifiedTime does not advance when content is written through COM and the hierarchy API can report a different value than the page",
  "itself, so it is informational only. Safety comes from the token's content fingerprint, and delete_page additionally requires the exact page name.",
].join("\n");

let sharedBridge: BridgeFn | undefined;

function defaultBridge(): BridgeFn {
  sharedBridge ??= createPowerShellBridge({ scriptPath: BRIDGE_SCRIPT.pathname, workSubdir: "onenote" });
  return sharedBridge;
}

export interface OneNoteServerOptions {
  bridge?: BridgeFn;
  tokens?: typeof defaultTokens;
}

/** Cheap and side-effect free: serveStdio may call the factory more than once per process. */
export function createOneNoteServer(options: OneNoteServerOptions = {}): McpServer {
  const { version } = JSON.parse(readFileSync(PACKAGE_JSON, "utf8")) as { version: string };
  const server = new McpServer({ name: "onenote", version }, { instructions: INSTRUCTIONS });
  registerOneNoteTools(server, {
    bridge: options.bridge ?? defaultBridge(),
    tokens: options.tokens ?? defaultTokens,
  });
  return server;
}
