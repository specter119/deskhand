#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { createOneNoteServer } from "./onenote/server.js";

const APPS: Record<string, () => ReturnType<typeof createOneNoteServer>> = {
  onenote: () => createOneNoteServer(),
};

function usage(): string {
  return [
    "Usage: deskhand <app>",
    "",
    "Apps:",
    ...Object.keys(APPS).map((name) => `  ${name}`),
    "",
    "Each app speaks MCP over stdio.",
  ].join("\n");
}

const [app, ...rest] = process.argv.slice(2);

if (app === "--version" || app === "-v") {
  const { version } = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
    version: string;
  };
  console.log(version);
} else if (!app || app === "--help" || app === "-h") {
  console.log(usage());
} else if (!(app in APPS) || rest.length) {
  console.error(app in APPS ? `Unexpected argument: ${rest[0]}` : `Unknown app: ${app}`);
  console.error(usage());
  process.exitCode = 2;
} else {
  // stdout carries JSON-RPC only; everything else goes to stderr.
  const handle = serveStdio(APPS[app], { onerror: (error) => console.error(`[deskhand ${app}]`, error) });
  const stop = () => void handle.close().finally(() => process.exit(0));
  process.once("SIGINT", stop);
  process.once("SIGTERM", stop);
}
