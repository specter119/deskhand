// Replays tests/parity/scenario.json against the built TypeScript server with a scripted bridge and
// writes what a client sees. The Python port replays the same scenario and must match.
//
//   npm run build && node tests/parity/dump.mjs
import { createHash } from "node:crypto";
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { TokenTable } from "../../dist/core/apply-token.js";
import { BridgeError } from "../../dist/core/powershell.js";
import { INSTRUCTIONS } from "../../dist/onenote/server.js";
import { connect, fakeBridge } from "../unit/helpers.mjs";

const scenario = JSON.parse(readFileSync(new URL("./scenario.json", import.meta.url), "utf8"));
const out = new URL("./fixtures/", import.meta.url);
mkdirSync(out, { recursive: true });

const tmp = mkdtempSync(join(tmpdir(), "deskhand-parity-"));
for (const [name, base64] of Object.entries(scenario.files)) {
  writeFileSync(join(tmp, name), Buffer.from(base64, "base64"));
}

const normalize = (value) =>
  JSON.parse(
    JSON.stringify(value)
      .replaceAll(tmp, "<TMP>")
      .replace(/dh1_[0-9a-f]{32}/g, "<TOKEN>"),
  );

const substitute = (value) => {
  if (typeof value === "string") return value.replace("$TMP", tmp);
  if (Array.isArray(value)) return value.map(substitute);
  if (value && typeof value === "object")
    return Object.fromEntries(Object.entries(value).map(([key, member]) => [key, substitute(member)]));
  return value;
};

const redactPayload = (args) => {
  const copy = structuredClone(args);
  if (typeof copy.base64 === "string") {
    copy.base64 = `sha256:${createHash("sha256").update(copy.base64).digest("hex")}`;
  }
  return copy;
};

const tokens = new TokenTable();
const issued = {};
const transcripts = [];
let toolList;

for (const step of scenario.steps) {
  const handlers = Object.fromEntries(
    Object.entries(step.bridge).map(([op, entry]) => [
      op,
      (args) => {
        if (entry.error) throw new BridgeError(entry.error.message, entry.error.code);
        if (entry.any) return entry.any;
        return args.apply === true ? entry.apply : entry.dry;
      },
    ]),
  );
  const bridge = fakeBridge(handlers);
  const t = await connect({ bridge, tokens });
  toolList ??= { tools: (await t.client.listTools()).tools, instructions: t.client.getInstructions() };

  const args = substitute(step.args);
  if (step.applyFrom) Object.assign(args, { dryRun: false, applyToken: issued[step.applyFrom] });

  const record = { name: step.name, tool: step.tool };
  try {
    const result = await t.call(step.tool, args);
    record.isError = result.isError === true;
    if (!record.isError) {
      record.text = result.content.map((block) => block.text).join("\n");
      record.structuredContent = result.structuredContent;
      issued[step.name] = result.structuredContent?.applyToken;
    }
  } catch (error) {
    record.isError = true;
    record.thrown = String(error.message).slice(0, 120);
  }
  record.bridgeCalls = bridge.calls.map((call) => ({ op: call.op, args: redactPayload(call.args) }));
  transcripts.push(normalize(record));
  await t.close();
}

if (toolList.instructions !== INSTRUCTIONS) throw new Error("instructions differ from the exported constant");
writeFileSync(new URL("ts-tools.json", out), `${JSON.stringify(normalize(toolList), null, 2)}\n`);
writeFileSync(new URL("ts-transcripts.json", out), `${JSON.stringify(transcripts, null, 2)}\n`);
console.log(`wrote ${transcripts.length} transcripts and ${toolList.tools.length} tools`);
