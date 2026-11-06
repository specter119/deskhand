import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { TokenTable } from "../../dist/core/apply-token.js";
import { createOneNoteServer } from "../../dist/onenote/server.js";

export const PAGE = {
  id: "page-1",
  name: "Probe",
  level: "0",
  lang: "en-US",
  created: "2026-01-01T00:00:00.000Z",
  lastModified: "2026-01-01T00:00:00.000Z",
  lines: [
    {
      index: 0,
      objectID: "oe-1",
      text: "hello",
      inTitle: false,
      inTable: false,
      hasImage: false,
      hasTable: false,
      hasInk: false,
      list: "",
      depth: 0,
      outlineID: "ol-1",
      quickStyle: "",
      alignment: "",
    },
  ],
  outlines: [
    { objectID: "ol-1", x: null, y: null, z: null, width: null, height: null, inTitle: false, lineCount: 1 },
  ],
  images: [],
  tables: [],
};

/** A bridge that records every call and answers from `handlers[op]`. */
export function fakeBridge(handlers = {}) {
  const calls = [];
  const bridge = async (op, args) => {
    calls.push({ op, args: structuredClone(args) });
    const handler = handlers[op];
    if (!handler) throw new Error(`no fake handler for ${op}`);
    return handler(args);
  };
  bridge.calls = calls;
  return bridge;
}

export const previewHandler =
  (fingerprint = "fp-1") =>
  (args) => ({
    applied: args.apply === true,
    fingerprint,
    page: PAGE,
    changes: [],
  });

export async function connect(options = {}) {
  const tokens = options.tokens ?? new TokenTable();
  const bridge = options.bridge ?? fakeBridge();
  const server = createOneNoteServer({ bridge, tokens });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "0" });
  await server.connect(serverSide);
  await client.connect(clientSide);
  return {
    client,
    bridge,
    tokens,
    call: (name, args = {}) => client.callTool({ name, arguments: args }),
    close: async () => {
      await client.close();
      await server.close();
    },
  };
}

export const textOf = (result) => result.content.map((block) => block.text).join("\n");
export const applyCalls = (bridge) => bridge.calls.filter((call) => call.args.apply === true);
