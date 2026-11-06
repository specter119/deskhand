import assert from "node:assert/strict";
import { test } from "node:test";
import { BridgeError } from "../../dist/core/powershell.js";
import { INSTRUCTIONS } from "../../dist/onenote/server.js";
import { applyCalls, connect, fakeBridge, PAGE, previewHandler, textOf } from "./helpers.mjs";

const TOOLS = [
  "notebooks",
  "pages",
  "search",
  "read_page",
  "create_page",
  "append_lines",
  "insert_table",
  "insert_image",
  "edit",
  "delete_page",
];
const WRITE_TOOLS = ["create_page", "append_lines", "insert_table", "insert_image", "edit", "delete_page"];

test("listTools exposes the ten tools with titles, annotations and schemas", async () => {
  const t = await connect();
  const { tools } = await t.client.listTools();
  assert.deepEqual(tools.map((tool) => tool.name).sort(), [...TOOLS].sort());
  for (const tool of tools) {
    assert.ok(tool.title, `${tool.name} has a title`);
    assert.equal(tool.inputSchema.type, "object");
    assert.equal(tool.outputSchema?.type, "object");
    assert.ok(tool.annotations, `${tool.name} has annotations`);
    assert.equal(tool.annotations.readOnlyHint === true, !WRITE_TOOLS.includes(tool.name), tool.name);
  }
  for (const name of WRITE_TOOLS) {
    const tool = tools.find((candidate) => candidate.name === name);
    assert.ok(tool.inputSchema.properties.dryRun && tool.inputSchema.properties.applyToken, name);
    assert.match(tool.description, /applyToken/);
  }
  assert.equal(tools.find((tool) => tool.name === "edit").annotations.destructiveHint, true);
  assert.equal(tools.find((tool) => tool.name === "delete_page").annotations.destructiveHint, true);
  assert.equal(t.client.getInstructions(), INSTRUCTIONS);
  await t.close();
});

test("read tools validate against their output schemas, including PowerShell quirks", async () => {
  const bridge = fakeBridge({
    notebooks: () => ({
      notebooks: [{ id: "n1", name: "NB", path: "p", sections: { id: "s1", name: "Sec", path: "p2" } }],
    }),
    pages: () => ({
      sectionId: "s1",
      section: "Sec",
      pages: { id: "p1", name: "Probe", level: "0", created: "t" },
    }),
    search: () => ({
      total: 1,
      returned: 1,
      pages: { pageId: "p1", page: "Probe", section: "Sec", notebook: "NB" },
    }),
    read_page: () => ({ page: { ...PAGE, lines: PAGE.lines[0], outlines: PAGE.outlines[0] } }),
  });
  const t = await connect({ bridge });
  for (const [name, args] of [
    ["notebooks", {}],
    ["pages", { sectionId: "s1" }],
    ["search", { query: "pro" }],
    ["read_page", { pageId: "p1" }],
  ]) {
    const result = await t.call(name, args);
    assert.notEqual(result.isError, true, `${name}: ${textOf(result)}`);
    assert.ok(result.structuredContent, name);
  }
  const page = (await t.call("read_page", { pageId: "p1" })).structuredContent.page;
  assert.ok(Array.isArray(page.lines) && Array.isArray(page.outlines));
  assert.equal(
    bridge.calls.some((call) => call.args.apply !== undefined),
    false,
  );
  await t.close();
});

test("a dry run puts the token on the first line and never applies", async () => {
  const bridge = fakeBridge({ append: previewHandler() });
  const t = await connect({ bridge });
  const result = await t.call("append_lines", { pageId: "p1", lines: ["a", "b"] });
  assert.notEqual(result.isError, true, textOf(result));
  const first = textOf(result).split("\n")[0];
  assert.match(first, /^applyToken: dh1_[0-9a-f]{32}$/);
  assert.equal(result.structuredContent.status, "dry-run");
  assert.equal(result.structuredContent.applyToken, first.slice("applyToken: ".length));
  assert.equal(applyCalls(bridge).length, 0);
  assert.deepEqual(
    bridge.calls.map((call) => call.args.apply),
    [false],
  );
  await t.close();
});

test("an apply without a valid token is refused before the bridge is reached", async () => {
  const bridge = fakeBridge({ append: previewHandler() });
  const t = await connect({ bridge });
  const args = { pageId: "p1", lines: ["a"] };
  const token = (await t.call("append_lines", args)).structuredContent.applyToken;

  const attempts = [
    { ...args, dryRun: false },
    { ...args, dryRun: false, applyToken: "dh1_forged" },
    { ...args, lines: ["changed"], dryRun: false, applyToken: token },
    { pageId: "other", lines: ["a"], dryRun: false, applyToken: token },
  ];
  for (const attempt of attempts) {
    const result = await t.call("append_lines", attempt);
    assert.equal(result.isError, true, JSON.stringify(attempt));
  }
  assert.equal(applyCalls(bridge).length, 0, "the bridge never saw an apply");
  await t.close();
});

test("a valid token applies once, with the previewed fingerprint, and cannot be replayed", async () => {
  const bridge = fakeBridge({ append: previewHandler("fp-A") });
  const t = await connect({ bridge });
  const args = { pageId: "p1", lines: ["a"] };
  const token = (await t.call("append_lines", args)).structuredContent.applyToken;

  const applied = await t.call("append_lines", { ...args, dryRun: false, applyToken: token });
  assert.notEqual(applied.isError, true, textOf(applied));
  assert.equal(applied.structuredContent.status, "applied");
  const sent = applyCalls(bridge);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].args.expectFingerprint, "fp-A");
  assert.deepEqual(Object.keys(sent[0].args).sort(), ["apply", "expectFingerprint", "lines", "pageId"]);

  const replay = await t.call("append_lines", { ...args, dryRun: false, applyToken: token });
  assert.equal(replay.isError, true);
  assert.equal(applyCalls(bridge).length, 1, "the replay never reached the bridge");
  await t.close();
});

test("a token issued for one tool does not apply another", async () => {
  const bridge = fakeBridge({ append: previewHandler(), delete_page: previewHandler() });
  const t = await connect({ bridge });
  const token = (await t.call("append_lines", { pageId: "p1", lines: ["a"] })).structuredContent.applyToken;
  const result = await t.call("delete_page", {
    pageId: "p1",
    expectName: "Probe",
    dryRun: false,
    applyToken: token,
  });
  assert.equal(result.isError, true);
  assert.equal(applyCalls(bridge).length, 0);
  await t.close();
});

test("every write tool follows the same dry-run then apply flow", async () => {
  const cases = {
    create_page: [{ sectionId: "s1", title: "T", lines: ["x"] }, "create_page"],
    append_lines: [{ pageId: "p1", lines: ["x"] }, "append"],
    insert_table: [{ pageId: "p1", rows: [["a", "b"]] }, "insert_table"],
    edit: [{ pageId: "p1", set: [{ objectID: "oe-1", text: "new" }] }, "edit"],
    delete_page: [{ pageId: "p1", expectName: "Probe" }, "delete_page"],
  };
  for (const [tool, [args, op]] of Object.entries(cases)) {
    const bridge = fakeBridge({ [op]: previewHandler(`fp-${tool}`) });
    const t = await connect({ bridge });
    const preview = await t.call(tool, args);
    assert.notEqual(preview.isError, true, `${tool}: ${textOf(preview)}`);
    const applied = await t.call(tool, {
      ...args,
      dryRun: false,
      applyToken: preview.structuredContent.applyToken,
    });
    assert.notEqual(applied.isError, true, `${tool}: ${textOf(applied)}`);
    assert.equal(applyCalls(bridge)[0].args.expectFingerprint, `fp-${tool}`, tool);
    await t.close();
  }
});

test("edit accepts only the whitelisted fields, so raw xml cannot be smuggled in", async () => {
  const bridge = fakeBridge({ edit: previewHandler() });
  const t = await connect({ bridge });
  const smuggled = [
    { pageId: "p1", append: [{ objectID: "ol-1", lines: ["x"], xml: ["<one:OE/>"] }] },
    { pageId: "p1", set: [{ objectID: "oe-1", text: "x", xml: "<one:OE/>" }] },
  ];
  for (const args of smuggled) {
    const result = await t.call("edit", args);
    assert.equal(result.isError, true, JSON.stringify(args));
  }
  assert.equal(bridge.calls.length, 0, "rejected by schema validation, before any bridge call");
  await t.close();
});

test("bridge failures surface as tool errors with their code", async () => {
  const bridge = fakeBridge({
    append: (args) => {
      if (args.apply)
        throw new BridgeError("OneNote changed_since_preview: the target changed", "changed_since_preview");
      return previewHandler()(args);
    },
  });
  const t = await connect({ bridge });
  const args = { pageId: "p1", lines: ["a"] };
  const token = (await t.call("append_lines", args)).structuredContent.applyToken;
  const result = await t.call("append_lines", { ...args, dryRun: false, applyToken: token });
  assert.equal(result.isError, true);
  assert.match(textOf(result), /changed_since_preview/);
  await t.close();
});

test("insert_image binds the token to the file contents", async () => {
  const { writeFileSync, mkdtempSync } = await import("node:fs");
  const { tmpdir } = await import("node:os");
  const { join } = await import("node:path");
  const dir = mkdtempSync(join(tmpdir(), "deskhand-"));
  const path = join(dir, "a.png");
  writeFileSync(path, Buffer.from("first"));

  const bridge = fakeBridge({ insert_image: previewHandler() });
  const t = await connect({ bridge });
  const args = { pageId: "p1", path };
  const token = (await t.call("insert_image", args)).structuredContent.applyToken;
  writeFileSync(path, Buffer.from("second"));
  const result = await t.call("insert_image", { ...args, dryRun: false, applyToken: token });
  assert.equal(result.isError, true, "the file changed after the preview");
  assert.equal(applyCalls(bridge).length, 0);
  await t.close();
});
