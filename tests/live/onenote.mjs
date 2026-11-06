// Live checks against a real OneNote desktop client. Needs WSL with Windows interop.
//
//   DESKHAND_LIVE_NOTEBOOK="<existing notebook name>" npm run test:live
//
// Every page it creates is titled "deskhand-live-<pid>-..." in the section DESKHAND_LIVE_SECTION
// (default "deskhand-probe", created if missing) and is deleted again, also when a check fails.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { deflateSync } from "node:zlib";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

const run = promisify(execFile);
const NOTEBOOK = process.env.DESKHAND_LIVE_NOTEBOOK;
const SECTION = process.env.DESKHAND_LIVE_SECTION ?? "deskhand-probe";
if (!NOTEBOOK) {
  console.error("Set DESKHAND_LIVE_NOTEBOOK to the name of an existing notebook.");
  process.exit(2);
}
const CLI = fileURLToPath(new URL("../../dist/cli.js", import.meta.url));
const SCRIPT = readFileSync(new URL("./ensure-section.ps1", import.meta.url), "utf8");
const tag = `deskhand-live-${process.pid}-${Date.now().toString(36)}`;

const quote = (value) => `'${value.replace(/'/g, "''")}'`;

async function ensureSection() {
  const script = `$Notebook = ${quote(NOTEBOOK)}; $Section = ${quote(SECTION)}\n${SCRIPT}`;
  const { stdout } = await run("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-EncodedCommand",
    Buffer.from(script, "utf16le").toString("base64"),
  ]);
  return stdout.trim();
}

async function startServer() {
  const client = new Client({ name: "deskhand-live", version: "0" });
  await client.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, "onenote"] }));
  const call = async (name, args = {}) => {
    const result = await client.callTool({ name, arguments: args });
    return { ...result, text: result.content.map((block) => block.text ?? "").join("\n") };
  };
  // Dry run, then apply with the token it issued. Returns both results.
  const write = async (name, args) => {
    const dry = await call(name, args);
    assert.ok(!dry.isError, `${name} dry run failed: ${dry.text}`);
    assert.equal(dry.structuredContent.status, "dry-run");
    const token = dry.structuredContent.applyToken;
    assert.match(
      dry.text.split("\n")[0],
      new RegExp(`^applyToken: ${token}$`),
      "token must be the first line",
    );
    const applied = await call(name, { ...args, dryRun: false, applyToken: token });
    return { dry, applied, token };
  };
  return { client, call, write, close: () => client.close() };
}

function png() {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const byte of buf) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const body = Buffer.concat([Buffer.from(type), data]);
    const out = Buffer.alloc(body.length + 8);
    out.writeUInt32BE(data.length, 0);
    body.copy(out, 4);
    out.writeUInt32BE(crc(body), body.length + 4);
    return out;
  };
  const size = 16;
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header.set([8, 2, 0, 0, 0], 8);
  const rows = Buffer.concat(
    Array.from({ length: size }, () => Buffer.concat([Buffer.from([0]), Buffer.alloc(size * 3, 0x80)])),
  );
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    chunk("IDAT", deflateSync(rows)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const results = [];
async function check(name, fn) {
  try {
    await fn();
    results.push([name, true]);
    console.log(`ok   ${name}`);
  } catch (error) {
    results.push([name, false]);
    console.log(`FAIL ${name}\n     ${error.stack?.split("\n").slice(0, 4).join("\n     ")}`);
  }
}

const created = [];
const where = { notebook: NOTEBOOK, section: SECTION };
const sectionState = await ensureSection();
console.log(`section ${SECTION}: ${sectionState.split(" ")[0]}`);

const a = await startServer();
const b = await startServer();

const pageIdOf = async (title) => {
  const listing = await a.call("pages", where);
  return listing.structuredContent.pages.find((page) => page.name === title)?.id;
};
const readPage = async (server, pageId) =>
  (await server.call("read_page", { pageId })).structuredContent.page;

try {
  await check("lists ten tools with output schemas", async () => {
    const { tools } = await a.client.listTools();
    assert.equal(tools.length, 10);
    for (const tool of tools) assert.ok(tool.outputSchema, `${tool.name} has no outputSchema`);
  });

  await check("notebooks shows the live notebook and the probe section", async () => {
    const { structuredContent } = await a.call("notebooks");
    const notebook = structuredContent.notebooks.find((item) => item.name === NOTEBOOK);
    assert.ok(notebook, "notebook missing");
    assert.ok(
      notebook.sections.some((section) => section.name === SECTION),
      "section missing",
    );
  });

  const title = `${tag}-main`;
  let pageId;

  await check(
    "create_page: dry run writes nothing, apply without token is refused, apply creates",
    async () => {
      const dry = await a.call("create_page", { ...where, title, lines: ["first line", "keep line"] });
      assert.ok(!dry.isError, dry.text);
      assert.equal(await pageIdOf(title), undefined, "dry run created the page");

      const refused = await a.call("create_page", {
        ...where,
        title,
        lines: ["first line", "keep line"],
        dryRun: false,
      });
      assert.ok(refused.isError, "apply without a token must fail");
      assert.equal(await pageIdOf(title), undefined, "refused apply created the page");

      const applied = await a.call("create_page", {
        ...where,
        title,
        lines: ["first line", "keep line"],
        dryRun: false,
        applyToken: dry.structuredContent.applyToken,
      });
      assert.ok(!applied.isError, applied.text);
      pageId = await pageIdOf(title);
      assert.ok(pageId, "page not found after apply");
      created.push(pageId);

      const reused = await a.call("create_page", {
        ...where,
        title,
        lines: ["first line", "keep line"],
        dryRun: false,
        applyToken: dry.structuredContent.applyToken,
      });
      assert.ok(reused.isError, "a token must be single use");
      const count = (await a.call("pages", where)).structuredContent.pages.filter(
        (p) => p.name === title,
      ).length;
      assert.equal(count, 1, "reused token created a second page");
    },
  );

  await check("read_page returns paragraphs that validate against the output schema", async () => {
    const page = await readPage(a, pageId);
    assert.equal(page.name, title);
    assert.ok(page.lines.some((line) => line.text === "first line"));
  });

  await check("append_lines adds a paragraph and leaves the old ones", async () => {
    const { applied } = await a.write("append_lines", { pageId, lines: ["second line"] });
    assert.ok(!applied.isError, applied.text);
    const texts = (await readPage(a, pageId)).lines.map((line) => line.text);
    assert.ok(texts.includes("first line") && texts.includes("second line"), texts.join("|"));
  });

  const lineByText = async (text) => (await readPage(a, pageId)).lines.find((line) => line.text === text);

  await check("a second dry run of an unchanged page does not invalidate the first token", async () => {
    const target = await lineByText("keep line");
    const args = { pageId, set: [{ objectID: target.objectID, text: "kept" }] };
    const one = await a.call("edit", args);
    const two = await a.call("edit", args);
    assert.ok(!one.isError && !two.isError, one.text + two.text);
    assert.notEqual(one.structuredContent.applyToken, two.structuredContent.applyToken);
    const applied = await a.call("edit", {
      ...args,
      dryRun: false,
      applyToken: one.structuredContent.applyToken,
    });
    assert.ok(!applied.isError, applied.text);
    assert.ok(await lineByText("kept"));
  });

  await check("a token is refused when the outline changed after the dry run", async () => {
    // The fingerprint covers the outline text that would be written, so the concurrent change must
    // touch another paragraph of the same outline.
    const target = await lineByText("first line");
    const neighbour = await lineByText("kept");
    const staleArgs = { pageId, set: [{ objectID: target.objectID, text: "STALE" }] };
    const stale = await a.call("edit", staleArgs);
    const other = await a.write("edit", {
      pageId,
      set: [{ objectID: neighbour.objectID, text: "neighbour" }],
    });
    assert.ok(!other.applied.isError, other.applied.text);
    const refused = await a.call("edit", {
      ...staleArgs,
      dryRun: false,
      applyToken: stale.structuredContent.applyToken,
    });
    assert.ok(refused.isError);
    assert.match(refused.text, /changed_since_preview/);
    assert.ok(!(await lineByText("STALE")));
    assert.ok(await lineByText("neighbour"));
  });

  await check("a token for different arguments is refused", async () => {
    const dry = await a.call("append_lines", { pageId, lines: ["bound"] });
    const refused = await a.call("append_lines", {
      pageId,
      lines: ["something else"],
      dryRun: false,
      applyToken: dry.structuredContent.applyToken,
    });
    assert.ok(refused.isError);
    assert.ok(!(await readPage(a, pageId)).lines.some((line) => line.text === "something else"));
  });

  await check("edit: set replaces text, remove deletes a paragraph, raw xml is rejected", async () => {
    const first = await lineByText("first line");
    const kept = await lineByText("neighbour");
    assert.equal(first.outlineID, kept.outlineID, "both lines should share the first outline");
    const removeResult = await a.write("edit", { pageId, remove: [kept.objectID] });
    assert.ok(!removeResult.applied.isError, removeResult.applied.text);
    assert.ok(!(await lineByText("neighbour")));
    assert.ok(await lineByText("first line"));
    const rejected = await a.call("edit", {
      pageId,
      set: [{ objectID: first.objectID, text: "x", xml: "<one:OE/>" }],
    });
    assert.ok(rejected.isError, "raw xml must be rejected");
  });

  await check("insert_table adds a table", async () => {
    const { applied } = await a.write("insert_table", {
      pageId,
      headers: ["a", "b"],
      rows: [
        ["1", "2"],
        ["3", "4"],
      ],
      hasHeaderRow: true,
      bordersVisible: true,
    });
    assert.ok(!applied.isError, applied.text);
    const page = await readPage(a, pageId);
    assert.equal(page.tables.length, 1);
    assert.equal(page.tables[0].rows, 3);
    assert.equal(page.tables[0].columns, 2);
  });

  await check("insert_image adds an image bound to the file contents", async () => {
    const dir = mkdtempSync(join(tmpdir(), "deskhand-live-"));
    const file = join(dir, "probe.png");
    writeFileSync(file, png());
    const { applied } = await a.write("insert_image", { pageId, path: file, alt: "probe" });
    assert.ok(!applied.isError, applied.text);
    const page = await readPage(a, pageId);
    assert.equal(page.images.length, 1);
  });

  await check("two server processes work on separate pages at the same time", async () => {
    const titles = [`${tag}-a`, `${tag}-b`];
    const [ra, rb] = await Promise.all([
      a.write("create_page", { ...where, title: titles[0], lines: ["from a"] }),
      b.write("create_page", { ...where, title: titles[1], lines: ["from b"] }),
    ]);
    assert.ok(!ra.applied.isError, ra.applied.text);
    assert.ok(!rb.applied.isError, rb.applied.text);
    const ids = await Promise.all(titles.map(pageIdOf));
    assert.ok(ids.every(Boolean), "a page is missing");
    created.push(...ids);
    const reads = await Promise.all([
      ...Array.from({ length: 3 }, () => a.call("read_page", { pageId: ids[0] })),
      ...Array.from({ length: 3 }, () => b.call("read_page", { pageId: ids[1] })),
      a.call("pages", where),
      b.call("notebooks"),
    ]);
    for (const read of reads) assert.ok(!read.isError, read.text);
    const appends = await Promise.all([
      a.write("append_lines", { pageId: ids[0], lines: ["a2"] }),
      b.write("append_lines", { pageId: ids[1], lines: ["b2"] }),
    ]);
    for (const { applied } of appends) assert.ok(!applied.isError, applied.text);
    assert.ok((await readPage(a, ids[0])).lines.some((line) => line.text === "a2"));
    assert.ok((await readPage(b, ids[1])).lines.some((line) => line.text === "b2"));
  });

  await check("delete_page checks the name, then removes the page", async () => {
    const wrong = await a.call("delete_page", { pageId, expectName: "not the title" });
    assert.ok(wrong.isError, "a wrong expectName must fail");
    assert.ok(await pageIdOf(title));
    const { applied } = await a.write("delete_page", { pageId, expectName: title });
    assert.ok(!applied.isError, applied.text);
    assert.equal(await pageIdOf(title), undefined);
  });
} finally {
  for (const id of created) {
    const name = (await a.call("pages", where)).structuredContent.pages.find((page) => page.id === id)?.name;
    if (!name) continue;
    try {
      const { applied } = await a.write("delete_page", { pageId: id, expectName: name });
      if (applied.isError) console.log(`cleanup of ${name} failed: ${applied.text}`);
    } catch (error) {
      console.log(`cleanup of ${name} failed: ${error.message}`);
    }
  }
  await Promise.allSettled([a.close(), b.close()]);
}

const failed = results.filter(([, ok]) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} passed`);
process.exit(failed ? 1 : 0);
