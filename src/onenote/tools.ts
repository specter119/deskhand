import { readFileSync, statSync } from "node:fs";
import { extname } from "node:path";
import type { McpServer, ToolAnnotations } from "@modelcontextprotocol/server";
import { Type } from "typebox";
import type { TokenTable } from "../core/apply-token.js";
import type { BridgeFn } from "../core/powershell.js";
import { defineTool, type ToolContext, type ToolOutput } from "../core/server.js";
import {
  asArray,
  asString,
  describeLines,
  formatEditChanges,
  formatPage,
  formatPageList,
  normalizePage,
  truncate,
} from "./format.js";
import {
  NotebookSchema,
  PageMetaSchema,
  PageSchema,
  TARGET_PARAMS,
  WRITE_PARAMS,
  WriteResultSchema,
} from "./schemas.js";

/** The largest image verified end to end was 5.4 MB of base64; cap well above that, not at it. */
const MAX_IMAGE_BYTES = 32 * 1024 * 1024;

const READ_ONLY: ToolAnnotations = { readOnlyHint: true, destructiveHint: false, openWorldHint: false };
const ADDITIVE: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
  openWorldHint: false,
};
/** Resends whole outlines, so a wrong request would delete paragraphs: treat as destructive. */
const REWRITING: ToolAnnotations = {
  readOnlyHint: false,
  destructiveHint: true,
  idempotentHint: false,
  openWorldHint: false,
};

const WRITE_FLOW =
  "Two steps: call it without dryRun first (nothing is written; the reply starts with a line 'applyToken: <token>' and shows the exact change), " +
  "then call it again with the same arguments plus dryRun=false and that applyToken. A token is single use, expires after 10 minutes, " +
  "and is refused if the arguments differ or the target changed in OneNote after the dry run.";

export interface OneNoteDeps {
  bridge: BridgeFn;
  tokens: TokenTable;
}

type Params = Record<string, unknown>;
type WriteStatus = { status: "dry-run" | "applied"; [key: string]: unknown };

function pickTarget(params: Params): Params {
  const target: Params = {};
  for (const key of ["pageId", "notebook", "section", "page"]) {
    if (typeof params[key] === "string" && params[key]) target[key] = params[key];
  }
  return target;
}

function targetLabel(target: Params): string {
  return (
    asString(target.pageId) ||
    `${asString(target.notebook) || "?"} / ${asString(target.section) || "?"} / ${asString(target.page) || "?"}`
  );
}

/** OneNote's image format pattern is auto|png|emf|jpg, so jpeg is spelled jpg and anything else gets auto. */
function imageFormatFor(path: string, requested?: string): string {
  if (requested) return requested.toLowerCase() === "jpeg" ? "jpg" : requested.toLowerCase();
  switch (extname(path).toLowerCase()) {
    case ".png":
      return "png";
    case ".jpg":
    case ".jpeg":
      return "jpg";
    case ".emf":
      return "emf";
    default:
      return "auto";
  }
}

interface WritePlan {
  tool: string;
  op: string;
  /** Exactly what is sent to the bridge, minus apply and expectFingerprint. The token is bound to it. */
  request: Params;
  intent: string;
  describe: (data: Params, applied: boolean) => string;
}

async function runWrite(
  deps: OneNoteDeps,
  ctx: ToolContext,
  plan: WritePlan,
  params: { dryRun?: boolean; applyToken?: string },
): Promise<ToolOutput<WriteStatus>> {
  if (params.dryRun !== false) {
    const data = await deps.bridge(plan.op, { ...plan.request, apply: false }, ctx.signal);
    const fingerprint = asString(data.fingerprint);
    if (!fingerprint) throw new Error("The OneNote bridge returned no fingerprint for this preview");
    const token = deps.tokens.issue(plan.tool, plan.request, fingerprint);
    const preview = plan.describe(data, false);
    const structured: WriteStatus = { status: "dry-run", applyToken: token, preview };
    if (data.page) structured.page = normalizePage(data.page);
    if (data.changes) structured.changes = asArray(data.changes);
    return {
      // The token goes first: long results are cut in the middle by some clients.
      text: [
        `applyToken: ${token}`,
        "Dry run, nothing was written.",
        plan.intent,
        "",
        preview,
        "",
        `To apply, call ${plan.tool} again with the same arguments plus dryRun=false and applyToken="${token}". The token is single use and expires in 10 minutes.`,
      ].join("\n"),
      structured,
    };
  }

  const fingerprint = deps.tokens.consume(params.applyToken, plan.tool, plan.request);
  const data = await deps.bridge(
    plan.op,
    { ...plan.request, apply: true, expectFingerprint: fingerprint },
    ctx.signal,
  );
  const page = data.page ? normalizePage(data.page) : undefined;
  const structured: WriteStatus = { status: "applied" };
  if (page) structured.page = page;
  if (data.changes) structured.changes = asArray(data.changes);
  return {
    text: truncate(page ? `${plan.describe(data, true)}\n\n${formatPage(page)}` : plan.describe(data, true)),
    structured,
  };
}

export function registerOneNoteTools(server: McpServer, deps: OneNoteDeps): void {
  defineTool(server, {
    name: "notebooks",
    title: "OneNote notebooks",
    description:
      "List the Microsoft OneNote notebooks and sections available in the user's Windows OneNote desktop client.",
    input: Type.Object({}),
    output: Type.Object({ notebooks: Type.Array(NotebookSchema) }, { additionalProperties: true }),
    annotations: READ_ONLY,
    async run(_args, ctx) {
      const data = await deps.bridge("notebooks", {}, ctx.signal);
      const notebooks = asArray<Params>(data.notebooks).map(
        (notebook): Params => ({ ...notebook, sections: asArray(notebook.sections) }),
      );
      const out = notebooks
        .map((notebook) => {
          const head = `${asString(notebook.name)}  ${asString(notebook.id)}`;
          const body = asArray<Params>(notebook.sections)
            .map((section) => `  - ${asString(section.name)}  ${asString(section.id)}`)
            .join("\n");
          return body ? `${head}\n${body}` : head;
        })
        .join("\n\n");
      return { text: truncate(out || "no notebooks"), structured: { notebooks } as never };
    },
  });

  defineTool(server, {
    name: "pages",
    title: "OneNote pages",
    description:
      "List the pages in one OneNote section. Pass sectionId, or notebook plus section names. Returns names, ids and creation times; for a page's live modification time use read_page.",
    input: Type.Object({
      sectionId: Type.Optional(Type.String({ description: "Section ID from notebooks" })),
      notebook: Type.Optional(Type.String({ description: "Notebook name" })),
      section: Type.Optional(
        Type.String({ description: "Section name, required when sectionId is omitted" }),
      ),
    }),
    output: Type.Object(
      { sectionId: Type.String(), section: Type.String(), pages: Type.Array(PageMetaSchema) },
      { additionalProperties: true },
    ),
    annotations: READ_ONLY,
    async run(args, ctx) {
      const data = await deps.bridge("pages", { ...args }, ctx.signal);
      const pages = asArray(data.pages);
      return {
        text: truncate(`section: ${asString(data.section)}\n${formatPageList(pages)}`),
        structured: { ...data, pages } as never,
      };
    },
  });

  defineTool(server, {
    name: "search",
    title: "Search OneNote",
    description:
      "Find OneNote pages by name across every notebook. Matching is a case-insensitive substring match on the page name; page bodies are not searched.",
    input: Type.Object({
      query: Type.String({ description: "Substring to look for in page names" }),
      limit: Type.Optional(Type.Number({ description: "Maximum matches to return (default 25)" })),
    }),
    output: Type.Object(
      {
        total: Type.Number(),
        returned: Type.Number(),
        pages: Type.Array(
          Type.Object(
            { pageId: Type.String(), page: Type.String(), section: Type.String(), notebook: Type.String() },
            { additionalProperties: true },
          ),
        ),
      },
      { additionalProperties: true },
    ),
    annotations: READ_ONLY,
    async run(args, ctx) {
      const data = await deps.bridge("search", { query: args.query, limit: args.limit }, ctx.signal);
      const pages = asArray<Params>(data.pages);
      const out = pages.length
        ? pages
            .map(
              (hit) =>
                `- ${asString(hit.notebook)} / ${asString(hit.section)} / ${asString(hit.page)}  ${asString(hit.pageId)}`,
            )
            .join("\n")
        : "no pages matched";
      const note = data.total !== data.returned ? `\n(${data.total} matches, showing ${data.returned})` : "";
      return { text: truncate(out + note), structured: { ...data, pages } as never };
    },
  });

  defineTool(server, {
    name: "read_page",
    title: "Read OneNote page",
    description:
      "Read one OneNote page: its paragraphs as text, each with the objectID needed to reference it in edit, plus its outlines, tables and images. The reported lastModified is informational and must not be used as a concurrency check.",
    input: Type.Object({
      ...TARGET_PARAMS,
      includeXml: Type.Optional(
        Type.Boolean({
          description:
            "Also return the raw OneNote XML. Large, and it includes the author's account metadata.",
        }),
      ),
    }),
    output: Type.Object({ page: PageSchema }, { additionalProperties: true }),
    annotations: READ_ONLY,
    async run(args, ctx) {
      const data = await deps.bridge(
        "read_page",
        { ...pickTarget(args), includeXml: args.includeXml === true },
        ctx.signal,
      );
      const page = normalizePage(data.page);
      const raw = asString(page.rawXml);
      return {
        text: formatPage(page) + (raw ? `\n\nraw xml:\n${raw}`.slice(0, 40_000) : ""),
        structured: { page } as never,
      };
    },
  });

  defineTool(server, {
    name: "create_page",
    title: "Create OneNote page",
    description: `Create a new page in a OneNote section with a title and optional body paragraphs, each its own paragraph in one outline. The dry run validates the section but cannot show the page, which does not exist yet. ${WRITE_FLOW}`,
    input: Type.Object({
      sectionId: Type.Optional(Type.String({ description: "Section ID from notebooks" })),
      notebook: Type.Optional(Type.String({ description: "Notebook name" })),
      section: Type.Optional(
        Type.String({ description: "Section name, required when sectionId is omitted" }),
      ),
      title: Type.String({ description: "Page title" }),
      lines: Type.Optional(
        Type.Array(Type.String({ description: "One paragraph" }), {
          description: "Body paragraphs, in order",
        }),
      ),
      ...WRITE_PARAMS,
    }),
    output: WriteResultSchema,
    annotations: ADDITIVE,
    async run(args, ctx) {
      const lines = asArray<string>(args.lines).filter((line) => typeof line === "string");
      const where = args.sectionId
        ? `section ${args.sectionId}`
        : `${args.notebook ?? "?"} / ${args.section ?? "?"}`;
      const intent = [
        `Create page "${args.title}" in ${where}`,
        lines.length
          ? `with ${lines.length} body paragraph(s):\n${describeLines(lines)}`
          : "with no body text",
      ].join("\n");
      return runWrite(
        deps,
        ctx,
        {
          tool: "create_page",
          op: "create_page",
          request: {
            sectionId: args.sectionId,
            notebook: args.notebook,
            section: args.section,
            title: args.title,
            lines,
          },
          intent,
          describe: (data, applied) =>
            applied
              ? `Created page ${asString((data.page as Params | undefined)?.id)}`
              : `Section "${asString(data.section)}" exists; the page would be created there.`,
        },
        args,
      );
    },
  });

  defineTool(server, {
    name: "append_lines",
    title: "Append to OneNote page",
    description: `Append paragraphs to a page as a new outline at the end. Existing content is not touched. Use this for adding to the end of a page; use edit to change text inside an existing outline. ${WRITE_FLOW}`,
    input: Type.Object({
      ...TARGET_PARAMS,
      lines: Type.Array(Type.String({ description: "One paragraph" }), {
        description: "Paragraphs to append, in order",
      }),
      ...WRITE_PARAMS,
    }),
    output: WriteResultSchema,
    annotations: ADDITIVE,
    async run(args, ctx) {
      const lines = asArray<string>(args.lines).filter((line) => typeof line === "string");
      if (!lines.length) throw new Error("append_lines needs at least one line");
      const target = pickTarget(args);
      const intent = `Append ${lines.length} paragraph(s) to ${targetLabel(target)} as a new outline:\n${describeLines(lines)}`;
      return runWrite(
        deps,
        ctx,
        {
          tool: "append_lines",
          op: "append",
          request: { ...target, lines },
          intent,
          describe: (data, applied) => {
            const page = data.page as Params;
            const paragraphs = asArray(page?.lines).length;
            const outlines = asArray(page?.outlines).length;
            return applied
              ? `Appended. The page now has ${paragraphs} paragraph(s) across ${outlines} outline(s).`
              : `Would add one outline with ${lines.length} paragraph(s). The page currently has ${paragraphs} paragraph(s) across ${outlines} outline(s).`;
          },
        },
        args,
      );
    },
  });

  defineTool(server, {
    name: "insert_table",
    title: "Insert OneNote table",
    description: `Insert a table into a page: give headers and rows as arrays of cell text. Without outline it becomes a new outline at the end of the page; with outline it is appended inside that existing outline. Cells are single paragraphs and cannot contain nested tables. ${WRITE_FLOW}`,
    input: Type.Object({
      ...TARGET_PARAMS,
      headers: Type.Optional(
        Type.Array(Type.String({ description: "One header cell" }), { description: "Optional header row" }),
      ),
      rows: Type.Array(Type.Array(Type.String({ description: "One cell" })), {
        description: "Rows of cell text",
      }),
      bordersVisible: Type.Optional(
        Type.Boolean({ description: "Draw cell borders; OneNote's default is false" }),
      ),
      hasHeaderRow: Type.Optional(Type.Boolean({ description: "Render the first row as a header row" })),
      columnWidths: Type.Optional(
        Type.Array(Type.Number({ description: "Width of one column" }), {
          description: "Optional per-column widths",
        }),
      ),
      outline: Type.Optional(
        Type.String({ description: "objectID of an existing outline to append the table to" }),
      ),
      ...WRITE_PARAMS,
    }),
    output: WriteResultSchema,
    annotations: ADDITIVE,
    async run(args, ctx) {
      const headers = asArray<string>(args.headers).filter((value) => typeof value === "string");
      const rows = asArray<unknown>(args.rows).map((row) =>
        asArray<string>(row).map((cell) => (typeof cell === "string" ? cell : String(cell ?? ""))),
      );
      if (!rows.length && !headers.length) throw new Error("insert_table needs headers or rows");
      const target = pickTarget(args);
      const columnCount = Math.max(headers.length, ...rows.map((row) => row.length));
      const rowCount = (headers.length ? 1 : 0) + rows.length;
      const intent = [
        `Insert a ${rowCount}x${columnCount} table into ${targetLabel(target)}`,
        args.outline ? `inside outline ${args.outline}` : "as a new outline at the end of the page",
        headers.length ? `header: ${headers.join(" | ")}` : "",
        ...rows.map((row) => `row: ${row.join(" | ")}`),
      ]
        .filter(Boolean)
        .join("\n");
      return runWrite(
        deps,
        ctx,
        {
          tool: "insert_table",
          op: "insert_table",
          request: {
            ...target,
            headers,
            rows,
            bordersVisible: args.bordersVisible === true,
            hasHeaderRow: args.hasHeaderRow === true,
            columnWidths: args.columnWidths,
            outline: args.outline,
          },
          intent,
          describe: (data, applied) => {
            const tables = asArray((data.page as Params | undefined)?.tables).length;
            return applied
              ? `Inserted the table. The page now has ${tables} table(s).`
              : `Would insert a ${rowCount}x${columnCount} table. The page currently has ${tables} table(s).`;
          },
        },
        args,
      );
    },
  });

  defineTool(server, {
    name: "insert_image",
    title: "Insert OneNote image",
    description: `Insert an image file into a page, read from a path this server can reach (a WSL path, or /mnt/c/... for a file on the Windows side). Without outline it becomes a new outline at the end of the page; with outline it is appended inside that existing outline. Without width and height OneNote picks the size. ${WRITE_FLOW}`,
    input: Type.Object({
      ...TARGET_PARAMS,
      path: Type.String({ description: "Image file path readable from this server" }),
      format: Type.Optional(
        Type.String({ description: "png, jpg, emf or auto; derived from the extension when omitted" }),
      ),
      width: Type.Optional(Type.Number({ description: "Rendered width; needs height too" })),
      height: Type.Optional(Type.Number({ description: "Rendered height; needs width too" })),
      alt: Type.Optional(Type.String({ description: "Alternative text stored on the image" })),
      outline: Type.Optional(
        Type.String({ description: "objectID of an existing outline to append the image to" }),
      ),
      ...WRITE_PARAMS,
    }),
    output: WriteResultSchema,
    annotations: ADDITIVE,
    async run(args, ctx) {
      const size = statSync(args.path).size;
      if (size > MAX_IMAGE_BYTES) {
        throw new Error(
          `${args.path} is ${size} bytes, above the ${MAX_IMAGE_BYTES}-byte limit this tool imposes`,
        );
      }
      const format = imageFormatFor(args.path, args.format);
      const target = pickTarget(args);
      const intent = [
        `Insert ${args.path} (${(size / 1024).toFixed(0)} KB, format ${format}) into ${targetLabel(target)}`,
        args.outline ? `inside outline ${args.outline}` : "as a new outline at the end of the page",
        args.width && args.height ? `rendered at ${args.width}x${args.height}` : "at OneNote's chosen size",
      ].join("\n");
      return runWrite(
        deps,
        ctx,
        {
          tool: "insert_image",
          op: "insert_image",
          request: {
            ...target,
            base64: readFileSync(args.path).toString("base64"),
            format,
            width: args.width,
            height: args.height,
            alt: args.alt,
            outline: args.outline,
          },
          intent,
          describe: (data, applied) => {
            const images = asArray<Params>((data.page as Params | undefined)?.images);
            const embedded = images.filter((image) => image.embedded === true).length;
            return applied
              ? `Inserted the image. OneNote reports ${images.length} image(s), ${embedded} with embedded bits.`
              : `Would insert one image. The page currently has ${images.length} image(s), ${embedded} with embedded bits.`;
          },
        },
        args,
      );
    },
  });

  defineTool(server, {
    name: "edit",
    title: "Edit OneNote page text",
    description: `Change paragraphs inside existing outlines of a page: replace their text with set, delete them with remove, or add paragraphs to an outline with append. Paragraph objectIDs come from read_page, and outline objectIDs may be used as an append target. OneNote replaces an outline with exactly what the caller sends, so this tool resends each affected outline in full, echoing back OneNote's own XML for everything it is not changing - that is why paragraph XML is never accepted from callers. ${WRITE_FLOW}`,
    input: Type.Object({
      ...TARGET_PARAMS,
      set: Type.Optional(
        Type.Array(
          Type.Object(
            {
              objectID: Type.String({ description: "objectID of the paragraph to replace" }),
              text: Type.String({ description: "New text for that paragraph" }),
            },
            { additionalProperties: false },
          ),
          { description: "Paragraphs whose text is replaced" },
        ),
      ),
      remove: Type.Optional(
        Type.Array(Type.String({ description: "objectID of a paragraph to delete" }), {
          description: "Paragraphs to delete",
        }),
      ),
      append: Type.Optional(
        Type.Array(
          Type.Object(
            {
              objectID: Type.String({
                description: "objectID of the outline (or of any paragraph in it) to add to",
              }),
              lines: Type.Array(Type.String({ description: "One paragraph" }), {
                description: "Paragraphs to add at the end of that outline",
              }),
            },
            { additionalProperties: false },
          ),
          { description: "Paragraphs to add inside an existing outline" },
        ),
      ),
      ...WRITE_PARAMS,
    }),
    output: WriteResultSchema,
    annotations: REWRITING,
    async run(args, ctx) {
      const set = asArray<{ objectID: string; text: string }>(args.set).map((item) => ({
        objectID: item.objectID,
        text: item.text,
      }));
      const remove = asArray<string>(args.remove).filter((value) => typeof value === "string");
      const append = asArray<{ objectID: string; lines: string[] }>(args.append).map((item) => ({
        objectID: item.objectID,
        lines: asArray<string>(item.lines),
      }));
      if (!set.length && !remove.length && !append.length)
        throw new Error("edit needs at least one of set, remove or append");
      const target = pickTarget(args);
      const intent = [
        `Edit ${targetLabel(target)}:`,
        ...set.map((item) => `  replace ${item.objectID} with ${JSON.stringify(item.text)}`),
        ...remove.map((objectID) => `  delete ${objectID}`),
        ...append.map((item) => `  add ${item.lines.length} paragraph(s) to ${item.objectID}`),
      ].join("\n");
      return runWrite(
        deps,
        ctx,
        {
          tool: "edit",
          op: "edit",
          request: { ...target, set, remove, append },
          intent,
          describe: (data, applied) =>
            `${applied ? "Changed" : "Would change"}: ${formatEditChanges(data)}\n\n${applied ? "OneNote now reports" : "The page currently has"} ${asArray((data.page as Params | undefined)?.lines).length} paragraph(s).`,
        },
        args,
      );
    },
  });

  defineTool(server, {
    name: "delete_page",
    title: "Delete OneNote page",
    description: `Delete a OneNote page. It goes to the notebook's recycle bin, so it can be restored from OneNote. expectName is required and must match the page the target resolves to, so a stale or wrong id cannot delete a different page. ${WRITE_FLOW}`,
    input: Type.Object({
      ...TARGET_PARAMS,
      expectName: Type.String({ description: "Exact page name the target must have; required safety check" }),
      ...WRITE_PARAMS,
    }),
    output: WriteResultSchema,
    annotations: REWRITING,
    async run(args, ctx) {
      const target = pickTarget(args);
      const intent = `Delete OneNote page ${targetLabel(target)}, which must be named "${args.expectName}". It goes to the recycle bin.`;
      return runWrite(
        deps,
        ctx,
        {
          tool: "delete_page",
          op: "delete_page",
          request: { ...target, expectName: args.expectName },
          intent,
          describe: (data, applied) =>
            applied
              ? `Deleted page "${asString(data.name)}" (${asString(data.pageId)}). It is in the notebook recycle bin.`
              : `Page "${asString(data.name)}" (${asString(data.pageId)}) would be deleted.`,
        },
        args,
      );
    },
  });
}
