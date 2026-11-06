export const MAX_TEXT_CHARS = 40_000;

export function asArray<T>(value: unknown): T[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? (value as T[]) : [value as T];
}

export function asString(value: unknown): string {
  return typeof value === "string" ? value : "";
}

export function truncate(text: string): string {
  if (text.length <= MAX_TEXT_CHARS) return text;
  return `${text.slice(0, MAX_TEXT_CHARS)}\n... truncated at ${MAX_TEXT_CHARS} characters`;
}

/** PowerShell 5.1 collapses one-element arrays in places; make every list a list again. */
export function normalizePage(raw: unknown): Record<string, unknown> {
  const page = (raw ?? {}) as Record<string, unknown>;
  return {
    ...page,
    lines: asArray(page.lines),
    outlines: asArray(page.outlines),
    images: asArray(page.images),
    tables: asArray(page.tables),
  };
}

interface PageLine {
  index: number;
  objectID?: string;
  text: string;
  inTitle?: boolean;
  inTable?: boolean;
  hasImage?: boolean;
  hasTable?: boolean;
  hasInk?: boolean;
  list?: string;
  depth?: number;
}

interface EditChange {
  outlineId: string;
  changed?: Array<{ objectID: string; before: string; after: string }>;
  removed?: Array<{ objectID: string; text: string }>;
  after?: string[];
}

export function formatPage(page: Record<string, unknown>): string {
  const lines: string[] = [
    `page: ${asString(page.name)}`,
    `id: ${asString(page.id)}`,
    `lastModified: ${asString(page.lastModified)} (informational)`,
  ];

  for (const table of asArray<Record<string, unknown>>(page.tables)) {
    lines.push(
      `table ${asString(table.objectID)}: ${table.rows}x${table.columns} cells, in outline ${asString(table.outlineID) || "(page level)"}`,
    );
  }

  for (const outline of asArray<Record<string, unknown>>(page.outlines)) {
    const geometry = outline.x !== null && outline.x !== undefined ? ` at (${outline.x},${outline.y})` : "";
    lines.push(
      `outline ${asString(outline.objectID)}${outline.inTitle ? " (title)" : ""}${geometry}, ${outline.lineCount} paragraph(s)`,
    );
  }

  for (const image of asArray<Record<string, unknown>>(page.images)) {
    const size = image.width ? ` ${Math.round(Number(image.width))}x${Math.round(Number(image.height))}` : "";
    lines.push(
      `image${size}${asString(image.alt) ? ` alt="${asString(image.alt)}"` : ""} in outline ${asString(image.outlineID) || "(page level)"}${image.embedded ? "" : " - NO BITS"}`,
    );
  }

  lines.push("content:");
  const pageLines = asArray<PageLine>(page.lines);
  for (const line of pageLines) {
    const marks: string[] = [];
    if (line.inTitle) marks.push("title");
    if (line.inTable) marks.push("table-cell");
    if (line.hasTable) marks.push("table");
    if (line.hasImage) marks.push("image");
    if (line.hasInk) marks.push("ink");
    if (line.list) marks.push(line.list);
    if (line.depth) marks.push(`depth=${line.depth}`);

    const indent = "  ".repeat(Math.min(line.depth ?? 0, 6));
    const body = line.text.includes("\n") ? line.text.replace(/\n/g, " / ") : line.text;
    const suffix = [marks.length ? `[${marks.join(",")}]` : "", line.objectID ? `oe=${line.objectID}` : ""]
      .filter(Boolean)
      .join(" ");
    lines.push(`  [${line.index}] ${indent}${body || "(no text)"}${suffix ? `   ${suffix}` : ""}`);
  }
  if (!pageLines.length) lines.push("  (empty page)");

  return truncate(lines.join("\n"));
}

export function formatPageList(pages: unknown): string {
  const list = asArray<Record<string, unknown>>(pages);
  if (!list.length) return "no pages";
  // No lastModified here on purpose: OneNote's hierarchy API reports a stale one.
  return list
    .map((page) => `- ${asString(page.name)}  ${asString(page.id)}  created=${asString(page.created)}`)
    .join("\n");
}

export function formatEditChanges(data: Record<string, unknown>): string {
  const changes = asArray<EditChange>(data.changes);
  if (!changes.length) return "no change";
  const out: string[] = [];
  for (const change of changes) {
    out.push(`outline ${change.outlineId}`);
    for (const item of asArray<{ before: string; after: string }>(change.changed))
      out.push(`  ~ "${item.before}"  ->  "${item.after}"`);
    for (const item of asArray<{ text: string }>(change.removed)) out.push(`  - delete "${item.text}"`);
    out.push(
      `  resulting paragraphs: ${asArray<string>(change.after)
        .map((entry) => JSON.stringify(entry))
        .join(", ")}`,
    );
  }
  return out.join("\n");
}

export function describeLines(lines: string[]): string {
  return lines.map((line, index) => `  ${index + 1}. ${line}`).join("\n");
}
