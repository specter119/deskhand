export const MAX_TEXT_CHARS = 40_000;
export function asArray(value) {
    if (value === undefined || value === null)
        return [];
    return Array.isArray(value) ? value : [value];
}
export function asString(value) {
    return typeof value === "string" ? value : "";
}
export function truncate(text) {
    if (text.length <= MAX_TEXT_CHARS)
        return text;
    return `${text.slice(0, MAX_TEXT_CHARS)}\n... truncated at ${MAX_TEXT_CHARS} characters`;
}
/** PowerShell 5.1 collapses one-element arrays in places; make every list a list again. */
export function normalizePage(raw) {
    const page = (raw ?? {});
    return {
        ...page,
        lines: asArray(page.lines),
        outlines: asArray(page.outlines),
        images: asArray(page.images),
        tables: asArray(page.tables),
    };
}
export function formatPage(page) {
    const lines = [
        `page: ${asString(page.name)}`,
        `id: ${asString(page.id)}`,
        `lastModified: ${asString(page.lastModified)} (informational)`,
    ];
    for (const table of asArray(page.tables)) {
        lines.push(`table ${asString(table.objectID)}: ${table.rows}x${table.columns} cells, in outline ${asString(table.outlineID) || "(page level)"}`);
    }
    for (const outline of asArray(page.outlines)) {
        const geometry = outline.x !== null && outline.x !== undefined ? ` at (${outline.x},${outline.y})` : "";
        lines.push(`outline ${asString(outline.objectID)}${outline.inTitle ? " (title)" : ""}${geometry}, ${outline.lineCount} paragraph(s)`);
    }
    for (const image of asArray(page.images)) {
        const size = image.width ? ` ${Math.round(Number(image.width))}x${Math.round(Number(image.height))}` : "";
        lines.push(`image${size}${asString(image.alt) ? ` alt="${asString(image.alt)}"` : ""} in outline ${asString(image.outlineID) || "(page level)"}${image.embedded ? "" : " - NO BITS"}`);
    }
    lines.push("content:");
    const pageLines = asArray(page.lines);
    for (const line of pageLines) {
        const marks = [];
        if (line.inTitle)
            marks.push("title");
        if (line.inTable)
            marks.push("table-cell");
        if (line.hasTable)
            marks.push("table");
        if (line.hasImage)
            marks.push("image");
        if (line.hasInk)
            marks.push("ink");
        if (line.list)
            marks.push(line.list);
        if (line.depth)
            marks.push(`depth=${line.depth}`);
        const indent = "  ".repeat(Math.min(line.depth ?? 0, 6));
        const body = line.text.includes("\n") ? line.text.replace(/\n/g, " / ") : line.text;
        const suffix = [marks.length ? `[${marks.join(",")}]` : "", line.objectID ? `oe=${line.objectID}` : ""]
            .filter(Boolean)
            .join(" ");
        lines.push(`  [${line.index}] ${indent}${body || "(no text)"}${suffix ? `   ${suffix}` : ""}`);
    }
    if (!pageLines.length)
        lines.push("  (empty page)");
    return truncate(lines.join("\n"));
}
export function formatPageList(pages) {
    const list = asArray(pages);
    if (!list.length)
        return "no pages";
    // No lastModified here on purpose: OneNote's hierarchy API reports a stale one.
    return list
        .map((page) => `- ${asString(page.name)}  ${asString(page.id)}  created=${asString(page.created)}`)
        .join("\n");
}
export function formatEditChanges(data) {
    const changes = asArray(data.changes);
    if (!changes.length)
        return "no change";
    const out = [];
    for (const change of changes) {
        out.push(`outline ${change.outlineId}`);
        for (const item of asArray(change.changed))
            out.push(`  ~ "${item.before}"  ->  "${item.after}"`);
        for (const item of asArray(change.removed))
            out.push(`  - delete "${item.text}"`);
        out.push(`  resulting paragraphs: ${asArray(change.after)
            .map((entry) => JSON.stringify(entry))
            .join(", ")}`);
    }
    return out.join("\n");
}
export function describeLines(lines) {
    return lines.map((line, index) => `  ${index + 1}. ${line}`).join("\n");
}
