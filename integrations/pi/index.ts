/**
 * deskhand gate - asks the user before a OneNote write made through the deskhand MCP server runs.
 *
 * Every write tool on that server is two-step: a dry run returns an `applyToken` plus the exact
 * change set, and the apply repeats the call with `dryRun: false` and that token. The server
 * already refuses applies without a valid token, so this handler is a human confirmation layer,
 * not the safety mechanism: it answers the apply request and shows what the dry run said.
 */

import type {
	ExtensionAPI,
	ExtensionContext,
	ToolCallEvent,
	ToolResultEvent,
} from "@earendil-works/pi-coding-agent";

const TOOL_PREFIX = "mcp__onenote__";
const WRITE_TOOL_PATTERN =
	/^mcp__onenote__(create_page|append_lines|insert_table|insert_image|edit|delete_page|batch)$/;
const DIALOG_TITLE = "deskhand: apply OneNote write?";
const DECLINE_REASON =
	"The user declined this OneNote write. Do not retry unless the user asks you to.";

/** The dry-run reply starts with this line; anything else is not a cached preview. */
const APPLY_TOKEN_LINE = /^applyToken: (dh1_[0-9a-f]+)$/;

const PREVIEW_TTL_MS = 15 * 60 * 1000;
const PREVIEW_CACHE_LIMIT = 64;
const PREVIEW_MAX_CHARS = 3_000;
const STRING_MAX_CHARS = 300;
const MESSAGE_MAX_CHARS = 4_000;

// --- preview cache ----------------------------------------------------------

interface PreviewEntry {
	text: string;
	expiresAt: number;
}

/** Insertion order is age order, so the first key is the oldest entry. */
const previews = new Map<string, PreviewEntry>();

function pruneExpired(now: number): void {
	for (const [token, entry] of previews) {
		if (entry.expiresAt <= now) previews.delete(token);
	}
}

function rememberPreview(token: string, text: string): void {
	const now = Date.now();
	previews.delete(token);
	pruneExpired(now);
	while (previews.size >= PREVIEW_CACHE_LIMIT) {
		const oldest = previews.keys().next();
		if (oldest.done) break;
		previews.delete(oldest.value);
	}
	previews.set(token, { text, expiresAt: now + PREVIEW_TTL_MS });
}

function lookupPreview(token: string): string | undefined {
	const entry = previews.get(token);
	if (!entry) return undefined;
	if (entry.expiresAt <= Date.now()) {
		previews.delete(token);
		return undefined;
	}
	return entry.text;
}

/**
 * Pulls the dry-run preview out of a write tool result: the text without its first line (the
 * token, which must never reach the dialog) and without the trailing "To apply, call ..." line.
 */
function previewFromResult(
	event: ToolResultEvent,
): { token: string; preview: string } | undefined {
	if (!WRITE_TOOL_PATTERN.test(event.toolName)) return undefined;

	const parts: string[] = [];
	for (const block of event.content) {
		if (block.type === "text") parts.push(block.text);
	}
	const lines = parts.join("\n").split("\n");

	const match = APPLY_TOKEN_LINE.exec((lines[0] ?? "").trim());
	if (!match) return undefined;

	const body = lines.slice(1);
	while (body.length && body[body.length - 1].trim() === "") body.pop();
	if (body.length && body[body.length - 1].trim().startsWith("To apply, call "))
		body.pop();

	const preview = clip(body.join("\n").trim(), PREVIEW_MAX_CHARS);
	if (!preview) return undefined;
	return { token: match[1], preview };
}

// --- summary ----------------------------------------------------------------

function clip(text: string, max = STRING_MAX_CHARS): string {
	return text.length <= max
		? text
		: `${text.slice(0, max)}... (${text.length} characters)`;
}

function quoted(value: unknown): string | undefined {
	return typeof value === "string" ? JSON.stringify(clip(value)) : undefined;
}

/** Quoted when it is a string, stringified otherwise, so a JSON escape can never leak. */
function quoteAny(value: unknown): string {
	return quoted(value) ?? JSON.stringify(clip(String(value)));
}

function ids(values: unknown): string[] {
	if (!Array.isArray(values)) return [];
	return values
		.filter((value): value is string => typeof value === "string")
		.map((value) => clip(value));
}

function listBlock(label: string, items: string[]): string[] {
	return items.length ? [`${label}:`, ...items.map((item) => `  ${item}`)] : [];
}

function inputOf(event: ToolCallEvent): Record<string, unknown> {
	return event.input as unknown as Record<string, unknown>;
}

/** Target of the call, as an allowlist so no token or payload can leak into the dialog. */
function describeTarget(input: Record<string, unknown>): string | undefined {
	const pageId = quoted(input.pageId);
	if (pageId) return `target: pageId ${pageId}`;

	const target = [
		["notebook", quoted(input.notebook)],
		["section", quoted(input.section)],
		["page", quoted(input.page)],
	]
		.filter(([, value]) => value !== undefined)
		.map(([label, value]) => `${label} ${value}`);

	if (target.length) return `target: ${target.join(" / ")}`;

	const sectionId = quoted(input.sectionId);
	return sectionId ? `target: sectionId ${sectionId}` : undefined;
}

function describeLines(value: unknown): string[] {
	if (!Array.isArray(value)) return [];
	return value.map((item) => quoteAny(item));
}

function summarizeRecordList(
	value: unknown,
	render: (item: Record<string, unknown>) => string,
): string[] {
	if (!Array.isArray(value)) return [];
	const out: string[] = [];
	for (const item of value) {
		if (item === null || typeof item !== "object") continue;
		out.push(render(item as Record<string, unknown>));
	}
	return out;
}

/** Compact argument summary, used when no cached dry run explains the change. */
function summarizeWrite(
	toolName: string,
	input: Record<string, unknown>,
): string {
	const lines: string[] = [];
	const target = describeTarget(input);
	if (target) lines.push(target);

	switch (toolName) {
		case "create_page": {
			const title = quoted(input.title);
			if (title) lines.push(`title: ${title}`);
			lines.push(...listBlock("lines", describeLines(input.lines)));
			break;
		}
		case "append_lines": {
			lines.push(...listBlock("lines", describeLines(input.lines)));
			break;
		}
		case "insert_table": {
			const headers = describeLines(input.headers);
			if (headers.length) lines.push(`headers: ${headers.join(" | ")}`);
			const rows = Array.isArray(input.rows) ? input.rows : [];
			lines.push(
				...listBlock(
					`rows (${rows.length})`,
					rows.map((row) => describeLines(row).join(" | ")),
				),
			);
			break;
		}
		case "insert_image": {
			const path = quoted(input.path);
			if (path) lines.push(`path: ${path}`);
			const outline = quoted(input.outline);
			if (outline) lines.push(`outline: ${outline}`);
			break;
		}
		case "edit": {
			lines.push(
				...listBlock(
					"set",
					summarizeRecordList(
						input.set,
						(item) =>
							`${clip(String(item.objectID))} -> ${quoteAny(item.text)}`,
					),
				),
			);
			lines.push(...listBlock("remove", ids(input.remove)));

			const append: string[] = [];
			for (const item of Array.isArray(input.append) ? input.append : []) {
				if (item === null || typeof item !== "object") continue;
				const entry = item as Record<string, unknown>;
				append.push(`  ${clip(String(entry.objectID))}:`);
				for (const line of describeLines(entry.lines))
					append.push(`    ${line}`);
			}
			if (append.length) lines.push("append:", ...append);
			break;
		}
		case "delete_page": {
			const expectName = quoted(input.expectName);
			if (expectName) lines.push(`expectName: ${expectName}`);
			break;
		}
		case "batch": {
			// Only when the dry-run preview is not cached: the preview already lists every change.
			const ops = Array.isArray(input.ops) ? input.ops : [];
			lines.push(`ops (${ops.length}):`);
			for (const [index, op] of ops.entries()) {
				if (op === null || typeof op !== "object") continue;
				const entry = op as Record<string, unknown>;
				const tool = typeof entry.tool === "string" ? entry.tool : "?";
				const target = describeTarget(entry);
				lines.push(`  ${index + 1}. ${tool}${target ? ` (${target})` : ""}`);
			}
			break;
		}
	}

	return lines.join("\n");
}

function buildMessage(event: ToolCallEvent): string {
	const shortName = event.toolName.slice(TOOL_PREFIX.length);
	const input = inputOf(event);
	const token = typeof input.applyToken === "string" ? input.applyToken : "";
	const preview = token ? lookupPreview(token) : undefined;
	const message = [shortName, preview ?? summarizeWrite(shortName, input)].join(
		"\n\n",
	);
	if (message.length <= MESSAGE_MAX_CHARS) return message;
	return `${message.slice(0, MESSAGE_MAX_CHARS)}\n... [message truncated]`;
}

// --- confirmation -----------------------------------------------------------

/**
 * Codemode scripts may issue nested MCP calls concurrently, so the dialogs are queued: the next
 * one starts only after the previous one settles, and a rejected queue link never breaks the chain.
 */
let dialogQueue: Promise<unknown> = Promise.resolve();

function serializeDialog<T>(task: () => Promise<T>): Promise<T> {
	const run = dialogQueue.then(task, task);
	dialogQueue = run.then(
		() => undefined,
		() => undefined,
	);
	return run;
}

async function askUser(
	ctx: ExtensionContext,
	message: string,
): Promise<boolean> {
	try {
		return (await ctx.ui.confirm(DIALOG_TITLE, message)) === true;
	} catch {
		// A dialog that cannot be shown must not count as approval.
		return false;
	}
}

function isApply(event: ToolCallEvent): boolean {
	if (!WRITE_TOOL_PATTERN.test(event.toolName)) return false;
	const input = inputOf(event);
	if (input.dryRun === false) return true;
	return typeof input.applyToken === "string" && input.applyToken !== "";
}

export default function (pi: ExtensionAPI) {
	pi.on("tool_result", (event) => {
		const found = previewFromResult(event);
		if (found) rememberPreview(found.token, found.preview);
	});

	pi.on("tool_call", async (event, ctx) => {
		// Dry runs, read tools and every other server pass through untouched.
		if (!isApply(event)) return;
		// A non-interactive run counts as already approved, as the desktop extension did.
		if (!ctx.hasUI) return;

		const message = buildMessage(event);
		if (await serializeDialog(() => askUser(ctx, message))) return;
		return { block: true, reason: DECLINE_REASON };
	});
}
