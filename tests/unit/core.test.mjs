import assert from "node:assert/strict";
import { test } from "node:test";
import { hashRequest, stableStringify, TokenError, TokenTable } from "../../dist/core/apply-token.js";
import { newCallStem } from "../../dist/core/powershell.js";

test("call stems never collide across calls and carry the pid", () => {
  const stems = new Set(Array.from({ length: 2000 }, () => newCallStem()));
  assert.equal(stems.size, 2000);
  for (const stem of stems) assert.ok(stem.startsWith(`${process.pid}-`));
});

test("stableStringify ignores key order and undefined members", () => {
  assert.equal(
    stableStringify({ b: 1, a: [{ d: 1, c: undefined, e: 2 }] }),
    stableStringify({ a: [{ e: 2, d: 1 }], b: 1 }),
  );
  assert.notEqual(hashRequest("edit", { a: 1 }), hashRequest("append_lines", { a: 1 }));
});

test("a token is single use, bound to its request, and expires", () => {
  let now = 1_000;
  const table = new TokenTable(10_000, () => now);
  const request = { pageId: "p", lines: ["x"] };
  const token = table.issue("append_lines", request, "fp");

  assert.throws(
    () => table.consume(undefined, "append_lines", request),
    (e) => e instanceof TokenError && e.reason === "missing",
  );
  assert.throws(
    () => table.consume("dh1_forged", "append_lines", request),
    (e) => e.reason === "unknown",
  );
  assert.throws(
    () => table.consume(token, "append_lines", { ...request, lines: ["y"] }),
    (e) => e.reason === "mismatch",
  );
  assert.throws(
    () => table.consume(token, "edit", request),
    (e) => e.reason === "mismatch",
  );

  assert.equal(table.consume(token, "append_lines", request), "fp");
  assert.throws(
    () => table.consume(token, "append_lines", request),
    (e) => e.reason === "used",
  );

  const late = table.issue("append_lines", request, "fp");
  now += 10_001;
  assert.throws(
    () => table.consume(late, "append_lines", request),
    (e) => e.reason === "expired",
  );
});

test("the table stays bounded", () => {
  const table = new TokenTable();
  const first = table.issue("t", { n: 0 }, "fp");
  for (let n = 1; n < 400; n++) table.issue("t", { n }, "fp");
  assert.throws(
    () => table.consume(first, "t", { n: 0 }),
    (e) => e.reason === "unknown",
  );
});
