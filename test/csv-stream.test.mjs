import test from "node:test";
import assert from "node:assert/strict";
import { createCSVStream, sheetToCSV } from "../src/csv.js";
import { normalizeSheet } from "../src/json.js";

/** Drain a stream into a string, counting chunks. */
async function collect(stream) {
  let text = "";
  let chunks = 0;
  for await (const chunk of stream) {
    text += chunk;
    chunks++;
  }
  return { text, chunks };
}

const SHEET = normalizeSheet([
  ["name", "score", "when", "ok"],
  ["ada", 42, "2024-03-15", true],
  ['says "hi", twice', 7.5, "1999-12-31", false],
  [null, null, null, null],
  ["=1+1", -2.5, "", null],
]);

test("createCSVStream: byte-identical to sheetToCSV under every option set", async () => {
  const optionSets = [
    {},
    { blankrows: false },
    { alwaysQuote: true },
    { quotedStrings: true },
    { delimiter: ";" },
    { delimiter: "\t", quotedStrings: true },
    { escapeFormulae: false },
    { skipHidden: true },
    { bom: true },
    { header: ["A", "B", "C", "D"] },
    { defval: "n/a" },
  ];
  for (const opts of optionSets) {
    const want = sheetToCSV(SHEET, opts);
    const { text } = await collect(createCSVStream(SHEET, opts));
    assert.equal(text, want, "stream must match string for " + JSON.stringify(opts));
  }
});

test("createCSVStream: chunkRows does not change the bytes", async () => {
  const want = sheetToCSV(SHEET);
  for (const chunkRows of [1, 2, 3, 4, 512, 100000]) {
    const { text } = await collect(createCSVStream(SHEET, {}, { chunkRows }));
    assert.equal(text, want, "chunkRows " + chunkRows);
  }
});

test("createCSVStream: respects chunkRows, and terminates the last row", async () => {
  const rows = 20;
  const sheet = normalizeSheet(Array.from({ length: rows }, (_, i) => [String(i)]));
  const { text, chunks } = await collect(createCSVStream(sheet, {}, { chunkRows: 5 }));
  assert.equal(chunks, 4, "20 rows at 5 per chunk is 4 chunks");
  assert.ok(text.endsWith("\n"), "the last row is terminated, as RFC 4180 recommends");
  assert.equal(text.split("\n").filter(Boolean).length, rows);
});

test("createCSVStream: a quoted newline stays inside its field", async () => {
  const sheet = normalizeSheet([["a\nb", "plain"], ["c", "d"]]);
  const want = sheetToCSV(sheet);
  const { text } = await collect(createCSVStream(sheet, {}, { chunkRows: 1 }));
  assert.equal(text, want, "a field containing a newline must not be split across chunks");
  // Chunks break on row boundaries, so no chunk ends mid-field.
  const chunks = [];
  for await (const c of createCSVStream(sheet, {}, { chunkRows: 1 })) chunks.push(c);
  for (const c of chunks.slice(0, -1)) {
    assert.ok(c.endsWith("\n") || c.endsWith("\r\n"), "chunk ends on a row boundary: " + JSON.stringify(c));
  }
});

test("createCSVStream: refuses styled CSV, like the string form", async () => {
  await assert.rejects(async () => collect(createCSVStream(SHEET, { style_text: true })), /styled CSV/);
});

test("createCSVStream: an empty sheet yields nothing, or just the BOM", async () => {
  const empty = normalizeSheet([]);
  assert.equal((await collect(createCSVStream(empty))).text, "");
  assert.equal((await collect(createCSVStream(empty, { bom: true }))).text, "\ufeff");
});

test("createCSVStream: the output is produced in chunks, not as one string", async () => {
  // Memory is the claim, but heap numbers without a forced collection are noise,
  // and --expose-gc is not how the suite runs. Assert the structural property
  // that MAKES the memory claim true: the output arrives in bounded chunks, none
  // of which is the whole document.
  const rows = 40000;
  const sheet = normalizeSheet(Array.from({ length: rows }, (_, r) => [r, "x".repeat(40), r * 2]));
  let bytes = 0;
  let biggest = 0;
  let chunks = 0;
  for await (const chunk of createCSVStream(sheet, {}, { chunkRows: 256 })) {
    bytes += chunk.length;
    if (chunk.length > biggest) biggest = chunk.length;
    chunks++;
  }
  assert.ok(bytes > 1e6, "should have produced more than a megabyte: " + bytes);
  assert.ok(chunks > 10, "expected many chunks, saw " + chunks);
  // No single chunk may be anything like the whole output.
  assert.ok(biggest < bytes / 4, "largest chunk " + biggest + " of " + bytes + " bytes");
});
