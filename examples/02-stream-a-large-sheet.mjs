/**
 * Stream a worksheet row by row instead of materialising it.
 *
 * Shows: the eachRow() event shape, stopping early, range/sheet selection, why
 * row numbers have gaps, and the difference between streaming a large sheet and
 * parsing it whole.
 *
 * The first half reads a committed fixture. The second half builds a
 * 100,000-row workbook into examples/out/ (gitignored) and streams that, so the
 * example does not depend on any large file being checked in.
 *
 * Run from anywhere:
 *   node examples/02-stream-a-large-sheet.mjs
 *   node --expose-gc examples/02-stream-a-large-sheet.mjs   # stable heap numbers
 */
import { eachRow, read, WorkbookWriter } from "gridxcel";
import { mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const FIXTURE = new URL("../test/fixtures/golden.xlsx", import.meta.url);
const OUT = fileURLToPath(new URL("./out/", import.meta.url));

/* ------------------------------------------------------- the event shape */

// eachRow() is an async iterable of { sheet, r, cells }. r is a 1-based sheet
// row number, so it survives gaps in the file. cells[0] is always column A,
// even when the batch reader would have rebased to originCol.
let rows = 0;
let first = null;
for await (const event of eachRow(FIXTURE, { sheets: ["Volume"] })) {
  if (first === null) first = event;
  rows++;
}
console.log("Volume events     :", rows);
console.log("first event       : sheet=" + first.sheet + " r=" + first.r);
console.log("first event cells :", JSON.stringify(first.cells.map((c) => (c ? c.v : null))));

/* -------------------------------------------------------- stopping early */

// Breaking out of the loop tears the reader down. You never inflate the rest
// of the archive, so cost tracks what you consume rather than what is stored.
let seen = 0;
for await (const event of eachRow(FIXTURE, { sheets: ["Volume"] })) {
  seen++;
  if (event.r >= 3) break;
}
console.log("\nrows consumed before break:", seen);

/* ---------------------------------------------------- range and sheets  */

// sheets picks by name or 0-based index; range bounds the scan.
const narrow = [];
for await (const event of eachRow(FIXTURE, { sheets: [2], range: "A1:B3" })) {
  narrow.push(event.r + ":" + JSON.stringify(event.cells.map((c) => (c ? c.v : null))));
}
console.log("range A1:B3       :", narrow.join("  "));

/* ------------------------------------------ only rows that are present  */

// Layout has data on rows 1, 7 and 8. A batch read returns eight entries with
// nulls in between; the stream yields three events and r jumps. Treat r as the
// row number and never as an array index.
const layoutRows = [];
for await (const event of eachRow(FIXTURE, { sheets: ["Layout"] })) layoutRows.push(event.r);
console.log("Layout row numbers:", layoutRows.join(", "), "(gaps are real gaps)");

/* ------------------------------------------------- a file worth streaming */

mkdirSync(OUT, { recursive: true });
const OUT_FILE = OUT + "large-100k.xlsx";

const t0 = performance.now();
const writer = new WorkbookWriter();
const sheetWriter = writer.addSheet("Rows");
for (let r = 1; r <= 100000; r++) {
  // addRow(index, cells). Rows must ascend; undefined auto-increments.
  sheetWriter.addRow(r, [r, "row " + r, r / 4, r % 2 === 0]);
}
sheetWriter.end();
const bytes = writer.finish();
writeFileSync(OUT_FILE, bytes); // finish() hands back the whole archive, so it is written in one go
const writeMs = performance.now() - t0;
console.log("\nbuilt", OUT_FILE.replace(/\\/g, "/").split("/examples/")[1], "-", bytes.length, "bytes in", writeMs.toFixed(0) + " ms");

// Stream it. Memory stays flat no matter how many rows the file has.
const gc = globalThis.gc;
if (gc) gc();
const heapBefore = process.memoryUsage().heapUsed;
const t1 = performance.now();
let cells = 0;
let lastRow = null;
for await (const event of eachRow(OUT_FILE)) {
  cells += event.cells.filter(Boolean).length;
  lastRow = event;
}
const streamMs = performance.now() - t1;
const streamHeap = process.memoryUsage().heapUsed - heapBefore;
console.log("streamed", lastRow.r, "rows /", cells, "cells in", streamMs.toFixed(0) + " ms");
console.log("peak heap growth while streaming :", (streamHeap / 1048576).toFixed(1), "MiB");

// Now hold the same sheet in memory, for contrast.
if (gc) gc();
const heapBefore2 = process.memoryUsage().heapUsed;
const t2 = performance.now();
const book = await read(OUT_FILE);
const held = book.sheets[0].rows.filter(Boolean).reduce((n, r) => n + r.filter(Boolean).length, 0);
const readMs = performance.now() - t2;
const holdHeap = process.memoryUsage().heapUsed - heapBefore2;
console.log("parsed  ", book.sheets[0].rowCount, "rows /", held, "cells in", readMs.toFixed(0) + " ms");
console.log("peak heap growth while parsing   :", (holdHeap / 1048576).toFixed(1), "MiB");

console.log("\nSame file, same values. The streamed pass never held the sheet; the");
console.log("parsed pass has to, and pays for it in heap for as long as it keeps it.");
console.log(gc ? "" : "Run with --expose-gc for comparable heap numbers.");
