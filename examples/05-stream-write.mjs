/**
 * Write a large workbook without building it in memory first.
 *
 * Shows: WorkbookWriter - addSheet/addRow/finish, automatic row numbering, the
 * row-order rule, and what the writer does and does not bound.
 *
 * Output goes to examples/out/ (gitignored). Run from anywhere:
 *   node examples/05-stream-write.mjs
 *   node --expose-gc examples/05-stream-write.mjs   # stable heap numbers
 */
import { WorkbookWriter } from "gridxcel";
import { mkdirSync, statSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const OUT = fileURLToPath(new URL("./out/", import.meta.url));
mkdirSync(OUT, { recursive: true });
const FILE = OUT + "streamed-250k.xlsx";
const ROWS = 250000;

const gc = globalThis.gc;

/* ------------------------------------------------------- append 250k rows */

if (gc) gc();
const before = process.memoryUsage().heapUsed;

const t0 = performance.now();
let writer = new WorkbookWriter();
let sheet = writer.addSheet("Rows");

for (let r = 1; r <= ROWS; r++) {
  // addRow(rowIndex, cells). Rows must arrive in ascending order.
  // Passing undefined auto-increments from the previous row, which is what a
  // loop over a source cursor normally wants.
  sheet.addRow(undefined, [r, "description " + r, r * 1.5, r % 3 === 0]);
}
const appendMs = performance.now() - t0;

if (gc) gc();
const afterRows = process.memoryUsage().heapUsed;
const rowHeap = afterRows - before;

// finish() assembles the OPC parts into one archive and returns its bytes.
const t1 = performance.now();
sheet.end();
let bytes = writer.finish();
const finishMs = performance.now() - t1;
writeFileSync(FILE, bytes);

if (gc) gc();
const afterFinish = process.memoryUsage().heapUsed;

console.log("rows written   :", sheet.rowCount);
console.log("append time    :", appendMs.toFixed(0), "ms  (" + Math.round(ROWS / (appendMs / 1000)).toLocaleString("en-US"), "rows/s)");
console.log("finish time    :", finishMs.toFixed(0), "ms");
console.log("archive        :", statSync(FILE).size, "bytes (" + (bytes.length / 1048576).toFixed(2), "MiB)");
console.log("");
console.log("heap while appending rows  :", (rowHeap / 1048576).toFixed(1), "MiB");
console.log("heap after finish()        :", ((afterFinish - before) / 1048576).toFixed(1), "MiB");

// Drop the writer and the whole build is collectable again.
bytes = null;
sheet = null;
writer = null;
if (gc) { gc(); gc(); }
const retained = process.memoryUsage().heapUsed - before;
console.log("retained after release     :", (retained / 1048576).toFixed(1), "MiB");
console.log(gc ? "" : "(run with --expose-gc for stable numbers)");
console.log("");
console.log("Rows are serialised to XML as they arrive, so no Cell object per cell is");
console.log("ever retained: peak heap tracks the bytes emitted, not a grid of cells,");
console.log("and all of it is collectable once the writer is dropped. That bounds the");
console.log("sheet representation - not the archive, which finish() still assembles");
console.log("in memory. See docs/LIMITATIONS.md.");

/* --------------------------------------------------- the ordering rule */

try {
  const w = new WorkbookWriter();
  const s = w.addSheet("Bad");
  s.addRow(5, [1]);
  s.addRow(2, [2]); // out of order
} catch (e) {
  console.log("\nout-of-order addRow throws", e.name + " " + e.code);
}
