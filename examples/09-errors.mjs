/**
 * Handle the ways a read can fail, and the one guard that stops an untrusted
 * sheet from allocating gigabytes.
 *
 * Shows: the GridxcelError hierarchy and its stable code strings, the
 * difference between a warning and a throw, strict mode, and RANGE_TOO_LARGE.
 *
 * Run from anywhere:
 *   node examples/09-errors.mjs
 */
import {
  BadZipError,
  FormatError,
  GridxcelError,
  StrictWarningError,
  ZipError,
  read,
  sheetToAOA,
} from "gridxcel";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const FIXTURE_DIR = fileURLToPath(new URL("../test/fixtures/", import.meta.url));

// One helper so every case prints the same shape. code is stable across
// releases; message is written for a human and may change.
async function attempt(label, fn) {
  try {
    await fn();
    console.log("  " + label.padEnd(24), "-> no throw");
  } catch (err) {
    const kind = err instanceof GridxcelError ? err.name + " / " + err.code : "UNTYPED " + err.name;
    console.log("  " + label.padEnd(24), "->", kind);
  }
}

console.log("-- not a workbook --");
await attempt("missing file", () => read(FIXTURE_DIR + "nope.xlsx"));
await attempt("HTML error page", () => read(Buffer.from("<!doctype html><html>error page</html>")));
await attempt("zero bytes", () => read(new Uint8Array(0)));

// A real workbook that was cut short mid-download. This is the one that
// actually happens in production, and it must not look like an empty sheet.
const whole = readFileSync(FIXTURE_DIR + "golden.xlsx");
await attempt("truncated at 900 bytes", () => read(whole.subarray(0, 900)));

console.log("\n-- the hierarchy --");
try {
  await read(whole.subarray(0, 900));
} catch (err) {
  console.log("  err.name                       :", err.name);
  console.log("  err.code                       :", err.code);
  console.log("  instanceof FormatError         :", err instanceof FormatError);
  console.log("  instanceof GridxcelError       :", err instanceof GridxcelError);
  console.log("  instanceof BadZipError         :", err instanceof BadZipError);
  console.log("  instanceof Error               :", err instanceof Error);
  console.log("  Every error gridxcel raises is a GridxcelError, so one catch and one");
  console.log("  switch on code covers container, schema, format and file failures alike.");
}
console.log("  exported:", [GridxcelError, BadZipError, ZipError, FormatError, StrictWarningError].map((c) => c.name).join(", "));

console.log("\n-- damaged but readable: warnings, not throws --");
// A file damaged in a recoverable way is still read, and every problem is
// reported on book.warnings. One bad cell does not cost you the other rows.
const damaged = await read(FIXTURE_DIR + "reader-badcells.xlsx");
console.log("  reader-badcells.xlsx ->  read ok,", damaged.warnings.length, "warnings");
for (const w of damaged.warnings) console.log("      " + w.code + ": " + w.message);
console.log("  warnings is always an array; on an intact file it is empty.");
console.log("  Use onWarning to collect them yourself instead of reading the array.");

console.log("\n-- strict mode turns the first warning into a throw --");
try {
  await read(FIXTURE_DIR + "reader-badcells.xlsx", { strict: true });
  console.log("  strict -> no throw");
} catch (err) {
  console.log("  strict ->", err.name + " / " + err.code);
  console.log("      warning:", JSON.stringify(err.warning));
  console.log("      StrictWarningError.warning is the Warning the collection would have held.");
}

console.log("\n-- a sheet that claims to be enormous --");
// This sheet says it has five million rows. Densifying that would allocate
// gigabytes from a one-kilobyte file, so the dense converters refuse and say
// what to do instead. The guard is on untrusted input, not on real sheets.
const hostile = {
  name: "Hostile",
  rows: [[{ v: "one cell", t: "string" }]],
  originRow: 1,
  originCol: 0,
  rowCount: 5000000,
  columnCount: 1,
};
try {
  await sheetToAOA(hostile);
} catch (err) {
  console.log("  sheetToAOA(hostile) ->", err.name + " / " + err.code);
  console.log("      " + err.message.split(";")[0]);
  console.log("      details:", JSON.stringify(err.details));
  console.log("      fixes: pass a bounded range, raise maxRows, or stream with eachRow().");
}

// sheetToAOA() is a dense conversion; eachRow() materialises one row at a time
// and is bounded by what you consume, so the same header is not a problem there.
let streamed = 0;
for await (const event of eachRowOver()) streamed++;
async function* eachRowOver() {
  const { eachRow } = await import("gridxcel");
  // Streamed from a fixture; the point is that no dense range is ever built.
  yield* eachRow(new URL("../test/fixtures/golden.xlsx", import.meta.url), { sheets: ["Volume"], range: "A1:B10" });
}
console.log("\n  eachRow over a bounded range yielded", streamed, "rows without densifying anything.");
