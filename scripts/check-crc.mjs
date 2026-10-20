#!/usr/bin/env node
/**
 * Regression check: are stored (uncompressed) zip parts CRC-checked on read?
 *
 * History: on 2026-10-18 gridxcel returned a corrupted stored value with no
 * warning (gap). The lead then made src/zip.js verify every part's CRC on read
 * (ZipArchive({verifyCrc:false}) opts out), so corruption now raises a typed
 * ZipError/CRC_MISMATCH.
 *
 * The script rewrites one digit of a stored worksheet part so the cell value
 * changes but the entry's CRC no longer matches, then reports which state holds:
 *   FIXED     - gridxcel throws a typed CRC error (exit 1; expected today)
 *   WARNED    - gridxcel returns the value but reports it via onWarning (exit 1)
 *   GAP       - gridxcel returns the corrupted value silently (exit 0; bad)
 *   INCONCLUSIVE - the fixture copy is not actually corrupt (exit 2)
 *
 * It is a report, not part of the test suite: the robustness law in
 * test/fuzz.test.mjs ("typed error or warned workbook, never a hang or a
 * non-Error") is satisfied either way.
 *
 * Usage: node scripts/check-crc.mjs
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { pathToFileURL } from "node:url";

import { locateEntry } from "./zip-surgery.mjs";
import { selftest } from "./oracle-client.mjs";

const ROOT = process.cwd();
const SRC = path.join(ROOT, "test", "fixtures", "openpyxl-stored.xlsx");
const OUT = path.join(ROOT, "bench", "tmp", "crc-corrupt.xlsx");

const buf = new Uint8Array(readFileSync(SRC));
const entry = locateEntry(buf, "xl/worksheets/sheet1.xml");
if (!entry) {
  console.error("fixture layout changed: xl/worksheets/sheet1.xml not found in " + SRC);
  process.exit(2);
}
const text = Buffer.from(buf.slice(entry.dataStart, entry.dataEnd)).toString("utf8");
const at = text.indexOf("<v>42</v>");
if (at < 0) {
  console.error("fixture layout changed: <v>42</v> not found in sheet1.xml");
  process.exit(2);
}
mkdirSync(path.dirname(OUT), { recursive: true });
const corrupted = buf.slice();
corrupted[entry.dataStart + at + 4] = "3".charCodeAt(0); // 42 -> 43, same length, stale CRC
writeFileSync(OUT, corrupted);

const oracle = selftest(ROOT, OUT);
const { readWorkbook } = await import(pathToFileURL(path.join(ROOT, "src", "read.js")).href);
const warnings = [];
let book = null;
let error = null;
try {
  book = await readWorkbook(new Uint8Array(corrupted), { onWarning: (w) => warnings.push(w) });
} catch (err) {
  error = err;
}
const cell = book?.sheets?.[0]?.rows?.[0]?.[1] ?? null;

console.log("corrupted copy:      " + path.relative(ROOT, OUT));
console.log("python testzip():    " + JSON.stringify(oracle.testzip));
console.log("gridxcel result:     " + (error
  ? error.name + " [" + (error.code ?? "no code") + "] " + error.message
  : JSON.stringify(cell)));
console.log("gridxcel warnings:   " + JSON.stringify(warnings.map((w) => w.code)));

if (oracle.testzip === null) {
  console.log("");
  console.log("INCONCLUSIVE: the written copy is not corrupt (fixture layout changed?).");
  process.exit(2);
}
if (error) {
  console.log("");
  console.log("FIXED: corruption rejected with a typed error (" + (error.code ?? error.name) + ").");
  process.exit(1);
}
if (warnings.some((w) => /CRC/i.test(String(w.code)))) {
  console.log("");
  console.log("FIXED (warned): corruption reported through onWarning.");
  process.exit(1);
}
if (warnings.length > 0) {
  console.log("");
  console.log("WARNED, but not as a CRC warning: " + JSON.stringify(warnings.map((w) => w.code)));
  process.exit(1);
}
console.log("");
console.log("GAP CONFIRMED: a part with a stale CRC is read without any warning.");
process.exit(0);
