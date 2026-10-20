/**
 * Keep formulas and get the display text a spreadsheet would show.
 *
 * Shows: cell.f and the cellFormula option, reading the rendered value with
 * cellText, and rendering a number yourself with formatValue/cellFormatCode.
 *
 * Run from anywhere:
 *   node examples/07-formulas-and-formats.mjs
 */
import { cellFormatCode, formatValue, isDateFormat, read, write } from "gridxcel";

const FIXTURE = new URL("../test/fixtures/golden.xlsx", import.meta.url);

/* -------------------------------------------------------------- formulas */

// A formula cell carries its text in cell.f, without the leading "=".
// cell.v is whatever was cached in the file, which may be nothing at all.
const book = await read(FIXTURE, { cellDates: true });
const types = book.sheets[0];
console.log("A2:", JSON.stringify(types.rows[1][0]));
console.log("A2 formula text:", JSON.stringify(types.rows[1][0].f), "| cached value:", types.rows[1][0].v);
console.log("B2 formula text:", JSON.stringify(types.rows[1][1].f));

// cellFormula: false strips the formula and leaves the cached value alone.
// A cell that held only a formula disappears from the sheet entirely.
const stripped = await read(FIXTURE, { cellFormula: false });
console.log("with cellFormula:false, A2 is:", stripped.sheets[0].rows[1][0]);

// Formulas survive a write: gridxcel writes cell.f back out unchanged and does
// not recalculate. The 0.1.0 reader never evaluates a formula.
const roundTripped = await read(await write(book));
console.log("after read -> write -> read, A2 =", JSON.stringify(roundTripped.sheets[0].rows[1][0].f));

/* ---------------------------------------------------------- display text */

// cellText fills cell.w with the number-format-rendered string. This is the
// text a spreadsheet shows, computed from the format code, not from the file.
const shown = await read(FIXTURE, { cellText: true });
const row = shown.sheets[0].rows[0];
console.log("\n-- value vs display text (row 1 of Types) --");
for (const cell of row) {
  if (!cell) continue;
  const raw = cell.t === "date" ? cell.serial : cell.v;
  console.log("  " + String(cell.t).padEnd(7), String(JSON.stringify(raw)).padEnd(20), "->", JSON.stringify(cell.w));
}

/* ------------------------------------------------- render it yourself */

// You do not have to ask the reader to do it. cellFormatCode resolves the
// format a cell was stored with, and formatValue renders any value with any
// code - including codes from a workbook you are generating.
const percent = row[4];
const code = cellFormatCode(book, percent);
console.log("\nE1 format code:", JSON.stringify(code));
console.log("formatValue(0.125, code)  :", formatValue(0.125, code));
console.log("formatValue(0.125, \"0.0%\"):", formatValue(0.125, "0.0%"));
console.log("formatValue(45366.6046875, \"yyyy-mm-dd hh:mm:ss\"):", formatValue(45366.6046875, "yyyy-mm-dd hh:mm:ss"));
console.log("formatValue(1234.5, \"#,##0.00\"):", formatValue(1234.5, "#,##0.00"));
console.log("isDateFormat(\"0.00%\") :", isDateFormat("0.00%"), "| isDateFormat(\"yyyy-mm-dd\"):", isDateFormat("yyyy-mm-dd"));
console.log("\nThe format code decides everything: 45366.6046875 is a date only");
console.log("because the cell says so. Swap the code and the same number is money.");
