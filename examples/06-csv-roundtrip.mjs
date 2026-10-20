/**
 * Convert between CSV and XLSX in both directions.
 *
 * Shows: csvToSheet(), sheetToCSV(), the delimiter/quote/BOM options, the CSV
 * injection guard, and a CSV -> XLSX -> CSV round trip that is asserted rather
 * than eyeballed.
 *
 * Output goes to examples/out/ (gitignored). Run from anywhere:
 *   node examples/06-csv-roundtrip.mjs
 */
import { csvToSheet, read, sheetToCSV, write } from "gridxcel";
import assert from "node:assert";
import { mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const OUT = fileURLToPath(new URL("./out/", import.meta.url));
mkdirSync(OUT, { recursive: true });

/* ---------------------------------------------------------- CSV -> sheet */

// RFC 4180 quoting: a field containing a quote or a comma is wrapped in
// quotes, and an embedded quote is doubled.
const text = [
  "region,units,note",
  'north,1200,"says ""hi"", then leaves"',
  "south,950,",
  "east,1800,multi",
].join("\n") + "\n";

const fromCSV = await csvToSheet(text);
console.log("csvToSheet rows:", fromCSV.rowCount, "cols:", fromCSV.columnCount);
for (const row of fromCSV.rows) {
  console.log("  ", JSON.stringify(row.map((c) => (c ? c.v : null))));
}
console.log("embedded comma survived:", fromCSV.rows[1][2].v === 'says "hi", then leaves');
console.log("types inferred: units is", fromCSV.rows[1][1].t, "(not a string)");

/* ---------------------------------------------------------- sheet -> CSV */

console.log("\nplain      :", JSON.stringify(await sheetToCSV(fromCSV)));
console.log("alwaysQuote:", JSON.stringify(await sheetToCSV(fromCSV, { alwaysQuote: true })));
console.log("delimiter  :", JSON.stringify(await sheetToCSV(fromCSV, { delimiter: ";" })));
console.log("bom        :", JSON.stringify((await sheetToCSV(fromCSV, { bom: true })).slice(0, 4)));

/* ---------------------------------------------------- CSV injection guard */

// Text starting with = + - @ is executed as a formula when a spreadsheet opens
// the CSV. sheetToCSV prefixes it with an apostrophe by default, so a hostile
// CSV field cannot become a live formula. escapeFormulae: false turns the
// guard off and is only appropriate for a consumer that is not a spreadsheet.
const risky = await csvToSheet("name,value\nsum,=1+1\n", { header: true });
console.log("\ncell read back :", JSON.stringify(risky.rows[1][1]));
console.log("guarded (default):", JSON.stringify(await sheetToCSV(risky)));
console.log("escapeFormulae:false:", JSON.stringify(await sheetToCSV(risky, { escapeFormulae: false })));
console.log("The guard is one-way: csvToSheet does not strip the apostrophe again,");
console.log("exactly like Excel's own CSV output.");

/* -------------------------------------------------------- the round trip */

// CSV -> sheet -> XLSX on disk -> read it back -> CSV again.
const XLSX = OUT + "csv-roundtrip.xlsx";
writeFileSync(XLSX, await write({ sheets: [fromCSV], sheetNames: [fromCSV.name] }));

const back = await read(XLSX);
const direct = await sheetToCSV(fromCSV);
const again = await sheetToCSV(back.sheets[0]);

console.log("\nwrote examples/out/csv-roundtrip.xlsx -", (await import("node:fs")).statSync(XLSX).size, "bytes");
console.log("CSV in :", JSON.stringify(text));
console.log("CSV out:", JSON.stringify(again));
assert.strictEqual(again, direct);
console.log("round trip is byte-identical for this input.");
console.log("It will not stay byte-identical for every input: an empty leading or");
console.log("trailing field becomes a cell that a workbook is free not to store, and");
console.log("numbers lose the spelling they arrived with (1.50 comes back as 1.5).");
console.log("If you need CSV bytes preserved exactly, do not convert them.");
