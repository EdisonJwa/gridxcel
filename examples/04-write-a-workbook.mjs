/**
 * Build a workbook and write it to disk.
 *
 * Shows: write() from an array of rows and from objects, per-cell types,
 * a style object, writing to a path, and reading the result back to prove it.
 *
 * Output goes to examples/out/ (gitignored). Run from anywhere:
 *   node examples/04-write-a-workbook.mjs
 */
import { read, write, writeFile, jsonToSheet, formatValue } from "gridxcel";
import { mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

const OUT = fileURLToPath(new URL("./out/", import.meta.url));
mkdirSync(OUT, { recursive: true });

/* --------------------------------------------------- an array of rows */

// A nested array is a sheet. Scalars become cells; a Date becomes a date cell
// with the date format applied; an object sets anything the scalar cannot
// express (type, format, formula, style).
const rows = [
  ["region", "units", "when", "note"],
  ["north", 1200, new Date(2024, 0, 15), "=B2*2"],
  ["south", 950, new Date(2024, 1, 2), null],
  ["east", { v: 1800, t: "number", z: "#,##0" }, new Date(2024, 2, 20), { v: "ok", s: { font: { b: true } } }],
];

// write() returns the encoded bytes. options.path also writes them to disk,
// and refuses to replace an existing file unless overwrite is set - a typo in
// a path should not destroy a previous run's output.
const bytes = await write(rows, { path: OUT + "regions.xlsx", overwrite: true });
console.log("wrote", bytes.length, "bytes to examples/out/regions.xlsx");

// What happens on the second attempt without the flag.
try {
  await write(rows, { path: OUT + "regions.xlsx" });
} catch (e) {
  console.log("write() again without overwrite ->", e.name + " / " + e.code);
}

/* ------------------------------------------------------- read it back */

const book = await read(OUT + "regions.xlsx", { cellDates: true, cellText: true });
const sheet = book.sheets[0];
console.log("sheet:", sheet.name, "| rows", sheet.rowCount, "| cols", sheet.columnCount);

for (const [i, row] of sheet.rows.entries()) {
  const shown = row.map((c) => {
    if (!c) return "";
    if (c.t === "date") return c.zf + " -> " + c.w;
    if (c.f) return "=" + c.f;
    return c.w ?? String(c.v);
  });
  console.log("  row " + (i + sheet.originRow) + ":", JSON.stringify(shown));
}

/* -------------------------------------------- objects rather than arrays */

// jsonToSheet() takes the keys of the first object as the header row, so the
// field order is the key order and every later object follows it.
const people = [
  { name: "Ada", born: 1815, score: 99.5 },
  { name: "Grace", born: 1906, score: 98.25 },
];
const peopleSheet = await jsonToSheet(people, { sheet: "People" });

await writeFile({ sheets: [peopleSheet], sheetNames: ["People"] }, OUT + "people.xlsx", { overwrite: true });
console.log("\nwrote examples/out/people.xlsx");
console.log("  rows:", JSON.stringify((await read(OUT + "people.xlsx")).sheets[0].rows.map((r) => r.map((c) => (c ? c.v : null)))));

/* ------------------------------------------------------- number formats */

// A format code renders a value the way a spreadsheet would show it.
console.log("\nformatValue:", formatValue(1234.5, "#,##0.00"), "|", formatValue(0.125, "0.00%"));
