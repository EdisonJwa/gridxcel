/**
 * Turn a worksheet into JSON objects.
 *
 * Shows: sheetToJSON() with its header forms, sheetToAOA() for a dense grid,
 * and readJSON() when you only want the first sheet.
 *
 * Run from anywhere:
 *   node examples/03-convert-to-json.mjs
 */
import { read, readJSON, sheetToAOA, sheetToJSON } from "gridxcel";

const FIXTURE = new URL("../test/fixtures/pandas.xlsx", import.meta.url);

const book = await read(FIXTURE);
const sheet = book.sheets[0];

console.log("sheet     :", sheet.name, "rows", sheet.rowCount, "cols", sheet.columnCount);
console.log("header row:", JSON.stringify(sheet.rows[0].map((c) => (c ? c.v : null))));
console.log("first data:", JSON.stringify(sheet.rows[1].map((c) => (c ? c.v : null))));

/* ------------------------------------------------- the default: header row */

// With no options the first row of the range supplies the keys and every row
// after it becomes an object. 101 rows in the file, 100 objects out.
const objects = await sheetToJSON(sheet);
console.log("\ndefault            :", objects.length, "objects, first =", JSON.stringify(objects[0]));

/* ------------------------------------------------------- the header forms */

// header: 1 keeps every row, as arrays. Nothing is treated as a header.
const arrays = await sheetToJSON(sheet, { header: 1 });
console.log("header: 1          :", arrays.length, "arrays, first =", JSON.stringify(arrays[0]));

// header: "A" uses column letters, so duplicate or missing labels cannot bite.
const letters = await sheetToJSON(sheet, { header: "A" });
console.log('header: "A"        :', JSON.stringify(letters[0]));

// A header array names the columns. Columns past the end of the array fall
// back to the sheet's own header text for that column.
const named = await sheetToJSON(sheet, { header: ["left", "middle", "right"] });
console.log("header: [...]      :", JSON.stringify(named[0]));

// header: n (n > 1) means "row n is the header row". Here row 2, so row 3 is
// the first object and the keys are whatever row 2 held.
const fromRow2 = await sheetToJSON(sheet, { header: 2 });
console.log("header: 2          :", JSON.stringify(fromRow2[0]));

// defval decides what a hole becomes. Without it a missing cell is dropped.
const filled = await sheetToJSON(sheet, { header: 1, defval: null });
console.log("header: 1 + defval :", JSON.stringify(filled[0]));

/* ------------------------------------------------------- dense AOA output */

// sheetToAOA() is the same walk without the object layer. It is dense: holes
// become null, so a sparse row costs memory proportional to its width.
const aoa = await sheetToAOA(sheet, { range: "A1:C4" });
console.log("\nsheetToAOA A1:C4   :");
for (const row of aoa) console.log("  ", JSON.stringify(row));

/* ------------------------------------------------------ one call, one sheet */

// readJSON() is read + sheetToJSON on the first sheet, for the common case.
const quick = await readJSON(FIXTURE, { header: ["a", "b", "c"] });
console.log("\nreadJSON           :", quick.length, "objects, first =", JSON.stringify(quick[0]));
console.log("readJSON {raw:true} :", (await readJSON(FIXTURE, { raw: true })).length, "objects (header row kept as data)");
