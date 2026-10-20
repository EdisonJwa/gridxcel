/**
 * Read a workbook and walk its data model.
 *
 * Shows: read(), sheet metadata, the six cell types, how dates are represented
 * (cellDates), and why rows are sparse arrays offset by originRow/originCol.
 *
 * Run from anywhere - the fixture is resolved relative to this file:
 *   node examples/01-read-a-workbook.mjs
 */
import { read, encodeCol } from "gridxcel";

const FIXTURE = new URL("../test/fixtures/golden.xlsx", import.meta.url);

// read() takes a path, a file: URL, bytes, a stream or an async iterable.
// Nothing is parsed until the promise resolves; a malformed file rejects here.
const book = await read(FIXTURE, { cellDates: true });

console.log("sheets     :", book.sheets.map((s) => s.name).join(", "));
console.log("date system:", book.dateSystem);
console.log("warnings   :", book.warnings.length);

/* ------------------------------------------------------------------ types */

const types = book.sheets[0];

// originRow/originCol say where the used range starts. Here it starts at A1,
// so rows[0] is sheet row 1. A sheet that starts at C3 would report 3 and 2.
console.log("\nTypes starts at row " + types.originRow + ", column " + types.originCol);

console.log("\n-- one cell of every type (row 1) --");
for (let c = 0; c < types.rows[0].length; c++) {
  const cell = types.rows[0][c];
  if (!cell) continue; // a hole means "no cell here", not "empty string"
  // A date cell prints as the wall clock it stands for, not as an ISO string.
  const shown = cell.t === "date" ? fmtLocal(cell.v) + " (Date)" : JSON.stringify(cell.v);
  console.log(
    "  " + encodeCol(c + types.originCol) + types.originRow +
    "  t=" + cell.t.padEnd(7) +
    "  v=" + shown.slice(0, 40),
  );
}

/* ------------------------------------------------------------------ dates */

// A date cell carries three things: the value you asked for (a Date, because
// cellDates was set), the raw Excel serial, and the resolved format code.
const d = types.rows[0][3];
console.log("\n-- the date trap --");
console.log("  serial        :", d.serial);
console.log("  local fields  :", fmtLocal(d.v));
console.log("  toISOString() :", d.v.toISOString());
console.log("  resolved code :", d.zf);
console.log("  The Date holds the wall clock a spreadsheet displays. toISOString()");
console.log("  converts to UTC and shifts by the host offset. Compare local fields,");
console.log("  or compare serials, never the ISO string.");

/* ----------------------------------------------------------------- sparse */

// Rows are arrays indexed by column, not objects. A row that uses A and C has
// a hole at index 1. Trailing holes are trimmed, so length is the last column.
const layout = book.sheets[1];
console.log("\n-- sparse rows --");
console.log("  Layout rowCount   :", layout.rowCount, " columnCount:", layout.columnCount);
console.log("  rows[6] (sheet row 7):", JSON.stringify((layout.rows[6] || []).map((c) => (c ? c.v : null))));
console.log("  index 1 is", layout.rows[6][1], "- a hole, written as null by JSON");
console.log("  merges            :", layout.merges.join(" "));

// Iterating a sparse row without filter(Boolean) hands you undefined cells.
const used = layout.rows[6].filter(Boolean).length;
console.log("  cells actually present in row 7:", used, "of", layout.rows[6].length);

/* -------------------------------------------------------------- metadata */

// Cheap metadata first: sheetNames() does not parse any worksheet.
console.log("\n-- per-sheet metadata --");
for (const sheet of book.sheets) {
  console.log("  " + sheet.name.padEnd(8), "rows", String(sheet.rowCount).padStart(4), " cols", String(sheet.columnCount).padStart(3), " dims", sheet.dims ?? "-");
}

function fmtLocal(date) {
  const p = (n) => String(n).padStart(2, "0");
  return date.getFullYear() + "-" + p(date.getMonth() + 1) + "-" + p(date.getDate()) +
    " " + p(date.getHours()) + ":" + p(date.getMinutes()) + ":" + p(date.getSeconds());
}
