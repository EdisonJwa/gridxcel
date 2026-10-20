# Quick start

From `npm install` to reading, writing and streaming a workbook. This page
assumes you know JavaScript and nothing about XLSX, OOXML or spreadsheets as a
file format.

Every example below is executed by `npm run check:docs`, so none of it can drift
from the implementation. Fixture paths such as `../../test/fixtures/golden.xlsx`
resolve from where that checker runs the examples; substitute a path to your own
file.

| I want to | Go to |
| --- | --- |
| Read a file into memory | [§2 Read a workbook](#2-read-a-workbook) |
| Write a file | [§3 Write a workbook](#3-write-a-workbook) |
| Handle a file too large for memory | [§4 Stream a workbook](#4-stream-a-workbook) |
| Understand the data model | [Five things that trip people up](#five-things-that-trip-people-up) |

## 1. Install

```sh
npm install gridxcel
```

Node 18 or newer. The package is ESM only - there is no CommonJS build and no
build step, so `src/*.js` is exactly what ships. It has no runtime
dependencies and the tests need no network.

The whole public surface is one import:

```js doc-example
import { read, write, eachRow, VERSION } from "gridxcel";

console.log("gridxcel", VERSION);
console.log(typeof read, typeof write, typeof eachRow);
```

## 2. Read a workbook

An XLSX file is a ZIP of XML parts. `read` hides that: give it a path, a file
URL, bytes, a stream or an async iterable, and get a workbook back.

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/pandas.xlsx");
assert.strictEqual(book.sheets.length, 1);

const sheet = book.sheets[0];
console.log("sheet   :", sheet.name);
console.log("rows    :", sheet.rowCount);
console.log("columns :", sheet.columnCount);
console.log("header  :", sheet.rows[0].map((c) => c.v));
console.log("row 2   :", sheet.rows[1].map((c) => c.v));
```

That is most of the data model already. A workbook has `sheets`; a sheet has
rows; a row is an array of cells; a cell is a small object whose `v` is the
value and whose `t` says what kind of value it is.

```js doc-example
import { read } from "gridxcel";

const book = await read("../../test/fixtures/golden.xlsx", { cellDates: true });
const sheet = book.sheets[0];

for (let c = 0; c < sheet.rows[0].length; c++) {
  const cell = sheet.rows[0][c];
  const shown = cell.t === "date" ? cell.v.toDateString() : JSON.stringify(cell.v);
  console.log("  t=" + String(cell.t).padEnd(7), "v=" + shown);
}
```

Six types exist: `string`, `number`, `boolean`, `date`, `error` and `empty`.
A type you did not expect is usually the interesting bug.

## 3. Write a workbook

The shortest write is an array of arrays. Strings, numbers, booleans and dates
are understood as they are; wrap a value in an object when you need to say more
about it.

```js doc-example
import { read, write } from "gridxcel";
import assert from "node:assert";

const bytes = await write([
  ["region", "units", "checked"],
  ["north", 1200, new Date(2024, 0, 15)],
  ["south", { v: 950, t: "number", z: "#,##0" }, new Date(2024, 1, 2)],
]);
console.log("wrote", bytes.length, "bytes");
assert.ok(bytes.length > 0);

// Read it back: the only proof that a written workbook is a workbook.
const book = await read(bytes, { cellDates: true, cellText: true });
for (const row of book.sheets[0].rows) {
  console.log("  " + row.map((c) => (c.t === "date" ? c.zf + " " + c.w : c.v)).join(" | "));
}
```

Writing to a path is the same call with `path` set. It refuses to replace an
existing file unless you pass `overwrite: true`.

## 4. Stream a workbook

The `read` function holds a sheet in memory. `eachRow` does not: it yields
one row at a time, so a file far larger than memory is fine as long as you
handle each row and let it go.

```js doc-example
import { eachRow } from "gridxcel";

let rows = 0;
let cells = 0;
for await (const { sheet, r, cells: row } of eachRow("../../test/fixtures/golden.xlsx", { sheets: ["Volume"] })) {
  rows++;
  cells += row.filter(Boolean).length;
  if (r === 1) console.log("first row of", sheet, "->", row[0].v, "...", row[25].v);
}
console.log("streamed", rows, "rows and", cells, "cells without holding the sheet");
```

Writing has the same shape. `WorkbookWriter` serialises each row as it is
appended instead of building a grid first.

```js doc-example
import { WorkbookWriter } from "gridxcel";

const writer = new WorkbookWriter();
const sheet = writer.addSheet("Rows");
for (let r = 1; r <= 1000; r++) sheet.addRow(r, [r, "row " + r, r / 2]);
sheet.end();
const bytes = writer.finish();

console.log("wrote", sheet.rowCount, "rows,", bytes.length, "bytes");
```

## Five things that trip people up

### 1. Dates are local wall clock, not UTC

A spreadsheet stores a date as a number of days, with no timezone. gridxcel
turns that back into a `Date` whose *local* fields are the ones the
spreadsheet displays. Never compare one of these with `toISOString()`, which
converts to UTC and shifts by the host offset.

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx", { cellDates: true });
const cell = book.sheets[0].rows[0][3];

console.log("serial            :", cell.serial);
console.log("local fields      :", cell.v.getFullYear() + "-" + (cell.v.getMonth() + 1) + "-" + cell.v.getDate());
console.log("getUTCFullYear()  :", cell.v.getUTCFullYear(), "  <- differs whenever the offset is negative");

// The wall clock is preserved: reading then writing then reading gives the
// same fields back, whatever the host timezone is.
const { write } = await import("gridxcel");
const again = await read(await write(book), { cellDates: true });
const back = again.sheets[0].rows[0][3].v;
assert.strictEqual(back.getFullYear(), cell.v.getFullYear());
assert.strictEqual(back.getHours(), cell.v.getHours());
assert.strictEqual(back.getMinutes(), cell.v.getMinutes());
console.log("round trip keeps the wall clock:", back.getHours() + ":" + back.getMinutes());
```

If you would rather work in numbers, do not ask for `cellDates` at all: read
the serial from `cell.serial` or `cell.v` instead, and convert it yourself
with `serialToDate` or `dateToSerial`.

### 2. Rows are sparse arrays offset by originRow and originCol

Rows are arrays indexed by column, not objects. A row that uses A and Z has two
cells plus a run of holes, and trailing holes are trimmed. The sheet records
where the used range starts in `originRow` and `originCol`, so `rows[0]` is the
sheet row numbered `originRow`, not necessarily row 1.

```js doc-example
import { read, encodeCol } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx");
const layout = book.sheets[1];

console.log("originRow:", layout.originRow, "originCol:", layout.originCol);
const row = layout.rows[6];             // sheet row 7
console.log("row 7 length:", row.length, "cells present:", row.filter(Boolean).length);
assert.strictEqual(row[1], undefined);  // a hole, not an empty string
console.log("index 1 is", row[1]);

// The A1 address of a cell is derived, not stored: origin, row, column index.
const address = encodeCol(layout.originCol + 2) + (layout.originRow + 6);
console.log("the cell at index 2 of that row is", address, "=", row[2].v);
```

Iterating a row without `filter(Boolean)` hands you `undefined` for every
hole. When you want a rectangular grid instead, `sheetToAOA` fills the holes.

### 3. Every cell carries an explicit t

gridxcel does not guess a type at the point of use. `cell.t` is decided while
reading and everything downstream respects it, which is why a value that looks
like a number can still be a string.

```js doc-example
import { csvToSheet } from "gridxcel";

// CSV has no types of its own, so the parser infers them from the text.
const sheet = await csvToSheet("id,code,when\n1,007,2024-03-15\n");
for (const cell of sheet.rows[1]) {
  console.log("  t=" + String(cell.t).padEnd(7), "v=" + JSON.stringify(cell.v));
}
console.log("A ZIP code that reads as a number is a number; ask for a string by");
console.log("setting t yourself when you write the cell.");
```

### 4. Malformed input throws - it never returns an empty workbook

Bytes that are not a workbook are an error, not an empty sheet list. A typo in
a path, a truncated download and an HTML error page all fail loudly, and each
failure is a typed error with a stable `code`.

```js doc-example
import { GridxcelError, read } from "gridxcel";
import assert from "node:assert";

const html = new TextEncoder().encode("<!doctype html><html>404</html>");
try {
  await read(html);
  assert.fail("this should not have been accepted");
} catch (err) {
  console.log(err.name + " / " + err.code);
  assert.ok(err instanceof GridxcelError);
}
```

A file that is damaged but still parseable is a different case: it reads, and
every problem is reported on `book.warnings`. Pass `strict: true` to turn the
first warning into a throw.

### 5. Converting an untrusted sheet to a dense grid is refused above maxRows

A worksheet is untrusted input: one `<row r="40000000">` makes the reader
report a forty-million-row extent. Converting that to an array of arrays would
allocate gigabytes from a one-kilobyte file, so the dense converters refuse
above 262,144 rows and tell you what to do instead.

```js doc-example
import { sheetToAOA } from "gridxcel";
import assert from "node:assert";

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
  assert.fail("this should have been refused");
} catch (err) {
  console.log(err.name + " / " + err.code);
  console.log(err.message.split(";")[0]);
}

// Passing a bounded range takes you at your word: the guard only refuses the
// extent the sheet declared for itself.
const bounded = await sheetToAOA(hostile, { range: "A1:A2" });
console.log("with range A1:A2 the same sheet converts:", JSON.stringify(bounded));
```

This is a guard on input you did not produce, not a limit on real sheets. If
you genuinely have five million rows, stream them with `eachRow` or pass a
range.

## Where to go next

| Document | Read it for |
| --- | --- |
| [API.md](./API.md) | Every function, option and error, with runnable examples. |
| [COOKBOOK.md](./COOKBOOK.md) | Task-shaped recipes: JSON conversion, CSV round trips, large files, formulas, HTTP. |
| [LIMITATIONS.md](./LIMITATIONS.md) | What is not supported, what is unverified, and the 0.1.0 caveats. |
| [MIGRATION.md](./MIGRATION.md) | Moving from SheetJS, including differences that change results. |

The [examples](https://github.com/EdisonJwa/gridxcel/tree/main/examples) directory holds eleven runnable programs,
and `node examples/run-all.mjs` runs every one of them.
