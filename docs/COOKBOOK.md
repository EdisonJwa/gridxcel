# Cookbook

Task-shaped recipes. Each one is a complete, runnable program: every code block
marked **doc-example** is executed by `npm run check:docs` from the repository
root, so nothing here can rot. Copy one, change the path, run it.

Fixture paths such as `../../test/fixtures/golden.xlsx` resolve from where that
checker runs; substitute your own file. Recipes that write files write them into
the working directory - point them somewhere real.

| Recipe | Use when |
| --- | --- |
| [1. Read a sheet into JSON objects](#1-read-a-sheet-into-json-objects) | You want records, not cells. |
| [2. Stream a file too big for memory](#2-stream-a-file-too-big-for-memory) | The sheet does not fit in RAM. |
| [3. Write a large file without buffering it](#3-write-a-large-file-without-buffering-it) | The rows arrive from a cursor, not an array. |
| [4. Convert CSV to XLSX and back](#4-convert-csv-to-xlsx-and-back) | Something downstream wants the other format. |
| [5. Read only some sheets](#5-read-only-some-sheets) | The workbook has twenty tabs and you need two. |
| [6. Keep formulas](#6-keep-formulas) | The workbook is a model, not a report. |
| [7. Get display text for a number](#7-get-display-text-for-a-number) | You need what the cell shows, not what it holds. |
| [8. Build a workbook from arrays](#8-build-a-workbook-from-arrays) | You are generating a file, not reading one. |
| [9. Append rows to a workbook](#9-append-rows-to-a-workbook) | You are adding to an existing file. |
| [10. Handle a file that fails to parse](#10-handle-a-file-that-fails-to-parse) | The input comes from users or the network. |
| [11. Read a workbook from an HTTP response](#11-read-a-workbook-from-an-http-response) | The file is served, not stored. |

---

## 1. Read a sheet into JSON objects

The `sheetToJSON` helper turns the first row of the range into keys and every
row after it into an object. It is async because it may have to materialise a
range first.

```js doc-example
import { read, sheetToJSON } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/pandas.xlsx");
const sheet = book.sheets[0];

const rows = await sheetToJSON(sheet);
assert.strictEqual(rows.length, 100);
console.log(rows.length, "objects");
console.log(rows[0]);
console.log(rows[99]);

// header: 1 keeps every row and gives you arrays, header: "A" uses column
// letters, header: [...] names the columns yourself.
const asArrays = await sheetToJSON(sheet, { header: 1 });
const named = await sheetToJSON(sheet, { header: ["first", "second", "third"] });
console.log("arrays:", asArrays.length, "| named:", JSON.stringify(named[0]));
```

Watch out for:

* With no options the first row of the range is consumed as the header. A sheet
  with no header row gives you keys taken from its data. Use `header: 1` when
  the first row is data.
* Duplicate header labels are disambiguated with a `_1` suffix rather than
  silently overwriting each other.
* A missing cell is absent from the object. Pass `defval` to make every key
  present, which is what most downstream code actually wants.
* If you only want the first sheet, `readJSON` is `read` plus `sheetToJSON`
  in one call.

## 2. Stream a file too big for memory

The `eachRow` iterator yields `{ sheet, r, cells }` one row at a time. Nothing
accumulates unless you accumulate it, so the loop can run over a file far
larger than the heap.

```js doc-example
import { eachRow } from "gridxcel";

const totals = new Map();

for await (const { sheet, r, cells } of eachRow("../../test/fixtures/golden.xlsx", { sheets: ["Volume"] })) {
  // cells[0] is column A even though the batch reader would rebase to
  // originCol. Sum column B into buckets of fifty rows.
  const bucket = "rows " + (Math.floor((r - 1) / 50) * 50 + 1) + "-" + (Math.floor((r - 1) / 50) * 50 + 50);
  totals.set(bucket, (totals.get(bucket) ?? 0) + (cells[1]?.v ?? 0));
  if (r === 1) console.log("first row:", cells[0].v, cells[1].v, cells[25].v);
}

console.log("aggregated", totals.size, "buckets without holding the sheet");
for (const [bucket, sum] of totals) console.log("  " + bucket + ":", sum);
```

Watch out for:

* `cells` is indexed from column A even when a batch read would rebase to
  `originCol`. Streaming and batch reads do not agree on the index; they agree
  on `cell.a` when it is present.
* `r` is the sheet row number, not an array index. Rows that do not exist in
  the file are not yielded, so `r` jumps.
* Breaking out of the loop early is supported and stops the work: you never
  inflate the rest of the archive.
* Hyperlinks are not available on a streamed cell; read the sheet in batch mode
  if you need them.
* The archive's central directory is still read into memory before any part can
  be located, so a huge *file* still costs its own size in bytes. Streaming
  bounds the sheet representation, not the input.

## 3. Write a large file without buffering it

The `WorkbookWriter` class serialises each row as it is appended. Nothing holds
a grid of cells, so the source can be a database cursor or a generator.

```js doc-example
import { WorkbookWriter } from "gridxcel";
import assert from "node:assert";

const writer = new WorkbookWriter();
const sheet = writer.addSheet("Rows", { columns: [{ width: 8 }, { width: 24 }] });

for (let r = 1; r <= 20000; r++) {
  // Row numbers must ascend. undefined means "the next row".
  sheet.addRow(undefined, [r, "row " + r, r * 1.5]);
}
sheet.end();

const bytes = writer.finish();
assert.ok(bytes.length > 0);
console.log("wrote", sheet.rowCount, "rows as", bytes.length, "bytes");
```

Watch out for:

* `addRow` throws `ROW_ORDER` if a row arrives out of order and `BAD_ROW` for
  a non-positive or non-integer index. Streaming writers cannot go back.
* `finish()` returns the whole archive, so the output bytes are in memory even
  though the sheet never was. This is why writing XLSX straight to a
  non-seekable sink is not supported.
* Write the buffer to a file in one call, or hand it to whatever wants it; there
  is no incremental sink.
* Once `end()` has been called the sheet writer is closed. Adding a row after
  that is an error, not a silent no-op.

## 4. Convert CSV to XLSX and back

The `csvToSheet` parser handles RFC 4180 CSV with the extensions real
pipelines need, and `sheetToCSV` writes it back. Types are inferred on the way
in from the text itself.

```js doc-example
import { csvToSheet, read, sheetToCSV, write } from "gridxcel";
import assert from "node:assert";

const csv = [
  "region,units,note",
  'north,1200,"says ""hi"", then leaves"',
  "south,950,",
].join("\n") + "\n";

const sheet = await csvToSheet(csv);
assert.strictEqual(sheet.rowCount, 3);
console.log("parsed:", sheet.rowCount, "rows,", sheet.columnCount, "columns");
console.log("types :", sheet.rows[1].map((c) => c.t).join(", "));

// CSV -> XLSX -> CSV.
const bytes = await write({ sheets: [sheet], sheetNames: [sheet.name] });
const back = await read(bytes);
console.log("round trip identical:", (await sheetToCSV(back.sheets[0])) === (await sheetToCSV(sheet)));

// Options you will actually need.
console.log("semicolons :", JSON.stringify(await sheetToCSV(sheet, { delimiter: ";" })));
console.log("bom        :", JSON.stringify((await sheetToCSV(sheet, { bom: true })).slice(0, 3)));
console.log("crlf       :", JSON.stringify(await sheetToCSV(sheet, { newline: "\r\n" })));
```

Watch out for:

* `csvToSheet` infers types: a field of digits becomes a number, so a ZIP code
  or a zero-padded id loses its leading zeros. Pass `raw: true` for every field
  as text, or `quotedStrings: true` to take quoted fields literally.
* `sheetToCSV` escapes text starting with `=`, `+`, `-`, `@`, a tab, a CR or an LF with an
  apostrophe by default, so a hostile CSV field cannot become a live formula.
  `escapeFormulae: false` disables that guard.
* An empty trailing field is a cell, and a workbook is free not to store an
  empty cell, so a round trip can lose them. Do not round trip CSV when the
  bytes must be preserved.
* `sheetToCSV` is O(output): it builds the whole string. There is no
  incremental CSV writer. See [LIMITATIONS.md](./LIMITATIONS.md).

## 5. Read only some sheets

Reading a workbook you will not use wastes the parse. Ask for the sheet names
first, then read the ones you need.

```js doc-example
import { read, sheetNames } from "gridxcel";

// sheetNames does not parse any worksheet, so it is cheap enough to call first.
const names = await sheetNames("../../test/fixtures/golden.xlsx");
console.log("declared sheets:", names.join(", "));

// sheets takes names or 0-based indexes. sheetNames still reports everything,
// so you can tell "not asked for" apart from "not there".
const book = await read("../../test/fixtures/golden.xlsx", { sheets: ["Types", "Volume"] });
console.log("read       :", book.sheets.map((s) => s.name).join(", "));
console.log("declared   :", book.sheetNames.join(", "));

const byIndex = await read("../../test/fixtures/golden.xlsx", { sheets: [1] });
console.log("by index   :", byIndex.sheets[0].name);

// sheets: false reads metadata only - names, styles and properties, no cells.
const meta = await read("../../test/fixtures/golden.xlsx", { sheets: false });
console.log("metadata   :", meta.sheets.length, "sheets loaded,", meta.sheetNames.length, "declared");
```

Watch out for:

* Indexes are 0-based and positional: inserting a sheet in Excel changes what
  index 1 means. Prefer names in anything long-lived.
* Reading one sheet of a workbook still reads the archive, so it saves the
  parse, not the download.
* `sheets: false` is the right call when you only need `sheetNames`, the
  workbook properties or the style table.

## 6. Keep formulas

A formula lives in `cell.f`, without the leading `=`. gridxcel never
evaluates one: it reads the text and the cached result if the file has one, and
writes the text back unchanged.

```js doc-example
import { read, write } from "gridxcel";

const book = await read("../../test/fixtures/golden.xlsx");
const cell = book.sheets[0].rows[1][0];

console.log("formula text :", cell.f);
console.log("cached value :", cell.v, "(t=" + cell.t + ")");
console.log("the cached result is whatever the writing application last stored,");
console.log("and it is often absent - the file may never have been opened.");

// Formulas survive a read -> write -> read cycle untouched.
const again = await read(await write(book));
console.log("after a round trip:", again.sheets[0].rows[1][0].f);

// cellFormula: false drops the formula, and a cell that held only a formula
// disappears from the sheet entirely rather than becoming an empty cell.
const stripped = await read("../../test/fixtures/golden.xlsx", { cellFormula: false });
console.log("with cellFormula:false, that cell is", stripped.sheets[0].rows[1][0]);
```

Watch out for:

* Changing a formula does not change the cached value, and reading does not
  recalculate. A spreadsheet application recalculates when it opens the file.
* `cell.v` is `null` when the file carried no cached result, which is normal
  for a file written by a library rather than saved by Excel.
* Writing a formula means setting `f` on a cell object. A plain string
  beginning with `=` is written as that string unless you say otherwise.

## 7. Get display text for a number

The stored value and the shown value are different things. `0.125` with a
percent format is `12.50%`. Ask for the rendered text with `cellText`, or
render any value yourself with `formatValue`.

```js doc-example
import { cellFormatCode, formatValue, isDateFormat, read } from "gridxcel";

const book = await read("../../test/fixtures/golden.xlsx", { cellText: true });
const row = book.sheets[0].rows[0];

for (const cell of row) {
  if (!cell) continue;
  console.log("  " + String(cell.t).padEnd(7), JSON.stringify(cell.v).padEnd(22), "->", JSON.stringify(cell.w));
}

// The rendered text is computed from the cell's format code, which you can
// read back and reuse on a value you are generating.
const code = cellFormatCode(book, row[4]);
console.log("format code:", JSON.stringify(code));
console.log("formatValue(0.125, code):", formatValue(0.125, code));
console.log("isDateFormat(\"yyyy-mm-dd\"):", isDateFormat("yyyy-mm-dd"), "| isDateFormat(\"0.00%\"):", isDateFormat("0.00%"));
```

Watch out for:

* `cell.w` is only filled when you ask for it with `cellText: true`. It is
  never filled by default, because rendering every cell is work you may not want.
* Numbers render through a format-code engine, not through the host locale.
  Locale-aware rendering is a different job; do it after extracting the value.
* `cell.z` is polymorphic by direction: writing accepts a code or an id,
  reading returns the numeric id. The resolved code is always available as
  `cell.zf`, and `cellFormatCode` is the supported way to ask for it.

## 8. Build a workbook from arrays

A nested array is a sheet. Objects become cells when a scalar cannot say
enough, and `jsonToSheet` goes the other way: records in, a sheet with a header
row out.

```js doc-example
import { aoaToSheet, jsonToSheet, read, write } from "gridxcel";
import assert from "node:assert";

// An array of arrays: strings, numbers, booleans and Dates are understood.
// aoaToSheet turns one into a named sheet.
const data = aoaToSheet([
  ["region", "units", "checked"],
  ["north", 1200, true],
  ["south", 950, false],
], { sheet: "Data" });

// An object says what a scalar cannot: type, number format, formula, style.
const styled = aoaToSheet([
  ["region", "units"],
  ["north", { v: 1200, t: "number", z: "#,##0" }],
  ["south", { v: 1200, f: "B2*2", t: "number" }],
], { sheet: "Styled" });

// Records in, sheet out. Key order is the header order.
const people = await jsonToSheet([{ name: "Ada", born: 1815 }, { name: "Grace", born: 1906 }], { sheet: "People" });

const bytes = await write({ sheets: [data, styled, people] });
const book = await read(bytes);
assert.deepStrictEqual(book.sheets.map((s) => s.name), ["Data", "Styled", "People"]);
console.log("sheets:", book.sheetNames.join(", "));
console.log("styled:", JSON.stringify(book.sheets[1].rows.map((r) => r.map((c) => c?.f ?? c?.v))));
console.log("people:", JSON.stringify(book.sheets[2].rows.map((r) => r.map((c) => c.v))));
```

Watch out for:

* Passing a bare array to `write` gives you one sheet called `Sheet1`. For
  several sheets or a chosen name, pass `{ sheets, sheetNames }`.
* A sheet object needs `name`, `rows`, `originRow`, `originCol`,
  `rowCount` and `columnCount`. `normalizeSheet` fills in what it can if you
  are building one by hand.
* Sheet names have rules: no more than 31 characters, and none of `: \ / ? * [ ]`.

## 9. Append rows to a workbook

There is no append mode. Read the workbook, add to the sheet, write it back.
That is fine for anything that fits in memory; for bigger inputs, stream the old
rows into a `WorkbookWriter` and append the new ones as you go.

```js doc-example
import { eachRow, read, WorkbookWriter, write } from "gridxcel";
import assert from "node:assert";

const SOURCE = "../../test/fixtures/pandas.xlsx";

// In memory: read, extend, write.
const book = await read(SOURCE);
const sheet = book.sheets[0];
const before = sheet.rowCount;
sheet.rows.push([{ v: "new", t: "string" }, { v: 1, t: "number" }, { v: "s-new", t: "string" }]);
sheet.rowCount = sheet.rows.length;

const bytes = await write(book);
const after = await read(bytes);
assert.strictEqual(after.sheets[0].rowCount, before + 1);
console.log("rowCount", before, "->", after.sheets[0].rowCount);
console.log("last row:", JSON.stringify(after.sheets[0].rows.at(-1).map((c) => c.v)));

// Bounded memory: copy the old rows through a writer, then add the new ones.
const writer = new WorkbookWriter();
const target = writer.addSheet(sheet.name);
let copied = 0;
for await (const { r, cells } of eachRow(SOURCE)) {
  target.addRow(r, cells);
  copied++;
}
target.addRow(copied + 1, [{ v: "streamed", t: "string" }, { v: 2, t: "number" }, { v: "s-streamed", t: "string" }]);
target.end();

const written = await read(writer.finish());
console.log("streamed copy:", copied, "rows ->", written.sheets[0].rowCount);
```

Watch out for:

* Update `rowCount` when you push onto `rows` yourself. The writers trust it,
  and a stale count truncates the output.
* The streaming copy appends rows in the order `eachRow` yields them, which is
  ascending by construction. A row inserted out of order throws `ROW_ORDER`.
* Merges, column widths, styles and hyperlinks are sheet-level and survive a
  read and write; they are not carried by the row stream, so a streamed copy
  starts from a default layout.

## 10. Handle a file that fails to parse

Input from users or the network fails in two ways: it is not a workbook at all,
or it is a workbook with damaged parts. The first throws; the second reads and
reports.

```js doc-example
import { GridxcelError, read } from "gridxcel";
import { readFileSync } from "node:fs";

const FIXTURES = "../../test/fixtures/";

// Not a workbook. A typo, a truncated download and an HTML error page all land
// here, and all of them throw rather than returning an empty workbook.
for (const [label, input] of [
  ["HTML error page", new TextEncoder().encode("<!doctype html><html>404</html>")],
  ["truncated download", readFileSync(FIXTURES + "golden.xlsx").subarray(0, 900)],
  ["zero bytes", new Uint8Array(0)],
]) {
  try {
    await read(input);
    console.log(label, "-> accepted (unexpected)");
  } catch (err) {
    const typed = err instanceof GridxcelError;
    console.log(label.padEnd(20), "->", err.name + " / " + err.code, "| typed:", typed);
  }
}

// Damaged but readable: the file opens, and every problem is on book.warnings.
const damaged = await read(FIXTURES + "reader-badcells.xlsx");
console.log("\ndamaged file:", damaged.warnings.length, "warnings");
for (const w of damaged.warnings) console.log("  " + w.code + ": " + w.message);

// strict turns the first warning into a throw, when a partial read is worse
// than no read at all.
try {
  await read(FIXTURES + "reader-badcells.xlsx", { strict: true });
} catch (err) {
  console.log("\nstrict ->", err.name + " / " + err.code, "on", err.warning.code);
}
```

Watch out for:

* Every error gridxcel raises has `GridxcelError` in its chain, so one
  `instanceof` check catches container, schema, format and filesystem failures.
  The `code` property is the stable part; the message is not.
* Warnings are collected by default so one bad cell does not cost you the rest
  of the file. Use `onWarning` to log them where you already log.
* Check `response.ok` before reading an HTTP body. An error page is a
  perfectly good byte stream and will fail later, less clearly.

## 11. Read a workbook from an HTTP response

The `read` function accepts a `ReadableStream`, which is what a fetch response
body is, so the bytes go into the parser rather than into a buffer you
allocated.

```js doc-example
import { eachRow, read } from "gridxcel";
import { createServer } from "node:http";
import { readFileSync } from "node:fs";

// A local server so this recipe needs no network.
const body = readFileSync("../../test/fixtures/golden.xlsx");
const server = createServer((req, res) => {
  res.writeHead(200, {
    "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "content-length": body.length,
  });
  res.end(body);
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const url = "http://127.0.0.1:" + server.address().port + "/report.xlsx";

try {
  const response = await fetch(url);
  if (!response.ok) throw new Error("HTTP " + response.status);

  // Stream the body straight in.
  const book = await read(response.body);
  console.log("streamed from HTTP:", book.sheetNames.join(", "));

  // Or hand the bytes to the row reader without materialising the workbook.
  const second = await fetch(url);
  let rows = 0;
  for await (const event of eachRow(second.body, { sheets: ["Volume"] })) rows++;
  console.log("rows streamed:", rows);
} finally {
  server.close();
  await new Promise((resolve) => server.once("close", resolve));
}
```

Watch out for:

* A stream can only be consumed once. Issue the request again, or read the body
  into an `ArrayBuffer` first, if you need two passes.
* Check the status before reading. A 404 page is a byte stream too, and it will
  fail with `NOT_A_ZIP` only after you have read it.
* If the response is `Content-Encoding: gzip`, the fetch layer decodes it
  before you see it. Passing already-decoded bytes to `read` is correct; do not
  decode a second time.
