# API reference

Version 0.1.0. MIT. Pre-1.0: the public surface may still change before 1.0.

gridxcel is an ESM-only, zero-dependency XLSX/CSV reader and writer for Node
(>= 18) and the browser. There is no build step: `src/*.js` is what ships. `require()` is not
supported - every entry point is an ES module.

This page documents all 117 public symbols exported by the seven entry points
declared in the package export map. Every fence marked **doc-example** is executed as a
real ES module by `npm run check:docs`; if an example here stops working, the
check fails. Fixture paths in examples are relative to the repository root
(they resolve from the generated example directory).

---

## Contents

* [Install and import](#install-and-import)
* [Entry points](#entry-points)
* [Data model](#data-model)
* [gridxcel (main entry point)](#gridxcel-main-entry-point)
* [gridxcel/compat](#gridxcelcompat)
* [gridxcel/csv](#gridxcelcsv)
* [gridxcel/json](#gridxceljson)
* [gridxcel/numfmt](#gridxcelnumfmt)
* [gridxcel/xml](#gridxcelxml)
* [gridxcel/zip](#gridxcelzip)
* [Errors](#errors)
* [Guarantees and non-guarantees](#guarantees-and-non-guarantees)
* [TypeScript](#typescript)

---

## Install and import

```sh
npm install gridxcel
```

```js
import { read, write, eachRow } from "gridxcel";

const book = await read("report.xlsx", { cellDates: true });
await write(book, { path: "copy.xlsx" });
```

Do not mix module systems. `require("gridxcel")` throws, because the package
declares `type: "module"` and ships no CommonJS build; use `import` everywhere,
including inside a CommonJS project (via a dynamic `import()` if necessary).

Node builtins used at runtime: `node:zlib`, `node:stream`, `node:fs`,
`node:fs/promises` and `node:url`, all confined to `src/platform.js`. Web
globals used: `TextDecoder`, `TextEncoder`, `ReadableStream`,
`TransformStream`, `CompressionStream`, `DecompressionStream`, `Blob`, `URL`
and `btoa`/`atob`. The byte-level API (`read`, `eachRow`, `sheetToJSON`,
`sheetToCSV`, `write`, `WorkbookWriter`) runs in a browser with no build step
and no polyfill; the synchronous API and the path-based helpers are Node-only
and throw `PlatformError` `[NOT_SUPPORTED]`. See [BROWSER.md](./BROWSER.md).

---

## Entry points

| Specifier | File | Named exports | Purpose |
| --- | --- | --- | --- |
| `gridxcel` | `src/index.js` | 68 (+ default object) | The whole library, re-exported. |
| `gridxcel/compat` | `src/compat.js` | 16 (+ default object) | SheetJS-shaped surface; see [MIGRATION.md](./MIGRATION.md). |
| `gridxcel/csv` | `src/csv.js` | 7 | CSV/DSV parse and serialise. |
| `gridxcel/json` | `src/json.js` | 19 | Sheet/JSON/AOA conversion and cell helpers. |
| `gridxcel/numfmt` | `src/numfmt.js` | 13 | ECMA-376 number formats and date serials. |
| `gridxcel/xml` | `src/xml.js` | 9 | Incremental SAX parser and XML escaping. |
| `gridxcel/zip` | `src/zip.js` | 13 | ZIP container reader/writer, store and deflate. |
| `gridxcel/package.json` | `package.json` | - | Package metadata. |

68 + 16 + 7 + 19 + 9 + 13 + 13 = 145 export slots. 117 of those names are
distinct; the remainder are exported by more than one entry point (for example
`read`, `write` and `utils` appear in both `gridxcel` and
`gridxcel/compat`). Import from a subpath to keep a bundle small; every
subpath is standalone.

```js doc-example
import { VERSION } from "gridxcel";
import { readFileSync } from "node:fs";
import assert from "node:assert";

const pkg = JSON.parse(readFileSync("../../package.json", "utf8"));
assert.strictEqual(VERSION, pkg.version);
console.log("version:", VERSION, "| node:", process.versions.node);
```

---

### Counting the API

Three numbers appear in this project's documents, and they measure different
things:

| Count | Meaning |
| --- | --- |
| **117** | Distinct public symbols across all seven entry points. This is the size of the API. |
| **145** | Export slots: the same symbol exported from two entry points counts twice. |
| **69** | Exports from the root `gridxcel` entry point alone, including its default export. |

When a document says "the API is 117 symbols", that is the first row.

## Data model

Three plain objects carry everything. They are ordinary data: no classes, no
hidden state, safe to spread and to serialise (except `Map` fields).

```text
Cell = {
  v,        // value: string | number | boolean | Date | null
  t,        // "string" | "number" | "boolean" | "date" | "error" | "empty"
  a?,       // A1 address, when the cell came from a parsed file
  f?,       // formula text, without the leading "="
  z?,       // number-format id (read) or format code (write)
  zf?,      // resolved number-format code, added by the reader
  s?,       // style table index
  e?,       // error literal for t === "error", e.g. "#DIV/0!"
  serial?,  // raw Excel serial for t === "date"
  w?,       // rendered display text, when cellText was requested
  c?,       // comment text, when cellComments was requested
  l?,       // { target, tooltip } hyperlink
}

Sheet = {
  name, rows, originRow, originCol, columnCount, rowCount, merges, columns,
  rowInfo, freeze?, autoFilter?, dims?, tabColor?, hidden?, sparseRows?
}

Workbook = {
  sheets, sheetNames, dateSystem, numFmts, styles, props, definedNames,
  warnings, parts?
}
```

| Field | Type | Notes |
| --- | --- | --- |
| `v` | `string` \| `number` \| `boolean` \| `Date` \| `null` | For `t === "date"` this is a `Date` when `cellDates` was set, otherwise the serial. |
| `t` | type literal | Exactly one of the six literals listed in the block above; `empty` marks a gap that `sheetStubs` materialised. |
| `a` | string | Absent unless the reader produced it or the sheet was built by `aoaToSheet`. Derivable from `originRow`, `originCol` and the index. |
| `f` | string | Formula without the leading `=`. Present when `cellFormula` was not disabled. |
| `z` | number \| string | Polymorphic by direction: numeric id on read, code or id on write. |
| `zf` | string | Resolved format code. Present on batch **and** streaming reads. |
| `s` | number | Index into `Workbook.styles`. |
| `e` | string | Error literal such as `#DIV/0!` when `t === "error"`. |
| `serial` | number | Raw serial; kept next to the `Date` so nothing is lost. |
| `w` | string | Display text, only when `cellText` was requested. |
| `c` | string | Comment text, only when `cellComments` was requested. |
| `l` | `{ target, tooltip? }` | Hyperlink. Batch reads only; see the streaming notes. |

| Sheet field | Type | Notes |
| --- | --- | --- |
| `name` | string | Sheet name as declared in the workbook. |
| `rows` | `Cell[][]` | **Sparse** arrays indexed from column `originCol`; row `i` is sheet row `originRow + i`. |
| `originRow` | number | 1-based row number of `rows[0]`. Defaults to 1. |
| `originCol` | number | 0-based column of `cells[0]`. Defaults to 0. |
| `columnCount` | number | Highest column index present + 1. |
| `rowCount` | number | Highest row number present, as declared by the file. |
| `merges` | `string[]` | Merge ranges in A1 notation. |
| `columns` | `ColumnInfo[]` | Width, hidden flag and outline level per column. |
| `rowInfo` | `Map<number, RowInfo>` | Keyed by 1-based row number. |
| `freeze` | `{ row, col }` | Frozen pane, when present. |
| `autoFilter` | string | AutoFilter range, when present. |
| `dims` | string | Declared `<dimension>` extent. |
| `tabColor` | string | Normalised to `#RRGGBB`. |
| `hidden` | boolean | Sheet visibility. |
| `sparseRows` | `Map<number, Cell[]>` | Rows too far from their neighbours to keep densely; see [Guarantees](#guarantees-and-non-guarantees). |

Rows are sparse arrays, so a row using only columns A and Z costs two entries
rather than twenty-six. `originRow` and `originCol` say where the data
starts, and `cell.a` is derived from them plus the index rather than stored
per cell - an address string per cell is real memory on a 400,000-cell sheet.

`book.warnings` is always an array, empty for a clean file, and it is the same
array that is handed to `onWarning`. `book.sheetNames` is populated even when
`sheets: false` skipped the worksheet parsing.

**Dates use the local wall clock.** A cell displaying `2024-03-15 14:30:45`
reads back as a `Date` whose **local** fields are `2024-03-15 14:30:45`, in any
timezone. Never compare a parsed date with `toISOString()`, which converts to
UTC and shifts by the host offset; compare local fields or serials. The suite is
verified under `TZ=UTC`, `TZ=America/New_York` and `TZ=Asia/Tokyo`.

---

## gridxcel (main entry point)

`import ... from "gridxcel"` re-exports the whole library. The table below is
the complete named surface, alphabetically; the sections that follow give the
long form and a runnable example for each significant function.

| Export | Signature | Summary |
| --- | --- | --- |
| `BUILTIN_FORMATS` | `Record<number, string>` | ECMA-376 built-in numFmtId to format code. |
| `BadZipError` | `class extends GridxcelError` | Base class for container failures. |
| `Crc32` | `class` | Incremental CRC-32 over a sequence of buffers. |
| `FormatError` | `class extends GridxcelError` | The bytes are not a workbook at all. |
| `GridxcelError` | `class extends Error` | Base class of everything the package throws. |
| `NotImplementedError` | `class extends GridxcelError` | A format or option this build does not support. |
| `SchemaError` | `class extends GridxcelError` | Well-formed XML that violates the OOXML schema. |
| `SheetWriter` | `class` | One worksheet inside a `WorkbookWriter`. |
| `StrictWarningError` | `class extends GridxcelError` | `strict: true` turned a warning into a throw. |
| `StyleTable` | `class` | Interns fonts, fills, borders, xf records and format ids. |
| `VERSION` | `string` | Package version, `0.1.0` at the time of writing. |
| `WorkbookWriter` | `class` | Streaming xlsx writer; parts are emitted on `finish()`. |
| `XmlParser` | `class` | Incremental SAX parser with `write()` / `end()`. |
| `ZipArchive` | `class` | Read a ZIP container; verifies CRC-32 by default. |
| `ZipError` | `class extends BadZipError` | Container-level failure; carries `offset`. |
| `ZipWriter` | `class` | Build a ZIP container from parts. |
| `aoaToSheet` | `(rows, options?) -> Sheet` | Array of arrays to `Sheet`. |
| `buildParts` | `(book, opts?) -> Map<string, string>` | Every OPC part as raw XML. |
| `cellFormatCode` | `(book, cell) -> string \| number \| undefined` | Resolve a read cell back to its format code. |
| `countCells` | `(book) -> number` | Count non-empty cells in a workbook. |
| `crc32` | `(input, seed?) -> number` | One-shot CRC-32 of a buffer or list of buffers. |
| `createBookWriter` | `(book, opts?) -> WorkbookWriter` | Build a writer from a read-model workbook. |
| `createZip` | `(files, opts?) -> Uint8Array` | Build an archive from pairs or a plain object. |
| `csvToSheet` | `(text, options?) -> Promise<Sheet>` | Parse CSV text into a `Sheet`. |
| `dateToSerial` | `(date, date1904?) -> number` | `Date` to Excel serial, using local fields. |
| `decodeCell` | `(a1) -> { r, c }` | A1 address to zero-based row and column. |
| `decodeCol` | `(letters) -> number` | Column letters to a zero-based index. |
| `decodeEntities` | `(s) -> string` | Decode XML entities and character references. |
| `decodeRange` | `(ref) -> { s, e }` | Range text to a pair of addresses. |
| `eachRow` | `(input, options?) -> AsyncIterable` | Stream rows with bounded memory. |
| `encodeCell` | `(r, c) -> string` | Zero-based row/column to A1. |
| `encodeCol` | `(c) -> string` | Zero-based column to letters. |
| `encodeRange` | `(s, e) -> string` | Address pair to `A1:B2` text. |
| `encodeXmlAttr` | `(s) -> string` | Escape text for an XML attribute value. |
| `encodeXmlText` | `(s) -> string` | Escape text for an XML text node. |
| `formatNumber` | `(value, code?, opts?) -> string` | Render a number through a format code. |
| `formatValue` | `(value, code?, opts?) -> string` | Render any cell value for display. |
| `generalNumber` | `(v) -> string` | Excel-style `General` rendering. |
| `isDateFormat` | `(codeOrId) -> boolean` | True when a format renders a date/time. |
| `jsonToSheet` | `(rows, options?) -> Promise<Sheet>` | Objects to `Sheet`, keys in first-seen order. |
| `mergeBooks` | `(...books) -> Workbook` | Concatenate the sheets of several workbooks. |
| `normalizeBook` | `(input, options?) -> Workbook` | Coerce any accepted workbook shape. |
| `normalizeCell` | `(input) -> Cell \| undefined` | Coerce a cell input; undefined for empty. |
| `normalizeColor` | `(color) -> string \| object \| undefined` | Canonicalise `#RRGGBB` / `AARRGGBB` / theme colours. |
| `normalizeSheet` | `(input, name?) -> Sheet` | Coerce rows or a sheet object into a `Sheet`. |
| `numberLiteral` | `(n) -> string` | Shortest round-tripping decimal spelling for `<v>`. |
| `parseStyles` | `(xml) -> object` | Parse `styles.xml` into formats, styles and `isDate()`. |
| `parseTag` | `(inner) -> { name, local, attrs }` | Split a tag body into name and attributes. |
| `read` | `(input, options?) -> Promise<Workbook>` | Read a workbook. |
| `readFile` | `(path, options?) -> Promise<Workbook>` | Read a workbook from a path. |
| `readJSON` | `(input, options?) -> Promise<object[]>` | Read the first sheet as objects. |
| `readSync` | `(bytes, options?) -> Workbook` | Synchronous read from bytes. |
| `renameSheet` | `(book, from, to) -> Workbook` | Rename a sheet, rejecting duplicates and illegal names. |
| `serialToDate` | `(serial, date1904?) -> Date` | Serial to a `Date` in local wall-clock fields. |
| `sheetNames` | `(input) -> Promise<string[]>` | Sheet names without parsing worksheets. |
| `sheetToAOA` | `(sheet, options?) -> Promise<any[][]>` | Dense array of arrays. |
| `sheetToCSV` | `(sheet, options?) -> Promise<string>` | Serialise a sheet to CSV text. |
| `createCSVStream` | `(sheet, options?, streamOpts?) -> AsyncGenerator<string>` | The streaming form of `sheetToCSV`: same bytes, one row at a time. |
| `sheetToJSON` | `(sheet, options?) -> Promise<object[]>` | Sheet to objects (SheetJS `sheet_to_json` semantics). |
| `splitSheetRef` | `(ref) -> [string \| null, string]` | Split `Sheet1!A1:B2` into name and range. |
| `toBuffer` | `(book, options?) -> Uint8Array` | Alias of `writeWorkbookSync` for explicit naming. |
| `toRange` | `(range) -> { s, e } \| undefined` | Normalise any accepted range spelling. |
| `utils` | `object` | The SheetJS-shaped helper namespace. |
| `write` | `(book, options?) -> Promise<Uint8Array \| string \| undefined>` | Write a workbook to bytes or a path. |
| `writeFile` | `(book, path, options?) -> Promise<Uint8Array \| undefined>` | Write a workbook to a path. Resolves to the bytes for xlsx/xlsm, undefined for csv/json. |
| `writeRows` | `(rows, options?) -> Promise<Uint8Array \| string \| undefined>` | Write raw rows straight to xlsx bytes. |
| `writeWorkbook` | `(book, options?) -> Promise<Uint8Array \| string \| undefined>` | Async writer; does not normalise its input. |
| `writeWorkbookSync` | `(book, options?) -> Uint8Array \| string` | Synchronous writer. |

### Reading

#### `read(input, options?) -> Promise<Workbook>`

`input` may be a `Uint8Array`, an `ArrayBuffer`, a file path, a `URL`
(including `file:` URLs), a `ReadableStream`, a Node readable stream, or any
async iterable of byte chunks.

```js doc-example
import { read } from "gridxcel";
import { createReadStream } from "node:fs";
import assert from "node:assert";

const fromPath = await read("../../test/fixtures/golden.xlsx");
assert.deepStrictEqual(fromPath.sheets.map((s) => s.name), ["Types", "Layout", "Volume", "Empty"]);

const fromStream = await read(createReadStream("../../test/fixtures/pandas.xlsx"));
assert.strictEqual(fromStream.sheets.length, 1);
console.log("path:", fromPath.sheets.length, "sheets | stream:", fromStream.sheets[0].rowCount, "rows");
```

#### Read options

| Option | Type | Default | Meaning |
| --- | --- | --- | --- |
| `sheets` | boolean \| Array | `true` | `false` reads metadata only; an array selects sheets by name or 0-based index. |
| `sheetStubs` | boolean | `false` | Emit `{ t: "empty" }` for gaps in the used range. |
| `cellDates` | boolean | `false` | Return `Date` objects for date cells instead of serial numbers. |
| `cellText` | boolean | `false` | Fill `cell.w` with the number-format-rendered display text. |
| `cellFormula` | boolean | `true` | Include `cell.f`. Set false to strip formulas. |
| `cellComments` | boolean | `false` | Read cell comments into `cell.c`. |
| `range` | string \| object | whole sheet | Restrict to a range, e.g. `A1:D20` or `{ s, e }`. Cells outside it are dropped. |
| `strict` | boolean | `false` | Throw a `StrictWarningError` on the first problem instead of collecting it. |
| `onWarning` | function | - | Called for each non-fatal problem; the same objects land on `book.warnings`. |
| `keepParts` | boolean | `false` | Keep the raw package parts on `book.parts`. |

These are exactly the keys the reader consults; other keys are ignored. In
particular there is no `verifyCrc` read option: every part is verified against
the CRC-32 in the central directory, always. Opting out means constructing a
`ZipArchive` yourself with `{ verifyCrc: false }`.

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/xlsxwriter-full.xlsx", {
  sheets: ["Report"],
  cellDates: true,
  cellText: true,
  range: "A1:F5",
});
const sheet = book.sheets[0];
const row = sheet.rows[0];
assert.strictEqual(sheet.name, "Report");
assert.strictEqual(row[0].v, "Product");
assert.ok(row[0].w);
console.log(sheet.name, "| cells in row 1:", row.filter(Boolean).length);
console.log("A1:", JSON.stringify({ t: row[0].t, v: row[0].v, w: row[0].w }));
```

#### `readSync(bytes, options?) -> Workbook`

Synchronous read. Bytes only - a stream cannot be consumed synchronously - and
Node only. It resolves the reader through `require` of an ES module, so it is
for code paths that already hold the buffer and cannot await.

```js doc-example
import { readSync } from "gridxcel";
import { readFileSync } from "node:fs";
import assert from "node:assert";

const bytes = readFileSync("../../test/fixtures/pandas.xlsx");
const book = readSync(bytes);
assert.strictEqual(book.sheets.length, 1);
console.log("sheets:", book.sheets.length, "| rows on sheet 1:", book.sheets[0].rowCount);
```

#### `readFile(path, options?) -> Promise<Workbook>`

States the intent and gives filesystem problems a typed code:
`FILE_NOT_FOUND`, `FILE_ACCESS_DENIED`, `FILE_IS_DIRECTORY`.

```js doc-example
import { readFile, countCells } from "gridxcel";
import assert from "node:assert";

const book = await readFile("../../test/fixtures/pandas.xlsx");
assert.strictEqual(book.sheets.length, 1);
console.log("sheets:", book.sheets.length, "| cells:", countCells(book));
```

#### `sheetNames(input) -> Promise<string[]>`

Sheet names without parsing a worksheet, so it is cheap enough to call before
deciding what to read.

```js doc-example
import { sheetNames } from "gridxcel";
import assert from "node:assert";

const names = await sheetNames("../../test/fixtures/golden.xlsx");
assert.strictEqual(names.length, 4);
console.log(names.join(","));
```

#### `readJSON(input, options?) -> Promise<object[]>`

Reads the first sheet and converts it to objects in one call. `options` are
both read options and `sheetToJSON` options, so `raw`, `header`, `defval`
and `blankrows` work here too.

```js doc-example
import { readJSON } from "gridxcel";
import assert from "node:assert";

const rows = await readJSON("../../test/fixtures/pandas.xlsx", { raw: true });
// 101 rows: the header row is data because raw: true keeps row 1.
assert.strictEqual(rows.length, 101);
console.log(rows.length, JSON.stringify(rows[0]));
```

#### Streaming with `eachRow(input, options?)`

`eachRow` walks every selected row of every selected sheet as an
`AsyncIterable` of `{ sheet, r, cells }`, holding one row at a time instead
of materialising a sheet. It accepts the same read options as `read`. This is
the entry point for files that do not fit in memory.

```js doc-example
import { eachRow } from "gridxcel";
import assert from "node:assert";

let rows = 0;
let cells = 0;
for await (const ev of eachRow("../../test/fixtures/golden.xlsx")) {
  rows++;
  cells += ev.cells.filter(Boolean).length;
}
assert.strictEqual(rows, 205);
console.log("rows:", rows, "| cells:", cells);
```

Streaming differs from a batch read in four documented ways:

| Difference | Streaming (`eachRow`) | Batch (`read`) |
| --- | --- | --- |
| Column indexing | **Absolute**: `cells[0]` is column A. | Rebases to `originCol`. |
| `sheetStubs` gaps | Filled from column A. | Filled from `originCol`. |
| `cell.zf` | Always present on styled cells, because a stream event has no workbook to resolve `cell.z` against. | Present as well. |
| `cell.l` (hyperlinks) | **Not available**: `<hyperlinks>` follows `<sheetData>`, so honouring it in one pass would mean retaining rows. | Available. |

Comments are streamed. Throughput matches a full parse at a fraction of the
memory: on a 200,000 x 10 workbook, 77,309 rows/s at 13.0 MiB peak heap
(5.9 MiB retained) against 68,086 rows/s at 224.9 MiB peak heap (164.9 MiB
retained) for a full read. See [PERFORMANCE.md](./PERFORMANCE.md).

### Writing

#### `write(book, options?) -> Promise<Uint8Array | string | undefined>`

Accepts a `Workbook`, an array of rows, an array of sheets, a
`{ SheetNames, Sheets }` object, or a single sheet object. Returns bytes, or
text for the text book types, or writes to disk and resolves `undefined` when
`options.path` is set.

```js doc-example
import { write } from "gridxcel";
import assert from "node:assert";

const bytes = await write(
  [
    ["name", "score", "when", "active"],
    ["ada", 42, new Date(2024, 2, 15, 14, 30), true],
    ["grace", 37.5, new Date(1999, 11, 31), false],
  ],
  { bookType: "xlsx" },
);
assert.ok(bytes.length > 1000);
console.log("bytes:", bytes.length);
```

#### Write options

| Option | Type | Default | Meaning |
| --- | --- | --- | --- |
| `bookType` | `"xlsx"` \| `"csv"` \| `"json"` | `"xlsx"` | Output format. `xlsm` is accepted as an alias of `xlsx`. |
| `path` | string | - | Write to this file instead of returning the value. |
| `overwrite` | boolean | `false` | Allow replacing an existing file; the open is atomic (`wx`), so the check cannot race. |
| `compressionLevel` | 0-9 | `6` | 0 stores every part uncompressed. |
| `type` | string | - | Shape of the return value, see below. |
| `recalc` | boolean | `true` | Mark the workbook for recalculation when opened. |
| `deterministic` | boolean | `false` | Fixed timestamps, so repeated writes are byte-identical. |
| `inlineStrings` | boolean | `false` | Emit inline strings instead of a shared-string table. |
| `date1904` | boolean | `false` | Write the 1904 date system. |
| `activeTab` | number | `0` | Index of the selected sheet. |
| `defaultRowHeight` | number | - | Sheet default row height in points. |
| `defaultColWidth` | number | - | Sheet default column width in characters. |
| `application` | string | `gridxcel` | Value written to `docProps/app.xml`. |
| `indent` | number | `2` | Indentation for `bookType: "json"`. |
| `sheet` | string | `Sheet1` | Default sheet name when raw rows are passed. |

`type` accepts `buffer`, `array`, `base64`, `binary` and `string`. The
last two are latin1 text, matching what SheetJS returns for a spreadsheet.

```js doc-example
import { write } from "gridxcel";
import assert from "node:assert";

const rows = [["a", 1], ["b", 2]];
const seen = {};
for (const type of ["buffer", "array", "base64", "binary", "string"]) {
  const out = await write(rows, { type });
  seen[type] = out instanceof Uint8Array ? "Uint8Array"
    : out instanceof ArrayBuffer ? "ArrayBuffer"
    : typeof out;
}
assert.deepStrictEqual(seen, {
  buffer: "Uint8Array", array: "ArrayBuffer",
  base64: "string", binary: "string", string: "string",
});
console.log(JSON.stringify(seen));
```

#### `writeFile(book, path, options?) -> Promise<Uint8Array | undefined>`

Writes to `path` and resolves to the encoded bytes for `xlsx` and `xlsm`, or to
`undefined` for `csv` and `json`, which write text. Unlike the compat writer,
this one overwrites by default and throws `GridxcelError` / `FILE_EXISTS` when
`overwrite: false` is passed and the file is already there.

```js doc-example
import { read, writeFile } from "gridxcel";
import { existsSync, rmSync } from "node:fs";
import assert from "node:assert";

const book = await read("../../test/fixtures/pandas.xlsx");
const out = "../doc-example-written.xlsx";

// xlsx: the file is written and the encoded bytes come back.
const bytes = await writeFile(book, out, { overwrite: true });
assert.ok(existsSync(out));
assert.ok(bytes instanceof Uint8Array);

// csv: the file is written and nothing comes back.
const csv = await writeFile(book, "../doc-example-written.csv", { bookType: "csv", overwrite: true });
assert.strictEqual(csv, undefined);

console.log("xlsx bytes:", bytes.length, "| csv return:", csv, "| both written:", existsSync(out));
rmSync(out, { force: true });
rmSync("../doc-example-written.csv", { force: true });
```

#### `writeRows(rows, options?) -> Promise<...>`

Normalises an array of rows into a workbook and writes it in one call, so a
consumer that only has data does not have to learn the `Workbook` shape.

```js doc-example
import { writeRows, writeWorkbook, writeWorkbookSync, toBuffer } from "gridxcel";
import assert from "node:assert";

const rows = [["x", 1], ["y", 2]];
const asyncBytes = await writeRows(rows, { type: "buffer" });
const syncBytes = writeWorkbookSync([{ name: "S", rows }], {});
const arrayBuffer = toBuffer(rows, { type: "array" });
assert.ok(asyncBytes.length > 0 && syncBytes.length > 0 && arrayBuffer.byteLength > 0);
console.log("writeRows:", asyncBytes.length, "| sync:", syncBytes.length, "| array:", arrayBuffer.byteLength);
```

#### `writeWorkbook(book, options?)`, `writeWorkbookSync(book, options?)`, `toBuffer(book, options?)`

The low-level writers. They do **not** normalise their input: pass a
`Workbook`, or call `normalizeBook` first. `writeWorkbook` is the async
form, `writeWorkbookSync` the synchronous one, and `toBuffer` is an alias of
`writeWorkbookSync` for callers who want the name to say bytes. All three
honour `options.path`, `options.type`, `options.overwrite` and
`options.compressionLevel`; they emit xlsx only and throw
`NotImplementedError` for another `bookType`.

#### `WorkbookWriter` and `SheetWriter`

For workbooks too large to hold in memory, push rows into a sheet as they are
produced and let the writer emit parts as it goes. `WorkbookWriter` is
constructed with the same options as `write` (`compressionLevel`,
`deterministic`, `inlineStrings`, `date1904`, `recalc`, `application`,
`activeTab`, `defaultRowHeight`, `defaultColWidth`); `addSheet(name, opts)`
returns a `SheetWriter` whose `addRow(rowIndex, cells)` appends one row
(`rowIndex` is 1-based, rows must arrive in ascending order, `undefined`
auto-increments), `end()` finalises the sheet, `rowCount` reports how many
rows were written, and `finish()` returns the archive bytes. `stringParts()`
and `parts()` expose the XML before zipping.

```js
import { WorkbookWriter } from "gridxcel";

const wb = new WorkbookWriter({ compressionLevel: 6 });
const sheet = wb.addSheet("Data");
for (let r = 1; r <= 1_000_000; r++) {
  // Pass undefined to auto-increment from the previous row.
  sheet.addRow(undefined, [{ v: r, t: "number" }]);
}
sheet.end();
const bytes = wb.finish();
```

```js doc-example
import { WorkbookWriter, SheetWriter } from "gridxcel";
import assert from "node:assert";

const wb = new WorkbookWriter({ deterministic: true });
const sheet = wb.addSheet("Data");
assert.ok(sheet instanceof SheetWriter);
// addRow(rowIndex, cells): rowIndex is 1-based; cells are indexed from column A.
for (let r = 1; r <= 500; r++) sheet.addRow(r, [{ v: "row" + r, t: "string" }, { v: r, t: "number" }]);
sheet.end();
const bytes = wb.finish();
assert.ok(bytes.length > 1000);
console.log("rows:", sheet.rowCount, "| bytes:", bytes.length);
```

A writer is one-shot: adding a sheet after `finish()` throws `WRITER_CLOSED`,
and appending to an ended sheet throws `SHEET_CLOSED`.

#### `createBookWriter(book, opts?) -> WorkbookWriter` and `buildParts(book, opts?) -> Map<string, string>`

`createBookWriter` configures a `WorkbookWriter` from a read-model
`Workbook` and pushes every row into it, seeding the style table so numeric
`cell.s` indices keep working after the cellXfs table is deduplicated.
`buildParts` is `createBookWriter` plus `stringParts()`: it returns the raw
XML of every package part without touching the filesystem, which is what the
tests and the streaming API share.

```js doc-example
import { buildParts, createBookWriter, normalizeBook } from "gridxcel";
import assert from "node:assert";

const book = normalizeBook([["a", 1]]);
const parts = buildParts(book);
assert.ok(parts.has("xl/worksheets/sheet1.xml"));
assert.ok(parts.get("xl/_rels/workbook.xml.rels").includes("worksheets/sheet1.xml"));
const writer = createBookWriter(book);
assert.strictEqual(writer.parts().size, parts.size);
console.log("parts:", [...parts.keys()].join(","));
```

### Workbook and sheet helpers

#### `normalizeBook(input, options?) -> Workbook`

Coerces the many shapes a caller may pass into a `Workbook`: an array of
rows, an array of sheets, `{ sheets }`, SheetJS-shaped
`{ SheetNames, Sheets }`, or a single sheet object. `options.date1904`
selects the 1904 date system. Anything else throws `GridxcelError` with code
`BAD_BOOK`.

#### `normalizeSheet(input, name?) -> Sheet`

Coerces an array of rows (or an array of objects keyed by column letter) into
a `Sheet`. A sheet-shaped object passes through with its name preserved.
Objects that are neither throw `BAD_SHEET`.

`gridxcel/json` exports a wider `normalizeSheet` that also accepts SheetJS
worksheets (`!ref` / `!data` / `A1` keys) and whole books, and selects a
sheet with `options.sheet` or `options.sheetName`.

#### `normalizeCell(input) -> Cell | undefined`

Coerces a value into the internal `Cell` shape: `string`, `number`,
`boolean` and `Date` become the obvious type, `null`/`undefined` become
`undefined` (no cell), and an object carrying `v`, `f` or `t` passes
through with its type inferred when absent. An object with none of those keys
throws `BAD_CELL` - one rule, in one place, so every entry point answers the
same way.

#### `mergeBooks(...books) -> Workbook`

Concatenates the sheets of several workbooks into a new one. Metadata is not
merged: the result starts from an empty `props`, `styles` and `numFmts`,
and 1904 wins if any input used it.

#### `renameSheet(book, from, to) -> Workbook`

Renames in place, applying Excel's rules: at most 31 characters, none of
`\\ / ? * [ ] :`, and no duplicate name. Violations throw `NO_SHEET`,
`BAD_SHEET_NAME` or `DUPLICATE_SHEET`.

#### `countCells(book) -> number`

Counts cells whose `t` is set and not `empty`; handy in tests and
diagnostics.

```js doc-example
import { normalizeBook, normalizeSheet, normalizeCell, mergeBooks, renameSheet, countCells, GridxcelError } from "gridxcel";
import assert from "node:assert";

const sheet = normalizeSheet([["a", 1], ["b", 2]], "Data");
assert.strictEqual(normalizeCell(42).t, "number");
assert.strictEqual(normalizeCell(new Date(2024, 0, 1)).t, "date");

const book = normalizeBook({ sheets: [sheet] });
assert.strictEqual(countCells(book), 4);
renameSheet(book, "Data", "Renamed");
assert.strictEqual(book.sheets[0].name, "Renamed");

const merged = mergeBooks(book, normalizeBook([["z"]]));
assert.strictEqual(merged.sheets.length, 2);
try {
  normalizeCell({ a: 1 });
  assert.fail("should have thrown");
} catch (e) {
  assert.ok(e instanceof GridxcelError);
  console.log("cell without v rejected:", e.code);
}
console.log("sheets:", merged.sheets.map((s) => s.name).join(","), "| cells:", countCells(merged));
```

#### `utils`

The SheetJS-shaped helper namespace, for migration. It exists on both
`gridxcel` and `gridxcel/compat`; the two differ in behaviour for
`format_cell` (see [LIMITATIONS.md](./LIMITATIONS.md)).

| Member | Signature | Notes |
| --- | --- | --- |
| `aoa_to_sheet` | `(rows, opts?) -> Sheet` | Array of arrays to sheet. |
| `json_to_sheet` | `(rows, opts?) -> Promise<Sheet>` | Objects to sheet. |
| `sheet_to_json` | `(sheet, opts?) -> Promise<object[]>` | Sheet to objects. |
| `sheet_to_csv` | `(sheet, opts?) -> Promise<string>` | Sheet to CSV. |
| `sheet_to_aoa` | `(sheet, opts?) -> Promise<any[][]>` | Sheet to dense arrays. |
| `book_new` | `() -> Workbook` | Empty workbook. |
| `sheet_add_aoa` | `(sheet, rows, opts?) -> Sheet` | Append rows at `opts.origin`. |
| `book_append_sheet` | `(book, sheet, name?) -> Workbook` | Append a sheet. |
| `decode_range` | `(ref) -> { s, e }` | Same as `decodeRange`. |
| `encode_cell` | `({ r, c }) -> string` | Same as `encodeCell`. |
| `encode_range` | `(a, b?) -> string` | Accepts two addresses or one range string. |
| `format_cell` | `(cell, v?, opts?) -> string` | Display text through `formatValue`. |

```js doc-example
import { utils } from "gridxcel";
import assert from "node:assert";

const sheet = utils.aoa_to_sheet([["a", 1]]);
const book = utils.book_new();
utils.book_append_sheet(book, sheet, "Data");
assert.strictEqual(utils.encode_cell({ r: 0, c: 1 }), "B1");
assert.deepStrictEqual(utils.decode_range("A1:B2").e, { r: 1, c: 1 });
const rows = await utils.sheet_to_json(sheet, { header: 1 });
console.log("sheets:", book.sheets.map((s) => s.name).join(","), "| grid:", JSON.stringify(rows[0]));
```

### Conversion helpers on the root entry

The root entry re-exports the converters as async functions so one import
covers reading, conversion and writing. The subpath sections below document
them in full; this is the summary.

| Function | Returns | Notes |
| --- | --- | --- |
| `aoaToSheet(rows, options?)` | `Sheet` | Synchronous. Sheet name from `options.sheet`. |
| `sheetToAOA(sheet, options?)` | `Promise<any[][]>` | Dense grid; `null` for gaps unless `defval` is set. |
| `sheetToJSON(sheet, options?)` | `Promise<object[]>` | SheetJS header semantics, dense `header: 1`. |
| `jsonToSheet(rows, options?)` | `Promise<Sheet>` | Keys in first-seen order. |
| `sheetToCSV(sheet, options?)` | `Promise<string>` | RFC 4180 plus the extensions real files need. |
| `csvToSheet(text, options?)` | `Promise<Sheet>` | Type inference by content. |

```js doc-example
import { read, sheetToAOA } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx", { sheets: ["Volume"] });
const grid = await sheetToAOA(book.sheets[0], { range: "A1:B3" });
assert.strictEqual(grid.length, 3);
console.log(JSON.stringify(grid));
```

### Addressing

`encodeCol`, `decodeCol`, `encodeCell`, `decodeCell`, `encodeRange`,
`decodeRange`, `toRange` and `splitSheetRef` convert between A1 text and
zero-based `{ r, c }` addresses. Absolute markers (a dollar sign before the column or the row) are accepted and
ignored, trailing junk is rejected with `BAD_ADDRESS` rather than silently
dropped, and the maximum column is `XFD` (16383).

```js doc-example
import { encodeCol, decodeCol, encodeCell, decodeCell, encodeRange, decodeRange, toRange, splitSheetRef } from "gridxcel";
import assert from "node:assert";

assert.strictEqual(encodeCol(0), "A");
assert.strictEqual(encodeCol(16383), "XFD");
assert.strictEqual(decodeCol("AA"), 26);
assert.strictEqual(encodeCell(0, 0), "A1");
assert.deepStrictEqual(decodeCell("$B$7"), { r: 6, c: 1 });
assert.strictEqual(encodeRange({ r: 0, c: 0 }, { r: 9, c: 3 }), "A1:D10");
assert.deepStrictEqual(decodeRange("Sheet1!A1:B2").e, { r: 1, c: 1 });
assert.strictEqual(toRange("A1:B2").e.c, 1);
assert.deepStrictEqual(splitSheetRef("'My Sheet'!A1"), ["My Sheet", "A1"]);
console.log(encodeCol(16383), encodeRange({ r: 0, c: 0 }, { r: 9, c: 3 }), JSON.stringify(splitSheetRef("'My Sheet'!A1")));
```

#### `toRange(range)`

Accepts a range string (`"A1:B2"`), a bare row number (1-based; runs to the
last used column), or an already parsed `{ s, e }` object. Returns `undefined`
for a falsy input and throws `BAD_RANGE` for anything unusable.

### Number formats

| Function | Signature | Notes |
| --- | --- | --- |
| `formatNumber` | `(value, code?, opts?) -> string` | ECMA-376 grammar; section selection, percent, grouping, fractions, elapsed time, `AM/PM`. |
| `formatValue` | `(value, code?, opts?) -> string` | Dispatches on the runtime type; booleans render `TRUE`/`FALSE`. |
| `isDateFormat` | `(codeOrId) -> boolean` | Accepts a code or a built-in id; cached per code. |
| `generalNumber` | `(v) -> string` | Excel's 11-significant-digit `General` rendering. |
| `serialToDate` | `(serial, date1904?) -> Date` | Local wall-clock fields. |
| `dateToSerial` | `(date, date1904?) -> number` | Uses the `Date`'s local fields. |
| `BUILTIN_FORMATS` | `Record<number, string>` | The ECMA-376 built-in table. |
| `cellFormatCode` | `(book, cell) -> string \| number \| undefined` | Resolves `cell.s` through `book.styles`, then `cell.z` through `book.numFmts`, then the built-in table. |

```js doc-example
import { read, cellFormatCode, formatNumber, isDateFormat, dateToSerial, serialToDate } from "gridxcel";
import assert from "node:assert";

assert.strictEqual(formatNumber(1234.5678, "#,##0.00"), "1,234.57");
assert.strictEqual(formatNumber(0.125, "0.00%"), "12.50%");
assert.strictEqual(formatNumber(-1234.5, "#,##0.00;[Red](#,##0.00)"), "(1,234.50)");
assert.strictEqual(isDateFormat("yyyy-mm-dd"), true);
assert.strictEqual(isDateFormat("#,##0.00"), false);

const serial = dateToSerial(new Date(2024, 2, 15, 14, 30, 45));
assert.strictEqual(serial, 45366.6046875);
assert.strictEqual(serialToDate(serial).getHours(), 14);

const book = await read("../../test/fixtures/golden.xlsx", { sheets: ["Types"] });
const cell = book.sheets[0].rows[0][3];
assert.strictEqual(cellFormatCode(book, cell), "yyyy-mm-dd hh:mm:ss");
console.log(formatNumber(1234.5678, "#,##0.00"), "| D1 code:", cellFormatCode(book, cell));
```

`cell.z` is the numeric id; the code is reachable three other ways so a
caller never has to guess: `cell.zf`, `book.styles[cell.s].numFmt` and
`book.numFmts.get(cell.z)`. `cellFormatCode` picks the right one for you.

### Styles

| Export | Signature | Notes |
| --- | --- | --- |
| `StyleTable` | `class` | Interns fonts, fills, borders and xf records; `xf(style)` returns a cellXfs index, `numFmtId(code)` allocates a custom id from 164, `toXml()` emits `styles.xml`. |
| `parseStyles` | `(xml) -> object` | The one parser for the part. Returns `{ numFmts, styles, isDate, fonts, fills, borders, cellXfs }`, where each style carries both `numFmt` (raw id, for writing) and `numFmtCode` (resolved, for reading) and `numFmts` includes the built-in formats; `isDate` also knows the built-in ids, so numFmtId 14 classifies as a date even with no `numFmts` element. |
| `normalizeColor` | `(color) -> string \| object \| undefined` | `#RRGGBB` and `RRGGBB` become upper-case `AARRGGBB`; theme, indexed and auto references stay objects. Invalid colours throw `BAD_STYLE`. |
| `numberLiteral` | `(n) -> string` | Shortest decimal spelling that parses back to the identical double, with exponent notation expanded because Excel writes plain decimals in `<v>`. |

```js doc-example
import { StyleTable, parseStyles, normalizeColor, numberLiteral } from "gridxcel";
import assert from "node:assert";

const table = new StyleTable();
const xf = table.xf({ numFmt: "#,##0.00", font: { b: true } });
assert.ok(xf > 0);
assert.strictEqual(normalizeColor("#ff0000"), "FFFF0000");

const parsed = parseStyles('<?xml version="1.0"?><styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><cellXfs count="2"><xf numFmtId="164"/><xf numFmtId="14"/></cellXfs></styleSheet>');
assert.strictEqual(parsed.numFmts.get(164), "0.000");
assert.strictEqual(parsed.isDate(1), true, "numFmtId 14 is a built-in date format");
console.log("xf:", xf, "| colour:", normalizeColor("#ff0000"), "| 1e21:", numberLiteral(1e21));
```

### Container

| Export | Signature | Notes |
| --- | --- | --- |
| `crc32` | `(input, seed?) -> number` | One-shot CRC-32 of a buffer or a list of buffers treated as one stream. |
| `Crc32` | `class` | Incremental: `update(buf)`, `digest()`, `reset()`. |
| `ZipArchive` | `class` | Parses the central directory; `open(path)` reads from disk. |
| `ZipWriter` | `class` | `add(name, data, opts)` then `build()` (or `toBuffer()`). |
| `createZip` | `(files, opts?) -> Uint8Array` | Builds from `[name, data]` pairs, a `Map` or a plain object. |
| `ZipError` | `class extends BadZipError` | Codes `NOT_ZIP`, `BAD_LOCAL_HEADER`, `UNSUPPORTED_METHOD`, `CRC_MISMATCH`, `INFLATE_FAILED`; carries `offset`. |

`ZipArchive` exposes `entries`, `byName`, `comment`, `size`, `get`,
`has`, `read`, `readText`, `names()` and a default iterator. Each
`ZipEntry` carries `name`, `size`, `compressedSize`, `method`, `crc`,
`localOffset`, `isDirectory`, `bytes()`, `text()` and `stream()`.
Verification is on by default; pass `new ZipArchive(bytes, { verifyCrc: false })`
to opt out at a measured cost of about 1.2 GB/s of payload.

```js doc-example
import { ZipArchive, ZipWriter, createZip, crc32, Crc32 } from "gridxcel";
import { ZipEntry } from "gridxcel/zip";
import assert from "node:assert";

const bytes = createZip([["hello.txt", "hello world"], ["data/stored.bin", new Uint8Array([1, 2, 3])]], { compressionLevel: 6 });
const archive = new ZipArchive(bytes);
const entry = archive.get("hello.txt");
assert.ok(entry instanceof ZipEntry);
assert.strictEqual(archive.readText("hello.txt"), "hello world");

const writer = new ZipWriter({ deterministic: true });
writer.add("stored.bin", new Uint8Array([1, 2, 3]), { compress: false });
const stored = new ZipArchive(writer.build()).get("stored.bin");
assert.strictEqual(stored.method, 0, "compress:false stores the entry");

const payload = new TextEncoder().encode("123456789");
assert.strictEqual(crc32(payload), 0xcbf43926);
const rolling = new Crc32();
rolling.update(payload);
assert.strictEqual(rolling.digest(), 0xcbf43926);
console.log("entries:", archive.names().join(","), "| crc32:", crc32(payload).toString(16));
```

### Parsing primitives

| Export | Signature | Notes |
| --- | --- | --- |
| `XmlParser` | `new XmlParser({ onOpen, onClose, onText })` | Incremental SAX; `write(chunk)` / `end(chunk)`, `depth` and `current`. Handles namespace prefixes, self-closing tags, CDATA, numeric references and the `_xHHHH_` escapes. |
| `parseTag` | `(inner) -> { name, local, attrs }` | Splits a tag body without regex backtracking; the no-attribute case shares one frozen map. |
| `encodeXmlText` | `(s) -> string` | Escapes `&`, `<`, `>` and control characters (as `_xHHHH_`). |
| `encodeXmlAttr` | `(s) -> string` | As above plus quotes, newlines and tabs. |
| `decodeEntities` | `(s) -> string` | The five XML entities, the named extras and numeric character references. |

```js doc-example
import { XmlParser, parseTag, decodeEntities, encodeXmlText, encodeXmlAttr } from "gridxcel";
import assert from "node:assert";

const events = [];
const parser = new XmlParser({
  onOpen: (name, attrs) => events.push(["open", name, attrs.r || null]),
  onText: (text) => events.push(["text", text]),
  onClose: (name) => events.push(["close", name]),
});
parser.write('<row r="1"><c t="s"><v>a &amp; b</v></c></row>');
parser.end();

const tag = parseTag('c r="A1" t="s"');
assert.strictEqual(tag.local, "c");
assert.strictEqual(tag.attrs.r, "A1");
assert.strictEqual(decodeEntities("a &amp; b"), "a & b");
assert.strictEqual(encodeXmlText("<&>"), "&lt;&amp;&gt;");
assert.strictEqual(encodeXmlAttr('a"b'), "a&quot;b");
assert.strictEqual(parser.depth, 0);
console.log("events:", events.length, "| first:", JSON.stringify(events[0]), "| text:", JSON.stringify(events[3]));
```

---

## gridxcel/compat

The SheetJS-shaped surface, for migrating an existing codebase by changing one
import specifier. It imports the reader and writer statically, so the
synchronous entry points work immediately with no `await` in front of them - the
whole point of a drop-in replacement - and there is no top-level `await`, so the
module graph stays synchronous and the entry also builds as an IIFE (`<script>`).
The cost of that choice is that importing `gridxcel/compat` always pulls in the
reader and writer. See [MIGRATION.md](./MIGRATION.md).

| Export | Signature | Notes |
| --- | --- | --- |
| `read` | `(data, opts?) -> SheetJSBook` | Synchronous; bytes, base64/binary text, or CSV text. |
| `readAsync` | `(data, opts?) -> Promise<SheetJSBook>` | Same inputs through the async/streaming reader. |
| `readFile` | `(path, opts?) -> SheetJSBook` | Synchronous file read. |
| `readFileAsync` | `(path, opts?) -> Promise<SheetJSBook>` | Path or stream through the async reader. |
| `write` | `(wb, opts?) -> Uint8Array \| string` | xlsx bytes, or csv/txt/json text; shaped by `opts.type`. |
| `writeAsync` | `(wb, opts?) -> Promise<Uint8Array \| string>` | xlsx through the streaming writer; other book types delegate to `write`. |
| `writeFile` | `(wb, path, opts?) -> void` | Overwrites an existing file unless `overwrite: false` is passed, matching SheetJS. |
| `writeFileAsync` | `(wb, path, opts?) -> Promise<void>` | Streaming file write; same overwrite behaviour as `writeFile`. |
| `toWorksheet` | `(sheet) -> object` | gridxcel sheet to a SheetJS worksheet (`!ref`, A1 keys, `!merges`). |
| `toSheetJSBook` | `(book) -> object` | `{ SheetNames, Sheets, Props }`. |
| `attachViews` | `(book) -> SheetJSBook` | Adds `SheetNames`/`Sheets` getters in place. |
| `normalizeBookInput` | `(input) -> Workbook` | Accepts every workbook-ish input SheetJS users pass to `write`. |
| `available` | `() -> object` | `{ read, write, readError, writeError }`. Always `read: true, write: true`: the peers are imported statically, so this is a settled report rather than a probe. Kept because MIGRATION.md and `compat.d.ts` publish it. |
| `preload` | `() -> Promise<object>` | Resolves immediately with the same object as `available()`. Kept as a documented no-op so code written against the earlier design keeps working. |
| `utils` | `object` | SheetJS-shaped namespace (see the comparison in [LIMITATIONS.md](./LIMITATIONS.md)). |
| `XLSX_like` | `object` | Convenience bundle: `read`, `readFile`, `write`, `writeFile`, `utils`, `version`. |

A string, or bytes that are not a zip, is read as CSV/DSV: SheetJS users hand
CSV to `read` constantly, and a hard failure there is the classic migration
trap. Unlike the root entry, the compat writer overwrites an existing file
unless `overwrite: false` is passed, matching SheetJS.

```js doc-example
import { read, readFile, toWorksheet, toSheetJSBook, attachViews, normalizeBookInput, preload, available } from "gridxcel/compat";
import assert from "node:assert";

// No await first: the peers are imported statically, so the synchronous entry
// points are ready as soon as the module is. This is what makes the entry a
// drop-in replacement.
const fromText = read("a,b\n1,2\n");
assert.deepStrictEqual(fromText.SheetNames, ["Sheet1"]);
assert.strictEqual(fromText.Sheets.Sheet1.rows[1][0].v, 1);

const book = readFile("../../test/fixtures/pandas.xlsx");
const worksheet = toWorksheet(book.sheets[0]);
assert.strictEqual(worksheet["!ref"], "A1:C101");

const view = toSheetJSBook(book);
assert.deepStrictEqual(view.SheetNames, ["df"]);
attachViews(book);
assert.deepStrictEqual(normalizeBookInput(book).sheets.map((s) => s.name), ["df"]);
console.log("csv sheet:", fromText.SheetNames.join(","), "| xlsx ref:", worksheet["!ref"], "| peers:", JSON.stringify(available()));
```

```js doc-example
import { readFileAsync, writeAsync, writeFileAsync, XLSX_like, utils, preload } from "gridxcel/compat";
import { existsSync, rmSync } from "node:fs";
import assert from "node:assert";

await preload();
const book = await readFileAsync("../../test/fixtures/pandas.xlsx");
assert.strictEqual(book.SheetNames.length, 1);

const bytes = await writeAsync(utils.aoa_to_sheet([["a", 1]]), { type: "buffer" });
assert.ok(bytes.length > 0);

const out = "../doc-example-compat.xlsx";
await writeFileAsync([["a"]], out, { type: undefined });
assert.ok(existsSync(out));
assert.strictEqual(typeof XLSX_like.version, "string");
console.log("async bytes:", bytes.length, "| written:", existsSync(out), "| version:", XLSX_like.version);
rmSync(out, { force: true });
```

---

## gridxcel/csv

RFC 4180 plus the extensions real files need: a UTF-8 BOM, CRLF or a lone CR,
quoted fields containing separators, quotes or newlines, ragged rows padded to
the widest, and a configurable delimiter. Every function here is synchronous.

| Export | Signature | Notes |
| --- | --- | --- |
| `sheetToCSV` | `(sheet, opts?) -> string` | Serialise; every row is terminated, including the last. |
| `sheetToDSV` | `(sheet, delimiter, opts?) -> string` | The same function with an explicit separator. |
| `createCSVStream` | `(sheet, opts?, streamOpts?) -> AsyncGenerator<string>` | The streaming form. Yields the same bytes while holding one row, so a 200,000 x 10 sheet retains about 0.1 MB instead of 14.3 MB. Chunks break on row boundaries, because a quoted field may legally contain a newline. `streamOpts.chunkRows` defaults to 512. |
| `csvToSheet` | `(text, opts?) -> Sheet` | Parse; the content decides each field's type. |
| `dsvToSheet` | `(text, delimiter, opts?) -> Sheet` | The same function with an explicit separator. |
| `parseDSV` | `(text, opts?) -> object[][]` | Low-level parse into `{ text, quoted }` field records. |
| `toCSV` | `(input, opts?) -> string` | Accepts a book or a sheet; picks a sheet with `opts.sheet`. |
| `fromCSV` | `(text, opts?) -> Workbook` | CSV text to a one-sheet workbook, ready for `writeWorkbook`. |

| Option | Default | Meaning |
| --- | --- | --- |
| `delimiter` | `","` | Field separator. |
| `quote` | `'"'` | Quote character. |
| `lineEnding` / `newline` | `"\n"` | Written line separator. |
| `header` | `false` | Emit a header row from the first row, or use `opts.header` as the labels. |
| `bom` | `false` | Prepend a UTF-8 BOM. |
| `alwaysQuote` | `false` | Quote every field. |
| `escapeFormulae` | `true` | Prefix a leading `=`, `+`, `-`, `@`, tab, CR or LF in a **text** field with an apostrophe. Numbers and dates are untouched, so `-2.5` stays numeric. |
| `quotedStrings` | `false` | Parse quoted fields as literal strings instead of inferring a type; set it when a string that looks like a number must survive a round trip. |
| `dateNF` | `yyyy-mm-dd` | Format used to render `Date` values. |
| `date1904` | `false` | Use the 1904 date system when parsing dates. |
| `skipHidden` | `false` | Omit hidden rows and columns. |
| `range` | whole sheet | Export only this range. |
| `raw` | `false` | Do not apply number formats; keep raw values. |
| `name` | `Sheet1` | Sheet name produced by `csvToSheet`. |
| `blankrows` | `true` | Keep rows that are entirely empty. |
| `style_text` | `false` | Rejected with `NotImplementedError`: styled CSV output does not exist here. |

By default the content decides the type: `123` becomes a number, `TRUE` a
boolean, `2024-01-15` a date. That matches SheetJS and pandas.

```js doc-example
import { createCSVStream, normalizeSheet } from "gridxcel";
import assert from "node:assert";

const sheet = normalizeSheet([
  ["name", "score"],
  ["ada", 42],
  ['says "hi"', 7],
]);

// Streaming and string forms agree byte for byte; the stream just never holds
// the whole output.
let streamed = "";
let chunks = 0;
for await (const chunk of createCSVStream(sheet, {}, { chunkRows: 1 })) {
  streamed += chunk;
  chunks++;
}
assert.strictEqual(streamed.split("\n").filter(Boolean).length, 3);
assert.ok(chunks >= 1, "produced at least one chunk");
console.log("streamed", JSON.stringify(streamed.split("\n")[0]), "in", chunks, "chunks");
```

```js doc-example
import { read } from "gridxcel";
import { sheetToCSV, csvToSheet } from "gridxcel/csv";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx", { sheets: ["Types"] });
const csv = sheetToCSV(book.sheets[0], { header: true });
assert.ok(csv.split("\n")[0].startsWith("text,"));

const back = csvToSheet("a,b\n1,hello\n", {});
assert.strictEqual(back.rows[1][0].v, 1);
console.log(csv.split("\n")[0]);
console.log("parsed rows:", back.rowCount, "| A2:", JSON.stringify(back.rows[1][0].v));
```

```js doc-example
import { csvToSheet, parseDSV, dsvToSheet, sheetToDSV, toCSV, fromCSV } from "gridxcel/csv";
import assert from "node:assert";

const records = parseDSV('a;"b;c"\n1;2\n', { delimiter: ";" });
assert.strictEqual(records[0][1].text, "b;c");
assert.strictEqual(records[0][1].quoted, true);

const sheet = dsvToSheet("a;b\n1;2\n", ";");
assert.strictEqual(sheetToDSV(sheet, ";"), "a;b\n1;2\n");
assert.strictEqual(toCSV(fromCSV("x,y\n"), {}), "x,y\n");
console.log("records:", records.length, "| dsv:", JSON.stringify(sheetToDSV(sheet, ";")));
```

The injected-formula guard applies to text fields only:

```js doc-example
import { csvToSheet } from "gridxcel/csv";
import assert from "node:assert";

const sheet = csvToSheet("name,note\n=cmd,ok\n-2.5,plain\n", {});
assert.strictEqual(sheet.rows[1][0].t, "string", "a formula-looking text field stays text");
assert.strictEqual(sheet.rows[2][0].t, "number", "-2.5 stays numeric");
console.log("guarded text:", JSON.stringify(sheet.rows[1][0].v));
console.log("numeric kept:", JSON.stringify(sheet.rows[2][0].v), sheet.rows[2][0].t);
```

---

## gridxcel/json

Sheet to JSON/AOA and back, plus the cell-level helpers the converters are built
from. Everything here is synchronous; the root entry wraps the converters in
async functions of the same name.

| Export | Signature | Notes |
| --- | --- | --- |
| `sheetToJSON` | `(sheet, opts?) -> object[]` | SheetJS header semantics; dense `header: 1`. |
| `jsonToSheet` | `(rows, opts?) -> Sheet` | Objects to sheet, keys in first-seen order. |
| `sheetToAOA` | `(sheet, opts?) -> any[][]` | Dense grid. |
| `aoaToSheet` | `(rows, opts?) -> Sheet` | Array of arrays to sheet, with `cell.a` filled in. |
| `normalizeSheet` | `(input, opts?) -> Sheet` | Also accepts SheetJS worksheets and books. |
| `worksheetToSheet` | `(ws) -> Sheet` | SheetJS worksheet to a gridxcel sheet. |
| `blankSheet` | `(name?) -> Sheet` | Empty sheet with every documented field present. |
| `dateCell` | `(date, opts?) -> Cell` | Date cell with `serial` and `z` set. |
| `cellFromValue` | `(v, opts?) -> Cell \| undefined` | Plain value to cell; `NaN` becomes `#NUM!`. |
| `cellValue` | `(cell, opts?) -> unknown` | The value a consumer sees (`cell.w` wins unless `rawNumbers`). |
| `cellText` | `(cell, opts?) -> string` | Display text used by CSV output. |
| `cellAt` | `(sheet, r, c) -> Cell \| undefined` | Absolute, 0-based row and column. |
| `rowAt` | `(sheet, r) -> Cell[] \| undefined` | 0-based row index. |
| `sheetRange` | `(sheet) -> { s, e, empty }` | Used range derived from cells that hold content. |
| `resolveRange` | `(sheet, opts?) -> { s, e, empty }` | Resolves `opts.range` against the used range. |
| `isRowHidden` | `(sheet, r) -> boolean` | 0-based row index, stored key 1-based. |
| `isColHidden` | `(sheet, c) -> boolean` | 0-based column index. |
| `assertDenseRange` | `(ws, range, opts?) -> number` | Refuses a dense conversion whose span exceeds `maxRows`. |
| `DEFAULT_DATE_NF` | `string` | `"yyyy-mm-dd"`, the format used when nothing else applies. |

```js doc-example
import { blankSheet, dateCell, cellFromValue, cellValue, cellText, cellAt, rowAt, sheetRange, resolveRange, isRowHidden, isColHidden, worksheetToSheet, DEFAULT_DATE_NF } from "gridxcel/json";
import assert from "node:assert";

const sheet = blankSheet("S");
sheet.rows = [[{ v: "a", t: "string" }, { v: 1, t: "number" }], [undefined, { v: 2, t: "number" }]];
sheet.columnCount = 2;
sheet.rowCount = 2;
sheet.rowInfo = new Map([[2, { index: 2, hidden: true }]]);
sheet.columns = [{ index: 1, min: 1, max: 1, hidden: true }];

assert.strictEqual(cellValue(cellAt(sheet, 0, 0)), "a");
assert.strictEqual(rowAt(sheet, 0).length, 2);
assert.strictEqual(isRowHidden(sheet, 1), true);
assert.strictEqual(isColHidden(sheet, 1), true);
assert.strictEqual(sheetRange(sheet).e.r, 1);
assert.strictEqual(resolveRange(sheet, { range: "A1:B1" }).e.r, 0);
assert.strictEqual(dateCell(new Date(2024, 2, 15)).z, DEFAULT_DATE_NF);
assert.strictEqual(cellFromValue(NaN).e, "#NUM!");
assert.strictEqual(cellText(cellFromValue(2.5)), "2.5");

const back = worksheetToSheet({ "!ref": "A1:B2", A1: { v: "x", t: "s" }, B2: { v: 7, t: "n" } });
assert.strictEqual(cellValue(cellAt(back, 1, 1)), 7);
console.log("used range:", JSON.stringify(sheetRange(sheet)), "| default date format:", DEFAULT_DATE_NF);
```

#### The dense-range guard

A worksheet is untrusted input: one `<row r="40000000">` makes the reader
report a forty-million-row extent, and densifying that range would allocate
gigabytes from a one-kilobyte file. `assertDenseRange` caps the span at
262,144 rows by default and throws `RANGE_TOO_LARGE` with the span and limit in
`error.details`. Callers who pass an explicit `range` are taken at their
word; callers who inherited the sheet's own extent are not.

```js doc-example
import { assertDenseRange, blankSheet } from "gridxcel/json";
import { GridxcelError } from "gridxcel";
import assert from "node:assert";

const untrusted = { ...blankSheet("Untrusted"), rowCount: 5_000_000 };
try {
  assertDenseRange(untrusted, { s: { r: 0, c: 0 }, e: { r: 4_999_999, c: 9 } }, {});
  assert.fail("should have refused");
} catch (e) {
  assert.ok(e instanceof GridxcelError);
  assert.strictEqual(e.code, "RANGE_TOO_LARGE");
  console.log("refused:", e.code, "| span:", e.details.span, "| limit:", e.details.limit);
}

const bounded = assertDenseRange(blankSheet(), { s: { r: 0, c: 0 }, e: { r: 99, c: 9 } }, {});
console.log("bounded span accepted:", bounded);
```

---

## gridxcel/numfmt

| Export | Signature | Notes |
| --- | --- | --- |
| `formatNumber` | `(value, code?, opts?) -> string` | Numbers through a format code. |
| `formatValue` | `(value, code?, opts?) -> string` | Any value; dispatches on type. |
| `formatDateTime` | `(serial, code, opts?) -> string` | Date/time serials; the only place the month-versus-minute rule lives. |
| `isDateFormat` | `(codeOrId) -> boolean` | Cached per code. |
| `builtinFormat` | `(id) -> string \| undefined` | Code for a built-in id, `undefined` for custom ids. |
| `BUILTIN_FORMATS` | `Record<number, string>` | The full ECMA-376 built-in table. |
| `generalNumber` | `(v) -> string` | `General` rendering. |
| `serialToDate` | `(serial, date1904?) -> Date` | Local wall-clock `Date`. |
| `dateToSerial` | `(date, date1904?) -> number` | Local fields to serial. |
| `serialToUTCms` | `(serial, date1904?) -> number` | Serial to UTC milliseconds, without allocating a `Date`. |
| `serialParts` | `(serial) -> object` | `{ y, m, d, dow, h, mi, s, ms, serial }`. |
| `splitSections` | `(code) -> string[]` | Split on top-level semicolons, respecting quotes, brackets and escapes. |
| `pickSection` | `(sections, value, isText) -> number` | Which of the up-to-four sections applies. |

```js doc-example
import { builtinFormat, BUILTIN_FORMATS, formatDateTime, formatValue, generalNumber, serialParts, serialToUTCms, splitSections, pickSection, isDateFormat } from "gridxcel/numfmt";
import assert from "node:assert";

const sections = splitSections('0.00;(0.00);"-";@');
assert.strictEqual(sections.length, 4);
assert.strictEqual(pickSection(sections, -1, false), 1);
assert.strictEqual(pickSection(sections, 0, true), 3);
assert.strictEqual(formatDateTime(45366.6046875, "yyyy-mm-dd hh:mm:ss"), "2024-03-15 14:30:45");
assert.strictEqual(serialParts(45366.6046875).h, 14);
assert.strictEqual(serialToUTCms(45366.6046875), 1710513045000);
assert.strictEqual(generalNumber(1 / 3), "0.33333333333");
assert.strictEqual(builtinFormat(14), "m/d/yy");
assert.strictEqual(builtinFormat(999), undefined);
assert.strictEqual(BUILTIN_FORMATS[49], "@");
assert.strictEqual(formatValue(true, undefined), "TRUE");
assert.strictEqual(isDateFormat(14), true);
console.log("sections:", sections.length, "| dt:", formatDateTime(45366.6046875, "yyyy-mm-dd hh:mm:ss"), "| id 14:", builtinFormat(14));
```

---

## gridxcel/xml

| Export | Signature | Notes |
| --- | --- | --- |
| `XmlParser` | `class` | Incremental parser; `write`/`end`, `depth`, `current`. |
| `parseChunks` | `(chunks, handlers) -> Promise<XmlParser>` | Drives the parser over an async iterable of strings or bytes. |
| `parseTag` | `(inner) -> { name, local, attrs }` | Tag body to name and attribute map. |
| `scanTagEnd` | `(s, from, quote) -> { gt, quote, scanned, done }` | Resumable scan for the `>` that ends a tag. Returns `gt` `-1` when the tag is unfinished, plus the quote state and offset to resume from, so a tag split across chunks is scanned once rather than from the start each time. The worksheet reader uses it; it is exported because a caller writing its own scanner needs the same resumable behaviour to stay linear on a large attribute value. |
| `localName` | `(name) -> string` | Strips a namespace prefix. |
| `tagLocalIfNoAttrs` | `(body) -> string \| null` | One-pass local name for a tag with no attributes, or `null` when it has some. A worksheet is mostly `<v>`, `<t>` and `</c>`, and this answers the only question the scanner has for those without allocating a result or an attribute map. Use `parseTag` for the `null` case. |
| `NO_ATTRS` | `object` | The shared frozen, null-prototype empty attribute map. Handed to callbacks so an attribute-free tag does not allocate one; never mutate it. |
| `decodeEntities` | `(s) -> string` | Entities and numeric character references. |
| `decodeOoxmlEscapes` | `(s) -> string` | The `_xHHHH_` escapes Excel uses for control characters. |
| `encodeXmlText` | `(s) -> string` | Escape a text node. |
| `encodeXmlAttr` | `(s) -> string` | Escape an attribute value. |

```js doc-example
import { XmlParser, parseChunks, parseTag, scanTagEnd, localName, tagLocalIfNoAttrs, NO_ATTRS, decodeEntities, decodeOoxmlEscapes, encodeXmlText, encodeXmlAttr } from "gridxcel/xml";
import assert from "node:assert";

assert.strictEqual(localName("x:row"), "row");

// tagLocalIfNoAttrs answers the scanner's only question for an attribute-free
// tag, in one pass and without allocating. null means "has attributes, use parseTag".
assert.strictEqual(tagLocalIfNoAttrs("v"), "v");
assert.strictEqual(tagLocalIfNoAttrs("x:c"), "c");
assert.strictEqual(tagLocalIfNoAttrs('c r="A1"'), null);
assert.strictEqual(parseTag('c r="A1"').local, "c");
assert.strictEqual(Object.keys(NO_ATTRS).length, 0, "the shared empty map is empty");
assert.strictEqual(decodeEntities("a &amp; b"), "a & b");
assert.strictEqual(decodeOoxmlEscapes("_x000A_"), "\n");
assert.strictEqual(encodeXmlText("<&>"), "&lt;&amp;&gt;");
assert.strictEqual(encodeXmlAttr('a"b'), "a&quot;b");
assert.strictEqual(parseTag("<ignored>").name, "<ignored>");

// scanTagEnd is resumable: scan each piece, carry the quote state across the
// boundary, and join only once the tag is complete. Scanning the whole
// accumulated tag again on every chunk is what made it quadratic.
const pieces = ['c r="A', '1" t="s"', ">"];
const collected = [];
let quote = 0;
let body = null;
for (const piece of pieces) {
  const scan = scanTagEnd(piece, 0, quote);
  quote = scan.quote;
  if (scan.done) { body = collected.join("") + piece.slice(0, scan.gt); break; }
  collected.push(piece);
}
assert.strictEqual(body, 'c r="A1" t="s"', "the tag reassembles across the split");
assert.strictEqual(parseTag(body).attrs.r, "A1");

const seen = [];
const parser = await parseChunks(["<a>", "1</a>"], { onText: (t) => seen.push(t) });
assert.deepStrictEqual(seen, ["1"]);
assert.strictEqual(parser.depth, 0);
assert.ok(new XmlParser({}) instanceof XmlParser);
console.log("localName:", localName("x:row"), "| text:", JSON.stringify(seen), "| depth:", parser.depth);
```

`parseChunks` accepts strings or byte chunks and decodes them as UTF-8 with a
streaming decoder, so a multi-byte character split across two chunks is
reassembled rather than mangled.

---

## gridxcel/zip

| Export | Signature | Notes |
| --- | --- | --- |
| `ZipArchive` | `class` | Read a container; CRC-verified by default. |
| `ZipEntry` | `class` | One entry: metadata, `bytes()`, `text()`, `stream()`. |
| `ZipWriter` | `class` | Build a container; `add()` then `build()`. |
| `createZip` | `(files, opts?) -> Uint8Array` | Convenience builder. |
| `ZipError` | `class extends BadZipError` | Container failure with a code and an `offset`. |
| `METHOD_STORE` | `0` | Stored, no compression. |
| `METHOD_DEFLATE` | `8` | Raw deflate. |
| `SIG_LOCAL` | `0x04034b50` | Local file header. |
| `SIG_CENTRAL` | `0x02014b50` | Central directory header. |
| `SIG_EOCD` | `0x06054b50` | End of central directory. |
| `SIG_EOCD64` | `0x06064b50` | Zip64 end of central directory. |
| `SIG_EOCD64_LOC` | `0x07064b50` | Zip64 locator. |
| `SIG_DATA_DESC` | `0x08074b50` | Data descriptor. |

The signature constants are exported for callers who parse containers
themselves; they are not needed to read or write an archive. Sizes and CRCs are
computed before each local header is emitted, so no data descriptors and no
seekable sink are required - the same code path drives a buffer and a live
stream. Zip64 records are emitted automatically once an entry or an offset
exceeds the 32-bit limits.

```js doc-example
import { ZipWriter, ZipArchive, ZipEntry, createZip, METHOD_STORE, METHOD_DEFLATE, SIG_LOCAL, SIG_CENTRAL, SIG_EOCD } from "gridxcel/zip";
import assert from "node:assert";

const bytes = createZip({ "a.txt": "alpha", "b.bin": new Uint8Array([1, 2, 3]) });
const archive = new ZipArchive(bytes);
assert.ok(archive.has("a.txt"));
assert.strictEqual(archive.readText("a.txt"), "alpha");
assert.ok(archive.get("a.txt") instanceof ZipEntry);
assert.strictEqual(SIG_LOCAL, 0x04034b50);
assert.strictEqual(SIG_CENTRAL, 0x02014b50);
assert.strictEqual(SIG_EOCD, 0x06054b50);
assert.deepStrictEqual([...archive].map((e) => e.name), ["a.txt", "b.bin"]);

const writer = new ZipWriter({ deterministic: true });
writer.add("stored.bin", new Uint8Array([1, 2, 3]), { compress: false });
const stored = new ZipArchive(writer.build()).get("stored.bin");
assert.strictEqual(stored.method, METHOD_STORE);
assert.notStrictEqual(METHOD_STORE, METHOD_DEFLATE);
console.log("entries:", archive.names().join(","), "| stored:", stored.method, "| deflate:", METHOD_DEFLATE);
```

---

## Errors

Every failure derives from `GridxcelError`, so a single `instanceof` check
separates library problems from your own bugs, and `error.code` is stable
enough to branch on.

| Class | Codes | Raised for |
| --- | --- | --- |
| `GridxcelError` | many | Base class for everything below; carries `details`. |
| `BadZipError` | `ZIP_ERROR` | Base class for container failures. |
| `ZipError` | `NOT_ZIP`, `BAD_LOCAL_HEADER`, `UNSUPPORTED_METHOD`, `CRC_MISMATCH`, `INFLATE_FAILED` | Zip container unreadable, truncated or corrupt; carries `offset`. `INFLATE_FAILED` means a deflate stream failed to decode - `details.zlibCode` distinguishes corruption (`Z_DATA_ERROR`) from truncation (`Z_BUF_ERROR`) and the original error is on `cause`. |
| `FormatError` | `NOT_A_ZIP`, `NOT_A_WORKBOOK` | The bytes are not a workbook at all. |
| `SchemaError` | `BAD_SCHEMA` | Well-formed XML that violates the OOXML schema. |
| `StrictWarningError` | the warning's code | `strict: true` turned a warning into an exception; carries `warning`. |
| `NotImplementedError` | `NOT_IMPLEMENTED` | A format or option this build does not support. |
| `PlatformError` | `NOT_SUPPORTED` | A capability this platform does not have: the synchronous API (`readSync`, `writeWorkbookSync`, `ZipArchive.bytes()`) and the path helpers (`readFile`, `writeFile`) outside Node. Raised by browsers, workers and any non-Node host, naming the async alternative. |

Reader and writer codes a caller is likely to branch on: `NO_SHEETS`,
`FILE_NOT_FOUND`, `FILE_ACCESS_DENIED`, `FILE_IS_DIRECTORY`, `FILE_EXISTS`,
`MODULE_UNAVAILABLE`, `BAD_ADDRESS`, `BAD_RANGE`, `BAD_CELL`, `BAD_SHEET`,
`BAD_BOOK`, `BAD_ROW`, `ROW_ORDER`, `SHEET_CLOSED`, `WRITER_CLOSED`,
`BAD_VALUE`, `BAD_STYLE`, `RANGE_TOO_LARGE` and `NOT_SUPPORTED`.

Outside Node, the synchronous entry points (`readSync`, `writeWorkbookSync`,
`ZipArchive.bytes()`) and the path helpers (`readFile`, `writeFile`) throw
`PlatformError` with code `NOT_SUPPORTED`, naming the async alternative, rather
than a `ReferenceError`. `PlatformError` is a `GridxcelError` subclass exported
from `gridxcel` alongside its six siblings, so it can be caught by class or by
branching on `error.code === "NOT_SUPPORTED"`; Node and the browser share the
one class.

```js doc-example
import { read, GridxcelError, FormatError } from "gridxcel";
import assert from "node:assert";

try {
  await read(new Uint8Array(23));
  assert.fail("should have thrown");
} catch (e) {
  assert.ok(e instanceof GridxcelError);
  assert.strictEqual(e.code, "NOT_A_ZIP");
  console.log("caught by base class:", e instanceof GridxcelError);
  console.log("specific class:", e instanceof FormatError, "| code:", e.code);
}
```

The distinction that matters in production: **bytes that are not a workbook
throw; a workbook that is partly damaged warns.** A truncated worksheet still
returns the rows it could read, alongside a warning describing the loss.

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/reader-cut-sheet.xlsx");
assert.ok(book.warnings.some((w) => w.code === "BAD_XML"));
console.log("sheets recovered:", book.sheets.length);
console.log("warning codes:", book.warnings.map((w) => w.code).join(","));
console.log("rows in sheet 1:", book.sheets[0].rows.length);
```

---

## Guarantees and non-guarantees

What the library promises at 0.1.0:

* **Typed errors.** Nothing is thrown as a bare string or a plain `Error`. The
  fuzz suite drives 456 corrupted inputs and observes 402 typed
  `GridxcelError` instances, 0 plain errors and 0 hangs.
* **Local wall-clock dates.** `dateToSerial` and `serialToDate` round-trip
  the spreadsheet's wall clock in any timezone; the suite is verified under
  three zones. Never compare a parsed date with `toISOString()`.
* **CRC-32 verification on every part.** Corruption throws `CRC_MISMATCH`
  instead of yielding wrong cell values; the cost is about 1.2 GB/s of payload,
  negligible against inflate.
* **Bounded dense conversion.** `sheetToAOA`, `sheetToJSON` and
  `sheetToCSV` refuse a span above `maxRows` (default 262,144) with
  `RANGE_TOO_LARGE` rather than allocating from an untrusted extent.
* **Bounded extents.** A row more than `EMPTY_ROW_SLACK` (65,536) rows beyond
  its neighbours is kept in `sheet.sparseRows` behind a `SPARSE_ROW_GAP`
  warning instead of forcing a dense allocation; the writer emits it again, so a
  rewrite is lossless.
* **Deterministic writes.** `deterministic: true` produces byte-identical
  output across runs and processes.
* **Round trips.** Read, write and read again preserves values, types,
  formulas, number formats, merges and sheet names; numbers survive IEEE-754
  exactly.
* **Zero dependencies, no build.** the dependencies and devDependencies fields are
  both empty, src/*.js is what ships, and there is no dist directory.
* **The byte-level API runs in a browser.** `read`, `eachRow`, `sheetToJSON`,
  `sheetToCSV`, `write` and `WorkbookWriter` need no build step and no polyfill:
  `src/platform.browser.js` replaces the Node platform module through the
  package `browser` field, and the Web compression streams supply `deflate-raw`.
  The synchronous API and the filesystem helpers are Node-only and fail with a
  typed error. See [BROWSER.md](./BROWSER.md).
* **Speed.** On the published 50,000 x 8 fixture the range across repeated runs
  is 1.4x to 2.7x faster than SheetJS CE 0.20.3 for read, async read and
  streaming read; see [PERFORMANCE.md](./PERFORMANCE.md) for method and numbers.

What it does **not** promise:

* **No synchronous API or filesystem outside Node.** A browser has no
  synchronous inflate, so `readSync`, `writeWorkbookSync` and
  `ZipArchive.bytes()` throw `NOT_SUPPORTED` there, and a browser has no
  filesystem, so `readFile` and `writeFile` are Node-only. The Web
  compression streams must accept `deflate-raw` (Chrome 103+, Firefox 113+,
  Safari 16.4+); older engines get the typed error rather than wrong bytes. A
  browser write stores its parts instead of deflating them, so the workbook is
  valid but larger.
* **No other spreadsheet formats.** XLSB, XLS (BIFF), SpreadsheetML 2003,
  ODS/FODS, NUMBERS, DBF, SYLK, DIF, PRN, Lotus/Quattro and HTML/RTF tables are
  out of scope. No charts, images, pivot tables, VBA/macros, conditional
  formatting or data validation. Encrypted workbooks can be partially read but
  not written. There is no incremental CSV writer.
* **Byte-identical SheetJS output.** The compatibility layer matches SheetJS
  semantics, not bytes: `sheetToJSON({ header: 1 })` emits dense `null`
  instead of array holes, the root returns Promises where `gridxcel/compat`
  returns values, and `book_new()` returns the native workbook from the root
  and the SheetJS-shaped book from `compat`.
* **A frozen API.** 0.1.0 is pre-1.0 and the public surface may still change.

---

## gridxcel/formats

Text-based spreadsheet formats beyond XLSX and CSV: **SYLK** and **DIF**, the two
legacy interchange formats Excel still reads and writes.

| Export | Signature | Notes |
| --- | --- | --- |
| `sylkToSheet` | `(text, opts?) -> Sheet` | SYLK to a Sheet. An omitted `Y` or `X` reuses the previous coordinate, as real files rely on. |
| `sheetToSYLK` | `(sheet, opts?) -> string` | Sheet to SYLK. A date is written as its serial: SYLK has no date type. |
| `difToSheet` | `(text, opts?) -> Sheet` | DIF to a Sheet, per the 1981 specification. |
| `sheetToDIF` | `(sheet, opts?) -> string` | Sheet to DIF. A date is written as its serial: DIF has no date type. |

`opts.name` sets the sheet name; `opts.eol` sets the line ending (CRLF by
default, which is what Excel writes).

You rarely need this entry point directly: `gridxcel/compat` sniffs the format
from the content, so `read(sylkText)` works without saying which format it is.
Import these when you already know and want to avoid the guess.

```js doc-example
import { difToSheet, sheetToDIF, sheetToSYLK, sylkToSheet } from "gridxcel/formats";
import { aoaToSheet } from "gridxcel/json";
import assert from "node:assert";

const sheet = aoaToSheet([["name", "score"], ["ada", 42]]);

// SYLK: a record-based format, one "C;Y..;X..;K.." per cell.
const sylk = sheetToSYLK(sheet);
assert.match(sylk, /^ID;P/);
assert.deepStrictEqual(sylkToSheet(sylk).rows[1].map((c) => c.v), ["ada", 42]);

// DIF: a header section, then two-line data chunks.
const dif = sheetToDIF(sheet);
assert.match(dif, /^TABLE/);
assert.deepStrictEqual(difToSheet(dif).rows[1].map((c) => c.v), ["ada", 42]);

console.log("sylk bytes:", sylk.length, "| dif bytes:", dif.length);
```


---

## TypeScript

Hand-written declarations ship for all seven entry points:
`src/index.d.ts`, `compat.d.ts`, `csv.d.ts`, `json.d.ts`, `numfmt.d.ts`,
`xml.d.ts` and `zip.d.ts`, plus the shared contracts in `src/types.d.ts`.

The declarations are hand written so cell narrowing works without casts:

```ts
import { read, type Workbook, type Sheet, type Cell } from "gridxcel";

const book: Workbook = await read("report.xlsx", { cellDates: true });
const sheet: Sheet = book.sheets[0];
const cell: Cell | undefined = sheet.rows[0][0];
if (cell?.t === "date") {
  // cell.v is string | number | boolean | Date | null
}
```

---

See also: [MIGRATION.md](./MIGRATION.md) for the SheetJS mapping,
[LIMITATIONS.md](./LIMITATIONS.md) for the supported surface and
[PERFORMANCE.md](./PERFORMANCE.md) for measurements.