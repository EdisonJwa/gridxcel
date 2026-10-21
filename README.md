# gridxcel

**Streaming XLSX, CSV, SYLK and DIF for Node.js and the browser.**
Zero dependencies. Bounded memory. No build step.

[![npm](https://img.shields.io/npm/v/gridxcel.svg)](https://www.npmjs.com/package/gridxcel)
[![CI](https://github.com/EdisonJwa/gridxcel/actions/workflows/ci.yml/badge.svg)](https://github.com/EdisonJwa/gridxcel/actions/workflows/ci.yml)
[![licence](https://img.shields.io/badge/licence-MIT-blue.svg)](LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org)
[![dependencies](https://img.shields.io/badge/dependencies-0-brightgreen.svg)](package.json)

```js
import { read, sheetToJSON } from "gridxcel";

const book = await read("report.xlsx");
const rows = sheetToJSON(book.sheets[0]);
console.log(rows.length, "rows");
```

---

## Installation

```sh
npm install gridxcel
pnpm add gridxcel
yarn add gridxcel
```

There is nothing to build, nothing to configure, and no peer dependency. The
package is ESM-only: `import` works everywhere, and CommonJS can use
`await import("gridxcel")`.

**Supported runtimes**

| Runtime | Status |
| --- | --- |
| Node.js 18, 20, 22, 24 | Tested on Linux and Windows in CI |
| Browsers | Native ESM, no polyfill, no bundler required |
| Deno, Bun, Cloudflare Workers | Web-platform API only; the Node helpers are absent |

## Quick start

```js doc-example
import { read, write, eachRow, sheetToJSON } from "gridxcel";
import assert from "node:assert";

// Read from a path, bytes, a URL, a stream or an async iterable.
const book = await read("../../test/fixtures/golden.xlsx", { cellDates: true });
assert.deepStrictEqual(book.sheets.map((s) => s.name), ["Types", "Layout", "Volume", "Empty"]);

// A sheet becomes objects keyed by the header row, or arrays of arrays.
// The root entry point is promise-based on purpose, so both are awaited.
const asObjects = await sheetToJSON(book.sheets[1]);
const asArrays = await sheetToJSON(book.sheets[2], { header: 1 });
assert.ok(asObjects.length > 0);
assert.strictEqual(asArrays.length, 200);

// Write from arrays, objects or a full workbook.
const bytes = await write([["name", "score"], ["ada", 42]], { bookType: "xlsx" });
assert.ok(bytes.length > 0);

// Stream a sheet that does not fit in memory: one row at a time.
let streamed = 0;
for await (const { sheet, r, cells } of eachRow("../../test/fixtures/golden.xlsx")) streamed++;
assert.strictEqual(streamed, 205);

console.log("sheets:", book.sheetNames.join(", "));
console.log("rows:", asArrays.length, "| streamed:", streamed, "| wrote", bytes.length, "bytes");
```

Every block marked `doc-example` in this repository is **executed** by
`npm run check:docs`, so no example here can drift from the code.

## Why this library

Three properties decide the choice. Each one is measured, and each measurement
has a command that reproduces it.

### 1. Bounded memory at every stage

Reading a sheet and writing it out are both streamable. A 200,000 x 10 workbook
costs **5.9 MiB retained** to read as a stream against **164.9 MiB** for a full
parse, at 88% of the throughput:

```js doc-example
import { eachRow, read, write } from "gridxcel";
import assert from "node:assert";

// A workbook built in memory, so this example needs no fixture on disk.
const rowsData = Array.from({ length: 20000 }, (_, r) => [r, "row " + r, r * 1.5]);
const bytes = await write(rowsData);

// Full parse: the whole workbook is a value you can inspect.
const book = await read(bytes);
assert.strictEqual(book.sheets[0].rowCount, 20000);
console.log("materialised:", book.sheets[0].rowCount, "rows in memory");

// Streaming: one row exists at a time. Works on files larger than the heap.
let rows = 0;
for await (const row of eachRow(bytes)) rows++;
assert.strictEqual(rows, 20000);
console.log("streamed:", rows, "rows, one at a time");
```

### 2. It is fast

On a 50,000 x 8 workbook written by openpyxl, with both libraries reading the
same 400,008 cells in the same process, across nine passes:

| Operation | Median | Versus SheetJS CE 0.20.3 |
| --- | --- | --- |
| SheetJS `read` | 1,072 ms | - |
| gridxcel `readSync` | 315 ms | **3.20x**, range 2.99x - 3.32x |
| gridxcel `read` | 281 ms | **3.50x**, range 3.40x - 3.75x |
| gridxcel `eachRow` | 286 ms | **3.41x**, range 3.31x - 3.79x |

Reproduce with `npm run bench:scaling`. The **range** is the claim, not the best
sample: a single pass on a shared machine is a sample, not a result, and this
project publishes that distinction rather than hiding it. Full method in
[PERFORMANCE.md](docs/PERFORMANCE.md).

### 3. Failure is loud and typed

Bytes that are not a workbook throw a typed error instead of returning a
plausible empty sheet. Every part is CRC-checked on read. Every failure carries a
stable `code` you can branch on.

```js doc-example
import { read, readSync, GridxcelError, FormatError } from "gridxcel";
import assert from "node:assert";

// 8 bytes that are not a workbook. SheetJS 0.20.3 returns an empty Sheet1;
// this rejects, and the error is catchable by class or by code.
await assert.rejects(read(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])), (error) => {
  assert.ok(error instanceof GridxcelError, "one base class");
  assert.ok(error instanceof FormatError, "a specific subclass");
  assert.strictEqual(error.code, "NOT_A_ZIP", "and a stable code");
  return true;
});

// The synchronous twin throws rather than rejecting.
assert.throws(() => readSync(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])), (error) => error.code === "NOT_A_ZIP");

console.log("a non-workbook input is a typed error, not an empty sheet");
```

## Examples

### Read a workbook

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx");

assert.ok(Array.isArray(book.sheets));
assert.strictEqual(book.sheets.length, 4);
assert.strictEqual(book.dateSystem, 1900);

const sheet = book.sheets[0];
console.log("name:", sheet.name);
console.log("extent:", sheet.rowCount, "rows x", sheet.columnCount, "columns");

// A cell is plain data, not a class instance.
const cell = sheet.rows[1][0];
console.log("first data cell:", JSON.stringify(cell));
```

### Read only the sheets you need

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

// Parsing a 40-sheet workbook to look at one sheet is wasted work.
const book = await read("../../test/fixtures/golden.xlsx", { sheets: ["Layout"] });

assert.deepStrictEqual(book.sheets.map((s) => s.name), ["Layout"]);
// The declared names are still reported, so the file shape is not hidden.
assert.strictEqual(book.sheetNames.length, 4);
console.log("parsed:", book.sheets.map((s) => s.name).join(", "), "of", book.sheetNames.join(", "));
```

### Stream rows and stop early

```js doc-example
import { eachRow } from "gridxcel";
import assert from "node:assert";

// The generator is lazy: breaking out stops the parse and releases the buffers.
const firstThree = [];
for await (const { r, cells } of eachRow("../../test/fixtures/golden.xlsx")) {
  firstThree.push({ r, values: cells.map((c) => (c ? c.v : null)) });
  if (firstThree.length === 3) break;
}

assert.strictEqual(firstThree.length, 3);
console.log(JSON.stringify(firstThree, null, 2));
```

### Write a workbook

```js doc-example
import { write, read } from "gridxcel";
import assert from "node:assert";

const bytes = await write(
  {
    sheets: [
      {
        name: "Report",
        rows: [
          [{ v: "Region", t: "string" }, { v: "Total", t: "string" }],
          [{ v: "North", t: "string" }, { v: 1250.5, t: "number", z: "#,##0.00" }],
          [{ v: "South", t: "string" }, { v: 980, t: "number", z: "#,##0.00" }],
        ],
      },
    ],
  },
  { bookType: "xlsx" },
);

assert.ok(bytes instanceof Uint8Array);

// What you wrote is what you read back.
const back = await read(bytes);
assert.strictEqual(back.sheets[0].rows[1][1].v, 1250.5);
console.log("wrote and re-read", bytes.length, "bytes");
```

### Write a file larger than memory

```js doc-example
import { WorkbookWriter } from "gridxcel";
import assert from "node:assert";

const writer = new WorkbookWriter();
const sheet = writer.addSheet("Rows", { columns: [{ width: 8 }, { width: 24 }] });

for (let r = 1; r <= 20000; r++) {
  // undefined means "the next row". Row numbers must ascend.
  sheet.addRow(undefined, [r, "row " + r, r * 1.5]);
}
sheet.end();

const bytes = writer.finish();
assert.ok(bytes.length > 0);
console.log("wrote", sheet.rowCount, "rows without holding them:", bytes.length, "bytes");
```

### Convert CSV out without building one huge string

```js doc-example
import { read, createCSVStream } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx");

// sheetToCSV returns one string: O(output) in memory.
// createCSVStream yields the same bytes a row at a time.
let csvBytes = 0;
let chunks = 0;
for await (const chunk of createCSVStream(book.sheets[2], {}, { chunkRows: 64 })) {
  csvBytes += chunk.length; // write it to a file, a socket, a Response
  chunks++;
}

assert.ok(csvBytes > 0);
console.log("streamed", csvBytes, "CSV bytes in", chunks, "chunks");
```

### Convert to JSON

```js doc-example
import { read, sheetToJSON, sheetToAOA, aoaToSheet } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/golden.xlsx");

// Header row drives the keys.
const objects = await sheetToJSON(book.sheets[1]);
assert.ok(Array.isArray(objects));

// Or keep it as arrays, which never merges a blank header. Also awaited: every
// converter on the root entry point returns a Promise, so a synchronous and an
// asynchronous call site cannot be confused for one another.
const arrays = await sheetToAOA(book.sheets[1]);
assert.ok(arrays.length > 0);

// And back again, from either shape. aoaToSheet is synchronous everywhere.
const rebuilt = aoaToSheet(arrays);
assert.ok(rebuilt.rows.length > 0);

console.log("objects:", objects.length, "| arrays:", arrays.length);
```

### Format numbers the way a spreadsheet does

```js doc-example
import { formatNumber } from "gridxcel";
import assert from "node:assert";

// The ECMA-376 format engine, on its own. Useful for showing a sheet cell as
// text without opening a spreadsheet.
assert.strictEqual(formatNumber(1234.5, "#,##0.00"), "1,234.50");
assert.strictEqual(formatNumber(0.25, "0.0%"), "25.0%");
assert.strictEqual(formatNumber(-1234.5, "#,##0.00;[Red](#,##0.00)"), "(1,234.50)");

console.log("1234.5 as #,##0.00  ->", formatNumber(1234.5, "#,##0.00"));
console.log("0.25   as 0.0%      ->", formatNumber(0.25, "0.0%"));
```

### Read and write the legacy text formats

```js doc-example
import { sheetToSYLK, sylkToSheet, sheetToDIF, difToSheet } from "gridxcel/formats";
import { aoaToSheet } from "gridxcel/json";
import assert from "node:assert";

const sheet = aoaToSheet([["name", "score"], ["ada", 42]]);

// SYLK: record-based, typed cells.
const sylk = sheetToSYLK(sheet);
assert.match(sylk, /^ID;P/);
assert.deepStrictEqual(sylkToSheet(sylk).rows[1].map((c) => c.v), ["ada", 42]);

// DIF: header section, then two-line data chunks.
const dif = sheetToDIF(sheet);
assert.match(dif, /^TABLE/);
assert.deepStrictEqual(difToSheet(dif).rows[1].map((c) => c.v), ["ada", 42]);

console.log("sylk:", sylk.length, "bytes | dif:", dif.length, "bytes");
```

### Work in a browser

```js doc-example
// The browser build is selected by the package "browser" field; the same import
// works. No shim, no polyfill, no bundler required.
import { read, write, eachRow } from "gridxcel";
import assert from "node:assert";

// In a browser these take a Response, a Blob, a File or a fetch body.
const bytes = await write([["a", 1], ["b", 2]]);
const book = await read(bytes);
assert.strictEqual(book.sheets[0].rows.length, 2);

// Only the filesystem helpers and the synchronous API are Node-only.
console.log("byte-level API:", typeof read, typeof write, typeof eachRow);
```

## Entry points

Import from a subpath to keep a bundle small. Every subpath is standalone.

| Specifier | Use it for | Exports |
| --- | --- | --- |
| `gridxcel` | The main API: read, write, streaming, conversion, addressing | 69 |
| `gridxcel/compat` | A SheetJS-shaped surface for migrating existing code | 16 |
| `gridxcel/csv` | CSV/DSV parsing and serialisation, including streaming out | 8 |
| `gridxcel/json` | Sheet to objects/arrays and back, plus cell helpers | 19 |
| `gridxcel/numfmt` | The ECMA-376 number-format engine and date serials | 13 |
| `gridxcel/xml` | The SAX parser and entity codec | 11 |
| `gridxcel/zip` | The ZIP container reader and writer | 13 |
| `gridxcel/formats` | SYLK and DIF reading and writing | 4 |

153 export slots in total; 124 of those names are distinct. Both numbers are
checked by `npm run check:exports`, which derives the entry points from
`package.json`.

### Migrating from SheetJS

```js doc-example
import * as XLSX from "gridxcel/compat";
import assert from "node:assert";

// Same shape as the SheetJS API: utils, read, write, sheet_to_json.
const book = XLSX.read("a,b\n1,2\n");
const rows = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header: 1 });

assert.deepStrictEqual(rows, [["a", "b"], [1, 2]]);
console.log("SheetJS-shaped call sites keep working");
```

See [MIGRATION.md](docs/MIGRATION.md) for the differences that change results.

## Compatibility

| | |
| --- | --- |
| Node.js | 18, 20, 22, 24 - tested on Linux and Windows |
| Browsers | Native ESM; `browser` field swaps the platform module |
| Module format | ESM only. CommonJS uses dynamic `import()` |
| TypeScript | Hand-written declarations for all eight entry points |
| Dependencies | None, at runtime or in development |
| Build step | None. `src/*.js` is what ships and what runs |

## Performance and memory

Streaming a 200,000 x 10 workbook:

| Stage | rows/s | Peak heap | Retained |
| --- | --- | --- | --- |
| `eachRow` (streaming read) | 77,309 | 13.0 MiB | **5.9 MiB** |
| `readWorkbook` (full parse) | 68,086 | 224.9 MiB | 164.9 MiB |
| `createCSVStream` (streaming write) | ~590,000 | 1 row | **0.1 MiB** |
| `WorkbookWriter` (streaming write) | 58,032 | 375.9 MiB | **5.8 MiB** |

Streaming gives up roughly 12% of throughput to hold 3.6% of the heap. The
package is 204 kB packed from 656 kB unpacked, unminified.

Method, caveats, and the quadratic scanner defect found and fixed along the way
are in [PERFORMANCE.md](docs/PERFORMANCE.md).

## What this library does not do

Read [LIMITATIONS.md](docs/LIMITATIONS.md) before adopting it. In short:

* **No XLSB, XLS, ODS, NUMBERS, DBF or HTML tables.** XLSX, CSV, SYLK and DIF only.
* **No charts, images, pivot tables, macros, conditional formatting or data
  validation.** Nothing is preserved and nothing is written.
* **No formula evaluation.** A formula is text with an optional cached result.
* **No CommonJS build.** ESM only.
* **No synchronous API outside Node.** A browser has no synchronous inflate, so
  `readSync` and `writeSync` throw `PlatformError` with code `NOT_SUPPORTED`.
  Use the async twins.

If you need any of those, use [SheetJS](https://docs.sheetjs.com/). It covers far
more ground, and [COMPARISON.md](docs/COMPARISON.md) states plainly where it wins.
Choosing this library means choosing a few formats done well over fifteen done
adequately.

## Verification

Correctness is checked against an implementation that is not this one. Every read
and write path is diffed against Python **openpyxl** and **XlsxWriter**:

```sh
npm test                 # unit, integration, fuzz and scaling suites
npm run check:docs       # every documented symbol exists; every example executes
npm run check:exports    # every public export is documented
npm run check:clone      # the docs work from a checkout without untracked files
npm run check:fuzz       # a million random cases across ten parsers
npm run interop          # diff against openpyxl / XlsxWriter
```

| Check | Result |
| --- | --- |
| Tests | 324 run, 320 pass, 0 fail, 4 skipped |
| Oracle interop | 37,293 read checks and 380 write checks, 0 mismatches |
| Fuzzing | 456 corrupted containers plus 328 hostile text inputs: 0 untyped throws, 0 hangs, 0 unbounded allocations |
| CRC corruption | Detected and rejected with a typed error |
| CI | 8 legs: Linux and Windows, Node 18/20/22/24, plus Chromium, Firefox and WebKit |
| Random fuzzing | Ten parsers, millions of generated cases per run; fixed seeds reproduce any failure |

The suite needs no network access and installs nothing.

## Documentation

| Document | Contents |
| --- | --- |
| [QUICKSTART.md](docs/QUICKSTART.md) | Install to working code, assuming no XLSX knowledge |
| [API.md](docs/API.md) | Complete reference for all eight entry points |
| [COOKBOOK.md](docs/COOKBOOK.md) | Task-shaped recipes |
| [MIGRATION.md](docs/MIGRATION.md) | Moving from SheetJS, including the compat shim |
| [LIMITATIONS.md](docs/LIMITATIONS.md) | Unsupported formats and features |
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | How it is built and why |
| [PERFORMANCE.md](docs/PERFORMANCE.md) | Every benchmark and the method |
| [COMPARISON.md](docs/COMPARISON.md) | Measured comparison against SheetJS |
| [BROWSER.md](docs/BROWSER.md) | Bundlers, `<script>`, React, Preact, TypeScript |
| [examples/](examples/) | Runnable scripts, one per task |

## License

MIT. See [LICENSE](LICENSE).

Contributions welcome - see [CONTRIBUTING.md](CONTRIBUTING.md). Security reports
go through [SECURITY.md](SECURITY.md).
