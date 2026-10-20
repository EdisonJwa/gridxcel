# gridxcel

**Read and write Excel files in Node.js and the browser without giving up control of your memory, your types, or your error handling.**

[![npm version](https://img.shields.io/npm/v/gridxcel.svg)](https://www.npmjs.com/package/gridxcel)
[![licence](https://img.shields.io/badge/licence-MIT-blue.svg)](LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org)
[![dependencies](https://img.shields.io/badge/dependencies-0-brightgreen.svg)](package.json)
[![tests](https://img.shields.io/badge/tests-250%20passing-brightgreen.svg)](https://github.com/EdisonJwa/gridxcel/blob/main/test/README.md)
[![bundle](https://img.shields.io/badge/browser-native%20ESM%2C%20no%20build-blue.svg)](docs/BROWSER.md)

`gridxcel` is a streaming-first XLSX and CSV reader/writer for Node.js **and the
browser**. Zero dependencies, no build step, nothing vendored - the whole runtime
import list is five `node:` builtins, and only one of them is needed in a browser.

Three properties make it worth choosing over the alternatives:

* **Bounded memory at every stage.** `eachRow` streams a worksheet one row at a
  time; `createCSVStream` streams the CSV out the same way. A 200,000 x 10 sheet
  costs 5.9 MiB retained to read and 0.1 MiB to serialise, against 164.9 MiB and
  14.3 MiB for the eager forms - at the same throughput.
* **Fast.** `read` is **3.4x faster than SheetJS CE 0.20.3** on a 50,000 x 8
  workbook, measured in the same process on the same bytes across nine passes.
* **It tells you when something is wrong.** A malformed upload throws a typed
  error instead of returning a plausible empty sheet; every part is CRC-checked;
  every failure carries a stable `code` you can branch on.

And the three commitments it has kept since the first commit:

* **Bounded memory.** `eachRow` streams a worksheet one row at a time. On a
  200,000 x 10 workbook it retains 5.9 MiB against 164.9 MiB for a full parse, at
  the same throughput.
* **Loud failure.** Bytes that are not a workbook throw a typed error instead of
  silently returning an empty sheet. Corruption is caught by CRC rather than
  read past. Errors share one base class and carry a stable `code`.
* **Dates that survive.** A date cell keeps its raw serial next to the Date, and
  parsing uses the local wall clock rather than shifting the value by your
  timezone offset.

> **Status: 0.1.0, pre-1.0.** Functional safety is covered and the API is stable
> in practice, but the surface may still change before 1.0. Read
> [docs/LIMITATIONS.md](docs/LIMITATIONS.md) for the supported surface. Everything
> claimed below is checked by a command that names it, and the full gate runs on
> two operating systems and four Node versions.

---

## Contents

- [Install](#install)
- [Quick start](#quick-start)
- [Why gridxcel](#why-gridxcel)
- [What it does not do](#what-it-does-not-do)
- [Documentation](#documentation)
- [Verification](#verification)
- [Performance](#performance)
- [Requirements](#requirements)
- [Project](#project)

## Install

```sh
npm install gridxcel
```

There is nothing to build and nothing to configure. gridxcel is ESM-only.

## Quick start

```js doc-example
import { read, write, eachRow, sheetToJSON } from "gridxcel";
import assert from "node:assert";

// Read. Bytes, a path, a URL, a stream or an async iterable all work.
const book = await read("../../test/fixtures/golden.xlsx", { cellDates: true });
assert.deepStrictEqual(book.sheets.map((s) => s.name), ["Types", "Layout", "Volume", "Empty"]);

// A sheet becomes objects, or arrays of arrays.
const rows = await sheetToJSON(book.sheets[2], { header: 1 });
assert.strictEqual(rows.length, 200);

// Write. Anything from an array of rows to a full workbook.
const bytes = await write([["name", "score"], ["ada", 42]], { bookType: "xlsx" });
assert.ok(bytes.length > 0);

// Stream a file that does not fit in memory.
let n = 0;
for await (const { sheet, r, cells } of eachRow("../../test/fixtures/golden.xlsx")) n++;
assert.strictEqual(n, 205);

// Stream it back out as CSV, one row at a time - no 14 MB string.
import { createCSVStream } from "gridxcel";
let csvBytes = 0;
for await (const chunk of createCSVStream(book.sheets[2], {}, { chunkRows: 256 })) {
  csvBytes += chunk.length; // write it to a file, a socket, a Response...
}
assert.ok(csvBytes > 0);

console.log("sheets:", book.sheets.map((s) => s.name).join(", "));
console.log("volume rows:", rows.length, "| streamed rows:", n, "| wrote", bytes.length, "bytes");
console.log("csv bytes streamed:", csvBytes);
```

Every `js doc-example` block in this repository is executed against the real
implementation by `npm run check:docs`, so the examples above cannot drift from
the code. Continue with [docs/QUICKSTART.md](docs/QUICKSTART.md), or go straight
to [docs/API.md](docs/API.md).

## Why gridxcel

| What you need | How gridxcel provides it |
| --- | --- |
| Read a workbook too large for memory | Incremental inflate feeds a SAX scanner; `eachRow` holds one row, `createCSVStream` emits a row at a time |
| Know when an upload is not a spreadsheet | Typed `FormatError`, never an empty sheet |
| Trust the numbers you read | Every part is CRC-checked; a mismatch throws |
| Branch on the failure, not the prose | One `GridxcelError` base class, a stable `code` on every one |
| Dates that do not shift by timezone | Local wall clock, raw serial preserved on `cell.serial` |
| Types that catch mistakes | Hand-written declarations; a consumer compiles with `skipLibCheck: false` and no `@types/node` |
| Run it in a browser, in a worker, in a bundler | Native ESM with a `browser` field swap; no shim, no polyfill, no build |
| Not inherit someone else's CVE history | Zero dependencies, nothing vendored, MIT |

### How it compares

The honest summary: **SheetJS reads more formats; gridxcel reads XLSX and CSV
faster, with bounded memory and stricter failure.** If you need XLSB, XLS, ODS,
formulas with a calculation engine, or pivot tables, use SheetJS - see
[docs/COMPARISON.md](docs/COMPARISON.md), which states plainly where it wins.

Where gridxcel wins, measured rather than claimed:

The performance claim is measured, not asserted. On a 50,000 x 8 workbook written
by openpyxl, with both libraries reading the same 400,008 cells in the same
process, across seven passes:

| Operation | Median time | Versus SheetJS CE 0.20.3 |
| --- | --- | --- |
| SheetJS `read` | 961 ms | - |
| gridxcel `readSync` | 315 ms | **2.99x - 3.32x faster** |
| gridxcel `read` | 281 ms | **3.40x - 3.75x faster** |
| gridxcel `eachRow` | 286 ms | **3.31x - 3.79x faster** |

So roughly **3.0x to 3.8x faster, median about 3.4x**. The range is the claim: a
single pass on a shared machine can land anywhere inside it, which is why it is
published as a range and not as one number. Reproduce with
`npm run bench:scaling`. Method and caveats in
[docs/PERFORMANCE.md](docs/PERFORMANCE.md).

## What it does not do

Read [docs/LIMITATIONS.md](docs/LIMITATIONS.md) before adopting this. In short:

* **No XLSB, XLS, ODS, NUMBERS, DBF or HTML tables.** XLSX, CSV, SYLK and DIF only.
* **No charts, images, pivot tables, macros, conditional formatting or data
  validation.**
* **No CommonJS.** ESM only - though browser bundling works, see [docs/BROWSER.md](docs/BROWSER.md).
* **Nothing else is streamed out as one call.** `createCSVStream` streams CSV
  rows; XLSX output still assembles before `finish()` returns.
* **No synchronous API outside Node.** A browser has no synchronous inflate, so
  `readSync`/`writeSync` throw `PlatformError` `[NOT_SUPPORTED]` there. Use the
  async twins.

If you need any of those, use [SheetJS](https://docs.sheetjs.com/) - it covers far
more ground, and [docs/COMPARISON.md](docs/COMPARISON.md) states plainly where it
beats gridxcel. Choosing gridxcel means choosing XLSX and CSV done well over
fifteen formats done adequately.

## Documentation

| Document | What it covers |
| --- | --- |
| [QUICKSTART.md](docs/QUICKSTART.md) | Install to working code, assuming no XLSX knowledge |
| [API.md](docs/API.md) | Complete reference for all seven entry points |
| [COOKBOOK.md](docs/COOKBOOK.md) | Task-shaped recipes: stream a huge file, CSV to XLSX, and more |
| [MIGRATION.md](docs/MIGRATION.md) | Moving from SheetJS, including the `gridxcel/compat` shim |
| [LIMITATIONS.md](docs/LIMITATIONS.md) | Unsupported formats and features, and the 0.1.0 caveats |
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | How it is built and why, for people changing it |
| [PERFORMANCE.md](docs/PERFORMANCE.md) | Every benchmark, the method, and the quadratic defect found and fixed |
| [COMPARISON.md](docs/COMPARISON.md) | Measured comparison against SheetJS CE 0.20.3 |
| [BROWSER.md](docs/BROWSER.md) | Using it in a browser: bundlers, `<script>`, React, Preact, TypeScript |
| [examples/](https://github.com/EdisonJwa/gridxcel/tree/main/examples) | Runnable scripts, one per task |
| [test/README.md](https://github.com/EdisonJwa/gridxcel/blob/main/test/README.md) | What the test suites cover and how to run them |

### Entry points

| Import | Use it for |
| --- | --- |
| `gridxcel` | The main API: read, write, streaming, conversion, addressing |
| `gridxcel/compat` | A SheetJS-shaped surface for migrating existing code |
| `gridxcel/csv` | CSV/DSV parsing and serialisation on their own |
| `gridxcel/json` | Sheet to array-of-arrays or JSON, and back |
| `gridxcel/numfmt` | The ECMA-376 number-format engine |
| `gridxcel/xml` | The SAX parser and entity codec |
| `gridxcel/zip` | The ZIP container reader and writer |
| `gridxcel/formats` | SYLK and DIF reading and writing |

## Verification

Correctness is checked against an implementation that is not this one. Every read
and write path is diffed against Python **openpyxl** and **XlsxWriter** over the
fixtures in `test/fixtures/`:

```sh
npm test                 # 250 tests: unit, integration, fuzz, scaling
npm run check:docs       # every documented symbol exists; every example runs
npm run check:exports    # every public export is documented
npm run interop          # diff against openpyxl / XlsxWriter
```

| Check | Result |
| --- | --- |
| Tests | 250 run, 246 pass, 0 fail, 4 skipped |
| Oracle interop | 37,293 read checks and 190 write checks, 0 mismatches |
| Fuzzing | 456 corrupted inputs: 402 typed errors, 0 untyped throws, 0 hangs |
| CRC corruption | Detected and rejected with a typed error |
| Documentation | Every documented symbol must exist and every example must execute (`npm run check:docs`): 62 examples and 141 links at the last measurement |

The suite needs no network access and installs nothing. See
[test/README.md](https://github.com/EdisonJwa/gridxcel/blob/main/test/README.md).

## Performance

Streaming a 200,000 x 10 workbook:

| Stage | rows/s | Peak heap | Retained |
| --- | --- | --- | --- |
| `eachRow` (streaming read) | 77,309 | 13.0 MiB | **5.9 MiB** |
| `readWorkbook` (full parse) | 68,086 | 224.9 MiB | 164.9 MiB |
| `createCSVStream` (streaming write) | ~590,000 | 1 row | **0.1 MiB** |
| `WorkbookWriter` (streaming write) | 58,032 | 375.9 MiB | **5.8 MiB** |

Streaming gives up roughly 12% of throughput to hold 3.6% of the heap. The
synchronous path is no longer a slow path: an unbounded quote search in the tag
scanner made it quadratic in buffer size, leaving the synchronous reader 6.6x
slower than the asynchronous one until it was fixed; they are now at parity.

The package is 286 kB raw and 85 kB gzipped, unminified, across the whole of
`src/`.

## Requirements

* **Node.js 18 or newer.** No older runtime is tested or supported.
* **Browsers, with no build step and no polyfill.** The byte-level API runs
  anywhere with the Web platform; only the filesystem helpers and the
  synchronous API are Node-only. See [docs/BROWSER.md](docs/BROWSER.md).
* **ESM.** `import` only. In CommonJS, use dynamic `import()`.
* No build step, no peer dependencies, no native modules.

## Project

* [CHANGELOG.md](CHANGELOG.md) - what changed, and which defects are still open
* [CONTRIBUTING.md](CONTRIBUTING.md) - setup, the test commands, and the doc-example rule
* [SECURITY.md](SECURITY.md) - how to report a vulnerability, and what is already defended
* [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) - Contributor Covenant 2.1

MIT licensed. See [LICENSE](LICENSE).
