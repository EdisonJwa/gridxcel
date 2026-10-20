# gridxcel

**Read and write Excel files in Node.js and the browser without giving up control of your memory, your types, or your error handling.**

[![npm version](https://img.shields.io/npm/v/gridxcel.svg)](https://www.npmjs.com/package/gridxcel)
[![licence](https://img.shields.io/badge/licence-MIT-blue.svg)](LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org)
[![dependencies](https://img.shields.io/badge/dependencies-0-brightgreen.svg)](package.json)
[![tests](https://img.shields.io/badge/tests-137%20passing-brightgreen.svg)](https://github.com/gridxcel/gridxcel/blob/main/test/README.md)

`gridxcel` is a streaming-first XLSX and CSV reader/writer for Node.js and the
browser. It has
no dependencies, no build step, and no vendored third-party code - the entire
runtime import list is `node:zlib`, `node:stream`, `node:fs`, `node:url`,
`node:module` and `node:buffer`.

It is a **narrower** library than SheetJS, deliberately. It reads and writes XLSX
and CSV well rather than reading fifteen formats adequately, and it is built
around three commitments:

* **Bounded memory.** `eachRow` streams a worksheet one row at a time. On a
  200,000 x 10 workbook it retains 5.9 MiB against 164.9 MiB for a full parse, at
  the same throughput.
* **Loud failure.** Bytes that are not a workbook throw a typed error instead of
  silently returning an empty sheet. Corruption is caught by CRC rather than
  read past. Errors share one base class and carry a stable `code`.
* **Dates that survive.** A date cell keeps its raw serial next to the Date, and
  parsing uses the local wall clock rather than shifting the value by your
  timezone offset.

> **Status: 0.1.0, pre-1.0.** The public API may still change before 1.0. The
> duplication defects (DUP-1 to DUP-8) are closed and browser support has landed
> - see [docs/BROWSER.md](docs/BROWSER.md). Read [docs/LIMITATIONS.md](docs/LIMITATIONS.md)
> before depending on gridxcel in production: functional safety is covered, API
> stability is not. The commands below are the release gate, and each one names
> the claim it proves.

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

console.log("sheets:", book.sheets.map((s) => s.name).join(", "));
console.log("volume rows:", rows.length, "| streamed rows:", n, "| wrote", bytes.length, "bytes");
```

Every `js doc-example` block in this repository is executed against the real
implementation by `npm run check:docs`, so the examples above cannot drift from
the code. Continue with [docs/QUICKSTART.md](docs/QUICKSTART.md), or go straight
to [docs/API.md](docs/API.md).

## Why gridxcel

| The problem in production | What gridxcel does |
| --- | --- |
| A large workbook costs several times its size in heap | Incremental inflate feeds a SAX scanner; `eachRow` holds one row |
| A malformed upload returns a plausible empty result | Throws `FormatError`. SheetJS 0.20.3 returns an empty `Sheet1` for 23 random bytes |
| A corrupted file silently yields wrong numbers | Every part is CRC-checked on read; a mismatch throws |
| Errors are prose strings you cannot branch on | 95 typed errors under one `GridxcelError` base class, each with a stable `code` |
| Dates shift by the host timezone | Local wall clock, with the raw serial preserved on `cell.serial` |
| Types are loose enough to hide mistakes | Hand-written declarations, 3 `any` in the whole public surface |
| You inherit someone else's CVE history | No dependencies, no vendored code, MIT |

The performance claim is measured, not asserted. On a 50,000 x 8 workbook written
by openpyxl, with both libraries reading the same 400,008 cells in the same
process, across seven passes:

| Operation | Median time | Versus SheetJS CE 0.20.3 |
| --- | --- | --- |
| SheetJS `read` | 961 ms | - |
| gridxcel `readSync` | 519 ms | **1.57x - 2.37x faster** |
| gridxcel `read` | 520 ms | **1.42x - 2.62x faster** |
| gridxcel `eachRow` | 511 ms | **1.56x - 2.68x faster** |

So roughly **1.4x to 2.7x faster, median about 1.9x**. The range is the claim: a
single pass on a shared machine can land anywhere inside it, which is why it is
published as a range and not as one number. Reproduce with
`npm run bench:scaling`. Method and caveats in
[docs/PERFORMANCE.md](docs/PERFORMANCE.md).

## What it does not do

Read [docs/LIMITATIONS.md](docs/LIMITATIONS.md) before adopting this. In short:

* **No XLSB, XLS, ODS, NUMBERS, DBF, SYLK, DIF or HTML tables.** XLSX and CSV only.
* **No charts, images, pivot tables, macros, conditional formatting or data
  validation.**
* **No CommonJS.** ESM only - though browser bundling works, see [docs/BROWSER.md](docs/BROWSER.md).
* **No incremental CSV writer.** `sheetToCSV` returns one string.
* **No synchronous API outside Node.** A browser has no synchronous inflate, so
  `readSync`/`writeSync` throw `PlatformError` `[NOT_SUPPORTED]` there. Use the
  async twins.

If you need any of those, use [SheetJS](https://docs.sheetjs.com/). It covers far
more ground, and [docs/COMPARISON.md](docs/COMPARISON.md) states plainly where it
beats gridxcel.

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
| [examples/](https://github.com/gridxcel/gridxcel/tree/main/examples) | Runnable scripts, one per task |
| [test/README.md](https://github.com/gridxcel/gridxcel/blob/main/test/README.md) | What the test suites cover and how to run them |

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

## Verification

Correctness is checked against an implementation that is not this one. Every read
and write path is diffed against Python **openpyxl** and **XlsxWriter** over the
fixtures in `test/fixtures/`:

```sh
npm test                 # 225 tests: unit, integration, fuzz, scaling
npm run check:docs       # every documented symbol exists; every example runs
npm run check:exports    # every public export is documented
npm run interop          # diff against openpyxl / XlsxWriter
```

| Check | Result |
| --- | --- |
| Tests | 225 run, 221 pass, 0 fail, 4 skipped |
| Oracle interop | 37,293 read checks and 190 write checks, 0 mismatches |
| Fuzzing | 456 corrupted inputs: 402 typed errors, 0 untyped throws, 0 hangs |
| CRC corruption | Detected and rejected with a typed error |
| Documentation | Every documented symbol must exist and every example must execute (`npm run check:docs`): 138 symbols, 61 examples and 149 links at the last measurement |

The suite needs no network access and installs nothing. See
[test/README.md](https://github.com/gridxcel/gridxcel/blob/main/test/README.md).

## Performance

Streaming a 200,000 x 10 workbook:

| Stage | rows/s | Peak heap | Retained |
| --- | --- | --- | --- |
| `eachRow` (streaming read) | 77,309 | 13.0 MiB | **5.9 MiB** |
| `readWorkbook` (full parse) | 68,086 | 224.9 MiB | 164.9 MiB |
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
