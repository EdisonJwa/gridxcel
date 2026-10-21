# Limitations

What this library does, what it does not, and what is verified versus merely
implemented. Read the first section before adopting it.

## Supported formats

gridxcel reads and writes XLSX-family packages, delimited text, and the two
legacy text formats SYLK and DIF. It will not open anything else, it will not
guess, and there is no plugin to add a format.

| You have | Supported | Alternative |
| --- | --- | --- |
| XLSX, XLSM | Yes | - |
| CSV, TSV, any single-character delimiter | Yes | - |
| SYLK | Yes - `gridxcel/formats`, or transparently via the compat sniffer | - |
| DIF | Yes - `gridxcel/formats`, or transparently via the compat sniffer | - |
| XLSB | No | SheetJS |
| XLS (BIFF) | No | SheetJS |
| SpreadsheetML 2003 | No | SheetJS |
| ODS, FODS | No | SheetJS |
| NUMBERS | No | SheetJS |
| DBF, PRN | No | SheetJS |
| Lotus 1-2-3, Quattro Pro | No | SheetJS |
| HTML or RTF tables | No | SheetJS, or an HTML parser |

"XLSX family" means the OPC package with `xl/workbook.xml` and worksheet parts.
A file with an `.xlsx` extension that is really an XLSB or an ODS package fails
with a container or schema error, not with a wrong answer.

## Not implemented inside XLSX either

Supporting the container is not the same as supporting what can be put in it.
These are absent, not partially working:

* **Charts, images, pivot tables, VBA and macros.** Nothing is preserved and
  nothing is written. Reading a workbook that has them and writing it back drops
  them.
* **Conditional formatting and data validation.** Neither is read nor written.
* **Autofilter rules.** The `ref` of an autofilter is written; the rule set is
  not.
* **Frozen panes beyond a simple split.** A row and column split is written and
  read; anything more elaborate is not.
* **Tables and list objects.** The range is data; the table definition is not
  represented.
* **Encrypted workbooks.** Reading a password-protected workbook is partial at
  best. Writing one is not supported.
* **Streaming XLSX output.** `WorkbookWriter` bounds memory while rows are
  produced, but `finish()` returns the complete archive. CSV does stream out -
  see `createCSVStream`.
* **Formula evaluation.** A formula is text with an optional cached result.
  Nothing is recalculated, and writing does not update the cache.

## In the browser

The byte-level API runs with no build step and no polyfill: `read`, `eachRow`,
`sheetToJSON`, `sheetToCSV`, `write`, `createCSVStream` and
`WorkbookWriter`. Where a capability is missing the failure is a typed error
rather than a wrong answer:

* **No synchronous API.** A browser has no synchronous inflate, so `readSync`,
  `writeSync` and `ZipArchive.bytes()` throw `PlatformError` with code
  `NOT_SUPPORTED`. The async twins work everywhere.
* **No filesystem.** `readFile` and `writeFile` are Node-only and throw the same
  typed refusal. Pass bytes, a `Blob`, a `File` or a `Response` instead.
* **A browser write stores its parts.** With no synchronous deflate available,
  the synchronous writer stores parts instead of compressing them. The archive is
  valid, just larger than the one Node writes. `WorkbookWriter` (async) does
  compress.
* **`deflate-raw` has a floor.** The Web compression streams must accept that
  format. Node 20+, Chrome 80+ and Firefox 113+ do. The library feature-detects it
  rather than assuming, and falls back cleanly.

## Verified versus merely implemented

These work. Something specific about them is not independently proven. The
distinction matters if you are deciding what to rely on.

| Area | What is not proven |
| --- | --- |
| Rendered display text (`cell.w`) | Format codes are diffed against openpyxl, but openpyxl does not render numbers, so the rendered string has no outside oracle. The common path is exact; exotic codes rest on documented Excel semantics. |
| Fonts, fills, borders, alignment | They survive our own write-then-read round trip. No outside implementation has re-read the records we produce. |
| `gridxcel/compat` write path | The interop harness drives the native API. The shim is unit-tested but its write path has not been inspected by an independent reader. |
| Deliberately corrupt fixtures | openpyxl refuses to open them, so they have no oracle and are skipped in the interop harness. |
| Fuzzing | Fixed corpora plus a random fuzzer (millions of generated cases per run) prove no untyped throw, no hang and no unbounded allocation. They do not prove that a damaged file yields the correct *subset* of rows - only intact files are checked for content. |
| Browser | Executed in a sandbox with the Web platform and no `Buffer`/`process`/`require`, and in real Chromium, Firefox and WebKit (Playwright) in CI. Safari itself is approximated by WebKit, not tested directly. |

Evidence: `npm test` runs 303 tests (299 pass, 4 skipped);
`node scripts/interop.mjs` reports 190 write checks and 37,293 read checks
against openpyxl with 0 mismatches; `node scripts/check-crc.mjs` confirms a
corrupted part is rejected with `CRC_MISMATCH`.

## Operational limits

* **A very large file still needs its bytes in memory.** `eachRow` bounds the
  sheet representation, not the archive: the central directory must be read
  before any part can be located, so an input of N bytes costs at least N bytes.
  The input file is the floor.
* **A CPU-spinning parse cannot be interrupted from inside the process.** The
  fuzz suite races each case against a wall-clock timer, but a blocked event loop
  is only observable from outside. If you parse untrusted input in a server, do it
  in a worker or a child process that you can kill.
* **Dense conversion has a ceiling.** A range wider than the grid, or a sheet
  declaring an implausible number of rows, is refused with `RANGE_TOO_LARGE`
  rather than allocated. Pass an explicit `range` or raise `maxRows` when you
  mean it.
* **Writing is not cheaper than reading.** `WorkbookWriter` peaks at about
  375 MiB on a 200,000 x 10 sheet because the finished archive is produced at
  `finish()`. Retained memory is 5.8 MiB; peak is not.
* **The writer refuses what it cannot represent.** A font size outside 1-409
  points, a column width that is not finite and positive, an autofilter that is
  not a range, a line break in a SYLK or DIF cell: each is a typed error rather
  than a file that opens wrong somewhere else.

## Pre-1.0 caveats

* **The public API may still change.** Names, option defaults and return shapes
  are stable in intent, not by contract. Pin the version and read
  [CHANGELOG.md](../CHANGELOG.md) on upgrade.
* **Two behaviours are deliberate, not bugs.** The package root returns Promises
  where `gridxcel/compat` returns values, and `book_new()` returns the native
  Workbook from the root and the SheetJS-shaped book from `compat`. Tests assert
  both so they cannot drift.
* **Functional safety is covered; API stability is not.** Do not describe 0.1.0
  as production-ready without that caveat.

## Assumptions worth stating

* **Dates use the local wall clock.** `dateToSerial` and `serialToDate`
  round-trip the calendar fields a spreadsheet displays, in any timezone. The raw
  serial is preserved on `cell.serial`.
* **A cell carries an explicit type.** `cell.t` is one of `string`, `number`,
  `boolean`, `date`, `error`, `empty`. Nothing is inferred at read time from
  the visual format.
* **Numbers are IEEE doubles.** Every numeric value goes through a JavaScript
  number, so integers beyond 2^53 lose precision. This is the same limit SheetJS
  has.
* **Rows are sparse arrays.** A row using columns A and Z has two entries, not
  twenty-six. `originRow` and `originCol` say where the data starts.
* **An A1 address is not stored per cell.** It costs a string per cell, so it is
  derived on demand with `encodeCell`.
* **Nothing is recalculated.** A formula is text with an optional cached result.

## History

Defects found and fixed during the 0.1.0 build are recorded in
[CHANGELOG.md](../CHANGELOG.md) rather than here. This page describes the current
state; the changelog is the record of what was wrong.
