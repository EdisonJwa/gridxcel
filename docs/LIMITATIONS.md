# Limitations

## If your input is not XLSX or CSV, stop here

gridxcel reads XLSX-family packages and delimited text. Nothing else. It will
not open the file, it will not guess, and there is no plugin to add a format.
If the file you have is not one of the two below, this is the wrong library and
no amount of configuration changes that.

| You have | gridxcel | Use instead |
| --- | --- | --- |
| XLSX, XLSM | Yes | - |
| CSV, TSV, any single-character delimiter | Yes | - |
| XLSB | No | SheetJS |
| XLS (BIFF) | No | SheetJS |
| SpreadsheetML 2003 | No | SheetJS |
| ODS, FODS | No | SheetJS |
| NUMBERS | No | SheetJS |
| DBF, SYLK, DIF, PRN | No | SheetJS |
| Lotus 1-2-3, Quattro Pro | No | SheetJS |
| HTML or RTF tables | No | SheetJS, or a parser for that format |

"XLSX family" means the OPC package with `xl/workbook.xml` and worksheet parts.
A file with an `.xlsx` extension that is really an XLSB or an ODS package
fails with a container or schema error, not with a wrong answer.

## Not implemented inside XLSX either

Supporting the container is not the same as supporting what can be put in it.
These are absent, not partially working:

* **Charts, images, pivot tables, and VBA or macros.** Nothing is preserved and
  nothing is written. Reading a workbook that has them and writing it back drops
  them.
* **Conditional formatting and data validation.** Neither is read nor written.
* **Autofilter rules.** The `ref` of an autofilter is written; the rule set is
  not.
* **Frozen panes beyond a simple split.** A row and column split is written and
  read; anything more elaborate is not.
* **Tables and list objects.** The range is data to gridxcel; the table
  definition is not represented.
* **Encrypted workbooks.** Reading a password-protected workbook is partial at
  best. Writing one is not supported at all.
* **Writing XLSX to a non-seekable sink as one streamed response.**
  `WorkbookWriter` bounds memory while the rows are produced, but `finish()`
  returns the complete archive. There is no incremental output.
* **~~An incremental CSV writer.~~** Closed: `createCSVStream` yields the same
  bytes one row at a time, so a 200,000 x 10 sheet retains about 0.1 MB instead of
  the 14.3 MB the complete string costs, and it is ~20% faster because it never
  builds an array of lines to join.

## In the browser

The byte-level API runs in a browser with no build step and no polyfill:
`read`, `eachRow`, `sheetToJSON`, `sheetToCSV`, `write` and `WorkbookWriter`.
The limits are specific, and where a capability is missing the failure is a
typed error rather than a wrong answer:

* **No synchronous API.** A browser has no synchronous inflate - the Web
  compression API is asynchronous by construction - so `readSync`,
  `writeWorkbookSync` and `ZipArchive.bytes()` throw `PlatformError`
  `[NOT_SUPPORTED]`, naming the async twin to use instead. `read`,
  `writeWorkbook`, `eachRow` and `ZipEntry.bytesAsync()` all work.
* **No filesystem.** `readFile` and `writeFile` are Node-only and throw the same
  typed error. Use a `File`, a `Blob` or a `fetch` response instead.
* **`deflate-raw` has a floor.** The Web compression streams must accept
  `deflate-raw`: Chrome 103+, Firefox 113+, Safari 16.4+. On an older engine the
  library reports `NOT_SUPPORTED` rather than returning wrong bytes.
* **A browser write stores its parts.** With no synchronous deflate available,
  `ZipWriter` stores entries instead of compressing them, so the workbook is
  valid and readable but larger than one written in Node. A bigger file beats a
  failed write.

Compression is `node:zlib` on Node and `CompressionStream`/`DecompressionStream`
elsewhere, behind the package `browser` field. The full story, including
bundlers and React, is in [BROWSER.md](./BROWSER.md).

## 0.1.0 caveats

Read this section before shipping anything on 0.1.0.

* **The public API may still change before 1.0.** This is a pre-1.0 release.
  Names, option defaults and return shapes are stable in intent, not by
  contract. Pin the version and read the changelog on upgrade.
* **Two behaviours are deliberate, not bugs.** The package root returns Promises
  where `gridxcel/compat` returns values, and `book_new()` returns the native
  Workbook from the root and the SheetJS-shaped book from `compat`. Both
  differences are asserted by tests so they cannot drift.
* **DUP-1 to DUP-8 are fixed, with one duplicate left.** Every duplicated helper
  was collapsed onto one implementation and the divergences resolved: cell
  coercion is one rule that throws rather than silently inventing or dropping a
  value, `format_cell` answers the same from both entry points, `styles.xml` has
  one parser that serves the reader and the writer, and sheet-name rules, the A1
  decoder and the shared limits each live in one module, and `freeSheetName` -
  the last duplicate helper - was collapsed into `src/api.js` and is imported by
  its three former callers.
* **Do not describe 0.1.0 as production-ready without this caveat.** Functional
  safety is covered; API stability is not.

Two defects found during the 0.1.0 build are **fixed and regression-covered**,
and should not be described as open: a crafted row index that drove an unbounded
allocation (FS-1), and a quadratic tag scanner in the XML reader.

## Verified but incomplete

These work. Something specific about them is not proven.

| Area | What is missing |
| --- | --- |
| Rendered display text (`Cell.w`) | Format codes are diffed against openpyxl; the rendered string has no independent formatter to compare against, because openpyxl does not render numbers. The common path is exact - a date cell yields `"2024-03-15 14:30:45"` - and exotic codes rest on documented Excel semantics. |
| Fonts, fills, borders, alignment | They survive a write then read round trip, verified. They have not been written and re-read by a second implementation, so the OOXML record shape is unverified by an outside reader. |
| The compatibility entry point | `gridxcel/compat` is unit-tested against openpyxl for sheet names, the first data row and CSV line count. The interop harness drives the native API only, so the shim's own write path has not been inspected by an independent implementation. |
| Deliberately corrupt fixtures | `reader-badcells.xlsx`, `reader-cut-sheet.xlsx`, `reader-cut-sst.xlsx` and `reader-nosheets.xlsx` are robustness inputs. openpyxl refuses to open them, so they have no oracle and appear as skipped in the interop harness. |
| Fuzzing | 456 corrupted inputs prove no plain `Error`, no hang and damage reported: 402 typed errors, 54 recovered. They do not prove that a damaged file yields the correct subset of rows - only intact files are checked for content. |
| Browser support | The bundle is executed in a sandbox with the Web platform and no `Buffer`, `process`, `require` or module loader, and the same code path is checked against the Node backend. It has not been run in Chrome, Firefox or Safari, so the `deflate-raw` support floor and any engine-specific stream behaviour are untested here. |

The evidence for the rows above: `npm test` runs 225 tests (221 pass, 4
skipped); `node scripts/interop.mjs` reports 190 write checks and 37,293 read
checks against openpyxl, 0 mismatches; `node scripts/check-crc.mjs` confirms a
corrupted part is rejected with `CRC_MISMATCH`.

## Operational limits

* **A very large file still needs its bytes in memory.** `eachRow` bounds the
  sheet representation, not the archive: the central directory must be read
  before any part can be located, so an input of N bytes costs at least N bytes.
  For a 200,000 x 10 workbook the peak heap is 13.0 MiB streaming against
  224.9 MiB parsed, and the retained heap after a forced collection is 5.9 MiB
  against 164.9 MiB. The input file is the floor, and it is the same floor for
  both.
* **A CPU-spinning parse cannot be interrupted from inside the process.** The
  fuzz suite races each case against a wall-clock timer, but a blocked event
  loop is only observable from outside. Run `node --test --test-timeout=60000`
  in CI.
* **Writing is not cheaper than reading.** `WorkbookWriter` peaks at 375.9 MiB
  for a 200,000 x 10 workbook and retains 5.8 MiB; `writeWorkbookSync` peaks at
  534.1 MiB. The streaming writer bounds what the sheet costs, not what the
  archive costs.
* **Dense conversion has a ceiling.** The dense converters refuse a range wider
  than 262,144 rows unless you pass an explicit `range` or raise `maxRows`,
  because a worksheet is untrusted input and a single large row index would
  otherwise allocate gigabytes. Stream with `eachRow` when you really do have
  more rows than that.
* **Column width units are caller-supplied characters.** A caller migrating
  from XlsxWriter, whose column setter converts through pixels, will see a
  different number stored for the same intent. Both display identically; the
  stored value differs.
* **Bundle size is not a win.** The whole of `src/` is 285,581 bytes raw and
  85,290 bytes gzipped, unminified. SheetJS CE 0.20.3 is 1,008,308 bytes raw and
  254,382 bytes gzipped, but its mini build is 86,664 bytes gzipped - within
  about 2% of gridxcel. If size is your reason for choosing, that is not a
  reason.
* **Speed is a range, not a headline.** Read throughput is 1.4x to 2.7x
  SheetJS CE 0.20.3 on the same bytes in the same process, measured as a median
  each the median of three. The median is about 1.9x and a single pass can read as
  low as 1.42x; quote the range, never one run.

## Assumptions worth stating

* **Dates use the local wall clock.** `dateToSerial` and `serialToDate`
  round-trip the calendar fields a spreadsheet displays, in any timezone. Do not
  compare these `Date` objects with `toISOString()`, which converts to UTC and
  shifts by the host offset. Compare local fields, or compare serials.
* **A cell carries an explicit type.** `cell.t` is one of `string`,
  `number`, `boolean`, `date`, `error` or `empty`. Nothing downstream
  re-guesses it, so a value that looks like a number can still be a string and
  should be treated as one.
* **A cell's format is polymorphic by direction.** Writing accepts a format code
  or an id; reading returns the numeric id. The resolved code is always
  available as `Cell.zf`, from the workbook style table, or through
  `cellFormatCode`.
* **An A1 address is not stored per cell.** It costs a string per cell and is
  derivable from `originRow`, `originCol` and the index. Use `encodeCell` or
  `encodeCol` when you need one.
* **Rows are sparse arrays.** A row using columns A and Z has two entries, not
  twenty-six. Iterate with `filter(Boolean)`, or use `sheetToAOA` when you
  want a dense grid.
* **Streaming cells use absolute column indexing.** On a streamed row,
  `cells[0]` is column A. A batch sheet rebases to `originCol`. Hyperlinks are
  batch-only.
* **Numbers are IEEE doubles.** Every numeric value goes through a JavaScript
  number, so integers beyond 2^53 and long decimal fractions lose precision
  exactly as they do anywhere else in the language.
* **Two deliberate deviations from SheetJS** are listed in
  [MIGRATION.md](./MIGRATION.md); both are one-line explainable and both are
  asserted by tests.
* **Nothing is recalculated.** A formula is text with an optional cached result.
  gridxcel never evaluates one, and writing does not update the cache.

## Fixed during the 0.1.0 build

Recorded because the failure modes are instructive, not to pad a changelog.
Struck-through items remain visible in [test/README.md](https://github.com/gridxcel/gridxcel/blob/main/test/README.md).

* **Corrupt content was returned silently.** Stored parts were never
  CRC-checked. Every part is now verified on read, and corruption raises
  `ZipError` with `CRC_MISMATCH`. The check is a permanent regression script.
* **The ZIP error escaped the error hierarchy.** It extended `Error`, so a
  catch on `GridxcelError` missed container failures. The fuzz suite counted 4
  of 456 cases as untyped; it is now 402 typed and 0 plain.
* **The XML parser silently produced zero events** for any document whose last
  attribute before the closing bracket ended with a quote, which meant
  `xl/sharedStrings.xml` yielded nothing at all. Tag termination is now an
  explicit quote-state scan, with a chunk-boundary property test at every chunk
  size from 1 to 97 bytes.
* **Number formats did not survive a round trip.** Reading and writing a
  workbook downgraded a date's format to the default. Verified fixed: 823 cells,
  0 value differences and 0 number-format differences through openpyxl.
* **Not-a-workbook input returned an empty sheet list** instead of throwing,
  which turned a typo or a corrupt download into an apparent success. It now
  raises `FormatError` with `NOT_A_ZIP` or `NOT_A_WORKBOOK`.
