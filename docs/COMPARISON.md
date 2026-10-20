# gridxcel vs SheetJS: an evidence-based comparison

Measured 2026-10-18 on the development machine (Windows, Node v24.13.0) against
**SheetJS Community Edition 0.20.3**, downloaded from
<https://cdn.sheetjs.com/> (the vendor CDN distribution, referenced by the
official installation documentation). Nothing here is taken from a README: every
SheetJS figure comes from the artifact itself, and every gridxcel figure comes
from this repository.

Two ground rules:

* Every gridxcel claim is followed by the command that produced it.
* Every SheetJS claim is either measured from the downloaded artifact or linked
  to its documentation.

Correctness evidence for gridxcel is the interop and fuzz suites
([test/README.md](https://github.com/EdisonJwa/gridxcel/blob/main/test/README.md)); this document is about behaviour, cost and
scope.

## What gridxcel does NOT do (read this before the rest)

This section decides whether gridxcel is usable for you. If your input is not
XLSX or CSV, stop here.

* **XLSB, XLS (BIFF), SpreadsheetML 2003, ODS/FODS, NUMBERS, DBF, SYLK, DIF, PRN,
  Lotus/Quattro, HTML/RTF tables** - none are implemented. SheetJS reads all of
  them ([format table](https://docs.sheetjs.com/docs/miscellany/formats/)).
* **Charts, images, pivot tables, VBA/macros, conditional formatting, data
  validation** - not implemented.
* **Writing to a non-seekable sink as one streamed response.** `WorkbookWriter`
  bounds memory while rows are built, but `finish()` returns the whole archive.
* **An incremental CSV writer.** `sheetToCSV` returns one string.
* **Any input that is not a workbook throws.** If you relied on SheetJS returning
  an empty workbook, that behaviour is deliberately not reproduced. See below.
* **A ten-year bug-fix tail.** SheetJS has a decade of field reports behind it.
  gridxcel has 16 fixtures and one generated volume file.

## The headline difference: how each handles bad input

Same inputs, both libraries, measured:

| Input | SheetJS 0.20.3 | gridxcel 0.1.0 |
| --- | --- | --- |
| 23 random bytes | returns a workbook: `["Sheet1"]` | throws `FormatError` / `NOT_A_ZIP` |
| empty buffer | returns a workbook: `["Sheet1"]` | throws `FormatError` / `NOT_A_ZIP` |
| plain text | returns a workbook: `["Sheet1"]` | throws `FormatError` / `NOT_A_ZIP` |
| truncated xlsx | throws `Error`, no `code` | throws `FormatError` |
| valid zip with no workbook part | throws `Error` | throws `FormatError` / `NOT_A_WORKBOOK` |
| byte flipped in a stored part | **returns a workbook** | throws `ZipError` / `CRC_MISMATCH` |
| archive missing `xl/styles.xml` | throws `Error: Cannot find file xl/styles.xml in zip` | reads it; the part is optional |

That third row is not academic: `test/fixtures/reader-stored-big.xlsx` omits
`styles.xml`. Python `zipfile.testzip()` returns `None` and openpyxl reads it
(6000 rows), while SheetJS refuses it. Strictness cuts both ways.

The first three rows matter more in a service that ingests uploaded files. An
empty `Sheet1` for arbitrary bytes turns "this is not a spreadsheet" into a
successful parse, and the corruption row turns a damaged file into wrong numbers.

## Bundle and footprint

Artifacts downloaded from <https://cdn.sheetjs.com/xlsx-0.20.3/>, measured with
gzip level 9 per file on both sides.

| Artifact | raw | gzip |
| --- | --- | --- |
| SheetJS `xlsx.mjs` (ESM) | 1,008,308 B | 254,382 B |
| SheetJS `xlsx.js` (CJS) | 975,564 B | 247,159 B |
| SheetJS `dist/xlsx.full.min.js` | 951,904 B | 335,014 B |
| SheetJS `dist/xlsx.mini.min.js` | 279,523 B | **86,664 B** |
| gridxcel `src/*.js` (17 modules, unminified ESM) | **285,581 B** | **85,290 B** |

gridxcel is 3.5x smaller than `xlsx.mjs` raw and 3.0x smaller gzipped. It is
**not** smaller than `xlsx.mini.min.js`, which is a minified
smaller-format-support build at 86,664 B gzipped against gridxcel's 85,290 B
unminified - those two are within 2% of each other, and minifying gridxcel would
put it clearly ahead, but do not expect a dramatic win over the mini build.

Installed footprint: SheetJS 0.20.3 unpacks to **7.7 MiB** across 26 files
(including source maps and five format-specific bundles); gridxcel's published
set is about **394 kB**.

## Dependencies and supply chain

| | gridxcel 0.1.0 | SheetJS 0.20.3 |
| --- | --- | --- |
| Declared runtime dependencies | `"dependencies": {}` | `"dependencies": {}` |
| Vendor audits | none known (new code, unaudited) | 0.18.5 on the npm registry carries CVE-2023-30533 and CVE-2024-22363; the CDN 0.20.3 line resolves them |

**Correction to a common claim.** Comparisons of SheetJS usually cite seven
dependencies (cfb, ssf, wmf, word, crc-32, adler-32, codepage) and two CVEs. Both
are true of the **npm registry's 0.18.5**, which the project has said is stale.
The CDN 0.20.3 build compared here declares `"dependencies": {}` - the codecs
are bundled. On that version, "zero dependencies" is not a differentiator.

## Error reporting

| | gridxcel 0.1.0 | SheetJS 0.20.3 `xlsx.mjs` |
| --- | --- | --- |
| `throw new Error(...)` | 0 | 285 |
| Typed error classes thrown | 95 | 0 |
| `.code` assignments | 4 (one per class definition) | 0 |
| Caller can catch one base class | `instanceof GridxcelError` | not possible; message text only |
| `console.warn` / `console.log` in the library | 0 | 8 |

gridxcel additionally offers `strict: true` (turn the first warning into an
exception) and an `onWarning` sink, with the collected warnings always available
as `book.warnings`.

## TypeScript

| | gridxcel | SheetJS 0.20.3 |
| --- | --- | --- |
| Declaration files | 8 | 1 |
| Declaration lines | 1,131 | 1,066 |
| `any` occurrences | **7** | 32 |
| Exported declarations | 201 | 70 |
| Declarations for every entry point | yes, all 7 subpaths | single entry point |

Counted with `grep -c` equivalents over `src/*.d.ts` against
`types/index.d.ts` in the 0.20.3 tarball. The `any` count is the one figure
worth reading twice: SheetJS uses `any` 32 times in its public surface, mostly in
the option and cell unions, where gridxcel uses it 7 times.

## Performance, measured head to head

Input: `bench/fair-50000x8.xlsx`, 50,000 rows x 8 columns, 400,008 cells,
2.06 MiB, written by openpyxl so that neither library produced it. Same bytes,
same process, same machine; both libraries read all 400,008 cells. Reproduce
with `node --expose-gc scripts/bench-scaling.mjs --passes=7`.

Seven passes, each the median of three runs. The time is the median across
passes; the speedup is the range across those same passes.

| Operation | Median time | Versus SheetJS |
| --- | --- | --- |
| SheetJS `read(bytes)` | 961 ms | - |
| gridxcel `readSync(bytes)` | 315 ms | median 3.20x, range 2.99x - 3.32x |
| gridxcel `read(bytes)` (async) | 281 ms | median 3.50x, range 3.40x - 3.75x |
| gridxcel `eachRow(bytes)` (streaming) | 286 ms | median 3.41x, range 3.31x - 3.79x |
| SheetJS read + write | 1,427 ms | - |
| gridxcel `writeWorkbookSync` | 348 ms | write only |

**gridxcel reads roughly 3.0x to 3.8x faster, median about 3.4x.** The range is
the claim: repeating the identical command produced speedups from 1.42x to 4.66x,
so a single pass is a sample rather than a result. Streaming costs the same as a
full parse while holding one row instead of a workbook.

**A correction.** Two earlier versions of this document quoted
2.74x/3.11x/3.06x and then 4.01x/4.38x/4.66x. Both were single passes, and both
overstated the typical case. The first also reported gridxcel's synchronous read
as 3.7x **slower** than SheetJS, which was a genuine defect - an unbounded quote
search in the tag scanner made it quadratic in buffer size. That is fixed; the
cause, the measurements and the fix are in
[docs/PERFORMANCE.md](PERFORMANCE.md). The synchronous path is now 315 ms
against 281 ms for the async one - the sync reader no longer pays a penalty for
being synchronous.

Memory on the 200,000 x 10 workbook in [bench-200000x10.json](https://github.com/EdisonJwa/gridxcel/blob/main/bench/bench-200000x10.json):

| Stage | ms | rows/s | peak heap | retained |
| --- | --- | --- | --- | --- |
| `eachRow` (streaming read) | 2,587 | 77,309 | 13.0 MiB | 5.9 MiB |
| `readWorkbook` (full read) | 2,937 | 68,086 | 224.9 MiB | 164.9 MiB |
| `createCSVStream` (streaming write) | ~340 | ~590,000 | 1 row | **0.1 MiB** |
| `sheetToCSV` | 414 | 482,905 | 259.6 MiB | 184.2 MiB |
| `WorkbookWriter` (streaming write) | 3,446 | 58,032 | 375.9 MiB | 5.8 MiB |
| `writeWorkbookSync` | 2,769 | 72,239 | 534.1 MiB | 164.1 MiB |

Streaming retains 5.9 MiB against 164.9 MiB for a full parse, at comparable
throughput. SheetJS publishes the opposite model -
"For maximal compatibility, SheetJS API functions read entire files into memory"
([large datasets](https://docs.sheetjs.com/docs/demos/bigdata/stream/)) - and its
`XLSX.stream` helpers (`to_csv`, `to_json`, `to_html`, `to_xlml`) all
consume an already-parsed worksheet, so the peak has already been paid.

## Cell model

| | gridxcel | SheetJS 0.20.3 |
| --- | --- | --- |
| Strings | `{ t: "string", v: "text" }` | `{ t: "s", v: "text" }` |
| Numbers | `{ t: "number", v: 42 }` | `{ t: "n", v: 42 }` |
| Booleans | `{ t: "boolean", v: true }` | `{ t: "b", v: true }` |
| Dates | `t: "date"` with `serial` kept alongside the `Date` | `t: "n"` by default; `t: "d"` only with `cellDates` |
| Display text | opt-in via `cellText` -> `cell.w` | always computed |
| Formula | `cell.f` without the leading `=` | `cell.f` without the leading `=` |
| Frozen panes | `sheet.freeze = { row, col }` | not exposed |
| Merges | `sheet.merges` (A1 strings) | `ws["!merges"]` (range objects) |
| Dimensions | `originRow` / `originCol` + sparse rows | `ws["!ref"]` |

`cell.l` (hyperlinks) is read on the batch path only; `eachRow` cannot
attach it because `<hyperlinks>` follows `<sheetData>`.

## Date handling

A cell whose number format is a date, holding the wall clock 2024-03-15 14:30:45
(serial 45366.6046875), read with `cellDates: true` on a GMT+9 host:

| | Result |
| --- | --- |
| SheetJS 0.20.3 | `Fri Mar 15 2024 23:30:45` |
| gridxcel 0.1.0 | `Fri Mar 15 2024 14:30:45` |

SheetJS's own renderer agrees the value is 14:30:45 (`w` is
`"2024-03-15 14:30:45"` on the same cell), so the `Date` it returns is
inconsistent with the display text it reports. gridxcel keeps the local wall
clock and preserves the raw serial on `cell.serial` so the convention can always
be reconstructed. The gridxcel suite is verified timezone-independent under
`TZ=UTC`, `TZ=America/New_York` and `TZ=Asia/Tokyo`.

## Feature coverage

| Capability | gridxcel | SheetJS 0.20.3 |
| --- | --- | --- |
| XLSX read/write | yes | yes |
| CSV read/write | yes | yes |
| XLSB / XLS / ODS / others | no | yes |
| Streaming row read | `eachRow`, one row at a time | parses whole sheet first |
| Streaming row write | `WorkbookWriter` | CSV/HTML/JSON only |
| Shared strings, inline strings, rich text | yes | yes |
| Shared and array formulas | resolved by reference shifting | yes |
| Number formats (ECMA-376 grammar) | yes (own renderer) | yes (`SSF`) |
| Fonts, fills, borders, alignment | modelled; round-trip verified; not python-verified | yes |
| Charts, images, pivot tables | no | charts/pivots not written either |
| Encrypted workbooks | no | partial read |
| SheetJS-compatible entry point | `gridxcel/compat` | n/a |

`SSF` was spot-checked for agreement rather than assumed:
`SSF.format("0.00%", 0.125)` and `SSF.format("yyyy-mm-dd", 45292)` both match
gridxcel's renderer exactly.

## Summary

Choose gridxcel when: your input is XLSX or CSV, you want malformed input to fail
loudly, you want typed errors, you need bounded memory on large sheets, or you
want a small dependency-free package with complete TypeScript declarations.

Choose SheetJS when: you need any other spreadsheet format, you need charts or
pivot tables, you need a decade of edge-case fixes, or you depend on the exact
SheetJS call shapes and error behaviour.

The two are not equivalent replacements. gridxcel is a narrower library that is
stricter about correctness, roughly 2.7-3.1x faster on this input, and much
better at streaming; SheetJS is a wider one that is far more permissive about
input and covers a decade of formats gridxcel does not attempt.

## Open questions this comparison surfaced

1. **Synchronous parse speed.** Resolved: the sync path was quadratic in the
   buffer size because of an unbounded quote search in the tag scanner. Fixed and
   covered by test/scaling.test.mjs; sync is now faster than SheetJS and at parity
   with gridxcel's own async path.
2. **Compressed content is trusted after inflate.** CRC covers stored parts and
   all parts are checked on read, but the cost of that check on very large
   deflated parts has not been benchmarked separately.
3. **Rendered text (`Cell.w`) has no independent oracle.** openpyxl does not
   render numbers, so the renderer is checked against a curated golden set
   (48 cases) and against documented Excel semantics, not against a second
   implementation. `SSF` agreements above are spot checks.
4. **Fixtures the oracle cannot read**
   (`reader-badcells`, `reader-cut-sheet`, `reader-cut-sst`,
   `reader-nosheets`) are skipped in the interop diff because openpyxl refuses
   them. They are fuzz inputs, not interop inputs.
5. **The SheetJS shim is not driven by the interop harness.** `gridxcel/compat`
   is unit-tested against openpyxl, but `scripts/interop.mjs` covers the native
   API only, so the shim's write path has not been inspected by python.
