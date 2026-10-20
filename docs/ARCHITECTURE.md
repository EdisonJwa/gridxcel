# Architecture

gridxcel is a zero-dependency, streaming-first XLSX reader and writer for
JavaScript. This document describes how the modules fit together, the data flow
from bytes to cells, and the design decisions behind both - each with the
reason it was taken. It is a repository document; the published user-facing set
is [API.md](./API.md), [MIGRATION.md](./MIGRATION.md),
[LIMITATIONS.md](./LIMITATIONS.md), [PERFORMANCE.md](./PERFORMANCE.md) and
[BROWSER.md](./BROWSER.md).

## Contents

* [Design constraints](#design-constraints)
* [Module map](#module-map)
* [Data flow](#data-flow)
* [Sync and async pairs](#sync-and-async-pairs)
* [Design decisions](#design-decisions)
* [Verification](#verification)
* [Defect status at 0.1.0](#defect-status-at-010)
* [Publishing surface](#publishing-surface)

---

## Design constraints

Every decision below follows from these constraints, so they come first.

| Constraint | Consequence for the design |
| --- | --- |
| ESM only, no CommonJS build | Peer modules are resolved lazily through a table of literal `import()` thunks. Literal, because a bundler can only rewrite a specifier it can see: a computed `import(spec)` survives into a browser bundle as a runtime import that no bundler can resolve. Nothing awaits at module scope, so the graph stays synchronous and an IIFE (`<script>`) build is possible. |
| Zero runtime dependencies, ever | The ZIP container, the deflate glue, the XML parser, the number-format engine and the CRC are first-party. Compression is `node:zlib` on Node and the Web `CompressionStream`/`DecompressionStream` everywhere else, behind one seam (`src/platform.js`); no polyfill and no dependency either way. |
| No build step | `src/*.js` is exactly what ships. There is no `dist/`, no transpile and no bundler, so a stack trace points at the file in the repository. |
| Node >= 18, and browsers | Modern builtins are assumed: raw deflate/inflate, `TextDecoder`/`TextEncoder`, async iterators. Everything the byte-level API needs has a Web equivalent, so the same source runs in a browser; only the filesystem helpers and the synchronous API are Node-only, and they say so with a typed error. |
| Browsers supported, without a build | `src/platform.browser.js` replaces `src/platform.js` through the package `browser` field, which esbuild, Vite, webpack and Rollup all honour. The browser file contains no `node:` reference at all, and the whole library still ships as the authored source: there is no browser bundle to keep in sync. `integration/browser-bundle.test.mjs` proves it by executing the bundle in a sandbox with the Web platform and no `Buffer`, `process` or `require`. |
| Untrusted input | A worksheet is attacker-controlled. Nothing may allocate proportionally to a declared extent, and no scan may be super-linear in its input. |
| Bounded memory for large files | Streaming read and streaming write never materialise a whole sheet; a full read of a 200,000 x 10 sheet peaked at 224.9 MiB against 13.0 MiB for the streaming path. |

---

## Module map

`src/` holds 18 JavaScript modules and 8 hand-written declaration files, about
299 kB raw, unminified. Every module below has exactly one home for the idea it
owns; where two modules used to implement the same thing, the survivor is named
in the last column.

| Layer | Module | Responsibility |
| --- | --- | --- |
| Public surface | `src/index.js` | Re-exports the library and wraps the converters. |
| Public surface | `src/api.js` | Normalises caller input, dispatches to reader/writer, holds `VERSION`. **Canonical:** `validateSheetName`, `normalizeSheet`, `normalizeBook`. |
| Public surface | `src/types.d.ts` | The frozen data contracts shared by every module. |
| Container | `src/zip.js` | Central-directory parse, store/deflate, Zip64, writer. |
| Container | `src/crc32.js` | Slice-by-8 CRC-32 table and incremental update. |
| Package | `src/opc.js` | `[Content_Types].xml`, relationship resolution, workbook metadata. |
| Parse | `src/xml.js` | Incremental event parser, entity codec, tag splitting. **Canonical:** `encodeXmlAttr`, `localName`. |
| Parse | `src/sheet-reader.js` | Chunked worksheet scanner, `RowScanner`, streaming row reader. |
| Parse | `src/read.js` | Workbook assembly (async and sync), option normalisation, warnings. |
| Model | `src/address.js` | A1 addressing and range maths. **Canonical:** `decodeCell`. |
| Model | `src/constants.js` | Spreadsheet limits, in one place: sheet-name length, columns, rows, sheet count, default sheet name. |
| Model | `src/numfmt.js` | ECMA-376 number-format engine and date serials. |
| Model | `src/errors.js` | The `GridxcelError` hierarchy. |
| Serialise | `src/workbook-writer.js` | Package assembly, `SheetWriter`, `WorkbookWriter`, `buildParts`. |
| Serialise | `src/style-registry.js` | `styles.xml` parse/write, style interning, custom format ids. **Canonical:** `parseStyles`, used by the reader and the writer. |
| Serialise | `src/write.js` | Filesystem-facing writers and output shaping. |
| Conversion | `src/csv.js` | CSV/DSV read and write. |
| Conversion | `src/json.js` | Sheet to AOA/JSON and back, cell helpers, dense-range guard. **Canonical:** `cellFromValue`, the one value-to-cell coercion every other path delegates to. |
| Compatibility | `src/compat.js` | SheetJS-shaped entry point. |

The dependency direction is one-way: container knows nothing about worksheets,
the parser knows nothing about the writer, and only `index.js` and `api.js`
know about all of them. That is what keeps the subpath entry points
independently importable.

---

## Data flow

```
input bytes / path / URL / stream / async iterable
        |
        v
   [ src/zip.js ]        central directory, CRC-32 verification per part
        |
        v
   [ src/opc.js ]        [Content_Types].xml -> rels -> workbook.xml metadata
        |
        v
   [ src/sheet-reader.js ]  inflate chunks -> XmlParser -> RowScanner
        |                        |                     |
        |                        |                     +--> streaming: row events
        |                        v
        |                   shared strings, styles, dates
        v
   [ src/read.js ]       assembly: options, bounds, stubs, comments, warnings
        |
        v
   [ src/json.js ]       Sheet -> cells -> AOA / JSON / CSV
```

Read step by step:

| Step | What happens | Why it is its own layer |
| --- | --- | --- |
| 1. Bytes | `read.js` normalises `Uint8Array`, `ArrayBuffer`, a path, a `URL`, a `ReadableStream`, a Node stream or any async iterable of chunks. | One place decides what "input" means, so every entry point accepts the same set. |
| 2. Zip | `ZipArchive` finds the end-of-central-directory record (and the Zip64 variants), reads the directory once, and hands out entries. | The container is a separate concern from the spreadsheet; `gridxcel/zip` is usable on its own. |
| 3. Parts | `opc.js` resolves the content types, the package relationships and `xl/workbook.xml` into sheet list, props, defined names, date system and part paths. | Metadata is cheap and needed before any worksheet is inflated, which is what makes `sheets: false` and `sheetNames` fast. |
| 4. Scanner | `sheet-reader.js` inflates one worksheet at a time into chunks, feeds them to `XmlParser`, and turns rows into cells against the shared-string and style tables. | A worksheet may be hundreds of megabytes; it is consumed as a stream and never buffered whole. |
| 5. Sheet | `read.js` applies options (range bounds, stubs, dates, text, comments), collects warnings and assembles the `Sheet`. | Policy lives here, not in the scanner, so the scanner can stay a pure one-pass parser. |
| 6. Cells | Sparse rows indexed from `originCol`, with `originRow` recorded once. | A cell object per gap would cost more than the data itself. |

The write path mirrors it: `Workbook` -> `style-registry.js` (intern fonts,
fills, borders, xf records, format ids) -> `workbook-writer.js`
(`SheetWriter` serialises rows as they are added) -> `zip.js` (deflate and
assemble) -> bytes.

The streaming read path bypasses step 5 entirely:

```
bytes -> zip -> opc -> sheet-reader -> { sheet, r, cells } per row
```

`eachRow` never builds a `Sheet`, which is why it keeps 5.9 MiB retained
against 164.9 MiB for a full read of the same 200,000 x 10 workbook, and why it
cannot offer hyperlinks: `<hyperlinks>` follows `<sheetData>` in the part,
and honouring it in one pass would mean retaining rows.

A layered view in code, from bytes to parts:

```js doc-example
import { read } from "gridxcel";
import { ZipArchive } from "gridxcel/zip";
import { readFileSync } from "node:fs";

const zip = new ZipArchive(readFileSync("../../test/fixtures/golden.xlsx"));
const parts = zip.names().filter((name) => name.startsWith("xl/"));
console.log("container parts:", parts.slice(0, 4).join(","));

const book = await read("../../test/fixtures/golden.xlsx", { sheets: false });
console.log("declared sheets:", book.sheetNames.join(","), "| parsed sheets:", book.sheets.length);
```

---

## Sync and async pairs

The library ships both a promise-based and a synchronous form of every
filesystem-facing operation, because the two audiences are genuinely different.

| Operation | Async | Sync | Why both exist |
| --- | --- | --- | --- |
| Read | `read` / `readFile` | `readSync` | A stream or socket can only be read asynchronously; a CLI, a config loader or a test that already holds a buffer cannot await. |
| Write | `write` / `writeFile` / `writeWorkbook` | `writeWorkbookSync` / `toBuffer` | Same split on the way out; the sync writer also makes byte-for-byte assertions in tests trivial. |
| Rows | `eachRow` | - | Streaming is inherently asynchronous; there is no synchronous generator form. |
| Compat | `readAsync` / `readFileAsync` / `writeAsync` / `writeFileAsync` | `read` / `readFile` / `write` / `writeFile` | SheetJS users call a synchronous `read()` by default, so the compat entry point keeps that shape and adds explicit async names beside it. |

Three properties keep the pair honest:

* **One implementation.** `readWorkbook` and `readWorkbookSync` share the
  same helpers (options, warnings, assembly rules); only the byte source
  differs. `writeWorkbook`, `writeWorkbookSync` and `toBuffer` all funnel
  through `createBookWriter`, so a book written from disk or from a buffer
  produces byte-identical parts.
* **The sync path declares its limits.** `readSync` accepts bytes only, and
  synchronous inflate is a Node capability: a browser has
  `DecompressionStream`, which is asynchronous. So the synchronous API is
  Node-only, and outside Node it throws `PlatformError` with code
  `NOT_SUPPORTED` naming the async alternative, rather than pretending. Every
  synchronous entry point has an async twin, so no capability is lost.
* **The async path is the default in documentation and examples**, because it
  is the one that works for files that do not fit in memory.

---

## Design decisions

Each decision states what was chosen and why. The first twelve are the original
frozen decisions, tightened; the last two were promoted to first-class rules
after defects were found.

1. **Dates are local wall clock.** `dateToSerial` and `serialToDate` convert
   between a serial and a `Date` whose *local* fields are the spreadsheet's
   wall clock, in any timezone. **Why:** a cell reading `2024-03-15 14:30:45`
   is a wall-clock label, not an instant; converting through UTC would shift the
   displayed time by the host offset and silently corrupt every date in a
   round trip. Never compare these dates with `toISOString()`. The interop
   suite runs under `TZ=UTC`, `TZ=America/New_York` and `TZ=Asia/Tokyo`.

2. **`Cell.z` is polymorphic by direction.** On write it may be a format code
   or a numeric id; on read it is the numeric `numFmtId`. **Why:** SheetJS
   users write codes and the file format stores ids, and refusing either would
   break a drop-in migration. The resolved code is exposed three other ways so a
   caller never guesses: `Cell.zf` (present on both the batch and streaming
   paths), `Book.styles[cell.s].numFmt`, and `Book.numFmts.get(cell.z)`
   (which includes the built-in table). `cellFormatCode(book, cell)` resolves
   in one call. Ids below 164 are built-ins and must not be re-emitted as custom
   formats.

3. **Read results carry more than the format requires.** `Book.sheetNames`
   (populated on every path, including `sheets: false`), `Book.warnings`
   (always an array, the same array handed to `onWarning`), `Cell.zf`,
   `Cell.l` for hyperlinks, and `Sheet.tabColor` normalised to `#RRGGBB`.
   **Why:** callers should not have to opt into a callback to discover that a
   file was damaged, and the streaming path has no workbook to resolve a format
   id against, so the code has to travel with the cell.

4. **`Cell.a` is deliberately not populated by the reader.** **Why:** an
   address string per cell is real memory on a 400,000-cell sheet, and the
   address is derivable from `originRow`, `originCol` and the index. Readers
   that want addresses compute them; writers that produce them (`aoaToSheet`)
   fill them in.

5. **Container errors throw; damaged content warns.** Bytes that are not a ZIP
   (`FormatError` / `NOT_A_ZIP`), a ZIP with no workbook part
   (`NOT_A_WORKBOOK`) and filesystem failures (`FILE_NOT_FOUND`,
   `FILE_ACCESS_DENIED`, `FILE_IS_DIRECTORY`) throw typed errors. A real
   workbook with corrupt parts returns what it could read plus warnings.
   **Why:** the two situations need opposite handling. "This is not a
   spreadsheet" is a bug in the call; "row 40,000 is truncated" is a fact about
   the data, and a caller who can read 39,999 rows should get them. A genuine
   workbook declaring zero sheets returns an empty book with `NO_SHEETS` -
   the one case where an empty sheet list is truthful.

6. **Every part is CRC-verified on read.** `ZipArchive` checks each
   decompressed part against the CRC-32 in the central directory and throws
   `CRC_MISMATCH` instead of returning corrupt bytes; pass
   `new ZipArchive(bytes, { verifyCrc: false })` to opt out. **Why:** a
   flipped bit in a deflate stream usually decodes successfully into *wrong*
   values, and a silently wrong cell is worse than a failed read. The measured
   cost is about 1.2 GB/s of payload, negligible against inflate.

7. **`eachRow` uses absolute column indexing.** `cells[c]` is column `c`
   (`A` = 0) and `sheetStubs` therefore fills gaps from column A, while the
   batch `Sheet` rebases to `originCol`. **Why:** a single-pass scanner does
   not know the sheet's first populated column until the sheet is over, and
   retaining rows to find out is exactly what streaming exists to avoid. One
   documented asymmetry follows: `Cell.l` (hyperlinks) is batch-only.
   Comments are streamed.

8. **`sheetToJSON({ header: 1 })` returns dense rows.** Missing cells become
   `null` (or `defval`) instead of SheetJS's array holes; `header: 1`,
   `header: "A"`, `header: [...]` and `header: n` otherwise follow the
   SheetJS semantics exactly, including duplicate-label disambiguation.
   **Why:** a hole carries no information but changes `length` and iteration
   semantics for every consumer; the deviation is deliberate and documented.

9. **Input types are generous.** `read` accepts `Uint8Array`,
   `ArrayBuffer`, a path string, a `URL` (including `file:` URLs, via
   `fileURLToPath`), a `ReadableStream`, a Node readable stream and any
   async iterable of byte chunks. **Why:** callers arrive with whichever one
   they have - a fetch body, a buffer from a queue, a path from a CLI - and
   making them adapt first is friction with no benefit.

10. **One error hierarchy.** Everything thrown derives from `GridxcelError`:
    `BadZipError` and its subclass `ZipError` (container-level failures such
    as `NOT_ZIP`, `BAD_LOCAL_HEADER`, `CRC_MISMATCH`,
    `UNSUPPORTED_METHOD`, `INFLATE_FAILED`), `FormatError`, `SchemaError`,
    `NotImplementedError`, `StrictWarningError` and `PlatformError`
    (`NOT_SUPPORTED`, raised outside Node). All seven are exported from the root
    entry point and declared in `src/errors.d.ts`, and the Node and browser
    platform modules re-export the same class object per name. **Why:** a caller can
    wrap any call in a single `instanceof` check and get a stable
    machine-readable `code` for branching. The fuzz suite drives 456 corrupted
    inputs and observes 402 typed `GridxcelError` instances, 0 plain errors
    and 0 hangs.

11. **Every scan is bounded by the construct it scans.** Tag scanners find the
    end of a tag by walking it once with quote state; they never search for a
    delimiter with an unbounded `indexOf`. **Why:** an unbounded search is
    quadratic in the buffer size whenever the searched character is rare, and
    that is exactly what happened - `findTagEnd` looked for the next quote
    from every attribute-less tag in a worksheet containing none, and the reader
    ran 11x slower than it should have (3,558 ms to 310 ms once fixed).
    `test/scaling.test.mjs` enforces the rule by asserting parse time is flat
    across feed sizes. The same principle bounds the empty-row skip in the
    scanner. See [PERFORMANCE.md](./PERFORMANCE.md).

12. **Untrusted extents are bounded, not trusted.** A worksheet's declared
    dimension is attacker-controlled, so neither the reader nor the dense
    converters allocate proportionally to it. A row more than
    `EMPTY_ROW_SLACK` (65,536) rows beyond its neighbours is kept in
    `sheet.sparseRows` (a `Map`) behind a `SPARSE_ROW_GAP` warning while
    `sheet.rowCount` still reports the declared extent, and the writer emits
    `sparseRows` so a rewrite is lossless. **Why:** a one-kilobyte file
    declaring `<row r="40000000">` must not be able to allocate gigabytes;
    §13 handles the converter half of the same problem.

13. **Dense conversions are guarded.** `sheetToAOA`, `sheetToJSON` and
    `sheetToCSV` share `assertDenseRange`, which refuses a span above
    `maxRows` (default 262,144, about 16 MiB of dense output) with
    `RANGE_TOO_LARGE`, honouring an explicit `range` and an explicit
    `maxRows`. **Why:** the reader can bound what it *keeps*, but a converter
    asked to densify a declared extent would still allocate from untrusted
    metadata. Callers who pass an explicit range are taken at their word;
    callers who inherited the sheet's own extent are not, and the error message
    names both escapes (`range`/`maxRows`, or `eachRow`). Regression
    coverage lives in `test/hostile-input.test.mjs`.

14. **The style table is seeded on rewrite.** `createBookWriter` interns the
    source workbook's styles first and keeps a map from old xf index to new, so
    numeric `Cell.s` values read from a file still resolve after the cellXfs
    table has been deduplicated. **Why:** dedup is necessary to keep
    `styles.xml` small, but it renumbers indices, and a rewritten workbook
    whose cells point at the wrong style is a silent corruption.

---

## Verification

Nothing in this design is accepted on reasoning alone; each claim has a
mechanical check.

| Check | Command | What it proves |
| --- | --- | --- |
| Unit and integration tests | `npm test` | 262 tests: 258 pass, 4 skipped, 0 fail. |
| Documentation | `npm run check:docs` | Every documented symbol exists in `src/`, every doc example runs as a real ES module, every link resolves. |
| Inverse coverage | `node scripts/check-exports-doc.mjs` | Every public export across the eight entry points is named in the documentation. |
| Oracle interop | `node scripts/interop.mjs` | `status=pass`, 190 write checks, 37,293 read checks, 0 mismatches, `dateConvention=local`, against Python openpyxl 3.1.5 and XlsxWriter 3.2.9. |
| CRC corruption | `node scripts/check-crc.mjs` | A corrupted archive is rejected with `CRC_MISMATCH`. |
| Fuzzing | `node --test test/fuzz.test.mjs` | 456 cases: 402 typed errors, 0 plain errors, 54 recovered, 0 hangs. |
| Scaling | `node --test test/scaling.test.mjs` | Parse time is flat across feed sizes, so no scan is quadratic. |

The oracles are independent implementations; no gridxcel module is ever used to
check another gridxcel module. Test modules: 25. Fixtures: 16.

Performance, as last measured on `bench/fair-50000x8.xlsx` (50,000 rows x 8
columns, written by openpyxl so neither library produced it; same bytes, same
process, nine passes each the median of three): SheetJS CE 0.20.3 reads it in
1,072 ms; `readSync` in 315 ms, the async `read` in 281 ms and `eachRow` in
286 ms. That is a median speedup of about 3.4x, and the range across passes is
3.0x to 3.8x - quote the range, not the best single number, because one pass is
a sample rather than a result.

---

## Defect status at 0.1.0

Fixed and regression-covered: DUP-1 to DUP-8 (cell coercion, the two `utils`
namespaces, the two `styles.xml` parsers, the weaker escaper, the three
normalisation paths, the duplicated sheet-name rules, the unvalidated A1 decoder
and the repeated magic values), the crafted row index that caused unbounded
allocation (FS-1), the quadratic tag scanner (decision 11), and the quadratic
XML text collector found by the security review.

Still open, and deliberately recorded rather than hidden:

| Id | Defect | Consequence |
| --- | --- | --- |
| MINOR-4 | `_xHHHH_` is decoded in text nodes but not in attribute values. | A control character in a font name does not survive a round trip. |
| MINOR-5 | The style table converges after two round trips (7 -> 8 -> 9 -> 9 cellXfs). | Two bytes and one extra pass. |

See [LIMITATIONS.md](./LIMITATIONS.md) for the supported surface.

---

## Publishing surface

`package.json` `files` ships exactly `src`, `README.md`, `LICENSE` and
the published documents: [API.md](./API.md), [MIGRATION.md](./MIGRATION.md),
[LIMITATIONS.md](./LIMITATIONS.md), [BROWSER.md](./BROWSER.md) and
[PERFORMANCE.md](./PERFORMANCE.md). Tests, fixtures, benchmarks and this
architecture document stay out of the tarball. A consumer install was verified
against that exact file set, resolving the root entry point and the `compat`,
`zip` and `numfmt` subpaths.

Change control: an exported signature changes only together with
`src/index.d.ts`, the affected subpath declaration and [API.md](./API.md).
The two documentation checkers exist so that this cannot be forgotten.