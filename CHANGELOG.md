# Changelog

All notable changes to gridxcel are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

0.1.0 has not been published to npm, so nothing in this file is released yet.

## [Unreleased]

### Added
- **Browser support, with no build step and no polyfill.** The byte-level API -
  `read`, `eachRow`, `sheetToJSON`, `sheetToCSV`, `write`, `WorkbookWriter` -
  runs anywhere with the Web platform. Compression is `node:zlib` on Node and
  `CompressionStream`/`DecompressionStream` elsewhere, behind one seam:
  `src/platform.js` for Node, `src/platform.browser.js` for everything else,
  swapped by the package `browser` field. The browser file contains no `node:`
  reference at all. `src/*.js` is still exactly what ships - there is no browser
  bundle to fall out of sync. Importing a single function tree-shakes to 21 kB
  against 251 kB for the whole surface. Verified by a real esbuild IIFE bundle
  and a real Vite production build, each **executed** in a sandbox with no
  `Buffer`, `process`, `require` or module loader. Compression is not part of
  that claim in the writer: the Web streams are asynchronous and the synchronous
  writer cannot await them, so a browser write stores its parts instead of
  deflating them - a valid archive, larger than one written in Node.
- `ZipEntry.bytesAsync()` and `textAsync()`, the async twins of `bytes()` and
  `text()` for platforms without synchronous inflate.
- `src/errors.d.ts`, which was missing: `gridxcel/zip`'s declarations imported a
  module with no declaration file, so every consumer saw the whole error
  hierarchy as `any`.
- `PlatformError`, a `GridxcelError` subclass with code `NOT_SUPPORTED`, raised
  by the synchronous API and by the path-based helpers outside Node, naming the
  async alternative. It is exported from `gridxcel` (named and on the default
  object) exactly like its six siblings, and it is one class object shared by the
  Node and browser platform modules.
- [docs/BROWSER.md](docs/BROWSER.md), and `integration/` - a test harness with
  its own dev tooling (esbuild, Vite, React, Preact, TypeScript) that is never
  published.


- Read and write of XLSX workbooks, asynchronous and synchronous, from bytes, a
  path, a URL, a `ReadableStream` or an async iterable.
- Seven entry points with hand-written declarations: `gridxcel` (68 named
  exports plus a default object), `gridxcel/compat` (16 plus a default),
  `gridxcel/csv` (7), `gridxcel/json` (19), `gridxcel/xml` (9),
  `gridxcel/numfmt` (13) and `gridxcel/zip` (13) - 145 export slots and
  117 distinct public symbols in total.
- Streaming read through `eachRow`, which holds one row at a time, and streaming
  write through `WorkbookWriter` and `SheetWriter`.
- ZIP container support: stored and deflated parts, Zip64, and CRC-32
  verification of every part on read.
- A SheetJS-shaped compatibility layer for migration, with CSV/DSV, AOA/JSON and
  an ECMA-376 number-format engine underneath the native API.
- One typed error hierarchy rooted at `GridxcelError`; every failure carries a
  stable machine-readable `code`.

### Fixed


- **`XmlParser` scaled quadratically, and this one was reachable.** An incomplete
  tag, comment, CDATA section or processing instruction was rescanned from its
  own start on every chunk. 8 MB through 16 KB chunks: a long tag cost 5,273 ms
  and is now 31.6 ms, with scaling per 4x of input down from 14.8x to 3.4x, and
  the same treatment applied to comments (1,015 -> 3.9 ms), CDATA (1,033 ->
  3.6 ms), processing instructions (1,076 -> 3.5 ms) and text (1,228 ->
  6.5 ms). Unlike the earlier scanner defects this was reachable through the
  public `gridxcel/xml` entry point, because `parseChunks` accepts any async
  iterable. The non-obvious half is recorded in the file: making the scan
  resumable is not enough if the pieces are still appended to one growing
  string, because V8 re-flattens the cons string on every chunk.
- **The published declarations did not compile.** `src/index.d.ts` exported a
  type literal where a value was required (`export default { VERSION: string;
  ... }`), which every type-checking consumer saw as a syntax error;
  `mergeBooks`, `renameSheet` and `countCells` were exported at runtime and
  declared nowhere; and the declarations referenced `NodeJS.ReadableStream` and
  `Buffer`, so they required `@types/node` even for a browser consumer. None of
  this was caught, because the smoke test regex-scanned the file and the docs
  checker loads only `src/*.js`. A consumer now compiles with
  `skipLibCheck: false` and no `@types/node`, and `integration/types.test.mjs`
  keeps it that way. `ByteStream` replaces `NodeJS.ReadableStream`.
- **An unnamed sheet collided instead of scanning.** `WorkbookWriter.addSheet()`
  named an unnamed sheet `"Sheet" + (sheets.length + 1)`, so a book that already
  contained that name threw `BAD_SHEET_NAME` on a call that should succeed -
  while the same call through either `utils` surface worked. There were two other
  implementations of the same rule; `freeSheetName` in `src/api.js` is now the
  only one.
- **`ReadOptions.headerRow` was published and read by nothing.** A TypeScript
  caller who used the documented option silently got row 1 as the header. The
  field is removed from both declaration files; `header` already covers it.
- **The style table needed two round trips to converge.** `borderKey` keyed the
  reader's blank border apart from the writer's default, though `borderElement`
  emits identical XML for both, so a byte-identical second border was interned and
  its id produced a new `xf` on the next write (8 -> 9 -> 9). A new
  `edgeIsBlank` predicate, used by both the key and the emitter, makes three
  consecutive round trips byte-identical from the first.
- **`_xHHHH_` escapes were decoded in text but not in attribute values**, so a
  control character in a font name or a number-format code did not survive a
  round trip. Attribute values now decode exactly as text does, including the
  `_x005F_` rule for a literal escape, and the encoders write a literal
  escape as a backslash-free escaped form so it survives too: a literal
  underscore-x sequence is written with the `_x005F_` prefix in front of it.
- **`gridxcel/compat` required an `await preload()` before any synchronous
  call.** A top-level `await` in the entry point also made an IIFE
  (`<script>`) build impossible - esbuild refused it outright. The reader and
  writer are now imported statically, so there is no top-level await, no cycle,
  and the synchronous entry points work immediately, which is the whole point of
  a drop-in replacement. `available()` reports a settled `{read: true, write:
  true}` and `preload()` is a documented no-op; both stay exported.

### Changed

- The synchronous reader is no longer the slow path. `readWorkbookSync` reads
  `bench/fair-50000x8.xlsx` in 310 ms against 3,558 ms before the scanner fix,
  within about 5% of the asynchronous path. Measurements and reproduction
  commands are in [docs/PERFORMANCE.md](docs/PERFORMANCE.md).


- **FS-1** (high): a worksheet declaring a row far beyond its data caused
  unbounded allocation in every densifying helper, so a 1 kB file could kill the
  process with an out-of-memory abort instead of raising a catchable error. Rows
  more than 65,536 rows past the others are now kept in `sheet.sparseRows` with a
  `SPARSE_ROW_GAP` warning, the declared extent is still reported on
  `sheet.rowCount`, dense conversions share one `assertDenseRange` guard that
  refuses an oversized span with a typed `RANGE_TOO_LARGE` error (`range` and
  `maxRows` are the way through), and the writer emits `sparseRows` so a rewrite
  loses nothing. Covered by
  [test/hostile-input.test.mjs](test/hostile-input.test.mjs).
- **Quadratic tag scanner** (high): `findTagEnd` searched for a quote to the end
  of the buffer, which made it quadratic in the buffer size; the synchronous
  reader feeds 256 KiB slices, so it paid the worst of it and was several times
  slower than the streaming path on identical input. The scan is now a single
  quote-aware walk bounded by the tag - the algorithm `xml.js` already used - and
  `readWorkbookSync` went from 3,558 ms to 519 ms on `bench/fair-50000x8.xlsx`,
  bringing it to parity with the asynchronous path (520 ms) from 6.6x behind it. [test/scaling.test.mjs](test/scaling.test.mjs) constrains the
  ratio; the measurements are in [docs/PERFORMANCE.md](docs/PERFORMANCE.md).
- Stored (uncompressed) ZIP parts are CRC-verified too; a flipped byte raises
  `ZipError` / `CRC_MISMATCH` rather than being used.
- `ZipError` now extends `BadZipError`, so every container failure is a
  `GridxcelError` and one `instanceof` covers the whole taxonomy.

- **Duplication removed (DUP-1 to DUP-8).** Cell coercion, sheet-name rules, the
  A1 decoder and the spreadsheet limits each have one implementation now;
  `styles.xml` has one parser serving both the reader and the writer; and the
  writer no longer silently drops a cell it cannot interpret. The divergences are
  resolved rather than hidden: coercing `{a: 1}` throws `BAD_CELL` on every path,
  and `format_cell` answers identically from both entry points.
- **CDATA sections are decoded.** A text section written as `<![CDATA[...]]>` was
  returned with its delimiters included, because the marker was looked for at the
  wrong offset. Both the sync and streaming readers had the same bug.

### Security

Six findings from an independent review of parsing untrusted workbooks were
fixed:

- **Quadratic scanner (high, denial of service).** Two paths accumulated
  quadratically. A single long text node was re-concatenated and rescanned on
  every chunk: a 104 kB upload holding one 100 MB inline string cost 116 s of CPU
  and 1.8 GB of heap, and 8 MB took 1.2 s. A single long ATTRIBUTE value hit the
  tag scanner the same way: 16 MB cost 11.9 s from a 17.7 kB upload, growing
  ~3.9x per doubling. Text is now buffered and joined once when the close tag
  arrives, and a half-read tag is parked with its quote state and resumed rather
  than rescanned. 8 MB of text: 1,211 -> 51 ms. 16 MB attribute: 11,912 -> 81 ms,
  and both curves are linear.
- **Sheet-declaration flooding (medium, denial of service).** A 794 kB workbook
  declaring 100,000 sheets cost 5.9 s and 321 MB. `MAX_SHEETS` caps assembly at
  4,096 with a `TOO_MANY_SHEETS` warning; the same file now costs 115 ms.
- **Order-dependent sparse-row guard (medium).** The FS-1 guard treated the same
  two cells as a normal sheet or a hostile one depending on which appeared first.
  Rows are ordered before the gap check, so the verdict is a property of the data.
- **Inflate failures escaped the error contract (medium).** A corrupted deflate
  stream threw a raw zlib `Error` instead of a `GridxcelError`, so a caller
  checking `instanceof` crashed on a malformed upload. Now `ZipError` /
  `INFLATE_FAILED`, with the zlib code in `details` and the original on `cause`.
- **CSV injection guard gap (medium).** A leading tab, CR or LF bypassed the
  formula guard; all three are neutralised now, and `escapeFormulae: false`
  still opts out.
- **Local-header length overflow (low).** Same code path as the inflate finding.

Verified safe by the same review, with the tests that failed to break them: path
traversal on write, prototype pollution, ReDoS, XML entity expansion and XXE,
deep nesting, malformed XML, the filesystem and process surface, and error-path
information leaks.

### Known issues

Three of the duplication defects were correctness bugs wearing a DRY costume:
the duplicated helpers had diverged, so which entry point you called changed the
answer. They are recorded here because a changelog is a history of what was wrong,
not only of what is now right.

- **DUP-1** (high) - cell coercion exists three times and the copies disagree:
  the same object becomes a string through `json.js`, throws through `api.js`, and is
  silently dropped by the writer.
- **DUP-2** (high) - two `utils` namespaces with the same method names and
  different behaviour, so migrated SheetJS code depends on which import it uses.
- **DUP-3** (high) - two `styles.xml` parsers that disagree on shape: `numFmt` is a
  numeric id in one and a resolved code in the other.
- **DUP-4** to **DUP-8** - a weaker attribute escaper that can leave control
  characters raw, three book and sheet normalisation paths, duplicated
  sheet-name rules, an unvalidated A1 decoder, and repeated magic values.

The review lists every finding, including the ones not repeated here. 0.1.0 is
pre-1.0: the public API may still change before 1.0.

## [0.1.0] - unreleased

The first release, not yet published to npm. It contains everything listed under
Unreleased above, and there is no earlier version to upgrade from.
