# Changelog

All notable changes to gridxcel are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

0.1.0 has not been published to npm, so nothing in this file is released yet.

## [Unreleased]

### Added
- **Random fuzzing in the gate, and a nightly job.** Every suite in `test/` uses a
  fixed corpus, so it can only catch a regression. `scripts/fuzz-random.mjs` drives
  ten parsers with generated input for a time budget under a seed that reproduces
  any failure exactly; a 10-second run is 2.85 million cases. It runs for 5 seconds
  in `npm run check` and for 60 seconds per target nightly on Node 18 and 22.
- **The browser claim is now tested in three engines.** Chromium, Firefox and
  WebKit each run read, write, `eachRow`, `createCSVStream` and the synchronous
  path through a real esbuild bundle. Chromium alone cannot show an engine
  difference in the Web compression streams that replace `node:zlib`.
- **`createCSVStream`, a streaming CSV writer.** `sheetToCSV` is O(output) in
  retained memory - it builds every line and joins them - which on a 200,000 x 10
  sheet costs 14.3 MB for the string. The stream yields the same bytes a row at a
  time, retains about 0.1 MB, and is roughly 20% faster because it never allocates
  the line array. Chunks break on row boundaries rather than byte boundaries: a
  quoted field may legally contain a newline, and splitting there would force a
  consumer to reassemble before parsing. Both forms walk one shared generator, so
  they cannot drift. Verified byte-identical across eleven option sets and six
  chunk sizes; covered by [test/csv-stream.test.mjs](test/csv-stream.test.mjs).
- **`gridxcel/formats`: SYLK and DIF reading and writing.** Both are legacy text
  formats that systems predating XLSX still export, and both are implemented
  against their specifications rather than against this library's own writer - the
  tests use each specification's worked example. SYLK keeps a cell's type
  (number, string, TRUE/FALSE, NA) and its rule that an omitted `Y` or `X` reuses
  the previous coordinate. DIF follows the 1981 grammar: three-line header chunks,
  two-line data chunks, with a numeric value in the type pair rather than on the
  keyword line. Also reachable transparently: `gridxcel/compat` sniffs the format
  from the content, recognising SYLK by its `ID` record and DIF by `TABLE`. An
  explicit `bookType` overrides the sniff, which matters for a CSV whose first
  cell is literally "ID". Neither format has a date type, so a date is written as
  its serial and reads back as a number. Covered by
  [test/formats.test.mjs](test/formats.test.mjs).
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
- Eight entry points with hand-written declarations: `gridxcel` (69 named
  exports plus a default object), `gridxcel/compat` (16 plus a default),
  `gridxcel/csv` (8), `gridxcel/json` (19), `gridxcel/xml` (11),
  `gridxcel/numfmt` (13), `gridxcel/zip` (13) and `gridxcel/formats` (4) -
  153 export slots and 124 distinct public symbols in total.
- Streaming read through `eachRow`, which holds one row at a time, and streaming
  write through `WorkbookWriter` and `SheetWriter`.
- ZIP container support: stored and deflated parts, Zip64, and CRC-32
  verification of every part on read.
- A SheetJS-shaped compatibility layer for migration, with CSV/DSV, AOA/JSON and
  an ECMA-376 number-format engine underneath the native API.
- One typed error hierarchy rooted at `GridxcelError`; every failure carries a
  stable machine-readable `code`.

### Fixed

- **A sparse row index made the writer walk two billion holes.** `addRow` looped
  `for (j = 0; j < cells.length; j++)` on a row that is by contract a sparse
  array. Writing to index 2^31-1 made length 2^31: 38 seconds for one row, then a
  column reference no spreadsheet can open. Index 1e9 took 16 seconds. Bounded on
  the way in with `RANGE_TOO_LARGE`; a row using columns A and Z, and the highest
  legal column, are unchanged.
- **Three characters of number-format code aborted the process.** `*"` left the
  render loop spinning forever: `_.` and `*.` consume "this character and the
  next", and they ran after quoted runs had been replaced by a NUL-delimited
  marker, so they ate half of one and left a lone NUL. The loop then did
  `i = indexOf(NUL, i+1)` (which is -1 when unpaired) and its own `i++` put it
  back on the same character, appending nothing forever until the heap was gone -
  exit 134, four gigabytes, not catchable. Directives now resolve in Excel's
  precedence order before any marker exists, and an unpaired marker is skipped
  rather than rewound onto.
- **A fraction denominator drove a billion-iteration loop.** `maxDen` was
  `10^len - 1` straight from the code, so `# ?/?????????` meant 999,999,999
  iterations per cell. Capped at 999, which is Excel's own behaviour: 9 ms -> 0 ms.
- **A long format code cost seconds per cell.** 100,000 numeric placeholders took
  7.5 s. Excel caps a cell at 32,767 characters, so a longer code is refused with
  `BAD_FORMAT`.
- **A code with more than 100 decimal places threw a raw `RangeError`.** The count
  came from the code and went straight into `toFixed`, which throws past 100 and
  is not a `GridxcelError` - breaking the promise that every failure is typed.
  Clamped to 30 (Excel's limit); scientific notation had the same hole.
- **A lone surrogate was silently replaced with U+FFFD.** A JavaScript string can
  hold half of a UTF-16 pair; it has no UTF-8 encoding, and `TextEncoder`
  substitutes the replacement character without complaint - so `"a\\ud800b"` was
  written and read back as `"a\\ufffdb"`, a one-way door with no error at any point.
  Both XML encoders now escape a lone surrogate with the OOXML `_xHHHH_` escape,
  which already existed for control characters and has the right semantics for a
  code unit that stands alone. A well-formed pair is left intact.
- **`<sz val="NaN"/>` was written for a non-finite font size.** A `font.sz` of NaN
  or Infinity went straight into an attribute, producing a package Excel has to
  repair. Sizes are now validated as finite and within 1-409 points (`BAD_STYLE`).
- **A freeze pane outside the grid emitted a cell reference past the sheet.**
  `freeze: { row: 1e9 }` produced `topLeftCell="CFDGSXM1000000001"`. Now
  `RANGE_TOO_LARGE`.
- **An invalid autofilter was written through verbatim.** `autoFilter: "not a range"`
  emitted `<autoFilter ref="not a range"/>`, which is not valid XML. The ref is now
  validated and normalised, or `BAD_RANGE`.
- **A column width that was not a finite positive number was silently dropped.**
  `width: NaN`, `Infinity` or a negative number left the caller with a default
  column and no signal. Now `BAD_VALUE`; an over-wide one is clamped to Excel's 255.
- **Column options with no index were silently dropped, including the documented
  form.** `{ columns: [{ width: 8 }, { width: 24 }] }` - which COOKBOOK tells the
  reader to write - set neither width, because an entry with no `min`/`index` fell
  through the index check. Entry *i* is now column *i*, as the docs always said.
- **A malformed `<mergeCell ref>` was passed through verbatim.** `"NOTAREF"` reached
  `sheet.merges`, where a consumer would hand it to a range library and fail far
  from the file that caused it. Now dropped with a `BAD_MERGE` warning - the
  reader tolerates a bad file where the writer refuses to produce one.
- **`import.meta.dirname` broke every read on Node 18.** The integration suite
  used it, but it landed in Node 20.11 and the package declares `engines >= 18`,
  so on 18 the fixture path became "undefined/fixtures/golden.xlsx" and every read
  threw `ENOENT`. Now derived from `import.meta.url`.
- **A test asserted a capability instead of detecting it.**
  `test/platform.test.mjs` claimed `hasWebDeflate() === true`. Node 18 has
  `CompressionStream`/`DecompressionStream` but rejects `deflate-raw`, which
  arrived in Node 20 - so the feature detection was correct and the test was
  wrong. It now asserts that the answer matches the host, and skips the round trip
  with a diagnostic when the host cannot support it.
- **The integration harness could not run on a fresh clone.** Its suites import
  the published specifier "gridxcel" deliberately - it is what a consumer writes,
  so it checks the exports map, the types and the `browser` field - which resolved
  through a self-link that existed only on the development machine. On a clone
  seven suites failed with "Cannot find package 'gridxcel'". `integration/link-self.mjs`
  creates the link idempotently and `run-all.mjs` calls it first; verified by
  hiding the parent node_modules directory entirely.
- **The browser CI leg could pass while testing nothing.** The suite exited 0 with
  a SKIP when no browser was present, and the leg did not install one.
  GRIDXCEL_REQUIRE_BROWSER now makes absence a failure in CI, and the leg
  installs Chromium.
- **The exports checker could not see a new entry point.** It used a hand-written
  table of specifiers, so adding `gridxcel/formats` to the package did not add it
  to the check and its four exports were never verified as documented. The table
  is now derived from `package.json`.
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
  `bench/fair-50000x8.xlsx` in 315 ms against 3,558 ms before the scanner fix,
  and the asynchronous `read` in 281 ms. Measurements and reproduction commands
  are in [docs/PERFORMANCE.md](docs/PERFORMANCE.md).
- **A fast path for attribute-free tags in the worksheet scanner.** 79% of the
  2.1 million tags in a 50,000-row sheet carry no attributes and the scanner only
  wants their local name, but `parseTag` allocated a result object for every one.
  `tagLocalIfNoAttrs` answers both questions in one pass and returns `null` when
  the full parser is needed; an unprefixed name returns the input string, so the
  common case allocates nothing. Measured at -16.4% on tag parsing in isolation,
  which is a few percent end to end because the tag layer is about a quarter of
  the parse. Verified against 2,116,398 real tags with zero disagreements.
- **The published speed figures were corrected, in every document.** README,
  `PERFORMANCE.md` and `COMPARISON.md` all still claimed 519 ms and "1.4x to 2.7x,
  median 1.9x" - figures from *before* the quadratic scanner was fixed. The
  measured result is 315 ms and 3.0x to 3.8x, median about 3.4x. The library was
  understating itself by nearly half in every document a prospective user reads.
- **The README was rewritten for a reader who has not chosen a library yet.** It
  opened by comparing itself to SheetJS and listing what it cannot do, which
  answers a question a new visitor has not asked. It now leads with bounded
  memory, measured speed and typed failure, then a capability table, then an
  honest comparison that says plainly when to use SheetJS instead.


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
