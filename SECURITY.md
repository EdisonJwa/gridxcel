# Security

gridxcel reads files you did not write. Parsing untrusted workbooks - a service
accepting uploads, a pipeline ingesting a stranger's `.xlsx` - is a supported use
case, not an accident, and the properties in this file are the reason.

## Supported versions

| Version | Supported |
| --- | --- |
| 0.1.0 (pre-1.0) | Yes - the only version, and not yet published to npm |

There is no older line to patch and no LTS branch: fixes are made on the default
branch and ship in the next release. 0.1.0 is pre-1.0, so the public API may
still change.

## Reporting a vulnerability

Report privately through GitHub's advisory form on the repository:
https://github.com/EdisonJwa/gridxcel/security/advisories/new

Please do not open a public issue, and do not ask in a public issue whether
something is a vulnerability. If the advisory form is unavailable to you, open
an issue that says only that you have a security report and asks for a private
channel.

A useful report contains:

- the gridxcel version, the Node.js version and the platform;
- what you did, what happened, and what you expected instead;
- the smallest input that shows it - a few kB of crafted `.xlsx` is more useful
  than a description of one, and a workbook can be rebuilt with fake values;
- the typed error code or warning code if one was raised, and the stack if the
  process died rather than throwing.

This is a volunteer project with no support contract, so no response time is
promised. Fixes are described in the advisory.

## What gridxcel defends against

| Property | How it behaves |
| --- | --- |
| Corruption is detected before the bytes are used | Every ZIP part is CRC-32 verified on read, buffered and streaming; a mismatch raises `ZipError` / `CRC_MISMATCH`. `ZipArchive({ verifyCrc: false })` opts out. |
| Bytes that are not a workbook fail loudly | Random bytes, an empty buffer and plain text raise `FormatError` / `NOT_A_ZIP` rather than returning an empty workbook. |
| Malformed input fails with a typed error | Every failure is a subclass of `GridxcelError` carrying a stable `code`. The fuzz suite runs 456 corrupted inputs: 0 untyped throws, 0 hangs. |
| A crafted row index cannot allocate without bound | FS-1 is fixed. Rows more than 65,536 rows past the others are kept in `sheet.sparseRows` with a `SPARSE_ROW_GAP` warning, and dense conversions (`sheetToAOA`, `sheetToCSV`) refuse an oversized span with `RANGE_TOO_LARGE` unless `range` or `maxRows` allows it. A 1 kB file that used to kill a 3 GB heap now costs about 0.2 MiB. Regression: `test/hostile-input.test.mjs`. |
| No XML entity expansion | Custom entities are never expanded; a billion-laughs DOCTYPE resolves the entity to its literal name in milliseconds instead of expanding it. |
| No stack overflow from deep XML | The parser is iterative; 20,000 levels of nesting parse without error. |
| No prototype pollution | A cell reference of ```__proto__``` is rejected with `BAD_CELL_REF` and `Object.prototype` is left alone. |
| No path traversal | Hostile part names such as ```../../evil.txt``` are carried as names only. There is no extract-to-disk API, so no file is ever written from a workbook. |
| Little to supply-chain | Zero runtime and zero dev dependencies. On Node, five builtins (`node:fs`, `node:fs/promises`, `node:stream`, `node:url`, `node:zlib`) are imported, all confined to `src/platform.js`; on every other platform the library uses Web globals (`TextDecoder`, `TextEncoder`, `ReadableStream`, `CompressionStream`, `DecompressionStream`, `Blob`, `URL`, `btoa`, `atob`). Nothing is fetched at runtime on either platform, and a browser build has no runtime dependency either. |
| Unrepresentable values are refused | Writing `NaN` raises `GridxcelError` / `BAD_VALUE` instead of emitting invalid XML. |
| Concurrent reads are independent | 32 simultaneous reads resolve correctly; there is no shared mutable state between them. |

## The npm install warnings

npm's static analysis flags two things when this package is installed. Both are
accurate and both are expected; this section exists so nobody has to take that on
trust.

**Filesystem access.** `readFile`, `readSync` and `writeFile` are part of the
public API: they open the path you pass them and nothing else. There is no
directory walking, no scanning for files you did not name, no reading of
environment or credential paths, and no writes outside the path you asked for. If
you never call a path-taking function, no file is ever opened. In a browser build
the whole filesystem layer is replaced by `src/platform.browser.js`, which throws
`PlatformError [NOT_SUPPORTED]` instead of touching a disk.

**URL strings.** All 18 URL-shaped strings in the shipped code are XML namespace
identifiers required by the OOXML specification - `schemas.openxmlformats.org`,
`purl.org`, `w3.org`. They are string constants written into document parts, never
resolved, and never fetched. The package contains no fetch, no
XMLHttpRequest, no Node http or https import, and no socket of any kind: a
workbook never leaves the process unless your own code sends it somewhere. You can
confirm this in one command:

```sh
grep -rE 'fetch|XMLHttpRequest|node:https?|WebSocket' node_modules/gridxcel/src/
```

It prints nothing.

One change was made in response to the filesystem warning: `scripts/run-tests.mjs`
is no longer included in the published tarball. It is the maintainer's test runner,
it has no use in an installed package, and it was the only shipped file outside
`src/` that did anything beyond documentation. The tarball went from 50 files to
49 and now contains no JavaScript outside `src/`.

## What is not covered

- **Size is a denial-of-service vector you have to manage.** A full parse
  retains heap proportional to the cell count - about 165 MiB for 2 million
  cells - so bound what you accept, or stream with `eachRow`, which holds one
  row at a time. Formats, formulas and merged ranges have limits of their own;
  see [docs/LIMITATIONS.md](docs/LIMITATIONS.md).
- **Unsupported formats are rejected, not parsed.** XLSB, XLS (BIFF), ODS and the
  rest of the list in [docs/LIMITATIONS.md](docs/LIMITATIONS.md) are out of scope,
  so there is no parser there to attack - and none of the properties above are
  claimed for a format the library does not read.
- **The duplication defects were correctness bugs, not memory-safety bugs, and
  they are fixed.** The DUP-1 to DUP-8 defects could make the same input accepted,
  rejected or silently dropped depending on which entry point it reached, and
  DUP-4's weaker escaper could leave control characters raw in a file Excel then
  refuses; each is now collapsed onto one implementation with a regression test.
  Nothing in the library duplicates a helper any more. Remaining - `freeSheetName` existed in
  three agreeing copies (`src/api.js` canonical, plus `src/index.js` and
  `src/compat.js`) - which is maintenance risk, not a security property. If your
  own pipeline validates or escapes data, keep doing it.
- **The browser is a second platform, and a smaller surface.** The byte-level
  API runs in a browser through `src/platform.browser.js`; see
  [docs/BROWSER.md](docs/BROWSER.md). The threat model gains a host, not a
  capability: a browser has no filesystem access and no synchronous entry point
  (those raise `PlatformError` / `NOT_SUPPORTED`), and there is still no network
  request and no runtime dependency there. Parsing untrusted bytes is the same
  job under the same limits - bound what you accept, or stream with `eachRow` -
  but the decoder is the Web `DecompressionStream` rather than `node:zlib`, and
  engines older than Chrome 103 / Firefox 113 / Safari 16.4 report
  `NOT_SUPPORTED` instead of decoding. Nothing in the library writes to disk on
  either platform.
