# Running the gridxcel verification suite

Everything here is independent verification: the reference implementations are
python (openpyxl / XlsxWriter / zipfile / lxml), never another gridxcel module.
That is the point - a bug in the reader cannot make the test agree with itself.

## Quick start

    cd gridxcel

    npm test                          # node --test "test/**/*.test.mjs"
    node --test                       # bare form also runs the *.selftest.mjs files
    node --test test/interop.test.mjs # interop only
    node --test test/fuzz.test.mjs    # robustness only
    node scripts/interop.mjs          # master oracle harness, human table + JSON
    node --expose-gc scripts/bench.mjs --rows=200000 --cols=10
    node scripts/check-crc.mjs        # CRC regression check (exit 0 only if the gap is back)

**Note on `node --test <dir>`:** on this Node build (v24.13.0) a positional
directory argument is resolved as an entry point and fails with
`Cannot find module .../test`. `package.json` therefore uses an explicit glob
(`node --test "test/**/*.test.mjs"`), and bare `node --test` works too because it
auto-discovers. `npm test` runs the `*.test.mjs` files (137 tests); a bare
`node --test` in the repository root also picks up owner self-tests if any are
present.

The Python interpreter is found automatically, either bundled with the
environment or as an executable on PATH (openpyxl and XlsxWriter are the only
requirements);
override with `GRIDXCEL_PYTHON`. Suites that need it skip with a
`NOT YET AVAILABLE` reason when a module or the oracle is missing, so a partial
checkout still produces a green suite plus an explicit skip.

## What each file does

| File | Purpose |
| --- | --- |
| `scripts/oracle.py` | The oracle. `inspect` dumps one workbook (cells from the raw worksheet XML spine enriched by openpyxl, number formats, merges, geometry, zip CRCs via `testzip()`); `gen` writes the volume workbook with openpyxl; `selftest` is container-only. |
| `scripts/oracle-client.mjs` | Spawns the oracle with `child_process.execFileSync`, canonicalises both sides, and diffs them (values, types, formats, merges, geometry). |
| `scripts/specs.mjs` | Write-path specifications in the frozen Book/Cell shape plus the expected values derived from the same literals. |
| `scripts/golden.mjs` | Manifest checks against `test/fixtures/golden.json` (pre-recorded openpyxl output), for both the full read and the `eachRow` stream. |
| `scripts/zip-surgery.mjs` | Hand-rolled zip byte surgery (local headers, EOCD, CRC-32) used to build corrupt inputs without using src/zip.js. |
| `scripts/interop.mjs` | Master harness: (a) gridxcel writes, python verifies; (b) gridxcel reads the python fixtures, openpyxl diff; (c) one JSON summary + human table. Exit 0 pass, 1 mismatch, 2 nothing proven. |
| `test/interop.test.mjs` | The same checks as a node:test suite. |
| `test/fuzz.test.mjs` | 456 corrupted-input cases: truncations at many offsets, worksheet-XML byte flips, EOCD/central-directory damage, empty/1-byte/text/zero-filled input, valid-zip-not-a-workbook. |
| `scripts/bench.mjs` | Parse / stream-parse / write / CSV benchmarks with sampled peak heap and post-GC retained heap. |
| `scripts/check-crc.mjs` | CRC regression check: rebuilds a stored part with a stale CRC and asserts gridxcel rejects it. Exit 1 = rejected (fixed), 0 = silently accepted (gap is back), 2 = inconclusive. |

## Interpreting the harness

`node scripts/interop.mjs` prints a human table and then one machine-readable
line:

    #JSON# {"harness":"gridxcel-interop","status":"pass","write":{"checks":190,...},"read":{"checks":37293,...},"mismatches":[]}

* `status: "pass"` - every check ran and matched.
* `status: "fail"` - at least one mismatch; exit code 1. Each mismatch prints
  the sheet name, the cell reference and the raw `<c>` XML from the oracle.
* `status: "partial"` / `not-available` - a module or the oracle was missing;
  exit code 2, nothing was proven either way.
* Read mismatches of severity `warn` (number-format text differences from
  XML-only cells) are printed but do not fail the run.

## Benchmarks

`node --expose-gc scripts/bench.mjs --rows=200000 --cols=10` (the script
re-executes itself with `--expose-gc` if you forget). Input is generated once
by python/opencyxl into `bench/volume-<rows>x<cols>.xlsx` and reused. Results
are written to `bench/bench-<rows>x<cols>.json` and printed as a table plus a
`#JSON#` line.

Peak heap is sampled every 5 ms; the sampler cannot see inside a long
synchronous block, so the post-`global.gc()` retained column is the number to
trust for "does it fit in memory". Stage order matters: `eachRow` runs first
against a clean heap, then the full parse, then CSV, then the writers; each row
reports the excess over the baseline it started from.

## Current suite state (2026-10-18, final run)

    npm test           -> 137 tests, 133 pass, 0 fail, 4 skipped
                          137 tests, 133 pass, 0 fail, 4 skipped   (TZ=UTC)
                          137 tests, 133 pass, 0 fail, 4 skipped   (TZ=America/New_York)
                          137 tests, 133 pass, 0 fail, 4 skipped   (TZ=Asia/Tokyo)
    node scripts/interop.mjs -> pass, 190 write checks + 37,293 read checks, 0 mismatches
    node --test test/fuzz.test.mjs -> pass, 456 corruption cases, 402 typed GridxcelError,
                                     0 non-Error throws, 54 recovered, 0 hangs
    node scripts/check-crc.mjs -> pass, corruption rejected with ZipError / CRC_MISMATCH
    node examples/run-all.mjs  -> 9 of 9 examples pass
    node scripts/check-crc.mjs -> FIXED: ZipError [CRC_MISMATCH], exit 1

The 4 skips are fixtures python/openpyxl refuses to open (gap 4 below); the test
names are reported with the reason. The interop, fuzz, manifest and streaming
checks own every result in this file; other owners' suites are included in the
totals because `node --test` discovers all of test/.
## Current gaps (verified, not speculation)

Closed items are kept struck through so the finding and its resolution stay
visible. `docs/LIMITATIONS.md` is the user-facing version of this list.

1. ~~Stored (uncompressed) zip parts are not CRC-checked.~~ **CLOSED.** Every part
   is CRC-verified on read, buffered and streaming, and corruption raises
   `ZipError [CRC_MISMATCH]`; `ZipArchive({verifyCrc:false})` opts out.
   `node scripts/check-crc.mjs` rewrites one digit of a stored worksheet part and
   reports `FIXED: corruption rejected with a typed error (CRC_MISMATCH)`. The
   script stays as a regression check - exit 0 would mean the gap is back.
2. ~~Date convention on write is not in the frozen contract.~~ **CLOSED.** The
   convention is local wall clock and it is pinned in `docs/ARCHITECTURE.md`
   (Resolved decisions, item 1) and `docs/API.md`. Verified timezone-independent
   under TZ=UTC, TZ=America/New_York and TZ=Asia/Tokyo.
3. ~~`eachRow` does not expose the workbook number-format table.~~ **CLOSED.**
   Streamed cells carry `Cell.zf`, the resolved format code, so a streaming
   consumer never needs the workbook. Verified: 6 of 6 styled cells in
   `golden.xlsx` carry `zf` (sample `{ z: 0, zf: "General" }`).
4. **Fixtures openpyxl cannot read are skipped, not verified.**
   `reader-badcells.xlsx`, `reader-cut-sheet.xlsx`, `reader-cut-sst.xlsx` and
   `reader-nosheets.xlsx` are deliberately malformed; openpyxl refuses them, so
   they appear as `skip` in the harness and are exercised only by the fuzz
   suite. They are valid robustness inputs but have no oracle.
5. **Rendered text (`Cell.w`) has no independent formatter oracle. Partly closed.**
   `isDateFormat` is verified against openpyxl's own classifier (63/63 cases in
   `test/numfmt.test.mjs`) and the renderer against a curated golden set (48
   cases). openpyxl does not render numbers, so `Cell.w` for exotic format codes
   is asserted from documented Excel semantics rather than measured from a second
   implementation. The common path is exact: `golden.xlsx` D1 with
   `cellText: true` yields `"2024-03-15 14:30:45"`.
6. **Styles beyond number formats have no interop spec. Partly closed.**
   Fonts, fills, borders and alignment survive a full write -> read round trip
   (verified: bold+italic+colour font, solid fill with an explicit ARGB colour,
   thin top border and centred alignment all come back intact). They are still not
   written and re-read by *python*, so the OOXML shape of those records is not
   checked by an independent implementation.
7. **The SheetJS-shaped shim is unit-tested against openpyxl but not interop-tested.**
   `src/compat.js` is a reachable entry point (`gridxcel/compat`, plus `./csv`,
   `./json`, `./xml`, `./numfmt`, `./zip` subpaths and `src/compat.d.ts`), and its
   tests assert `SheetNames`, the first data row and the CSV line count against
   openpyxl on `test/fixtures/pandas.xlsx`. What is missing is a
   `scripts/interop.mjs` round trip through `XLSX_like.readFile`/`writeFile` - the
   harness covers the native API only, so the shim's own write path has not been
   inspected by python.
8. **CSV export is not streaming.** `sheetToCSV` builds the whole string
   (99 kB for a 6000-row sheet, 19 MB for 200k rows, measured); there is no
   incremental CSV writer.
9. **A CPU-spin hang cannot be interrupted in-process.** The fuzz suite races
   every case against a `GRIDXCEL_FUZZ_TIMEOUT` (default 10 s) timer, but a
   blocked event loop is only catchable from outside: use
   `node --test --test-timeout=60000` in CI.
10. **The 2000x8 volume file for the read diff is generated on first use** by
    python (cached in `bench/`). The 200k-row benchmark input is generated the
    same way and takes a couple of minutes the first time.
11. **Fuzz coverage is byte-level, not semantic.** It proves "no non-Error, no
    hang, damage reported" but does not prove that a truncated file yields the
    *correct subset* of rows; that is only checked for intact files.
12. ~~`ZipError` is not part of the `GridxcelError` hierarchy.~~ **CLOSED.**
    `src/zip.js` now declares `export class ZipError extends BadZipError`, so every
    container failure satisfies `instanceof GridxcelError`. The fuzz suite went
    from 398 typed with 4 plain `Error` to **402 typed, 0 plain Error**, 0 hangs.
