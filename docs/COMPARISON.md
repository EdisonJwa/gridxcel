# Comparison with other libraries

gridxcel reads and writes XLSX, CSV, SYLK and DIF. It does four formats well
rather than fifteen adequately, and it optimises for three things: bounded
memory, speed on large files, and failures you can branch on.

This page says plainly where the alternatives win. That is more useful to you
than a table where every row has the same winner.

Measurements below were taken against **SheetJS Community Edition 0.20.3** on the
same bytes, in the same process, on the same machine. Every figure has a command
that reproduces it, and the method is in [PERFORMANCE.md](PERFORMANCE.md).

## Choose another library if

Being direct about this first saves you an evaluation:

* **You need XLSB, XLS (BIFF), ODS/FODS, NUMBERS, DBF, PRN, Lotus or HTML
  tables.** Not implemented. SheetJS reads all of them - see its
  [format table](https://docs.sheetjs.com/docs/miscellany/formats/).
* **You need charts, images, pivot tables, macros, conditional formatting or data
  validation** preserved or written. None are.
* **You need formula calculation.** A formula is text with an optional cached
  result; nothing is recalculated.
* **You want one call that streams XLSX out to a socket.** `WorkbookWriter`
  bounds memory while rows are produced, but `finish()` returns the whole
  archive. CSV does stream out (`createCSVStream`).
* **You need CommonJS.** ESM only.
* **You want a decade of field reports behind every edge case.** SheetJS has that
  history. This library has a large test suite and a hostile-input fuzzer, which
  is not the same thing.

## Where gridxcel wins

### Bad input fails loudly

This is the difference that matters most in production. Feed both libraries bytes
that are not a workbook:

| Input | SheetJS 0.20.3 | gridxcel |
| --- | --- | --- |
| 23 random bytes | returns a workbook with an empty `Sheet1` | throws `FormatError` `NOT_A_ZIP` |
| A zip with no workbook part | returns an empty workbook | throws `FormatError` `NOT_A_WORKBOOK` |
| A truncated worksheet XML | returns a partial sheet, silently | returns what parsed, plus a `BAD_XML` warning |
| A part whose CRC does not match | read past | throws `CRC_MISMATCH` |

A plausible-looking empty result is the worst failure mode for a data pipeline:
it propagates. A typed error stops at the boundary where you can act on it.

```js
try {
  const book = await read(bytes);
} catch (error) {
  if (error.code === "NOT_A_ZIP") {
    // not a spreadsheet at all - reject the upload
  } else if (error.code === "NOT_A_WORKBOOK") {
    // a zip, but not an xlsx - maybe a .docx or a .zip of something else
  }
}
```

### Bounded memory

On a 200,000 x 10 workbook:

| | Retained | Peak heap |
| --- | --- | --- |
| `eachRow` (streaming read) | **5.9 MiB** | 13.0 MiB |
| A full parse | 164.9 MiB | 224.9 MiB |
| `createCSVStream` (streaming write) | **0.1 MiB** | one row |
| `WorkbookWriter` (streaming write) | **5.8 MiB** | 375.9 MiB |

SheetJS's `read` materialises the workbook; streaming requires a different
mental model and it does not offer the same ceiling on retained memory.

### Speed

On a 50,000 x 8 workbook written by openpyxl, both libraries reading the same
400,008 cells, nine passes, each the median of three:

| Operation | Median | Versus SheetJS |
| --- | --- | --- |
| SheetJS `read` | 1,072 ms | - |
| gridxcel `readSync` | 315 ms | 3.20x (range 2.99x - 3.32x) |
| gridxcel `read` | 281 ms | 3.50x (range 3.40x - 3.75x) |
| gridxcel `eachRow` | 286 ms | 3.41x (range 3.31x - 3.79x) |

Publish the range, not the best sample. A single pass on a shared machine moves
by a factor of two; three earlier revisions of this document quoted single passes
and each was wrong in a different direction. Reproduce with
`npm run bench:scaling`.

### Types that hold up

| | gridxcel | SheetJS 0.20.3 |
| --- | --- | --- |
| Declarations | Hand-written for all eight entry points | Bundled, largely `any`-tolerant |
| `any` in the public surface | none in the documented API | widespread in `Sheet`/`WorkBook` |
| A consumer compiling with `skipLibCheck: false` | passes | passes, but accepts almost anything |

Both compile. The difference is whether the declarations catch a mistake or
permit it.

### No dependencies

`package.json` lists none, at runtime or in development, and no third-party code
is vendored. There is no dependency CVE history to inherit and no lockfile to
audit. SheetJS CE is also dependency-free, so this is a tie rather than a win -
but the supply-chain argument for either is real.

### Dates keep the wall clock

A date cell keeps its raw serial beside the Date, and parsing uses the local wall
clock rather than shifting by the host offset. A spreadsheet shows `2024-03-15`;
you get `2024-03-15` local, and `cell.serial` if you need the number.

## Where they are equal

* **Bundle size is not a differentiator.** Unminified, the whole of `src/` is
  656 kB raw and 204 kB packed, against SheetJS at roughly 1 MB raw for the full
  build. But SheetJS's mini build is within a few percent of gridxcel gzipped. If
  bundle size is your reason for choosing, check the mini build first.
* **API stability.** SheetJS has a decade of stability; gridxcel is pre-1.0 and
  the surface may still change.
* **Format support for XLSX itself.** Both read ordinary XLSX correctly. The
  difference is what each does with the parts neither fully supports, and how
  each handles a file that is damaged.

## Feature coverage

| Feature | gridxcel | SheetJS CE 0.20.3 |
| --- | --- | --- |
| XLSX read/write | Yes | Yes |
| CSV/TSV read/write | Yes | Yes |
| SYLK read/write | Yes | Read only |
| DIF read/write | Yes | Read only |
| XLSB, XLS, ODS, NUMBERS, DBF | No | Yes |
| Streaming read | Yes (`eachRow`) | No equivalent |
| Streaming CSV write | Yes (`createCSVStream`) | No equivalent |
| Formula text and cached values | Yes | Yes |
| Formula evaluation | No | No (CE) |
| Number formats | Yes (`formatNumber`) | Yes (`SSF`) |
| Cell styles | Read and write | Read and write |
| Charts, images, pivots | No | Partial (read, not preserved) |
| Browser support | Yes, native ESM | Yes, via bundler |

## Migrating

If you are moving from SheetJS, `gridxcel/compat` provides the same function
shapes - `read`, `write`, `utils.sheet_to_json` - so most call sites keep
working. [MIGRATION.md](MIGRATION.md) lists the differences that change results,
which are the ones worth reading before you switch.
