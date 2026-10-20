# Examples

> These scripts run against the repository checkout: they read the fixtures in
> `test/fixtures/`, which are not part of the published npm package. Clone
> https://github.com/gridxcel/gridxcel to run them.


Nine runnable programs, one idea each. Every file is a plain ES module with no
dependencies beyond gridxcel itself and Node builtins.

```sh
npm install gridxcel

node examples/run-all.mjs          # run all of them, report failures
node examples/01-read-a-workbook.mjs   # or run one
```

Any example can be started from any directory: fixtures are resolved relative to
the file with `new URL(..., import.meta.url)`, not from the current directory.
Whatever an example writes goes into `examples/out/`, which is gitignored.

| Example | Shows |
| --- | --- |
| [01-read-a-workbook.mjs](./01-read-a-workbook.mjs) | `read()`, the six cell types, dates as local wall clock, sparse rows and `originRow`/`originCol` |
| [02-stream-a-large-sheet.mjs](./02-stream-a-large-sheet.mjs) | `eachRow()`: the event shape, stopping early, `range`, row-number gaps, and streaming a 100,000-row file against parsing it |
| [03-convert-to-json.mjs](./03-convert-to-json.mjs) | `sheetToJSON()` with every header form, `sheetToAOA()`, `readJSON()` |
| [04-write-a-workbook.mjs](./04-write-a-workbook.mjs) | `write()` from rows and from objects, cell objects, a style, writing to a path, reading the result back |
| [05-stream-write.mjs](./05-stream-write.mjs) | `WorkbookWriter`: 250,000 rows appended, heap measured at each stage, and the row-order rule |
| [06-csv-roundtrip.mjs](./06-csv-roundtrip.mjs) | `csvToSheet()` and `sheetToCSV()`, quoting, delimiters, BOM, the CSV injection guard, CSV to XLSX and back |
| [07-formulas-and-formats.mjs](./07-formulas-and-formats.mjs) | `cell.f`, `cellFormula`, display text with `cellText`, and rendering numbers with `formatValue()` |
| [08-http-read.mjs](./08-http-read.mjs) | reading from a `fetch()` body and an `ArrayBuffer`, a CSV response, and what an HTML error page does |
| [09-errors.mjs](./09-errors.mjs) | the `GridxcelError` hierarchy, warnings versus `strict`, and the `RANGE_TOO_LARGE` dense guard |

## What the examples use

Fixtures come only from `test/fixtures/`, which is committed:

| Fixture | Used for |
| --- | --- |
| `golden.xlsx` | four sheets covering every cell type, merges, hidden rows, formulas and 200 rows of column data |
| `pandas.xlsx` | one tidy sheet with a real header row (100 data rows) |
| `reader-badcells.xlsx` | a damaged file that reads with warnings instead of throwing |

Nothing here reads `bench/`, which is generated on demand and mostly gitignored.

## Two examples need a flag for clean numbers

Examples 02 and 05 report heap growth. Run them with `--expose-gc` and they
call `global.gc()` before each measurement:

```sh
node --expose-gc examples/02-stream-a-large-sheet.mjs
node --expose-gc examples/05-stream-write.mjs
```

Without the flag they still run and still print numbers; the numbers just
include whatever the collector had not got to yet, and each prints a reminder.

## These are checked

```sh
node examples/run-all.mjs
```

exits non-zero if any example fails, which is what CI runs.
