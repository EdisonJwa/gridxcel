# Migrating from SheetJS

The compatibility surface is a real entry point, so a migration can start by
changing one import specifier and nothing else.

```diff
- import * as XLSX from "xlsx";
+ import * as XLSX from "gridxcel/compat";
```

## What the shim provides

```js
import { readFile, writeFile, utils, XLSX_like } from "gridxcel/compat";

const wb = readFile("in.xlsx");                    // { SheetNames, Sheets }
const rows = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
utils.book_append_sheet(wb, utils.aoa_to_sheet([[1, 2]]), "More");
writeFile(wb, "out.xlsx", { bookType: "xlsx" });
```

```js doc-example
import { readFile, utils, available, XLSX_like } from "gridxcel/compat";
import assert from "node:assert";

// The reader and writer are static imports of this module, so the synchronous
// SheetJS entry points work on the next line, exactly as SheetJS does.
assert.deepStrictEqual(available().read, true);
const wb = readFile("../../test/fixtures/pandas.xlsx");
assert.deepStrictEqual(wb.SheetNames, ["df"]);
const rows = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
assert.strictEqual(rows.length, 100);
assert.strictEqual(typeof XLSX_like.read, "function");
console.log("SheetNames:", JSON.stringify(wb.SheetNames), "| rows:", rows.length);
```

| SheetJS entry point | Status in `gridxcel/compat` |
| --- | --- |
| `read`, `readFile`, `write`, `writeFile` | Synchronous, matching SheetJS. `*Async` variants also exist. |
| `utils.aoa_to_sheet`, `json_to_sheet`, `sheet_to_json`, `sheet_to_csv`, `sheet_to_aoa` | Present. |
| `utils.book_new`, `book_append_sheet`, `book_set_sheet_visibility` | Present. |
| `utils.decode_range`, `encode_cell`, `encode_range`, `format_cell` | Present. |
| `read(data, { type })` | `"buffer"`, `"array"`, `"base64"`, `"binary"`, `"string"`. A string or non-zip input is parsed as CSV. |
| `write(wb, { type })` | Same set. With no `type` it returns a `Uint8Array`; SheetJS returns a binary string, so pass `type: "binary"` if you depend on that. |
| `type: "file"` | Throws `NotImplementedError`. Use `writeFile`. |
| `bookType` `xlsb`/`xls`/`ods`/`html`/`prn` | Throws `NotImplementedError` rather than writing a wrong file. |
| `sheet_to_csv(sheet, { style_text: true })` | Throws `NotImplementedError`; styles are not rendered to text. |
| `XLSX.stream.*` | Not provided. Use `eachRow` from the native API, which streams rows rather than accepting an already-parsed sheet. |

`writeFile` overwrites by default, matching SheetJS. Pass
`{ overwrite: false }` to refuse.

## Migrating to the native API

The shim is a bridge, not the destination. The native API replaces three
awkward parts of the SheetJS surface.

| SheetJS | gridxcel |
| --- | --- |
| `XLSX.read(data, { type })` with a magic `type` string | `await read(data)` - bytes, a path, a `URL`, a stream and an async iterable are all detected. |
| `wb.SheetNames` then `wb.Sheets[name]` | `book.sheets` is an array of sheet objects carrying their own `name`. |
| `cell.t` in OOXML letters (`b`/`n`/`e`/`s`/`d`/`z`) | `cell.t` in words: `"string"`, `"number"`, `"boolean"`, `"date"`, `"error"`, `"empty"`. |
| `XLSX.stream.to_csv(ws)` on an already-parsed sheet | `for await (const row of eachRow(path))` - the sheet is never materialised. |
| `cellDates` decides whether the serial is lost | The serial is always kept on `cell.serial`. |
| Errors are `Error` with prose messages | Typed classes with stable `code` values, all under `GridxcelError`. |

Read a sheet as the library's own model:

```js doc-example
import { read } from "gridxcel";
import assert from "node:assert";

const book = await read("../../test/fixtures/pandas.xlsx");
const sheet = book.sheets[0];
assert.strictEqual(sheet.name, "df");
assert.strictEqual(sheet.rows[0][0].v, "a");
console.log("sheet:", sheet.name, "| rowCount:", sheet.rowCount, "| columnCount:", sheet.columnCount);
console.log("A1:", JSON.stringify({ t: sheet.rows[0][0].t, v: sheet.rows[0][0].v }));
```

## Behaviour differences to plan for

These are the ones that change results rather than names.

1. **Missing dates stay serials unless you ask.** SheetJS converts cells that
   look like dates when `cellDates` is set; so does gridxcel, but gridxcel always
   records the original serial in `cell.serial` and never guesses from the value -
   only the cell's number format decides that a cell is a date.

2. **A formula cell's cached number is not reclassified as a date** just because
   the number falls in a date range. The style decides. This is asserted by a
   dedicated test because SheetJS's behaviour here surprises people.

3. **`escapeFormulae` is on by default** for CSV export, and it applies to
   *text* fields only, so `-2.5` stays numeric. SheetJS leaves this to the caller.

4. **CSV type inference is content-based**, matching SheetJS and pandas: `123` becomes
   a number, `TRUE` a boolean, `2024-01-15` a date. Set `quotedStrings: true` when
   a string that *looks* numeric must survive a round trip.

5. **Filesystem errors are typed** (`FILE_NOT_FOUND` and friends) instead of a raw
   `ENOENT`, so an existing `catch` block that inspects `err.code` will see different
   values. This is the one change most likely to need attention in a codebase
   that branches on `err.code`.

6. **Not-a-workbook throws.** SheetJS returns a workbook-shaped object for some
   garbage inputs; gridxcel throws `FormatError`. If you relied on "empty sheets
   means the file was not a workbook", invert that check.

## Deliberate deviations

Two, both small, both documented where they are implemented:

1. **`sheetToJSON({ header: 1 })` is dense.** SheetJS leaves array holes for
   missing cells; gridxcel emits `null` (or `defval`). A hole carries no information
   but changes `length` and iteration semantics for every consumer.
2. **Zero-sheet workbooks report an empty list.** A real workbook that declares
   no sheets returns `sheets: []` with a `NO_SHEETS` warning. Bytes that are not a
   workbook throw instead, so an empty list is never ambiguous.

## Checklist

1. Replace the import with `gridxcel/compat` and run your test suite. Nothing else
   needs to change for the SheetJS call shapes listed above.
2. Grep for `bookType` and `type` values you actually use, and confirm none of
   them are in the `NotImplementedError` list.
3. Grep for `err.code` and `Array.isArray(cell.t)` style checks and adjust for the
   typed errors and word-shaped cell types.
4. Switch the hot paths to `await read()` and `eachRow()`; that is where the
   memory and correctness wins are.
5. Run `node scripts/interop.mjs` in your own repository if you want the same
   python-verified evidence this package uses on itself.
