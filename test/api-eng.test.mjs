/**
 * api-eng acceptance tests: CSV, JSON/AOA and the SheetJS compat shim.
 * Run: node --test test/api-eng.test.mjs
 *
 * The independent oracle is the bundled Python 3 (its csv module, datetime
 * arithmetic for serials, and pandas for type inference). No gridxcel module
 * is ever used to check another gridxcel module.
 *
 * Source strings here are deliberately ASCII: unicode data is written with
 * escapes so the file stays portable.
 */
import test from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { sheetToCSV, sheetToDSV, csvToSheet, toCSV, fromCSV } from "../src/csv.js";
import { sheetToJSON, jsonToSheet, aoaToSheet, sheetToAOA } from "../src/json.js";
import * as compat from "../src/compat.js";

const PYTHON = "C:\\Users\\admin\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe";

/** openpyxl describes test/fixtures/pandas.xlsx for the compat fixture test. */
const PY_PANDAS_SRC = [
  "import json, sys",
  "from openpyxl import load_workbook",
  "wb = load_workbook(sys.argv[1])",
  "ws = wb[wb.sheetnames[0]]",
  "rows = []",
  "for row in ws.iter_rows(min_row=1, max_row=2, max_col=ws.max_column, values_only=True):",
  "    rows.append(list(row))",
  "meta = {\"sheets\": wb.sheetnames, \"rows\": rows, \"max_row\": ws.max_row, \"max_column\": ws.max_column}",
  "print(json.dumps(meta))",
];

/** openpyxl reads back the file our writer produced: the independent oracle. */
const PY_OPENPYXL_SRC = [
  "import json, os, sys",
  "from openpyxl import load_workbook",
  "base = sys.argv[1]",
  "wb = load_workbook(os.path.join(base, \"roundtrip.xlsx\"))",
  "ws = wb[wb.sheetnames[0]]",
  "rows = []",
  "for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):",
  "    line = []",
  "    for cell in row:",
  "        v = cell.value",
  "        if hasattr(v, \"isoformat\"):",
  "            line.append(\"date:\" + v.isoformat() + \"|\" + str(cell.number_format))",
  "        else:",
  "            line.append(type(v).__name__ + \":\" + repr(v))",
  "    rows.append(line)",
  "meta = {\"names\": wb.sheetnames, \"dims\": ws.dimensions, \"rows\": rows}",
  "with open(os.path.join(base, \"openpyxl.json\"), \"w\", encoding=\"utf-8\") as fh:",
  "    json.dump(meta, fh, ensure_ascii=False)",
];

/** The oracle script: stdlib csv + datetime, pandas when present. */
const PY_SRC = [
  "import csv, json, os, sys",
  "from datetime import date, datetime",
  "",
  "base = sys.argv[1]",
  "with open(os.path.join(base, \"rows.json\"), encoding=\"utf-8\") as fh:",
  "    data = json.load(fh)",
  "",
  "rows = data[\"rows\"]",
  "with open(os.path.join(base, \"oracle.csv\"), \"w\", encoding=\"utf-8\", newline=\"\") as fh:",
  "    writer = csv.writer(fh, lineterminator=chr(10), quoting=csv.QUOTE_MINIMAL)",
  "    writer.writerows(rows)",
  "",
  "meta = {\"serials\": {}}",
  "for iso in data[\"dates\"]:",
  "    day = datetime.strptime(iso, \"%Y-%m-%d\").date()",
  "    meta[\"serials\"][iso] = (day - date(1899, 12, 30)).days",
  "",
  "try:",
  "    import pandas as pd",
  "    frame = pd.read_csv(os.path.join(base, \"oracle.csv\"), keep_default_na=False)",
  "    meta[\"pandas\"] = True",
  "    meta[\"columns\"] = list(frame.columns)",
  "    meta[\"names\"] = [str(v) for v in frame[\"name\"].tolist()]",
  "    meta[\"qty\"] = [float(v) for v in frame[\"qty\"].tolist()]",
  "    meta[\"flags\"] = list(frame[\"flag\"].tolist())",
  "    meta[\"flag_dtype\"] = str(frame[\"flag\"].dtype)",
  "except Exception as exc:",
  "    meta[\"pandas\"] = False",
  "    meta[\"pandas_error\"] = str(exc)",
  "",
  "with open(os.path.join(base, \"oracle.json\"), \"w\", encoding=\"utf-8\") as fh:",
  "    json.dump(meta, fh)",
];

const UNICODE = "h\u00e9llo \u2192 \u65e5\u672c\u8a9e \ud83c\udf89";

/** Comparable view of a cell: values, not identities. */
function plain(cell) {
  if (!cell) return null;
  return { t: cell.t, v: cell.v instanceof Date ? cell.v.getTime() : cell.v };
}

function plainRow(row) {
  return row.map(plain);
}

/* ------------------------------------------------------------------- CSV */

test("csv: round trip keeps commas, quotes, newlines, unicode, numbers, booleans and dates", () => {
  const sheet = aoaToSheet([
    ["name", "note", "qty", "ok", "when"],
    ["a,b", 'say "hi"', 3, true, new Date(2024, 0, 15)],
    ["multi\nline", UNICODE, -2.5, false, new Date(2024, 5, 30)],
  ]);
  const text = sheetToCSV(sheet, { quotedStrings: true });
  const expected = [
    "name,note,qty,ok,when",
    '"a,b","say ""hi""",3,TRUE,2024-01-15',
    '"multi\nline",' + UNICODE + ",-2.5,FALSE,2024-06-30",
    "",
  ].join("\n");
  assert.equal(text, expected);

  const back = csvToSheet(text, { quotedStrings: true });
  assert.equal(back.rows.length, sheet.rows.length);
  assert.equal(back.columnCount, 5);
  assert.deepEqual(back.rows.map(plainRow), sheet.rows.map(plainRow));

  // Dates carry the serial the writer needs, and the format code that shows it.
  const dateCell = back.rows[1][4];
  assert.equal(dateCell.t, "date");
  assert.ok(dateCell.v instanceof Date);
  assert.equal(typeof dateCell.serial, "number");
  assert.equal(dateCell.z, "yyyy-mm-dd");
  assert.equal(back.rows[1][2].t, "number");
  assert.equal(back.rows[1][3].t, "boolean");
  assert.equal(back.rows[1][3].v, true);
  assert.equal(back.rows[2][1].v, UNICODE);
});

test("csv: content-based inference by default, quotedStrings opt-in", () => {
  // Quoting does not change types unless asked, which is what SheetJS and
  // pandas do; sheetToCSV only adds the quotes when quotedStrings is on.
  const sheet = aoaToSheet([["123", "TRUE", "2024-01-15"]]);
  assert.equal(sheetToCSV(sheet), "123,TRUE,2024-01-15\n");
  assert.equal(sheetToCSV(sheet, { quotedStrings: true }), '"123","TRUE","2024-01-15"\n');
  const inferred = csvToSheet("123,TRUE,2024-01-15\n").rows[0];
  assert.deepEqual(inferred.map((c) => c.t), ["number", "boolean", "date"]);
  const literal = csvToSheet('"123","TRUE","2024-01-15"\n', { quotedStrings: true }).rows[0];
  assert.deepEqual(literal.map((c) => c.t), ["string", "string", "string"]);
  assert.deepEqual(literal.map((c) => c.v), ["123", "TRUE", "2024-01-15"]);
});

test("csv: hand-written CRLF + BOM + embedded newline parses to the exact cells", () => {
  const raw = "\ufeffid,text,flag\r\n1,\"line one\r\nline two\",TRUE\r\n2,\"a,b\",FALSE\r\n";
  const sheet = csvToSheet(raw);
  assert.equal(sheet.rows.length, 3);
  assert.equal(sheet.columnCount, 3);
  assert.deepEqual(sheet.rows.map(plainRow), [
    [{ t: "string", v: "id" }, { t: "string", v: "text" }, { t: "string", v: "flag" }],
    [{ t: "number", v: 1 }, { t: "string", v: "line one\r\nline two" }, { t: "boolean", v: true }],
    [{ t: "number", v: 2 }, { t: "string", v: "a,b" }, { t: "boolean", v: false }],
  ]);
  // The BOM is consumed, never part of the first cell.
  assert.equal(sheet.rows[0][0].v, "id");
});

test("csv: ragged rows pad to the widest, empty fields stay empty strings", () => {
  const sheet = csvToSheet("a,b,c\n1,2\n3\n");
  assert.equal(sheet.rowCount, 3);
  assert.equal(sheet.columnCount, 3);
  assert.deepEqual(sheet.rows[1].map((c) => c.v), [1, 2, ""]);
  assert.deepEqual(sheet.rows[2].map((c) => c.v), [3, "", ""]);
  assert.ok(sheet.rows[1].every((c) => c.v !== null));
  // Ragged padding matches the widest row, so every row is the same length.
  assert.ok(sheet.rows.every((r) => r.length === 3));
});

test("csv: a trailing newline does not invent a row, an extra blank line does", () => {
  assert.equal(csvToSheet("a\n").rows.length, 1);
  assert.equal(csvToSheet("a\n\n").rows.length, 2);
  assert.equal(csvToSheet("a\r\nb\r\n").rows.length, 2);
  // a lone CR is a row break too, exactly like RFC 4180's predecessors
  assert.deepEqual(csvToSheet("a\rb").rows.map((r) => r[0].v), ["a", "b"]);
  assert.deepEqual(csvToSheet("a\n\n", { blankrows: false }).rows.map((r) => r[0].v), ["a"]);
});

test("csv: escapeFormulae prefixes formula text with an apostrophe (default on)", () => {
  const sheet = aoaToSheet([["=1+1", "-2", "@cmd", "+SUM(A1)", "plain", -2, "=a,b"]]);
  assert.equal(sheetToCSV(sheet), "'=1+1,'-2,'@cmd,'+SUM(A1),plain,-2,\"'=a,b\"\n");
  assert.equal(sheetToCSV(sheet, { escapeFormulae: false }), "=1+1,-2,@cmd,+SUM(A1),plain,-2,\"=a,b\"\n");
});

test("csv: delimiter, lineEnding, bom and alwaysQuote options", () => {
  assert.equal(sheetToDSV(aoaToSheet([[1, 2], [3, 4]]), { delimiter: ";" }), "1;2\n3;4\n");
  assert.equal(sheetToCSV(aoaToSheet([[1, 2]]), { delimiter: "\t" }), "1\t2\n");
  assert.equal(sheetToCSV(aoaToSheet([[1, 2]]), { lineEnding: "\r\n", bom: true }), "\ufeff1,2\r\n");
  assert.equal(sheetToCSV(aoaToSheet([["a"]]), { alwaysQuote: true }), '"a"\n');
  assert.equal(csvToSheet("1;2\n", { delimiter: ";" }).rows[0][1].v, 2);
  assert.equal(csvToSheet("1\t2\n", { FS: "\t" }).rows[0][1].v, 2);
});

test("csv: dateNF drives both the written text and the parsed value", () => {
  const sheet = aoaToSheet([[new Date(2024, 0, 15)]]);
  assert.equal(sheetToCSV(sheet, { dateNF: "mm/dd/yy" }), "01/15/24\n");
  const back = csvToSheet("01/15/24\n", { dateNF: "mm/dd/yy" });
  assert.equal(back.rows[0][0].t, "date");
  assert.equal(back.rows[0][0].v.getTime(), new Date(2024, 0, 15).getTime());
  // An impossible date is text, not a silent rollover into March.
  assert.equal(csvToSheet("2024-02-31\n").rows[0][0].t, "string");
});

test("csv: toCSV/fromCSV bridge books and text", () => {
  const book = fromCSV("a,b\n1,2\n");
  assert.equal(book.sheets.length, 1);
  assert.equal(book.sheets[0].name, "Sheet1");
  assert.equal(toCSV(book), "a,b\n1,2\n");
  assert.deepEqual(sheetToJSON(book.sheets[0]), [{ a: 1, b: 2 }]);
});

/* ------------------------------------------------------------------ JSON */

test("json: the header row becomes the keys", () => {
  const sheet = aoaToSheet([["name", "qty"], ["a", 1], ["b", 2]]);
  assert.deepEqual(sheetToJSON(sheet), [{ name: "a", qty: 1 }, { name: "b", qty: 2 }]);
  // An explicit header list supplies the keys and leaves row 1 as data.
  assert.deepEqual(sheetToJSON(sheet, { header: ["n", "q"] }), [
    { n: "name", q: "qty" },
    { n: "a", q: 1 },
    { n: "b", q: 2 },
  ]);
  // header by 1-based row number
  const two = aoaToSheet([["junk"], ["name", "qty"], ["a", 1]]);
  assert.deepEqual(sheetToJSON(two, { header: 2 }), [{ name: "a", qty: 1 }]);
});

test("json: raw:true uses A/B/C keys and keeps the first row as data", () => {
  const sheet = aoaToSheet([["name", "qty"], ["a", 1]]);
  assert.deepEqual(sheetToJSON(sheet, { raw: true }), [
    { A: "name", B: "qty" },
    { A: "a", B: 1 },
  ]);
  // header:"A" is the modern SheetJS spelling of the same thing
  assert.deepEqual(sheetToJSON(sheet, { header: "A" }), [
    { A: "name", B: "qty" },
    { A: "a", B: 1 },
  ]);
});

test("json: header forms follow SheetJS (1 = arrays, A = letters, n > 1 = a header row)", () => {
  const sheet = aoaToSheet([["h1", "h2", "h3"], [1, 2, 3], [4, 5, 6]]);
  // absent: row 1 is the header row and is not data
  assert.deepEqual(sheetToJSON(sheet), [{ h1: 1, h2: 2, h3: 3 }, { h1: 4, h2: 5, h3: 6 }]);
  // header: 1 -> array of arrays, first row included
  assert.deepEqual(sheetToJSON(sheet, { header: 1 }), [["h1", "h2", "h3"], [1, 2, 3], [4, 5, 6]]);
  // header: "A" -> column letters, first row included
  assert.deepEqual(sheetToJSON(sheet, { header: "A" }), [
    { A: "h1", B: "h2", C: "h3" },
    { A: 1, B: 2, C: 3 },
    { A: 4, B: 5, C: 6 },
  ]);
  // raw:true is the older spelling of header:"A"
  assert.deepEqual(sheetToJSON(sheet, { raw: true }), sheetToJSON(sheet, { header: "A" }));
  // header: [...] -> these names, first row still data
  assert.deepEqual(sheetToJSON(sheet, { header: ["a", "b", "c"] }), [
    { a: "h1", b: "h2", c: "h3" },
    { a: 1, b: 2, c: 3 },
    { a: 4, b: 5, c: 6 },
  ]);
  // header: n (n > 1) -> row n is the header row, data follows
  assert.deepEqual(sheetToJSON(sheet, { header: 2 }), [{ 1: 4, 2: 5, 3: 6 }]);
  // a shorter list falls back to the sheet's own header text for the rest
  assert.deepEqual(Object.keys(sheetToJSON(sheet, { header: ["only"] })[0]), ["only", "h2", "h3"]);
});

test("json: repeated header labels are disambiguated like SheetJS", () => {
  const sheet = aoaToSheet([["S", "h", "e", "e", "t"], [1, 2, 3, 4, 5]]);
  assert.deepEqual(sheetToJSON(sheet), [{ S: 1, h: 2, e: 3, e_1: 4, t: 5 }]);
  // a caller-supplied list is used verbatim, as SheetJS documents: repeated
  // names are NOT disambiguated, so the later column wins
  assert.deepEqual(sheetToJSON(sheet, { header: ["e", "e"], range: "A1:B2" }), [{ e: "h" }, { e: 2 }]);
});

test("json: defval and blankrows behave in every header mode", () => {
  const sheet = aoaToSheet([["h1", "h2"], [1], [], [2, 3]]);
  // header:1 keeps blank rows by default (dense grid) and fills gaps with defval
  assert.deepEqual(sheetToJSON(sheet, { header: 1 }), [["h1", "h2"], [1, null], [null, null], [2, 3]]);
  assert.deepEqual(sheetToJSON(sheet, { header: 1, defval: "" }), [["h1", "h2"], [1, ""], ["", ""], [2, 3]]);
  assert.deepEqual(sheetToJSON(sheet, { header: 1, blankrows: false }), [["h1", "h2"], [1, null], [2, 3]]);
  // object modes drop blank rows by default and fill missing keys with defval
  assert.deepEqual(sheetToJSON(sheet, { header: ["a", "b"] }), [
    { a: "h1", b: "h2" },
    { a: 1 },
    { a: 2, b: 3 },
  ]);
  assert.deepEqual(sheetToJSON(sheet, { header: ["a", "b"], defval: null }), [
    { a: "h1", b: "h2" },
    { a: 1, b: null },
    { a: 2, b: 3 },
  ]);
  assert.deepEqual(sheetToJSON(sheet, { header: "A", blankrows: true, defval: 0 }), [
    { A: "h1", B: "h2" },
    { A: 1, B: 0 },
    { A: 0, B: 0 },
    { A: 2, B: 3 },
  ]);
  // header:2 with a blank row after it behaves the same way
  const shifted = aoaToSheet([["junk"], ["a", "b"], [1], [], [2, 3]]);
  assert.deepEqual(sheetToJSON(shifted, { header: 2 }), [{ a: 1 }, { a: 2, b: 3 }]);
  assert.deepEqual(sheetToJSON(shifted, { header: 2, blankrows: true }), [{ a: 1 }, {}, { a: 2, b: 3 }]);
});

test("json: defval fills missing cells, blankrows controls empty rows", () => {
  const sheet = aoaToSheet([["a", "b"], [1], [], [2, 3]]);
  assert.deepEqual(sheetToJSON(sheet), [{ a: 1 }, { a: 2, b: 3 }]);
  assert.deepEqual(sheetToJSON(sheet, { defval: null }), [{ a: 1, b: null }, { a: 2, b: 3 }]);
  assert.deepEqual(sheetToJSON(sheet, { blankrows: true }), [{ a: 1 }, {}, { a: 2, b: 3 }]);
  assert.deepEqual(sheetToJSON(sheet, { blankrows: true, defval: 0 }), [
    { a: 1, b: 0 },
    { a: 0, b: 0 },
    { a: 2, b: 3 },
  ]);
});

test("json: range limits the output and shifts the header", () => {
  const sheet = aoaToSheet([["skip", "me"], ["a", "b"], [1, 2], [3, 4]]);
  assert.deepEqual(sheetToJSON(sheet, { range: "A2:B4" }), [{ a: 1, b: 2 }, { a: 3, b: 4 }]);
  // A bare number is the 1-based first row (SheetJS spelling) and the row it
  // points at is still the header; an explicit header[] replaces the keys but
  // the first row of the range stays consumed.
  assert.deepEqual(sheetToJSON(sheet, { range: 2 }), [{ a: 1, b: 2 }, { a: 3, b: 4 }]);
  assert.deepEqual(sheetToJSON(sheet, { range: 2, header: ["x", "y"] }), [
    { x: "a", y: "b" },
    { x: 1, y: 2 },
    { x: 3, y: 4 },
  ]);
  assert.deepEqual(sheetToJSON(sheet, { range: "A2:B3", raw: true }), [{ A: "a", B: "b" }, { A: 1, B: 2 }]);
});

test("json: jsonToSheet keeps first-seen key order across every row", () => {
  const sheet = jsonToSheet([{ b: 1, a: 2 }, { a: 3, c: 4 }, { c: 5, b: 6 }]);
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["b", "a", "c"]);
  assert.deepEqual(sheetToJSON(sheet), [{ b: 1, a: 2 }, { a: 3, c: 4 }, { b: 6, c: 5 }]);
  assert.deepEqual(Object.keys(sheetToJSON(sheet)[2]), ["b", "c"]);
  // header option sets the order and keeps unlisted keys out of the header row
  const ordered = jsonToSheet([{ b: 1, a: 2 }], { header: ["a", "b", "z"] });
  assert.deepEqual(ordered.rows[0].map((c) => c.v), ["a", "b", "z"]);
  const noHeader = jsonToSheet([{ a: 1 }], { skipHeader: true });
  assert.deepEqual(noHeader.rows.map((r) => r.map((c) => c.v)), [[1]]);
});

test("json: types survive a JSON round trip, dates included", () => {
  const when = new Date(2024, 2, 15, 14, 30, 45);
  const sheet = jsonToSheet([{ s: "x", n: 1.5, b: false, d: when, nil: null }]);
  // The header row keeps every key, including the one whose only value is null.
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["s", "n", "b", "d", "nil"]);
  // A null value is an empty cell: the key comes back only when defval asks.
  assert.deepEqual(sheetToJSON(sheet), [{ s: "x", n: 1.5, b: false, d: when }]);
  assert.deepEqual(sheetToJSON(sheet, { defval: null }), [{ s: "x", n: 1.5, b: false, d: when, nil: null }]);
  const row = sheetToJSON(sheet)[0];
  assert.deepEqual(Object.keys(row), ["s", "n", "b", "d"]);
  assert.ok(row.d instanceof Date);
  assert.equal(row.d.getTime(), when.getTime());
});

test("json: sheetToAOA is dense, honours defval and blankrows", () => {
  const sheet = aoaToSheet([["a", "b"], [1], [], [2, 3]]);
  assert.deepEqual(sheetToAOA(sheet), [["a", "b"], [1, null], [null, null], [2, 3]]);
  assert.deepEqual(sheetToAOA(sheet, { defval: "" }), [["a", "b"], [1, ""], ["", ""], [2, 3]]);
  assert.deepEqual(sheetToAOA(sheet, { blankrows: false }), [["a", "b"], [1, null], [2, 3]]);
  assert.deepEqual(sheetToAOA(sheet, { range: "A2:B2" }), [[1, null]]);
});

test("json: NaN becomes an error cell rather than a number Excel cannot open", () => {
  const sheet = aoaToSheet([[NaN, Infinity]]);
  assert.deepEqual(sheet.rows[0].map((c) => c.t), ["error", "error"]);
  assert.deepEqual(sheet.rows[0].map((c) => c.e), ["#NUM!", "#NUM!"]);
});

/* ---------------------------------------------------------------- compat */

test("compat: utils surface behaves like SheetJS", () => {
  const wb = compat.utils.book_new();
  const ws = compat.utils.aoa_to_sheet([["n", "s"], [1, "a"], [2, "b"]]);
  compat.utils.book_append_sheet(wb, ws, "Data");
  assert.deepEqual(wb.SheetNames, ["Data"]);
  assert.equal(wb.Sheets.Data, ws);
  assert.equal(compat.utils.sheet_to_csv(wb.Sheets.Data), "n,s\n1,a\n2,b\n");
  assert.deepEqual(compat.utils.sheet_to_json(wb.Sheets.Data), [{ n: 1, s: "a" }, { n: 2, s: "b" }]);
  assert.equal(compat.utils.encode_cell({ r: 0, c: 1 }), "B1");
  assert.equal(compat.utils.decode_cell("B1").c, 1);
  assert.equal(compat.utils.encode_col(26), "AA");
  assert.equal(compat.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }), "A1:B2");
  assert.deepEqual(compat.utils.decode_range("A1:B2"), { s: { r: 0, c: 0 }, e: { r: 1, c: 1 } });
  assert.equal(compat.utils.format_cell({ t: "number", v: 3.14159, z: "0.00" }), "3.14");
  assert.equal(compat.utils.format_cell({ t: "boolean", v: true }), "TRUE");
  assert.equal(compat.utils.format_cell({ t: "number", v: 1 }, 42), "42");
  assert.equal(compat.utils.format_cell({ t: "string", v: "x", w: "shown" }), "shown");
  assert.equal(compat.utils.sheet_to_txt(compat.utils.aoa_to_sheet([[1, 2]])), "\ufeff1\t2\n");
  assert.throws(() => compat.utils.sheet_to_csv(ws, { style_text: true }), /NOT_IMPLEMENTED|not supported/);
  compat.utils.book_set_sheet_visibility(wb, "Data", 1);
  assert.equal(wb.Sheets.Data.hidden, true);
  compat.utils.book_set_sheet_visibility(wb, 0, 0);
  assert.equal(wb.Sheets.Data.hidden, false);
  assert.deepEqual(compat.utils.decode_range("B2:C3").e, { r: 2, c: 2 });
});

test("compat: write takes a SheetJS workbook and returns the documented types", () => {
  const ws = compat.utils.aoa_to_sheet([["n", "s"], [1, "a"]]);
  const sjswb = { SheetNames: ["S1"], Sheets: { S1: ws } };
  assert.equal(compat.write(sjswb, { bookType: "csv", type: "string" }), "n,s\n1,a\n");
  assert.ok(Buffer.isBuffer(compat.write(sjswb, { bookType: "csv", type: "buffer" })));
  assert.equal(compat.write(sjswb, { bookType: "csv", type: "array" }) instanceof Uint8Array, true);
  assert.equal(
    Buffer.from(compat.write(sjswb, { bookType: "csv", type: "base64" }), "base64").toString("utf8"),
    "n,s\n1,a\n",
  );
  assert.deepEqual(JSON.parse(compat.write(sjswb, { bookType: "json", type: "string" })), [{ n: 1, s: "a" }]);
  // "binary" is latin1 text, which the SheetJS-shaped reader accepts back
  const binary = compat.write(sjswb, { bookType: "csv", type: "binary" });
  assert.equal(binary, "n,s\n1,a\n");
  const backBinary = compat.read(binary, { type: "binary" });
  assert.equal(backBinary.Sheets[backBinary.SheetNames[0]].rows[0][0].v, "n");
  // A worksheet on its own is accepted too, and the SheetJS view round-trips.
  assert.equal(compat.write(ws, { bookType: "csv", type: "string" }), "n,s\n1,a\n");
  const view = compat.toWorksheet(ws);
  assert.equal(view["!ref"], "A1:B2");
  assert.deepEqual(view.A1, { v: "n", t: "s" });
  assert.deepEqual(compat.toSheetJSBook(compat.normalizeBookInput(sjswb)).SheetNames, ["S1"]);
  assert.throws(() => compat.write(sjswb, { bookType: "xlsb" }), /NOT_IMPLEMENTED|not support/);
});

test("compat: csv text and bytes go through read() like SheetJS", () => {
  const wb = compat.read("n,s\n1,a\n", { type: "string" });
  assert.deepEqual(wb.SheetNames, ["Sheet1"]);
  assert.deepEqual(compat.utils.sheet_to_json(wb.Sheets.Sheet1), [{ n: 1, s: "a" }]);
  const bytes = compat.read(Buffer.from("n,s\n1,a\n", "utf8"));
  assert.deepEqual(compat.utils.sheet_to_json(bytes.Sheets[bytes.SheetNames[0]]), [{ n: 1, s: "a" }]);
});

test("compat: xlsx write/read round trip when the peers are present", async (t) => {
  const state = await compat.preload();
  if (!state.read || !state.write) {
    console.log("SKIP compat xlsx round trip: read.js/write.js not present yet (" + JSON.stringify(state) + ")");
    t.skip("peer modules missing");
    return;
  }
  const wb = compat.utils.book_new();
  compat.utils.book_append_sheet(wb, compat.utils.aoa_to_sheet([["n", "s"], [1, "a"], [2.5, UNICODE]]), "Data");
  const bytes = compat.write(wb, { bookType: "xlsx", type: "buffer" });
  assert.ok(Buffer.isBuffer(bytes));
  assert.equal(bytes.subarray(0, 2).toString("latin1"), "PK");
  const back = compat.read(bytes, { type: "buffer" });
  assert.deepEqual(back.SheetNames, ["Data"]);
  assert.deepEqual(compat.utils.sheet_to_json(back.Sheets.Data), [{ n: 1, s: "a" }, { n: 2.5, s: UNICODE }]);
});

test("compat: readFile reads test/fixtures/pandas.xlsx and matches openpyxl", async (t) => {
  if (!(await compat.preload()).read) {
    console.log("SKIP compat readFile: src/read.js is not present yet");
    t.skip("no src/read.js");
    return;
  }
  const fixture = fileURLToPath(new URL("./fixtures/pandas.xlsx", import.meta.url));
  assert.ok(existsSync(fixture), "fixture exists: " + fixture);
  const wb = compat.readFile(fixture);
  const first = wb.Sheets[wb.SheetNames[0]];
  assert.ok(first && first.rows.length > 0, "the first sheet has rows");
  assert.equal(first.name, wb.SheetNames[0]);
  assert.ok(compat.utils.sheet_to_csv(first).length > 0);

  if (!existsSync(PYTHON)) {
    console.log("SKIP openpyxl cross-check of pandas.xlsx: bundled python is not present");
    return;
  }
  const dir = mkdtempSync(join(tmpdir(), "gridxcel-pandas-"));
  try {
    writeFileSync(join(dir, "oracle_pandas.py"), PY_PANDAS_SRC.join("\n"), "utf8");
    const meta = JSON.parse(execFileSync(PYTHON, [join(dir, "oracle_pandas.py"), fixture], { encoding: "utf8" }));
    assert.deepEqual(wb.SheetNames, meta.sheets);
    const expected = {};
    meta.rows[0].forEach((name, i) => { expected[String(name)] = meta.rows[1][i]; });
    assert.deepEqual(compat.utils.sheet_to_json(first)[0], expected);
    const lines = compat.utils.sheet_to_csv(first).trimEnd().split("\n");
    assert.equal(lines.length, meta.max_row);
    assert.equal(lines[0], meta.rows[0].join(","));
    assert.equal(lines[1], meta.rows[1].join(","));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

/* ------------------------------------------------- extra shape coverage */

test("json: a SheetJS worksheet can be consumed directly", () => {
  const ws = {
    "!ref": "A1:B2",
    A1: { t: "s", v: "name" },
    B1: { t: "s", v: "qty" },
    A2: { t: "s", v: "a" },
    B2: { t: "n", v: 1 },
  };
  assert.deepEqual(sheetToJSON(ws), [{ name: "a", qty: 1 }]);
  assert.equal(sheetToCSV(ws), "name,qty\na,1\n");
  assert.deepEqual(sheetToAOA(ws), [["name", "qty"], ["a", 1]]);
});

test("csv: skipHidden drops hidden rows and columns", () => {
  const sheet = aoaToSheet([["a", "b", "c"], [1, 2, 3], [4, 5, 6]]);
  sheet.rowInfo.set(3, { index: 3, hidden: true });
  sheet.columns = [{ index: 1, min: 1, max: 1, hidden: true }];
  assert.equal(sheetToCSV(sheet, { skipHidden: true }), "a,c\n1,3\n");
  assert.deepEqual(sheetToJSON(sheet, { skipHidden: true }), [{ a: 1, c: 3 }]);
  assert.equal(sheetToCSV(sheet), "a,b,c\n1,2,3\n4,5,6\n");
});

test("csv: the header option writes a row above the data on both sides", () => {
  assert.equal(sheetToCSV(aoaToSheet([[1, 2]]), { header: ["a", "b"] }), "a,b\n1,2\n");
  const parsed = csvToSheet("1,2\n", { header: ["a", "b"] });
  assert.deepEqual(sheetToJSON(parsed), [{ a: 1, b: 2 }]);
  assert.equal(sheetToCSV(parsed), "a,b\n1,2\n");
});

/* ------------------------------------------- end to end with the oracle */

/**
 * The whole chain: CSV text -> Sheet -> writeWorkbook -> xlsx bytes ->
 * readWorkbook, with openpyxl confirming what a real spreadsheet program sees
 * (including that the date cells carry both a Date and a usable serial).
 */
test("interop: csv -> sheet -> xlsx -> openpyxl -> sheet round trips", async (t) => {
  const state = await compat.preload();
  if (!state.read || !state.write || !existsSync(PYTHON)) {
    console.log("SKIP interop round trip: peers or python missing (" + JSON.stringify(state) + ")");
    t.skip("peer modules or python missing");
    return;
  }
  const dir = mkdtempSync(join(tmpdir(), "gridxcel-interop-"));
  try {
    const text = "name,note,qty,ok,when\n"
      + "\"a,b\",\"say \"\"hi\"\"\",3,TRUE,2024-01-15\n"
      + "\"multi\nline\"," + UNICODE + ",-2.5,FALSE,2024-06-30\n";
    const sheet = csvToSheet(text);
    assert.ok(sheet, "csvToSheet produced a sheet");
    sheet.name = "Data";
    const bytes = compat.write({ sheets: [sheet] }, { bookType: "xlsx", type: "buffer" });
    assert.equal(bytes.subarray(0, 2).toString("latin1"), "PK");
    writeFileSync(join(dir, "roundtrip.xlsx"), bytes);
    // Deliberately not named openpyxl.py: the script's directory is on
    // sys.path and would shadow the real package.
    writeFileSync(join(dir, "oracle_openpyxl.py"), PY_OPENPYXL_SRC.join("\n"), "utf8");
    execFileSync(PYTHON, [join(dir, "oracle_openpyxl.py"), dir]);

    const meta = JSON.parse(readFileSync(join(dir, "openpyxl.json"), "utf8"));
    assert.deepEqual(meta.names, ["Data"]);
    assert.equal(meta.dims, "A1:E3");
    assert.deepEqual(meta.rows[1], ["str:'a,b'", 'str:\'say "hi"\'', "int:3", "bool:True", "date:2024-01-15T00:00:00|yyyy-mm-dd"]);
    assert.deepEqual(meta.rows[2], ["str:'multi\\nline'", "str:'" + UNICODE + "'", "float:-2.5", "bool:False", "date:2024-06-30T00:00:00|yyyy-mm-dd"]);

    const back = compat.read(bytes, { type: "buffer", cellDates: true });
    const round = back.Sheets[back.SheetNames[0]];
    assert.equal(sheetToCSV(round), text);
    assert.equal(round.rows[1][4].t, "date");
    assert.equal(round.rows[1][4].serial, 45306);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("json: golden.xlsx header modes (raw === header A, header 1 is the grid)", async (t) => {
  if (!(await compat.preload()).read) {
    console.log("SKIP golden.xlsx header modes: src/read.js is not present yet");
    t.skip("no src/read.js");
    return;
  }
  const fixture = fileURLToPath(new URL("./fixtures/golden.xlsx", import.meta.url));
  if (!existsSync(fixture)) {
    console.log("SKIP golden.xlsx header modes: fixture is not present");
    t.skip("no fixture");
    return;
  }
  const wb = compat.readFile(fixture);
  const types = wb.Sheets.Types;
  if (!types) {
    console.log("SKIP golden.xlsx header modes: no Types sheet");
    t.skip("no Types sheet");
    return;
  }
  const letters = compat.utils.sheet_to_json(types, { header: "A" });
  // raw:true is accepted as an alias for header:"A", first row included
  assert.deepEqual(compat.utils.sheet_to_json(types, { raw: true }), letters);
  assert.deepEqual(Object.keys(letters[0]).slice(0, 4), ["A", "B", "C", "D"]);
  assert.equal(letters[0].A, "text");
  assert.equal(letters[0].B, 42);
  assert.equal(letters[0].C, true);
  // the emoji/unicode string survives byte-exactly
  assert.equal(letters[0].I, "\u65e5\u672c\u8a9e \ud83d\ude00 \u00e9\u00e8");
  // header:1 is the same grid as arrays, including the first row
  const grid = compat.utils.sheet_to_json(types, { header: 1 });
  assert.deepEqual(grid[0].slice(0, 4), [letters[0].A, letters[0].B, letters[0].C, letters[0].D]);
  assert.equal(grid.length, letters.length);
  // the Types sheet's first row is data, so the default mode reads it as the header
  assert.equal(compat.utils.sheet_to_json(types).length + 1, letters.length);
});

/* --------------------------------------------------- independent oracle */

/**
 * Python writes the same rows with its own csv module; the bytes must match
 * exactly. Python's datetime also supplies the expected Excel serials and
 * pandas independently infers the column types from the file we produced.
 */
test("csv: byte-identical to python csv.writer, types confirmed by pandas", (t) => {
  if (!existsSync(PYTHON)) {
    console.log("SKIP python oracle: " + PYTHON + " is not present");
    t.skip("no bundled python");
    return;
  }
  const dir = mkdtempSync(join(tmpdir(), "gridxcel-oracle-"));
  try {
    const rows = [
      ["name", "qty", "flag"],
      ["a,b", 3, "TRUE"],
      ['say "hi"', -2.5, "FALSE"],
      ["multi\nline", 0, "TRUE"],
      [UNICODE, 7, "FALSE"],
    ];
    const dates = ["2024-01-15", "2024-06-30"];
    writeFileSync(join(dir, "rows.json"), JSON.stringify({ rows, dates }), "utf8");
    writeFileSync(join(dir, "oracle.py"), PY_SRC.join("\n"), "utf8");
    execFileSync(PYTHON, [join(dir, "oracle.py"), dir]);

    const mine = Buffer.from(sheetToCSV(aoaToSheet(rows)), "utf8");
    const theirs = readFileSync(join(dir, "oracle.csv"));
    assert.deepEqual(mine, theirs);

    const meta = JSON.parse(readFileSync(join(dir, "oracle.json"), "utf8"));
    for (const iso of dates) {
      const cell = csvToSheet(iso).rows[0][0];
      assert.equal(cell.t, "date");
      assert.equal(cell.serial, meta.serials[iso], "serial for " + iso);
      assert.equal(cell.v.getTime(), new Date(iso + "T00:00:00").getTime());
    }

    if (!meta.pandas) {
      console.log("SKIP pandas type cross-check: " + meta.pandas_error);
      return;
    }
    assert.deepEqual(meta.columns, ["name", "qty", "flag"]);
    assert.equal(meta.flag_dtype, "bool");
    const parsed = csvToSheet(mine.toString("utf8"));
    assert.equal(parsed.rows[0].map((c) => c.v).join(","), "name,qty,flag");
    assert.deepEqual(parsed.rows.slice(1).map((r) => r[0].v), meta.names);
    assert.deepEqual(parsed.rows.slice(1).map((r) => r[1].v), meta.qty);
    assert.deepEqual(parsed.rows.slice(1).map((r) => r[2].v), meta.flags);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
