/**
 * De-duplication regressions: DUP-1, DUP-5, DUP-6, DUP-7 and DUP-8.
 *
 * Where two implementations disagreed, the disagreement was a live correctness
 * bug, so each agreed behaviour is pinned here rather than left to be
 * re-litigated in a later refactor.
 */
import test from "node:test";
import assert from "node:assert/strict";

import { GridxcelError } from "../src/errors.js";
import {
  validateSheetName,
  normalizeCell,
  normalizeSheet,
  normalizeBook,
  renameSheet,
} from "../src/api.js";
import {
  cellFromValue,
  normalizeSheet as jsonNormalizeSheet,
  worksheetToSheet,
  blankSheet,
  aoaToSheet,
  sheetToAOA,
} from "../src/json.js";
import { decodeCell, toRange, encodeCol } from "../src/address.js";
import { MAX_SHEET_NAME_LENGTH, MAX_COLUMNS, MAX_ROWS, DEFAULT_SHEET_NAME } from "../src/constants.js";

const hasCode = (code) => (e) => e instanceof GridxcelError && e.code === code;

/* ------------------------------------------------------------------ DUP-8 */

test("DUP-8: constants.js is the one home for the limits", () => {
  assert.equal(MAX_SHEET_NAME_LENGTH, 31);
  assert.equal(MAX_COLUMNS, 16384);
  assert.equal(MAX_ROWS, 1048576);
  assert.equal(DEFAULT_SHEET_NAME, "Sheet1");
  assert.equal(encodeCol(MAX_COLUMNS - 1), "XFD");
});

test("DUP-8: DEFAULT_SHEET_NAME is what an unnamed sheet gets", () => {
  assert.equal(blankSheet().name, DEFAULT_SHEET_NAME);
  assert.equal(aoaToSheet([[1]]).name, DEFAULT_SHEET_NAME);
  assert.equal(normalizeSheet([[1]]).name, DEFAULT_SHEET_NAME);
  assert.equal(jsonNormalizeSheet(null).name, DEFAULT_SHEET_NAME);
});

test("DUP-8: an A1 range ends at the last column constant, not a literal", () => {
  const r = toRange(1);
  assert.equal(r.e.c, MAX_COLUMNS - 1);
  assert.equal(r.s.r, 0);
});

test("DUP-8: MAX_COLUMNS is the column stride of the worksheet adapter", () => {
  // Two cells in the same row must not collide in the packed key map, and the
  // last column has to land at MAX_COLUMNS - 1 (row arrays are origin-relative).
  const wide = worksheetToSheet({ "!ref": "A1:XFD2", A1: { v: "left", t: "s" }, XFD2: { v: "right", t: "s" } });
  assert.equal(wide.originCol, 0);
  assert.equal(wide.rows[0][0].v, "left");
  assert.equal(wide.rows[1][MAX_COLUMNS - 1].v, "right");
  const edge = worksheetToSheet({ "!ref": "XFD1:XFD1", XFD1: { v: "edge", t: "s" } });
  assert.equal(edge.originCol, MAX_COLUMNS - 1);
  assert.equal(edge.rows[0][0].v, "edge");
});

/* ------------------------------------------------------------------ DUP-6 */

test("DUP-6: validateSheetName returns the name so callers can assign it", () => {
  assert.equal(validateSheetName("Data"), "Data");
  assert.equal(validateSheetName("x".repeat(MAX_SHEET_NAME_LENGTH)), "x".repeat(31));
});

test("DUP-6: every Excel rule is enforced with BAD_SHEET_NAME", () => {
  const bad = [
    "",
    "x".repeat(MAX_SHEET_NAME_LENGTH + 1),
    "back\\slash",
    "for/ward",
    "question?",
    "star*",
    "bracket[",
    "bracket]",
    "co:lon",
    "'leading",
    "trailing'",
    "History",
    "history",
    "HISTORY",
    null,
    42,
  ];
  for (const name of bad) {
    assert.throws(() => validateSheetName(name), hasCode("BAD_SHEET_NAME"), "should reject " + JSON.stringify(name));
  }
});

test("DUP-6: allowed punctuation is not collateral damage", () => {
  for (const name of ["a-b", "a_b", "a b", "a.b", "a,b", "a(b)", "a&b", "100%", "Sheet 1"]) {
    assert.equal(validateSheetName(name), name);
  }
});

test("DUP-6: taken accepts Set, array and Map, compared case-insensitively", () => {
  assert.equal(validateSheetName("Free", new Set(["Other"])), "Free");
  assert.equal(validateSheetName("Free", ["Other"]), "Free");
  assert.equal(validateSheetName("Free", new Map([["Other", 1]])), "Free");
  for (const taken of [new Set(["Taken"]), ["Taken"], new Map([["Taken", {}]])]) {
    assert.throws(() => validateSheetName("taken", taken), hasCode("BAD_SHEET_NAME"));
    assert.throws(() => validateSheetName("TAKEN", taken), hasCode("BAD_SHEET_NAME"));
  }
  assert.equal(validateSheetName("Free", undefined), "Free");
});

/* ------------------------------------------------------------------ DUP-6 call site */

test("DUP-6: renameSheet rejects an invalid name at the call site", () => {
  const book = normalizeBook({ sheets: [{ name: "Data", rows: [] }] });
  assert.throws(() => renameSheet(book, "Data", "bad/name"), hasCode("BAD_SHEET_NAME"));
  assert.throws(() => renameSheet(book, "Data", "x".repeat(32)), hasCode("BAD_SHEET_NAME"));
  assert.throws(() => renameSheet(book, "Data", ""), hasCode("BAD_SHEET_NAME"));
  // The old behaviour accepted these and failed later, at write time.
  assert.equal(book.sheets[0].name, "Data");
});

test("DUP-6: renameSheet keeps its documented duplicate code, case-insensitively", () => {
  const book = normalizeBook({ sheets: [{ name: "Data", rows: [] }, { name: "Other", rows: [] }] });
  assert.equal(renameSheet(book, "Data", "Renamed").sheets[0].name, "Renamed");
  assert.throws(() => renameSheet(book, "Renamed", "Other"), hasCode("DUPLICATE_SHEET"));
  assert.throws(() => renameSheet(book, "Renamed", "other"), hasCode("DUPLICATE_SHEET"));
  assert.throws(() => renameSheet(book, "Renamed", "History"), hasCode("BAD_SHEET_NAME"));
  assert.throws(() => renameSheet(book, "Missing", "Anything"), hasCode("NO_SHEET"));
});

/* ------------------------------------------------------------------ DUP-1 */

test("DUP-1: a plain object is not a cell - both implementations throw BAD_CELL", () => {
  const notCells = [{ a: 1 }, [], [1, 2], { nested: { deep: true } }, { a: 1, b: 2 }];
  for (const v of notCells) {
    assert.throws(() => cellFromValue(v), hasCode("BAD_CELL"), "json.cellFromValue " + JSON.stringify(v));
    assert.throws(() => normalizeCell(v), hasCode("BAD_CELL"), "api.normalizeCell " + JSON.stringify(v));
  }
  // json.js used to JSON.stringify these, writing text nobody supplied.
  assert.throws(() => aoaToSheet([[{ a: 1 }]]), hasCode("BAD_CELL"));
  assert.throws(() => normalizeSheet([[{ a: 1 }]]), hasCode("BAD_CELL"));
});

test("DUP-1: the two coercion entry points agree value for value", () => {
  const values = [undefined, null, "text", "", 0, -1.5, 42, true, false, new Date(2024, 0, 15)];
  for (const v of values) {
    assert.deepEqual(normalizeCell(v), cellFromValue(v), "disagreement on " + String(v));
  }
  const rich = { v: 1, t: "number", z: "0.00" };
  assert.deepEqual(normalizeCell(rich), cellFromValue(rich));
  assert.deepEqual(normalizeCell({ f: "A1+B1" }), { f: "A1+B1", t: "empty", v: null });
  assert.deepEqual(cellFromValue({ f: "A1+B1" }), normalizeCell({ f: "A1+B1" }));
});

test("DUP-1: non-finite numbers become the documented error cell in both", () => {
  for (const v of [NaN, Infinity, -Infinity]) {
    assert.deepEqual(cellFromValue(v), { v: null, t: "error", e: "#NUM!" });
    assert.deepEqual(normalizeCell(v), cellFromValue(v));
  }
});

test("DUP-1: rich cells keep their fields and gain a default type", () => {
  const cell = normalizeCell({ v: "2024-01-15", z: "yyyy-mm-dd" });
  assert.equal(cell.t, "string");
  assert.equal(cell.z, "yyyy-mm-dd");
  const date = normalizeCell({ v: new Date(2024, 0, 15), t: "date" });
  assert.equal(date.t, "date");
  assert.equal(typeof date.serial, "number");
});

/* ------------------------------------------------------------------ DUP-7 */

test("DUP-7: json.js decodes A1 through address.decodeCell", () => {
  // The deleted private decoder claimed decodeCell could not handle absolute
  // markers; pin the claim that it can, so the stale comment cannot come back.
  assert.deepEqual(decodeCell("$A$1"), { r: 0, c: 0 });
  assert.deepEqual(decodeCell("$B3"), { r: 2, c: 1 });
  const ws = worksheetToSheet({ "!ref": "A1:C3", a1: { v: "low", t: "s" }, B2: { v: "mid", t: "s" }, C3: { v: "high", t: "s" } });
  assert.equal(ws.originRow, 1);
  assert.equal(ws.originCol, 0);
  assert.equal(ws.rows[0][0].v, "low");
  assert.equal(ws.rows[1][1].v, "mid");
  assert.equal(ws.rows[2][2].v, "high");
});

test("DUP-7: worksheet keys that are not A1 references are skipped, not decoded", () => {
  const ws = worksheetToSheet({ "!ref": "A1:A1", A1: { v: 1, t: "n" }, $B$1: { v: 2, t: "n" }, "!merges": [] });
  assert.equal(ws.rows[0][0].v, 1);
  assert.equal(ws.rows[0][1], undefined);
});

/* ------------------------------------------------------------------ DUP-5 */

test("DUP-5: json.normalizeSheet is a shape adapter, not a second builder", () => {
  // gridxcel Sheet: passed through untouched.
  const sheet = aoaToSheet([[1, 2]]);
  assert.equal(jsonNormalizeSheet(sheet), sheet);
  // SheetJS spellings that api.normalizeSheet does not accept.
  assert.equal(jsonNormalizeSheet({ "!ref": "A1:A1", A1: { v: 5, t: "n" } }).rows[0][0].v, 5);
  assert.equal(jsonNormalizeSheet({ "!data": [[{ v: 6, t: "n" }]] }).rows[0][0].v, 6);
  // AOA and Book, as before.
  assert.equal(jsonNormalizeSheet([[7]]).rows[0][0].v, 7);
  assert.equal(jsonNormalizeSheet({ sheets: [aoaToSheet([[8]])] }).rows[0][0].v, 8);
  assert.equal(jsonNormalizeSheet(null).name, DEFAULT_SHEET_NAME);
  assert.throws(() => jsonNormalizeSheet({ nope: true }), hasCode("BAD_SHEET"));
});

test("DUP-5: api.normalizeSheet builds from raw values and keeps its name rule", () => {
  assert.equal(normalizeSheet([[1]], "S").name, "S");
  assert.equal(normalizeSheet({ rows: [[1]] }, "S").name, "S");
  assert.equal(normalizeSheet({ rows: [[1]], name: "Kept" }, "S").name, "Kept");
  const byKey = normalizeSheet([{ A: 1, C: 3 }]);
  assert.equal(byKey.rows[0][0].v, 1);
  assert.equal(byKey.rows[0][2].v, 3);
  assert.equal(byKey.columnCount, 3);
});

test("DUP-5: both normalisers describe an empty sheet identically", () => {
  assert.deepEqual(jsonNormalizeSheet(null), normalizeSheet(undefined));
  assert.deepEqual(jsonNormalizeSheet(null), blankSheet());
});

test("DUP-5: both normalisers share the one coercion rule", () => {
  assert.throws(() => normalizeSheet([[{ a: 1 }]]), hasCode("BAD_CELL"));
  assert.throws(() => jsonNormalizeSheet([[{ a: 1 }]]), hasCode("BAD_CELL"));
  // api.normalizeSheet does not stamp cell.a, json's builder does; the value
  // and type of every cell must still be identical.
  const bare = (sheet) => sheet.rows.map((row) => row.map((c) => (c ? { v: c.v, t: c.t } : c)));
  assert.deepEqual(bare(normalizeSheet([[1, "x"]])), bare(jsonNormalizeSheet([[1, "x"]])));
});

test("DUP-5: the read/convert path still round-trips", () => {
  const sheet = aoaToSheet([["a", "b"], [1, 2]]);
  assert.deepEqual(sheetToAOA(sheet), [["a", "b"], [1, 2]]);
});
