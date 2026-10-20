/**
 * The explicit-separator DSV API (docs/API.md "gridxcel/csv").
 *
 * These two exports were documented and type-declared but had no test, and
 * that gap was hiding a real defect: they were plain aliases of the CSV
 * functions, so the documented second-argument delimiter was ignored -
 * dsvToSheet("a;b\n1;2\n", ";") parsed ONE column and sheetToDSV(sheet, ";")
 * emitted commas. The doc example still passed because both sides ignored the
 * argument in the same way, so the round trip was self-consistent and wrong.
 *
 * Every expectation below is the documented contract, checked against the
 * implementation rather than against the other function.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { csvToSheet, dsvToSheet, parseDSV, sheetToCSV, sheetToDSV } from "../src/csv.js";
import { DEFAULT_SHEET_NAME } from "../src/constants.js";
import { aoaToSheet } from "../src/json.js";

/* ------------------------------------------------- dsvToSheet (documented) */

test("dsvToSheet(text, delimiter) honours the explicit separator", () => {
  const sheet = dsvToSheet("a;b\n1;2\n", ";");
  assert.equal(sheet.columnCount, 2, "a semicolon file must parse into two columns");
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["a", "b"]);
  assert.deepEqual(sheet.rows[1].map((c) => c.v), [1, 2], "types are still inferred");
  assert.equal(sheet.name, DEFAULT_SHEET_NAME);

  const tab = dsvToSheet("a\tb\n", "\t");
  assert.deepEqual(tab.rows[0].map((c) => c.v), ["a", "b"]);

  const pipe = dsvToSheet("a||b\n", "||");
  assert.deepEqual(pipe.rows[0].map((c) => c.v), ["a", "b"], "a multi-character separator works");

  // A quoted field keeps the separator as data.
  const quoted = dsvToSheet('a;"b;c"\n', ";");
  assert.deepEqual(quoted.rows[0].map((c) => c.v), ["a", "b;c"]);
});

test("dsvToSheet accepts options third, or alone", () => {
  const withOpts = dsvToSheet("a;b\n", ";", { header: ["x", "y"] });
  assert.deepEqual(withOpts.rows[0].map((c) => c.v), ["x", "y"]);
  assert.deepEqual(withOpts.rows[1].map((c) => c.v), ["a", "b"]);

  // The option-object form keeps working: it is what the rest of the module
  // takes and what callers were already passing.
  const optsForm = dsvToSheet("a;b\n", { delimiter: ";" });
  assert.equal(optsForm.columnCount, 2);
  const oneArg = dsvToSheet("a,b\n");
  assert.deepEqual(oneArg.rows[0].map((c) => c.v), ["a", "b"], "no second argument means comma");

  // A typed sheet still round-trips through the explicit form.
  const sheet = dsvToSheet("na;1\n", ";", { quotedStrings: true });
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["na", 1]);
});

/* ------------------------------------------------- sheetToDSV (documented) */

test("sheetToDSV(sheet, delimiter) emits the explicit separator", () => {
  const sheet = aoaToSheet([[1, 2], [3, 4]]);
  assert.equal(sheetToDSV(sheet, ";"), "1;2\n3;4\n");
  assert.equal(sheetToDSV(sheet, "\t"), "1\t2\n3\t4\n");
  assert.equal(sheetToDSV(sheet, "|"), "1|2\n3|4\n");
  assert.equal(sheetToDSV(sheet), "1,2\n3,4\n", "no separator means comma");
  // Fields containing the separator are quoted with the configured quote.
  assert.equal(sheetToDSV(aoaToSheet([["a;b", 1]]), ";"), '"a;b";1\n');
  // Explicit separator plus options.
  assert.equal(sheetToDSV(aoaToSheet([[1, 2]]), ";", { lineEnding: "\r\n" }), "1;2\r\n");
  // Option-object form, unchanged for existing callers.
  assert.equal(sheetToDSV(sheet, { delimiter: "|" }), "1|2\n3|4\n");
});

test("the documented round trip survives a real round trip", () => {
  const text = "name;qty\nwidget;2\n";
  const sheet = dsvToSheet(text, ";");
  assert.equal(sheetToDSV(sheet, ";"), text);
  // With the default comma separator the semicolon is ordinary data, so the
  // CSV form needs no quoting and the value survives as one field.
  const csv = sheetToCSV(sheet);
  assert.equal(csv, "name,qty\nwidget,2\n");
  // Converting the DSV to CSV keeps the table: two columns, same values.
  const back = csvToSheet(csv);
  assert.deepEqual(back.rows[0].map((c) => c.v), ["name", "qty"]);
  assert.deepEqual(back.rows[1].map((c) => c.v), ["widget", 2]);
  // A field that contains the separator is quoted on the way back out.
  const quoted = dsvToSheet('"a;b";1\n', ";");
  assert.deepEqual(quoted.rows[0].map((c) => c.v), ["a;b", 1]);
  assert.equal(sheetToDSV(quoted, ";"), '"a;b";1\n');
});

test("an unusable option is rejected even when nothing would be rendered", () => {
  const empty = aoaToSheet([]);
  assert.throws(() => sheetToDSV(empty, { quote: "xx" }), (err) => err.code === "BAD_OPTION");
  assert.throws(() => sheetToDSV(empty, { delimiter: "" }), (err) => err.code === "BAD_OPTION");
  assert.throws(() => sheetToDSV(aoaToSheet([[1]]), ";", { quote: "" }), (err) => err.code === "BAD_OPTION");
  assert.throws(() => dsvToSheet("a\n", { delimiter: "" }), (err) => err.code === "BAD_OPTION");
});

/* ---------------------------------------------------- parseDSV (documented) */

test("parseDSV returns { text, quoted } records", () => {
  const records = parseDSV('a;"b;c"\n1;2\n', { delimiter: ";" });
  assert.equal(records.length, 2);
  assert.deepEqual(records[0], [{ text: "a", quoted: false }, { text: "b;c", quoted: true }]);
  assert.deepEqual(records[1], [{ text: "1", quoted: false }, { text: "2", quoted: false }]);
  assert.deepEqual(parseDSV(""), [], "empty text is no records");
});

test("parseDSV: the tolerance the doc comment promises", () => {
  // \r\n, \n and a lone \r all end a record.
  assert.deepEqual(parseDSV("a\r\nb\r\n").map((r) => r[0].text), ["a", "b"]);
  assert.deepEqual(parseDSV("a\rb\r").map((r) => r[0].text), ["a", "b"]);
  assert.deepEqual(parseDSV("a\nb\n").map((r) => r[0].text), ["a", "b"]);
  // A trailing terminator does not invent a record; a blank line does.
  assert.equal(parseDSV("a\n").length, 1);
  assert.deepEqual(parseDSV("a\n\nb\n").map((r) => r.map((f) => f.text)), [["a"], [""], ["b"]]);
  // "" inside a quoted field is one quote.
  assert.equal(parseDSV('a,"b""c"')[0][1].text, 'b"c');
  assert.equal(parseDSV('a,"b""c"')[0][1].quoted, true);
  // A quote that does not start a field is literal text.
  assert.equal(parseDSV('a,b"c')[0][1].text, 'b"c');
  assert.equal(parseDSV('a,b"c')[0][1].quoted, false);
  // Text after a closing quote is kept, never silently dropped.
  assert.equal(parseDSV('a,"b"c')[0][1].text, "bc");
  assert.equal(parseDSV('a,"b"c')[0][1].quoted, true);
  // An unterminated quote runs to end of input.
  assert.deepEqual(parseDSV('a,"b')[0][1], { text: "b", quoted: true });
  // A custom quote character.
  assert.deepEqual(parseDSV("a,'b,c'", { quote: "'" })[0][1], { text: "b,c", quoted: true });
  // The separator can be several characters; an empty one is refused.
  assert.deepEqual(parseDSV("a||b", { delimiter: "||" })[0].map((f) => f.text), ["a", "b"]);
  assert.throws(() => parseDSV("a,b", { delimiter: "" }), (err) => err.code === "BAD_OPTION");
  // parseDSV is the raw layer: it does not strip a BOM (csvToSheet does).
  assert.equal(parseDSV("\ufeffa,b")[0][0].text, "\ufeffa");
  assert.equal(csvToSheet("\ufeffa,b").rows[0][0].v, "a");
});

/* --------------------------------------------------------------- F13 residue */

test("F13: the default sheet name comes from constants.js", () => {
  assert.equal(DEFAULT_SHEET_NAME, "Sheet1");
  assert.equal(csvToSheet("a\n").name, DEFAULT_SHEET_NAME);
  assert.equal(csvToSheet("a\n", { name: "Custom" }).name, "Custom");
  assert.equal(csvToSheet("a\n", { sheet: "FromSheetKey" }).name, "FromSheetKey");
  assert.equal(csvToSheet("a\n", { sheet: "" }).name, DEFAULT_SHEET_NAME, "an empty name falls back");
  assert.equal(dsvToSheet("a;b\n", ";").name, DEFAULT_SHEET_NAME);
  // And the literal is gone from the source, so the next divergence cannot hide.
  const src = readFileSync("src/csv.js", "utf8");
  assert.equal(/"Sheet1"/.test(src), false, 'csv.js must not contain the literal "Sheet1"');
  assert.ok(/DEFAULT_SHEET_NAME/.test(src), "csv.js must import the constant");
});
