/**
 * CSV export safety: the formula-injection guard (security review MED-5).
 *
 * sheetToCSV prefixes a text field with an apostrophe when its first character
 * is one of = + - @ TAB CR LF, so a spreadsheet opens the field as text instead
 * of evaluating it. The guard is on by default, is turned off with
 * escapeFormulae: false, and never applies to numeric or date cells.
 *
 * Run: node --test test/csv.test.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

import { csvToSheet, sheetToCSV } from "../src/csv.js";

/** The rendered field for a one-cell sheet, without the row terminator. */
function field(text, opts) {
  const out = sheetToCSV([[text]], opts);
  assert.ok(out.endsWith("\n"), "every row is terminated");
  return out.slice(0, -1);
}

/** Reverse RFC 4180 quoting, so a quoted field can be compared to its value. */
function unquote(text) {
  if (text.length >= 2 && text.startsWith('"') && text.endsWith('"')) {
    return text.slice(1, -1).split('""').join('"');
  }
  return text;
}

/**
 * Every character that makes a spreadsheet read a field as a formula, with a
 * payload that would be dangerous if it were evaluated.
 *
 * TAB, CR and LF are the MED-5 bypass: applications strip leading whitespace
 * before deciding that a cell is a formula, so "\t=1+1" is evaluated exactly
 * like "=1+1".
 */
const LEADS = [
  ["=", "=1+1"],
  ["+", "+1+1"],
  ["-", "-1+1"],
  ["@", "@SUM(1)"],
  ["TAB", "\t=1+1"],
  ["CR", "\r=1+1"],
  ["LF", "\n=1+1"],
];

/* ------------------------------------------------- one test per character */

for (const [name, payload] of LEADS) {
  test("csv: a leading " + name + " in a text field is neutralised", () => {
    const rendered = field(payload);
    // CR and LF force RFC 4180 quoting, so compare the field's *value*: what a
    // spreadsheet ends up with after unquoting. It must start with the
    // apostrophe, never with the dangerous character.
    const value = unquote(rendered);
    assert.equal(
      value[0],
      "'",
      "expected an apostrophe first, got " + JSON.stringify(rendered),
    );
    // The apostrophe is the only thing added: nothing of the payload is lost.
    assert.equal(value.slice(1), payload, "the payload must survive unchanged");
    assert.notEqual(rendered, payload, "the raw field must not be emitted verbatim");
  });
}

/* ------------------------------------------------------------ the bypass */

test("csv: the TAB bypass reported as MED-5 is closed", () => {
  // Before the fix this was emitted verbatim and Excel evaluated it.
  assert.notEqual(field("\t=1+1"), "\t=1+1");
  assert.equal(field("\t=1+1"), "'\t=1+1");

  // CR and LF were only quote-wrapped, which does not stop the evaluation:
  // the field content still began with the operator.
  assert.equal(field("\r=1+1"), '"\'\r=1+1"');
  assert.equal(field("\n=1+1"), '"\'\n=1+1"');

  // A run of them is a bypass too, so the guard tests the first character only
  // and does not try to count whitespace.
  assert.equal(field("\t\t=1+1"), "'\t\t=1+1");
  assert.equal(field("\r\n=1+1"), '"\'\r\n=1+1"');
});

test("csv: the guard is one-way, like Excel's own CSV output", () => {
  // Reading the escaped CSV back does not strip the apostrophe again; the
  // field is text from then on, which is the point of the guard.
  const back = csvToSheet(sheetToCSV([["\t=1+1"]]));
  assert.equal(back.rows[0][0].v, "'\t=1+1");
  assert.equal(back.rows[0][0].t, "string");
});

/* ------------------------------------------------------------- opt-out */

test("csv: escapeFormulae: false still disables the guard for every lead", () => {
  for (const [name, payload] of LEADS) {
    const rendered = field(payload, { escapeFormulae: false });
    assert.equal(unquote(rendered), payload, name + " must be written verbatim when the guard is off");
  }
});

test("csv: the default is on, and an explicit true matches it", () => {
  assert.equal(field("=1+1"), field("=1+1", { escapeFormulae: true }));
  assert.equal(field("\t=1+1"), field("\t=1+1", { escapeFormulae: true }));
});

/* --------------------------------------------- text only, numbers untouched */

test("csv: only text fields are guarded, so -2.5 stays numeric and unquoted", () => {
  const text = sheetToCSV([[-2.5, "-2.5", 42, -2, "+3"]]);
  assert.equal(
    text,
    "-2.5,'-2.5,42,-2,'+3\n",
    "a numeric cell keeps its exact shape; a string that looks numeric is text",
  );

  // The same values through a named-typed cell: t decides, not the text.
  assert.equal(sheetToCSV([[{ v: -2.5, t: "number" }]]), "-2.5\n");
  assert.equal(sheetToCSV([[{ v: "-2.5", t: "string" }]]), "'-2.5\n");
});

test("csv: only the first character triggers the guard", () => {
  // Operators and tabs anywhere else in the field are ordinary text.
  assert.equal(sheetToCSV([["a=b", "1+1", "x\ty", "a\tb"]]), "a=b,1+1,x\ty,a\tb\n");

  // A lead is a lead even when a space follows it: the first character is what
  // a spreadsheet looks at.
  assert.equal(sheetToCSV([["- 2"]]), "'- 2\n");

  // Leading spaces are NOT in the guarded set. The set is exactly = + - @ TAB
  // CR LF; a field that starts with a space starts with ordinary text.
  assert.equal(sheetToCSV([["  -2"]]), "  -2\n");
});

test("csv: an empty text field is written as an empty field", () => {
  assert.equal(sheetToCSV([[""]]), "\n");
  assert.equal(sheetToCSV([["", "=1+1"]]), ",'=1+1\n");
});
