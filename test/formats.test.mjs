import test from "node:test";
import assert from "node:assert/strict";
import { aoaToSheet } from "../src/json.js";
import { sylkToSheet, sheetToSYLK } from "../src/sylk.js";
import { difToSheet, sheetToDIF } from "../src/dif.js";

/** A sheet with one of every cell type, so a round trip has something to lose. */
function typedSheet() {
  return aoaToSheet([
    ["text", "number", "boolean", "date"],
    ["ada", 42, true, new Date(2024, 2, 15)],
    ['says "hi"; and more', -2.5, false, new Date(1999, 11, 31)],
  ]);
}

const flat = (s) => s.rows.map((r) => r.map((c) => (c ? c.t + "=" + String(c.t === "date" ? c.serial : c.v) : "")).join(",")).join("/");

/* ------------------------------------------------------------------- SYLK */

// The example from the SYLK specification, so the reader is tested against
// something a different implementation wrote, not against its own writer.
const SYLK_SPEC = [
  "ID;P",
  "B;Y2;X3",
  'C;Y1;X1;K"Name"',
  'C;Y1;X2;K"Score"',
  'C;Y1;X3;K"Pass"',
  'C;Y2;X1;K"ada"',
  "C;Y2;X2;K42",
  "C;Y2;X3;KTRUE",
  "E",
].join("\r\n") + "\r\n";

test("sylk: reads the specification's example with correct types", () => {
  const sheet = sylkToSheet(SYLK_SPEC);
  assert.equal(sheet.rowCount, 2);
  assert.equal(sheet.columnCount, 3);
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["Name", "Score", "Pass"]);
  assert.deepEqual(sheet.rows[1].map((c) => c.t), ["string", "number", "boolean"]);
  assert.equal(sheet.rows[1][1].v, 42);
  assert.equal(sheet.rows[1][2].v, true);
});

test("sylk: round trips types, and the output re-reads identically", () => {
  const first = sylkToSheet(sheetToSYLK(typedSheet()));
  assert.equal(flat(sylkToSheet(sheetToSYLK(first))), flat(first), "two round trips must agree");
  // Text, number and boolean survive; a date is written as its serial (SYLK has
  // no date type) and comes back as a number, which is the documented behaviour.
  assert.equal(first.rows[1][0].v, "ada");
  assert.equal(first.rows[1][1].v, 42);
  assert.equal(first.rows[1][2].v, true);
});

test("sylk: a quoted field may contain the separator and a doubled quote", () => {
  const sheet = sylkToSheet(sheetToSYLK(typedSheet()));
  assert.equal(sheet.rows[2][0].v, 'says "hi"; and more');
});

test("sylk: omitting Y or X reuses the previous coordinate, as real files do", () => {
  // "C;Y1;X1" then "C;X2" means row 1 column 2.
  const text = 'ID;P\r\nC;Y1;X1;K"a"\r\nC;X2;K"b"\r\nC;X3;K"c"\r\nE\r\n';
  const sheet = sylkToSheet(text);
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["a", "b", "c"]);
});

test("sylk: rejects input that is not a SYLK file, with a typed error", () => {
  for (const bad of ["", "hello", 'a,b\n1,2\n', "B;Y2;X2\r\nE\r\n"]) {
    assert.throws(() => sylkToSheet(bad), (e) => e.code === "BAD_SYMBOL", "should reject " + JSON.stringify(bad));
  }
});

/* -------------------------------------------------------------------- DIF */

// The example from the DIF technical specification.
const DIF_SPEC = [
  "TABLE", "0,1", '"EXCEL"',
  "VECTORS", "0,2", '""',
  "TUPLES", "0,2", '""',
  "DATA", "0,0", '""',
  "-1,0", "BOT",
  "1,0", '"Name"',
  "1,0", '"Age"',
  "-1,0", "BOT",
  "1,0", '"Bob"',
  "0,34", "V",
  "-1,0", "BOT",
  "1,0", '"Sheetal"',
  "0,22", "V",
  "-1,0", "EOD",
].join("\r\n") + "\r\n";

test("dif: reads the specification's example with correct types", () => {
  const sheet = difToSheet(DIF_SPEC);
  assert.equal(sheet.rowCount, 3);
  assert.deepEqual(sheet.rows[0].map((c) => c.v), ["Name", "Age"]);
  assert.deepEqual(sheet.rows[1].map((c) => c.v), ["Bob", 34]);
  assert.equal(sheet.rows[1][1].t, "number");
  assert.deepEqual(sheet.rows[2].map((c) => c.v), ["Sheetal", 22]);
});

test("dif: the numeric value comes from the pair, not the keyword line", () => {
  // 0,34 / V - a reader that takes the second line as the value gets 0.
  const sheet = difToSheet('TABLE\r\n0,1\r\n""\r\nVECTORS\r\n0,1\r\n""\r\nTUPLES\r\n0,1\r\n""\r\nDATA\r\n0,0\r\n""\r\n-1,0\r\nBOT\r\n0,123.5\r\nV\r\n-1,0\r\nEOD\r\n');
  assert.equal(sheet.rows[0][0].v, 123.5);
  assert.equal(sheet.rows[0][0].t, "number");
});

test("dif: round trips, and the output re-reads identically", () => {
  const first = difToSheet(sheetToDIF(typedSheet()));
  assert.equal(flat(difToSheet(sheetToDIF(first))), flat(first), "two round trips must agree");
  assert.equal(first.rows[1][0].v, "ada");
  assert.equal(first.rows[1][1].v, 42);
  assert.equal(first.rows[1][2].v, true);
});

test("dif: rejects input that is not a DIF file, with a typed error", () => {
  for (const bad of ["", "hello", 'a,b\n1,2\n', '0,1\r\n""\r\n']) {
    assert.throws(() => difToSheet(bad), (e) => e.code === "BAD_DIF", "should reject " + JSON.stringify(bad));
  }
});

/* ------------------------------------------------------- through the API */

test("formats: the compat surface sniffs SYLK and DIF, and still reads CSV", async () => {
  const { read, write } = await import("../src/compat.js");

  const sylk = read(SYLK_SPEC);
  assert.deepEqual(sylk.sheetNames, ["Sheet1"]);
  assert.equal(sylk.sheets[0].rows[1][1].v, 42);

  const dif = read(DIF_SPEC);
  assert.equal(dif.sheets[0].rows[1][1].v, 34);

  // CSV has no signature, so it must still be the fallback.
  const csv = read("a,b\n1,2\n");
  assert.deepEqual(csv.sheets[0].rows[1].map((c) => c.v), [1, 2]);
});

test("formats: read sniffs, and an explicit bookType overrides the sniff", async () => {
  const { read } = await import("../src/compat.js");
  // A CSV whose first cell is literally "ID" would otherwise look like SYLK.
  const looksSylk = 'ID,value\n1,2\n';
  const sniffed = read(looksSylk);
  assert.equal(sniffed.sheets[0].rows.length, 2, "mis-sniffed as SYLK");
  const forced = read(looksSylk, { bookType: "csv" });
  assert.deepEqual(forced.sheets[0].rows[1].map((c) => c.v), [1, 2]);
});

test("formats: write emits SYLK and DIF, and each reads back", async () => {
  const { read, write } = await import("../src/compat.js");
  const wb = [["a", "b"], [1, 2]];
  for (const bookType of ["sylk", "dif"]) {
    const out = write(wb, { bookType });
    assert.equal(typeof out, "string", bookType + " is text");
    const back = read(out, { bookType });
    assert.deepEqual(back.sheets[0].rows[0].map((c) => c.v), ["a", "b"], bookType);
    assert.deepEqual(back.sheets[0].rows[1].map((c) => c.v), [1, 2], bookType);
  }
});

/* ------------------------------------------------ regression: review findings */

test("sylk: a quoted K decides the type, not the characters", () => {
  // K"1e5" is the string "1e5"; K1e5 is the number 100000. Looking only at the
  // characters silently converted one into the other.
  const text = 'ID;P\r\nC;Y1;X1;K"1e5"\r\nC;Y1;X2;K1e5\r\nC;Y2;X1;K"42"\r\nC;Y2;X2;K42\r\nE\r\n';
  const sheet = sylkToSheet(text);
  assert.equal(sheet.rows[0][0].t, "string");
  assert.equal(sheet.rows[0][0].v, "1e5");
  assert.equal(sheet.rows[0][1].t, "number");
  assert.equal(sheet.rows[0][1].v, 100000);
  assert.equal(sheet.rows[1][0].t, "string", 'K"42" is text');
  assert.equal(sheet.rows[1][0].v, "42");
  assert.equal(sheet.rows[1][1].t, "number", "K42 is a number");
});

test("sylk: the writer quotes a string that would otherwise read as a number", () => {
  const sheet = aoaToSheet([["42", "1e5", "TRUE", "NA", "plain"]]);
  const back = sylkToSheet(sheetToSYLK(sheet));
  assert.deepEqual(back.rows[0].map((c) => c.t), ["string", "string", "string", "string", "string"]);
  assert.deepEqual(back.rows[0].map((c) => c.v), ["42", "1e5", "TRUE", "NA", "plain"]);
});

test("sylk: an empty string survives a round trip", () => {
  const sheet = aoaToSheet([["a", "", "c"]]);
  const back = sylkToSheet(sheetToSYLK(sheet));
  assert.equal(back.rows[0][1].t, "string");
  assert.equal(back.rows[0][1].v, "");
});

test("sylk: a coordinate that would allocate gigabytes is refused, not attempted", () => {
  // Twenty bytes claiming a billion rows. Before the bound this exhausted the
  // heap and aborted the process, which is not a catchable failure.
  for (const text of [
    "ID;P\r\nC;Y999999999;X1;K1\r\nE\r\n",
    "ID;P\r\nC;Y1;X999999999;K1\r\nE\r\n",
    "ID;P\r\nB;Y999999999;X1\r\nE\r\n",
  ]) {
    assert.throws(() => sylkToSheet(text), (e) => e.code === "RANGE_TOO_LARGE", "should refuse " + JSON.stringify(text.slice(0, 30)));
  }
});

test("dif: an empty string is a value, not an absent cell", () => {
  const sheet = aoaToSheet([["a", "", "c"], ["", "b", ""]]);
  const back = difToSheet(sheetToDIF(sheet));
  for (let r = 0; r < 2; r++) {
    for (let c = 0; c < 3; c++) {
      assert.ok(back.rows[r][c], "cell " + r + "," + c + " must exist");
      assert.equal(back.rows[r][c].t, "string");
    }
  }
  assert.equal(back.rows[0][1].v, "");
});

test("dif: a header count that would allocate gigabytes is refused", () => {
  for (const [header, count] of [["VECTORS", 999999999], ["TUPLES", 999999999]]) {
    const text = [
      "TABLE", "0,1", '""',
      header, "0," + count, '""',
      "DATA", "0,0", '""',
      "-1,0", "BOT", "0,1", "V", "-1,0", "EOD",
    ].join("\r\n") + "\r\n";
    assert.throws(() => difToSheet(text), (e) => e.code === "RANGE_TOO_LARGE", "should refuse " + header);
  }
});

test("formats: a row that ends in empty cells keeps its column count", () => {
  const sheet = aoaToSheet([["a", "b", "c"], ["", "", ""]]);
  for (const [to, from] of [[sheetToSYLK, sylkToSheet], [sheetToDIF, difToSheet]]) {
    const back = from(to(sheet));
    assert.equal(back.columnCount, 3, "columnCount must survive trailing empties");
  }
});
