/**
 * Robustness suite for the text-format parsers in src/sylk.js and src/dif.js.
 *
 * The law is the same one the container fuzz suite holds read.js to, because the
 * failure modes are the same kind:
 *
 *   for ANY input text, sylkToSheet / difToSheet either
 *     (1) throws a GridxcelError (a stable code is the contract),
 *     (2) returns a Sheet whose declared extent is bounded, or
 *     (3) -- never -- throws a non-Error, hangs, or aborts the process.
 *
 * Case (3) is why this file exists. Both formats carry row and column numbers as
 * unbounded integers, so a 29-byte file claiming a billion rows is a legal file by
 * the grammar and an out-of-memory abort by a naive reader. An abort is not
 * catchable, so no test in the suite could observe it - only a fuzzer can.
 */

import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { sylkToSheet, sheetToSYLK } from "../src/sylk.js";
import { difToSheet, sheetToDIF } from "../src/dif.js";
import { GridxcelError } from "../src/errors.js";

const CASE_TIMEOUT_MS = Number(process.env.GRIDXCEL_FUZZ_TIMEOUT ?? 10000);

/** A deterministic PRNG, so a failure is reproducible from the seed. */
function rng(seed) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

const FRAGMENTS = [
  "ID;P", "C;Y1;X1;K1", "C;Y1;X1;K", 'C;Y1;X1;K"', 'C;Y1;X1;K""', "E",
  "TABLE", "0,1", '""', "VECTORS", "TUPLES", "DATA", "-1,0", "BOT", "EOD",
  "0,1", "1,0", "V", "NA", "TRUE", "FALSE", "0,999999999", "C;Y999999999",
  "C;X999999999", "B;Y999999999", 'C;Y-1;X-1;K1', "C;Y0;X0;K1", ";", "\r\n",
  "\n", "\r", "\t", " ", '"', "\\", "\u0000", "\uFFFD", "K", "%", "&amp;",
];

/** Build a corpus: structured mutations plus pure noise. */
function corpus() {
  const out = [];
  const r = rng(0x9e3779b9);
  // Known-abusive shapes first, then random assemblies, then pure noise.
  out.push("ID;P\r\nC;Y999999999;X999999999;K1\r\nE\r\n");
  out.push('TABLE\r\n0,1\r\n""\r\nVECTORS\r\n0,999999999\r\n""\r\nDATA\r\n');
  out.push("ID;P\r\nC;Y1;X1;K\"unterminated\r\n");
  out.push("TABLE\r\n0,1\r\n\"\r\nDATA\r\n-1,0\r\n");
  for (let i = 0; i < 120; i++) {
    const parts = [];
    const n = 1 + Math.floor(r() * 12);
    for (let k = 0; k < n; k++) parts.push(FRAGMENTS[Math.floor(r() * FRAGMENTS.length)]);
    out.push(parts.join(r() < 0.5 ? ";" : ""));
  }
  for (let i = 0; i < 40; i++) {
    let s = "";
    const n = 1 + Math.floor(r() * 80);
    for (let k = 0; k < n; k++) s += String.fromCharCode(Math.floor(r() * 128));
    out.push(s);
  }
  return out;
}

/** Run one parser over one input, classifying the outcome. */
function probe(fn, text) {
  const started = Date.now();
  try {
    const sheet = fn(text);
    return {
      ok: true,
      sheet,
      ms: Date.now() - started,
      // A returned sheet must obey the same bounds the parser promises.
      bounded:
        typeof sheet.rowCount === "number" &&
        typeof sheet.columnCount === "number" &&
        sheet.rowCount >= 0 &&
        sheet.rowCount <= 1048576 &&
        sheet.columnCount >= 0 &&
        sheet.columnCount <= 16384 &&
        Array.isArray(sheet.rows) &&
        sheet.rows.length <= 1048576,
    };
  } catch (error) {
    return { ok: false, error, typed: error instanceof GridxcelError, ms: Date.now() - started };
  }
}

describe("fuzz: SYLK and DIF never abort, hang or throw a non-Error", () => {
  for (const [name, parse] of [["sylk", sylkToSheet], ["dif", difToSheet]]) {
    test(name + ": every corpus input is a typed error or a bounded sheet", { timeout: CASE_TIMEOUT_MS * 12 }, () => {
      const cases = corpus();
      const results = cases.map((text, i) => ({ name: name + "#" + i, text, ...probe(parse, text) }));

      const untyped = results.filter((r) => !r.ok && !r.typed);
      const unbounded = results.filter((r) => r.ok && !r.bounded);
      const slow = results.filter((r) => r.ms > 1000);

      assert.deepEqual(
        untyped.map((r) => r.name + ": " + (r.error && r.error.name)),
        [],
        "a non-GridxcelError escaped the parser",
      );
      assert.deepEqual(
        unbounded.map((r) => r.name + ": rows=" + r.sheet.rowCount + " cols=" + r.sheet.columnCount),
        [],
        "a returned sheet declared an unbounded extent",
      );
      assert.deepEqual(
        slow.map((r) => r.name + ": " + r.ms + "ms"),
        [],
        "an input took over a second, which suggests a quadratic scan",
      );

      const typed = results.filter((r) => !r.ok && r.typed).length;
      const recovered = results.filter((r) => r.ok).length;
      console.log("  " + name + " fuzz: " + cases.length + " cases, " + typed + " typed GridxcelError, " + recovered + " recovered sheet(s), 0 untyped, 0 unbounded, 0 slow");
    });
  }
});

test("fuzz: a writer round trip is total - anything it emits, it can read", { timeout: CASE_TIMEOUT_MS * 6 }, () => {
  // The other direction matters too: the parser must accept its own output for
  // arbitrary text, including the characters the quoting rules exist for.
  const r = rng(0x1234abcd);
  const NASTY = ['plain', '', '42', '1e5', 'TRUE', 'NA', '#N/A', 'a;b', 'a"b', 'a\nb', 'a\rb', '  pad  ', '\u00e9\u00fc', '\u0000', 'x'.repeat(200)];
  for (const write of [sheetToSYLK, sheetToDIF]) {
    const read = write === sheetToSYLK ? sylkToSheet : difToSheet;
    for (let i = 0; i < 60; i++) {
      const rows = [];
      const nRows = 1 + Math.floor(r() * 4);
      for (let y = 0; y < nRows; y++) {
        const row = [];
        const nCols = 1 + Math.floor(r() * 4);
        for (let x = 0; x < nCols; x++) {
          const pick = r();
          if (pick < 0.2) { row.push(undefined); continue; }
          if (pick < 0.4) { row.push({ t: "number", v: Math.round(r() * 1e6) / 100 }); continue; }
          if (pick < 0.5) { row.push({ t: "boolean", v: r() < 0.5 }); continue; }
          row.push({ t: "string", v: NASTY[Math.floor(r() * NASTY.length)] });
        }
        rows.push(row);
      }
      const sheet = { name: "S", rows, originRow: 1, originCol: 0, rowCount: rows.length, columnCount: 4, merges: [], columns: [], rowInfo: new Map() };
      let text;
      try {
        text = write(sheet);
      } catch (error) {
        // A documented, typed refusal is a valid outcome: SYLK cannot represent a
        // line break, and emitting something unparseable would be worse.
        assert.ok(error instanceof GridxcelError, "writer threw a non-typed error: " + error);
        continue;
      }
      const back = read(text);
      assert.equal(back.rows.length, rows.length, "row count survives for input #" + i + ": " + JSON.stringify(text.slice(0, 120)));
    }
  }
});
