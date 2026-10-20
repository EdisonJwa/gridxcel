import test from "node:test";
import assert from "node:assert/strict";
import { encodeXmlText, encodeXmlAttr, decodeOoxmlEscapes } from "../src/xml.js";
import { writeWorkbookSync } from "../src/write.js";
import { readWorkbookSync } from "../src/read.js";

/**
 * UTF-16 code units that do not form a character.
 *
 * JavaScript strings are UTF-16, so they can hold a lone surrogate - half of a
 * pair. It has no UTF-8 encoding, and Buffer/TextEncoder substitute U+FFFD
 * silently. That made a write a one-way door: "a\ud800b" came back "a\ufffdb"
 * with no error anywhere. These are the cases that have to survive.
 */

const LONE = {
  "high": "a\ud800b",
  "low": "a\udc00b",
  "leading high": "\ud800",
  "trailing low": "a\udc00",
  "both": "\ud800\udc00\ud800",
};

test("encodeXmlText: a lone surrogate is escaped, not replaced, and round trips", () => {
  for (const [label, value] of Object.entries(LONE)) {
    const encoded = encodeXmlText(value);
    assert.ok(!encoded.includes("\ufffd"), label + ": must not substitute U+FFFD");
    assert.match(encoded, /_x[Dd][89A-Fa-f][0-9A-Fa-f]{2}_/, label + ": should use the _xHHHH_ escape");
    assert.equal(decodeOoxmlEscapes(encoded), value, label + ": must round trip exactly");
  }
});

test("encodeXmlText: a well-formed pair is left alone", () => {
  // The escape exists for a LONE unit; escaping half of a valid pair would split
  // a character that had no problem.
  for (const value of ["\ud83d\ude00", "a\ud83d\ude00b", "\ud83d\ude00\ud83d\ude01"]) {
    assert.equal(encodeXmlText(value), value, JSON.stringify(value) + " should pass through");
  }
});

test("encodeXmlAttr: a lone surrogate is escaped too", () => {
  const encoded = encodeXmlAttr("x\ud800\n");
  assert.ok(!encoded.includes("\ufffd"));
  assert.match(encoded, /_xD800_/);
});

test("writer: a lone surrogate survives a write and read", () => {
  const value = "a\ud800b\udc00c";
  const book = {
    sheetNames: ["S"],
    sheets: [{ name: "S", rows: [[{ t: "string", v: value }]], originRow: 1, originCol: 0, rowCount: 1, columnCount: 1, merges: [], columns: [], rowInfo: new Map() }],
  };
  const back = readWorkbookSync(writeWorkbookSync(book));
  const got = back.sheets[0].rows[0][0].v;
  assert.equal(got, value, "the value must be byte-identical through a round trip");
  assert.deepEqual([...got].map((c) => c.codePointAt(0)), [...value].map((c) => c.codePointAt(0)));
});

test("writer: unicode, control characters and escape lookalikes all round trip", () => {
  const values = ["plain", "", "=1+1", "a\nb", "a\rb", "\u0000", "\u0001\u001f", "\ud83d\ude00", "x".repeat(33000), "<&>", "]]>", "_x0041_", "\ufeff", "  ", "\t", "\u00e9\u00fc\u4f60\u597d"];
  const rows = values.map((v) => [{ t: "string", v }]);
  const book = {
    sheetNames: ["S"],
    sheets: [{ name: "S", rows, originRow: 1, originCol: 0, rowCount: rows.length, columnCount: 1, merges: [], columns: [], rowInfo: new Map() }],
  };
  const back = readWorkbookSync(writeWorkbookSync(book));
  for (let i = 0; i < values.length; i++) {
    assert.equal(back.sheets[0].rows[i][0].v, values[i], "value " + i + " (" + JSON.stringify(values[i].slice(0, 12)) + ") must round trip");
  }
});
