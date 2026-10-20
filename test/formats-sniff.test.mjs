import test from "node:test";
import assert from "node:assert/strict";
import { read } from "../src/compat.js";

/**
 * The compat surface picks a text format from the content. Every case here is a
 * shape the sniffer got wrong or could plausibly get wrong.
 */

const SYLK = 'ID;P\r\nB;Y2;X2\r\nC;Y1;X1;K"a"\r\nC;Y2;X2;K1\r\nE\r\n';
const DIF = ['TABLE', '0,1', '"EXCEL"', 'VECTORS', '0,2', '""', 'TUPLES', '0,2', '""',
  'DATA', '0,0', '""', '-1,0', 'BOT', '1,0', '"a"', '1,0', '"b"',
  '-1,0', 'BOT', '0,1', 'V', '0,2', 'V', '-1,0', 'EOD'].join("\r\n") + "\r\n";

test("sniff: a BOM does not stop SYLK or DIF from being recognised", () => {
  // The sniffer compared against the BOM and fell through to CSV, so a
  // BOM-prefixed SYLK file was parsed as one long header row.
  for (const [label, text] of [["sylk", SYLK], ["dif", DIF]]) {
    const plain = read(text);
    const withBom = read("\ufeff" + text);
    assert.equal(withBom.sheetNames.length, 1, label);
    assert.deepEqual(
      withBom.sheets[0].rows.map((r) => r.map((c) => c && c.v)),
      plain.sheets[0].rows.map((r) => r.map((c) => c && c.v)),
      label + ": a BOM must not change the parse",
    );
  }
});

test("sniff: a CSV whose first field is ID or TABLE is still CSV", () => {
  // The sniff requires the full signature, not the bare word, so these stay CSV.
  for (const csv of ['ID,value\n1,2\n', 'TABLE,x\n1,2\n']) {
    const book = read(csv);
    assert.equal(book.sheets[0].rows.length, 2, JSON.stringify(csv));
  }
  // And a cell that is exactly "ID" with no separator is one row, not a workbook.
  assert.equal(read("ID").sheets[0].rows.length, 1);
  assert.equal(read("TABLE").sheets[0].rows.length, 1);
});

test("sniff: an explicit bookType overrides the guess", () => {
  // A CSV that really does begin "ID;" is the ambiguous case, and the option is
  // the documented way out - it must win over the sniff in both directions.
  const looksLikeSylk = "ID;P\r\nE\r\n";
  assert.equal(read(looksLikeSylk).sheets[0].rows.length, 0, "sniffed as SYLK");
  const forced = read(looksLikeSylk, { bookType: "csv" });
  assert.ok(forced.sheets[0].rows.length >= 1, "bookType csv overrides the sniff");

  // A SYLK body forced to DIF must fail as DIF, not silently fall back.
  assert.throws(() => read(SYLK, { bookType: "dif" }), (e) => e.code === "BAD_DIF");
});

test("sniff: an unknown bookType is refused with the supported list", () => {
  assert.throws(() => read("a,b\n1,2\n", { bookType: "xlsb" }), (e) => {
    assert.match(e.message, /csv, txt, sylk and dif/);
    return true;
  });
});

test("sniff: garbage is a typed error or a sheet, never a crash", () => {
  for (const text of ["", "   ", "\u0000\u0001\u0002", "{\"a\":1}", "not a spreadsheet", "\ufeff"]) {
    const book = read(text);
    assert.equal(book.sheetNames.length, 1, "always one sheet for text input");
  }
});
