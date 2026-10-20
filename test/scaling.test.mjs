/**
 * Parser scaling tests.
 *
 * The scanner used to call String#indexOf for the next '"' and '\'' from each
 * tag's start. Those searches are unbounded, so on a worksheet containing no
 * apostrophes - which is every worksheet written by Excel, openpyxl or this
 * library - every attribute-less tag scanned the whole remaining buffer for a
 * character that was not there. The cost was quadratic in the buffer size: the
 * same document parsed in 724 ms when fed in 16 KiB pieces and 18,730 ms when
 * fed in 1 MiB pieces.
 *
 * These tests pin the scaling rather than an absolute time, so they stay
 * meaningful on a slow or loaded machine.
 */

import { test } from "node:test";
import assert from "node:assert";
import { parseSheetXml, parseSheetXmlSync } from "../src/sheet-reader.js";

/** A worksheet of \`rows\` rows, 8 columns, no apostrophes anywhere. */
function worksheet(rows) {
  const out = ['<?xml version="1.0"?><worksheet><sheetData>'];
  for (let r = 1; r <= rows; r++) {
    out.push('<row r="' + r + '">');
    for (let c = 0; c < 8; c++) {
      // Alternate inline text and numeric cells so both the <is><t> and the
      // attribute-less <v> paths are exercised.
      if (c % 2 === 0) out.push('<c r="' + col(c) + r + '" t="inlineStr"><is><t>row' + r + "c" + c + "</t></is></c>");
      else out.push('<c r="' + col(c) + r + '"><v>' + (r * 8 + c) + "</v></c>");
    }
    out.push("</row>");
  }
  out.push("</sheetData></worksheet>");
  return out.join("");
}

function col(c) {
  return String.fromCharCode(65 + c);
}

async function* inChunks(text, size) {
  for (let i = 0; i < text.length; i += size) yield text.slice(i, i + size);
}

const opts = { name: "S", sharedStrings: [], styles: { isDate: () => false } };

const best = async (fn, n) => {
  await fn();
  let b = Infinity;
  for (let i = 0; i < n; i++) {
    const s = performance.now();
    await fn();
    const d = performance.now() - s;
    if (d < b) b = d;
  }
  return b;
};

test("a document with no apostrophes still parses every cell", () => {
  const xml = worksheet(200);
  const sheet = parseSheetXmlSync(xml, opts);
  assert.equal(sheet.rowCount, 200);
  assert.equal(sheet.rows.length, 200);
  assert.equal(sheet.rows[0][0].v, "row1c0");
  assert.equal(sheet.rows[0][1].v, 9);
  assert.equal(sheet.rows[199][0].v, "row200c0");
  assert.equal(sheet.rows[199][6].v, "row200c6");
});

test("a quoted attribute containing '>' does not end the tag early", () => {
  const xml = '<?xml version="1.0"?><worksheet><sheetData>'
    + '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>'
    + "</sheetData></worksheet>";
  const sheet = parseSheetXmlSync(xml, opts);
  assert.equal(sheet.rows[0][0].v, "x");
});

test("attribute values may contain '>' and still round-trip", () => {
  // A formula or string containing '>' must not truncate the tag. The scanner
  // only treats '>' as a terminator when it sits outside a quoted value.
  const xml = '<?xml version="1.0"?><worksheet><sheetData>'
    + '<row r="1"><c r="A1" t="str"><f>IF(A2&gt;1,"big","small")</f><v>big</v></c></row>'
    + "</sheetData></worksheet>";
  const sheet = parseSheetXmlSync(xml, opts);
  assert.equal(sheet.rows[0][0].v, "big");
});

test("parse time does not grow with the input chunk size", async () => {
  // 4000 rows is enough for the quadratic term to dominate if it returns.
  const xml = worksheet(4000);
  assert.ok(xml.length > 500_000, "fixture is large enough to be meaningful: " + xml.length + " chars");

  const small = await best(() => parseSheetXml(inChunks(xml, 16 * 1024), opts), 3);
  const large = await best(() => parseSheetXml(inChunks(xml, 1024 * 1024), opts), 3);

  // Before the fix this ratio was about 26x and grew with the buffer. A correct
  // linear scanner is flat, so anything close to 1 is expected; 4 leaves room
  // for noise without permitting a return of the quadratic term.
  const ratio = large / small;
  assert.ok(
    ratio < 4,
    "parsing in 1 MiB chunks took " + large.toFixed(0) + " ms against " + small.toFixed(0)
      + " ms in 16 KiB chunks (" + ratio.toFixed(1) + "x); the scanner is no longer linear",
  );
});

test("the synchronous and streaming readers agree on a large document", async () => {
  const xml = worksheet(1500);
  const sync = parseSheetXmlSync(xml, opts);
  const streamed = await parseSheetXml(inChunks(xml, 64 * 1024), opts);
  assert.equal(streamed.rowCount, sync.rowCount);
  assert.equal(streamed.rows.length, sync.rows.length);
  for (let r = 0; r < sync.rows.length; r++) {
    for (let c = 0; c < 8; c++) {
      const a = sync.rows[r][c];
      const b = streamed.rows[r][c];
      assert.equal(b && b.v, a && a.v, "cell " + col(c) + (r + 1) + " differs");
      assert.equal(b && b.t, a && a.t, "cell type " + col(c) + (r + 1) + " differs");
    }
  }
});
