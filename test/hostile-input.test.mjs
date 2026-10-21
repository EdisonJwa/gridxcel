/**
 * Hostile-input regression tests.
 *
 * Every case here was a real defect or a near miss found while hardening the
 * library against untrusted input. They are kept because the failure mode is
 * an unbounded allocation driven by a number inside an untrusted file, which no
 * amount of normal-fixture testing would catch.
 */

import { test } from "node:test";
import assert from "node:assert";
import { readWorkbook, readWorkbookSync } from "../src/read.js";
import { parseSheetXml } from "../src/sheet-reader.js";
import { createZip } from "../src/zip.js";
import { sheetToAOA, sheetToJSON } from "../src/json.js";
import { sheetToCSV } from "../src/csv.js";
import { writeWorkbookSync } from "../src/write.js";
import { GridxcelError } from "../src/errors.js";
import { eachRow } from "../src/read.js";

const RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
const WB = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>';
const WBRELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';

/** A one-kilobyte workbook whose second row claims an arbitrary position. */
function workbookWithRow(ref) {
  const sheet = '<?xml version="1.0"?><worksheet><sheetData>'
    + '<row r="1"><c r="A1"><v>1</v></c></row>'
    + '<row r="' + ref + '"><c r="A' + ref + '"><v>2</v></c></row>'
    + "</sheetData></worksheet>";
  return createZip([["_rels/.rels", RELS], ["xl/workbook.xml", WB], ["xl/_rels/workbook.xml.rels", WBRELS], ["xl/worksheets/sheet1.xml", sheet]]);
}

test("a crafted row index does not drive an unbounded allocation", () => {
  // 40 million is well past what any real sheet declares for four cells.
  const bytes = workbookWithRow(40000000);
  assert.ok(bytes.length < 4096, "the hostile input stays tiny: " + bytes.length + " B");

  const before = process.memoryUsage().heapUsed;
  const book = readWorkbookSync(bytes);
  const afterRead = process.memoryUsage().heapUsed;

  // Reading must not allocate proportional to the declared extent.
  assert.ok(afterRead - before < 64 * 1024 * 1024,
    "read allocated " + ((afterRead - before) / 1048576).toFixed(1) + " MiB for a 1 kB file");

  // The declared extent is still reported, so nothing is hidden from the caller.
  assert.equal(book.sheets[0].rowCount, 40000000, "declared rowCount is preserved");

  // The row array is sized to the rows present, not the declared extent.
  assert.ok(book.sheets[0].rows.length < 1000,
    "rows.length is " + book.sheets[0].rows.length + "; it must track rows present");

  // The far row is not discarded: it is kept out of the dense array.
  assert.ok(book.sheets[0].sparseRows instanceof Map, "far rows are kept in sheet.sparseRows");
  assert.equal(book.sheets[0].sparseRows.size, 1, "exactly one row was moved aside");
});

test("dense converters refuse an unbounded range with a typed, actionable error", async () => {
  const bytes = workbookWithRow(40000000);
  const book = readWorkbookSync(bytes);
  const sheet = book.sheets[0];

  for (const [name, fn] of [["sheetToAOA", () => sheetToAOA(sheet, {})], ["sheetToCSV", () => sheetToCSV(sheet, {})]]) {
    await assert.rejects(async () => fn(), (err) => {
      assert.ok(err instanceof GridxcelError, name + " threw " + err.constructor.name);
      assert.equal(err.code, "RANGE_TOO_LARGE", name + " code");
      assert.match(err.message, /range|eachRow/, name + " message points at a way forward");
      return true;
    }, name + " must refuse rather than allocate");
  }

  // sheetToJSON iterates the rows that exist rather than the declared extent, so
  // it completes without densifying and returns only the real rows.
  const rows = await sheetToJSON(sheet, {});
  assert.ok(rows.length <= 2, "sheetToJSON returned " + rows.length + " rows");
});

test("streaming the same hostile file stays bounded and yields only real rows", async () => {
  const bytes = workbookWithRow(40000000);
  let seen = 0;
  for await (const row of eachRow(bytes)) {
    seen++;
    assert.ok(row.r === 1 || row.r === 40000000, "unexpected row number " + row.r);
  }
  assert.equal(seen, 2, "eachRow emits the rows that exist, not the declared extent");
});

test("the bounded range still works when the caller asks for it", async () => {
  const bytes = workbookWithRow(40000000);
  const book = readWorkbookSync(bytes);
  // A bounded range makes the same sheet safe to densify.
  const aoa = await sheetToAOA(book.sheets[0], { range: "A1:B2" });
  assert.equal(aoa.length, 2);
  assert.deepEqual(aoa[0], [1, null]);
  // And the workaround is not needed for a normal sheet.
  const ok = await sheetToAOA(book.sheets[0], { range: "A1:A1" });
  assert.equal(ok.length, 1);
});

test("a crafted row index does not corrupt a round trip", () => {
  const book = readWorkbookSync(workbookWithRow(40000000));
  const out = writeWorkbookSync(book, {});
  const again = readWorkbookSync(out);
  // The far cell and the near cell both survive, and the extent is unchanged.
  assert.equal(again.sheets[0].rowCount, 40000000);
  assert.equal(again.sheets[0].rows[0][0].v, 1);
});

test("unparseable numbers are decoded without inventing values", () => {
  const sheet = '<?xml version="1.0"?><worksheet><sheetData><row r="1">'
    + '<c r="A1"><v>NaN</v></c><c r="B1"><v>Infinity</v></c><c r="C1"><v>1e999</v></c>'
    + "</row></sheetData></worksheet>";
  const bytes = createZip([["_rels/.rels", RELS], ["xl/workbook.xml", WB], ["xl/_rels/workbook.xml.rels", WBRELS], ["xl/worksheets/sheet1.xml", sheet]]);
  const row = readWorkbookSync(bytes).sheets[0].rows[0];
  // Whatever the decision, it must not be a number that silently means NaN.
  for (const cell of row) {
    assert.ok(cell !== undefined, "cell present");
    assert.ok(!(cell.t === "number" && !Number.isFinite(cell.v)),
      "a non-finite value must not be typed as a number: " + JSON.stringify(cell));
  }
  // Writing a real NaN must never produce invalid XML. Either it is refused with
  // a typed error, or the cell is dropped - but the output must stay parseable
  // and must not contain the literal "NaN".
  let out = null;
  try {
    out = writeWorkbookSync({ sheets: [[{ v: NaN, t: "number" }]] }, {});
  } catch (err) {
    assert.ok(err instanceof GridxcelError, "NaN write threw " + err.constructor.name);
    assert.equal(err.code, "BAD_VALUE");
  }
  if (out !== null) {
    const xml = Buffer.from(out).toString("latin1");
    assert.ok(!/<v>\s*(NaN|Infinity)/.test(xml), "written XML must not contain a non-finite literal");
    const again = readWorkbookSync(out);
    assert.equal(again.sheets.length, 1, "the written file is still readable");
  }
});

/* ---------------------------------------------------------------------------
 * Security review regression tests
 *
 * HIGH-1 and MED-3 were both found by the security review and both live in the
 * worksheet scanner. They are kept here because their failure mode is a
 * resource blow-up rather than a wrong answer, which no fixture-based test
 * would ever notice.
 * ------------------------------------------------------------------------- */

/** One cell whose inline string is `mb` megabytes of a single text node. */
function textBomb(mb) {
  const sheet = '<?xml version="1.0"?><worksheet><sheetData><row r="1">'
    + '<c r="A1" t="inlineStr"><is><t>' + "A".repeat(mb * 1024 * 1024) + "</t></is></c>"
    + "</row></sheetData></worksheet>";
  return createZip([["_rels/.rels", RELS], ["xl/workbook.xml", WB], ["xl/_rels/workbook.xml.rels", WBRELS], ["xl/worksheets/sheet1.xml", sheet]]);
}

/**
 * Fastest of `runs` attempts.
 *
 * A single sample is not a measurement here: this file runs alongside the rest
 * of the suite, so any one timing can be inflated by whatever else the machine
 * is doing. The minimum is the estimate least polluted by that, and it is the
 * right statistic for "is this algorithm linear" - a load spike can only make a
 * run slower, never faster.
 */
async function bestOf(runs, fn) {
  let best = Infinity;
  for (let i = 0; i < runs; i++) best = Math.min(best, await fn());
  return best;
}

test("HIGH-1: a long text node does not accumulate quadratically", async () => {
  // The scanner used to re-concatenate and rescan the whole accumulated text on
  // every chunk: 8 MB took 1.2 s and quadrupled per doubling, and a 104 kB
  // upload holding a 100 MB text node cost 116 s of CPU and 1.8 GB of heap.
  const time = async (mb) => {
    const bytes = textBomb(mb);
    const t0 = performance.now();
    const book = await readWorkbook(bytes);
    const dt = performance.now() - t0;
    assert.equal(book.sheets[0].rows[0][0].v.length, mb * 1024 * 1024, "text length preserved");
    return dt;
  };

  // The RATIO is what matters, and it is sensitive to the small sample: one slow
  // reading of the 2 MB case inflates it even when both are in fact linear. Five
  // samples per size, interleaved so a load spike cannot land only on the small
  // one, and the minimum of each taken. This suite runs in parallel with the rest.
  const smalls = [];
  const larges = [];
  for (let i = 0; i < 5; i++) {
    smalls.push(await time(2));
    larges.push(await time(8));
  }
  const small = Math.min(...smalls);
  const large = Math.min(...larges);

  // Four times the data must not cost sixteen times the time. A linear scanner
  // lands near 4x; the quadratic one was above 15x.
  const ratio = large / small;
  assert.ok(
    ratio < 10,
    "4x the text cost " + ratio.toFixed(1) + "x the time (" + small.toFixed(0) + " ms -> "
      + large.toFixed(0) + " ms); the text accumulator is no longer linear",
  );

  // And the absolute cost has to be sane, not merely better.
  assert.ok(large < 4000, "8 MB of text took " + large.toFixed(0) + " ms");
});

test("HIGH-1: a long attribute value does not accumulate quadratically", async () => {
  // A single attribute value longer than one chunk hits the tag scanner rather
  // than the text collector, and it needed its own fix: the unfinished tag was
  // concatenated and rescanned from its start on every chunk. Measured at the
  // time, a 16 MB attribute cost 11,912 ms from a 17.7 kB upload, growing ~3.9x
  // per doubling - roughly a 670,000x amplification.
  //
  // This asserts the SCALING, not an absolute time, because an absolute bound is
  // passed by the broken version on a fast enough machine: the old code answered
  // a 2 MB attribute in 317 ms, which any generous timeout would have allowed.
  const attribute = (mb) => {
    const sheet = '<?xml version="1.0"?><worksheet><sheetData><row r="1">'
      + '<c r="A1" s="' + "B".repeat(mb * 1024 * 1024) + '"><v>1</v></c>'
      + "</row></sheetData></worksheet>";
    return createZip([["_rels/.rels", RELS], ["xl/workbook.xml", WB], ["xl/_rels/workbook.xml.rels", WBRELS], ["xl/worksheets/sheet1.xml", sheet]]);
  };

  const time = async (mb) => {
    const bytes = attribute(mb);
    const t0 = performance.now();
    const book = await readWorkbook(bytes);
    const dt = performance.now() - t0;
    assert.equal(book.sheets[0].rows[0][0].v, 1, "the cell value still parses");
    return dt;
  };

  // Best of three, for the same reason as the text case above.
  const small = await bestOf(3, () => time(2));
  const large = await bestOf(3, () => time(8));

  // Four times the data must not cost sixteen times the time. Linear lands near
  // 4x; the quadratic version was 14.8x.
  const ratio = large / small;
  assert.ok(
    ratio < 10,
    "4x the attribute cost " + ratio.toFixed(1) + "x the time (" + small.toFixed(0) + " ms -> "
      + large.toFixed(0) + " ms); the tag scanner is no longer linear",
  );
});

test("a tag split across every possible chunk boundary still parses", async () => {
  // The resumable tag scan parks a half-read tag and continues from the next
  // chunk. Getting the hand-off wrong silently produces zero rows, so every
  // boundary of a small document is exercised rather than a representative few.
  const xml = '<?xml version="1.0"?><worksheet><sheetData><row r="1">'
    + '<c r="A1"><v>42</v></c><c r="B1"><v>7</v></c></row>'
    + '<row r="2"><c r="A2" t="inlineStr"><is><t>hi</t></is></c></row>'
    + "</sheetData></worksheet>";
  const opt = { name: "S", sharedStrings: [], styles: { isDate: () => false } };
  async function* pieces(size) {
    for (let i = 0; i < xml.length; i += size) yield xml.slice(i, i + size);
  }
  for (let size = 1; size <= xml.length; size++) {
    const sheet = await parseSheetXml(pieces(size), opt);
    assert.equal(sheet.rowCount, 2, "chunk size " + size + " lost a row");
    assert.equal(sheet.rows[0][0].v, 42, "chunk size " + size + " lost A1");
    assert.equal(sheet.rows[0][1].v, 7, "chunk size " + size + " lost B1");
    assert.equal(sheet.rows[1][0].v, "hi", "chunk size " + size + " lost A2");
  }
});

test("MED-3: the sparse-row guard does not depend on row order", () => {
  // The same two cells were treated as a normal sheet or a hostile one depending
  // on which appeared first: ascending gave rows.length 1 with a warning,
  // descending gave rows.length 1,048,576 with none.
  const mk = (rowsXml) => createZip([["_rels/.rels", RELS], ["xl/workbook.xml", WB], ["xl/_rels/workbook.xml.rels", WBRELS],
    ["xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet><sheetData>' + rowsXml + "</sheetData></worksheet>"]]);
  const row = (n) => '<row r="' + n + '"><c r="A' + n + '"><v>' + n + "</v></c></row>";

  const ascending = readWorkbookSync(mk(row(1) + row(1048576)));
  const descending = readWorkbookSync(mk(row(1048576) + row(1)));

  assert.equal(ascending.sheets[0].rows.length, descending.sheets[0].rows.length,
    "rows.length differs with row order");
  assert.equal(ascending.sheets[0].rowCount, descending.sheets[0].rowCount);
  assert.equal(
    ascending.sheets[0].sparseRows ? ascending.sheets[0].sparseRows.size : 0,
    descending.sheets[0].sparseRows ? descending.sheets[0].sparseRows.size : 0,
    "sparseRows differs with row order",
  );
  // A normal dense sheet must be untouched by the ordering fix.
  const normal = readWorkbookSync(mk(row(1) + row(2) + row(3)));
  assert.equal(normal.sheets[0].rows.length, 3);
  assert.equal(normal.sheets[0].sparseRows, undefined);
});

test("a CDATA section is unwrapped, whole or mixed into text", () => {
  // captureText tested for the CDATA marker at offset 1; it starts at offset 2,
  // so the delimiters were returned as part of the cell value.
  const mk = (text) => createZip([["_rels/.rels", RELS], ["xl/workbook.xml", WB], ["xl/_rels/workbook.xml.rels", WBRELS],
    ["xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>'
      + text + "</t></is></c></row></sheetData></worksheet>"]]);

  assert.equal(readWorkbookSync(mk("<![CDATA[raw <b> text]]>")).sheets[0].rows[0][0].v, "raw <b> text");
  assert.equal(readWorkbookSync(mk("before <![CDATA[<b>]]> after")).sheets[0].rows[0][0].v, "before <b> after");
});
