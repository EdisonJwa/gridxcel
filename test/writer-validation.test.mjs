import test from "node:test";
import assert from "node:assert/strict";
import { writeWorkbookSync } from "../src/write.js";
import { readWorkbookSync } from "../src/read.js";
import { ZipArchive } from "../src/zip.js";

/**
 * Writer input validation.
 *
 * The writer used to accept values it could not represent in XML and emit the
 * result anyway: a column width of NaN was silently dropped, an autofilter of
 * "not a range" was written through verbatim, and a freeze pane at row 1e9
 * produced the cell reference "CFDGSXM1000000001". All three produced a file the
 * consumer has to repair, and the failure appeared on the other side of the wire.
 *
 * Rule under test: anything the writer cannot represent is a typed error HERE.
 */

function sheetWith(extra) {
  return {
    name: "S",
    rows: [[{ t: "number", v: 1 }]],
    originRow: 1, originCol: 0, rowCount: 1, columnCount: 1,
    merges: [], columns: [], rowInfo: new Map(),
    ...extra,
  };
}

function write(extra, bookExtra = {}) {
  const sheet = sheetWith(extra);
  return writeWorkbookSync({ sheetNames: [sheet.name], sheets: [sheet], ...bookExtra });
}

/** A one-cell sheet carrying a style, for the style-registry tests. */
function writeSheetWithStyle(style) {
  const sheet = sheetWith({ rows: [[{ t: "number", v: 1, s: style }]] });
  return writeWorkbookSync({ sheetNames: ["S"], sheets: [sheet] });
}

function stylesXml(bytes) {
  return new ZipArchive(bytes).readText("xl/styles.xml");
}

function sheetXml(bytes) {
  return new ZipArchive(bytes).readText("xl/worksheets/sheet1.xml");
}

/* ------------------------------------------------------------- columns */

test("writer: a column width that is not a finite positive number is refused", () => {
  for (const width of [NaN, Infinity, -Infinity, -5, 0]) {
    assert.throws(
      () => write({ columns: [{ min: 0, max: 0, width }] }),
      (e) => e.code === "BAD_VALUE" && /finite positive/.test(e.message),
      "width " + String(width) + " must be refused, not silently dropped",
    );
  }
});

test("writer: an over-wide column is clamped to Excel's 255, not emitted raw", () => {
  const xml = sheetXml(write({ columns: [{ min: 0, max: 0, width: 1e9 }] }));
  assert.match(xml, /width="255"/);
});

test("writer: a valid width is written exactly", () => {
  const xml = sheetXml(write({ columns: [{ min: 0, max: 0, width: 12.5 }] }));
  assert.match(xml, /width="12\.5" customWidth="1"/);
});

test("writer: positional column options are applied, not silently dropped", () => {
  // [{ width: 8 }, { width: 24 }] means columns A and B - the documented form.
  // An entry with no explicit index used to fall through the index check and be
  // dropped, so a caller set two widths and got neither, with no error.
  const xml = sheetXml(write({ columns: [{ width: 8 }, { width: 24 }] }));
  assert.match(xml, /<col min="1" max="1" width="8"/, "column A");
  assert.match(xml, /<col min="2" max="2" width="24"/, "column B");
});

test("writer: a positional column with a bad width is still refused", () => {
  assert.throws(() => write({ columns: [{ width: NaN }] }), (e) => e.code === "BAD_VALUE");
});

test("writer: a column object outside an array still needs an index", () => {
  // A bare object has no position and no min, so it genuinely cannot be placed.
  assert.throws(() => write({ columns: { width: 10 } }), (e) => e.code === "BAD_VALUE");
});

test("writer: a column index outside the sheet is refused", () => {
  for (const index of [-1, 16384, 1e9]) {
    assert.throws(() => write({ columns: [{ min: index, max: index }] }), (e) => e.code === "RANGE_TOO_LARGE", "index " + index);
    assert.throws(() => write({ columns: [index] }), (e) => e.code === "RANGE_TOO_LARGE", "shorthand " + index);
  }
});

/* ---------------------------------------------------------- auto filter */

test("writer: a valid autofilter range is normalised and written", () => {
  for (const ref of ["A1:G501", "A1", "$A$1:$D$10", "Sheet1!B2:C3"]) {
    const xml = sheetXml(write({ autoFilter: ref }));
    assert.match(xml, /<autoFilter ref="/, "should emit an autoFilter for " + ref);
    assert.ok(!/ref="/.test(xml.split("<autoFilter")[1].split(">")[0].replace(/ref="[^"]*"/, "")), "no stray attribute");
  }
});

test("writer: true means the used range and is still accepted", () => {
  const xml = sheetXml(write({ autoFilter: true }));
  assert.match(xml, /<autoFilter ref="[A-Z]+[0-9]+:[A-Z]+[0-9]+"/);
});

test("writer: an invalid autofilter is refused, not written verbatim", () => {
  // Before this, autoFilter: "not a range" emitted <autoFilter ref="not a range"/>
  // - invalid XML that Excel repairs on open.
  for (const bad of ["not a range", "A1:", ":::", "A0:B2", "1:2"]) {
    assert.throws(() => write({ autoFilter: bad }), (e) => e.code === "BAD_RANGE", JSON.stringify(bad));
  }
});

/* -------------------------------------------------------------- freeze */

test("writer: a freeze pane outside the sheet is refused", () => {
  // The pane's topLeftCell is derived from the split, so an absurd value emitted
  // a reference past the end of the grid.
  for (const freeze of [{ row: 1e9, col: 0 }, { row: 0, col: 1e9 }, { row: 1048576, col: 0 }, { row: 0, col: 16384 }]) {
    assert.throws(() => write({ freeze }), (e) => e.code === "RANGE_TOO_LARGE", JSON.stringify(freeze));
  }
});

test("writer: a valid freeze pane writes a real cell reference", () => {
  const xml = sheetXml(write({ freeze: { row: 1, col: 1 } }));
  const at = xml.indexOf("topLeftCell=");
  assert.ok(at > 0, "should emit a pane");
  const ref = xml.slice(at, xml.indexOf('"', at + 14) + 1);
  assert.match(ref, /^topLeftCell="[A-Z]{1,3}[0-9]+"$/);
  // And it must address the sheet.
  const cell = ref.replace(/topLeftCell="|"/g, "");
  const col = cell.replace(/[0-9]/g, "");
  assert.ok(col.length <= 3, "column part " + col + " must be within the grid");
});

test("writer: a merge that is not a range is refused", () => {
  for (const bad of ["NOTAREF", "A1:", "::::"]) {
    assert.throws(() => write({ merges: [bad] }), (e) => e.code === "BAD_ADDRESS", JSON.stringify(bad));
  }
});

/* ------------------------------------------------------- still work */

test("writer: ordinary metadata still writes and reads back", () => {
  const bytes = write({
    columns: [{ min: 0, max: 2, width: 20, hidden: false }],
    autoFilter: "A1:C10",
    freeze: { row: 1, col: 0 },
    merges: ["A1:B1"],
  });
  const back = readWorkbookSync(bytes);
  assert.equal(back.sheets[0].rows[0][0].v, 1);
  assert.deepEqual(back.sheets[0].merges, ["A1:B1"]);
  assert.equal(back.sheets[0].freeze && back.sheets[0].freeze.row, 1);
});

/* ------------------------------------------------------ reader side */

test("reader: a malformed mergeCell ref is dropped with a warning, not passed through", async () => {
  // The ref was pushed onto sheet.merges verbatim, so "NOTAREF" reached the
  // caller and would explode in whatever consumed a range.
  const { parseSheetXmlSync } = await import("../src/sheet-reader.js");
  const xml = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
    '<mergeCells count="3"><mergeCell ref="NOTAREF"/><mergeCell ref="A1:B2"/><mergeCell ref="A1:B2:C3"/></mergeCells></worksheet>';
  const warnings = [];
  const sheet = parseSheetXmlSync(xml, { name: "S", warn: (code) => warnings.push(code) });
  assert.deepEqual(sheet.merges, ["A1:B2"], "only the valid ref survives");
  assert.ok(warnings.includes("BAD_MERGE"), "the drop must be reported: " + JSON.stringify(warnings));
});

test("reader: valid merge spellings are kept", async () => {
  const { parseSheetXmlSync } = await import("../src/sheet-reader.js");
  const refs = ["A1", "A1:B2", "$A$1:$B$2", "AA100:AB200"];
  const xml = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="' +
    refs.length + '">' + refs.map((r) => '<mergeCell ref="' + r + '"/>').join("") + "</mergeCells></worksheet>";
  const sheet = parseSheetXmlSync(xml, { name: "S" });
  assert.deepEqual(sheet.merges, refs);
});

test("reader: hostile metadata never aborts and never leaves the grid", async () => {
  // Same construction the writer-side tests use: a real workbook with its
  // worksheet XML surgically replaced.
  const { ZipArchive } = await import("../src/zip.js");
  const { storedZip } = await import("../scripts/zip-surgery.mjs");
  const base = writeWorkbookSync({ sheetNames: ["S"], sheets: [sheetWith({})] });
  const src = new ZipArchive(base);
  const original = src.readText("xl/worksheets/sheet1.xml");
  const variants = {
    "pane row 1e9": (x) => x.replace("</sheetViews>", '<pane ySplit="1000000000" topLeftCell="A1000000001" activePane="bottomLeft" state="frozen"/></sheetViews>'),
    "autoFilter not a range": (x) => x.replace("</worksheet>", '<autoFilter ref="NOTAREF"/></worksheet>'),
    "col min huge": (x) => x.replace("<sheetData>", '<cols><col min="999999999" max="999999999" width="10"/></cols><sheetData>'),
    "col width NaN": (x) => x.replace("<sheetData>", '<cols><col min="1" max="1" width="NaN"/></cols><sheetData>'),
    "dimension huge": (x) => x.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:ZZ999999999"/>'),
  };
  for (const [label, patch] of Object.entries(variants)) {
    const patched = patch(original);
    assert.notEqual(patched, original, label + ": patch did not apply");
    const entries = [];
    for (const e of src.entries) entries.push([e.name, e.name === "xl/worksheets/sheet1.xml" ? patched : e.bytes()]);
    // The law: a typed error or a bounded sheet. Never an abort, never an untyped throw.
    let book;
    try {
      book = readWorkbookSync(storedZip(entries));
    } catch (error) {
      assert.ok(error && typeof error.code === "string", label + ": an untyped error escaped: " + error);
      continue;
    }
    assert.equal(book.sheets.length, 1, label + ": the sheet must still parse");
    const sheet = book.sheets[0];
    assert.ok(sheet.rowCount >= 0 && sheet.rowCount < 2 ** 21, label + ": rowCount " + sheet.rowCount);
    assert.ok(sheet.columnCount >= 0 && sheet.columnCount < 2 ** 20, label + ": columnCount " + sheet.columnCount);
  }
});

/* ------------------------------------------------------------- styles */

test("writer: a font size that is not a finite number is refused", () => {
  // <sz val="NaN"/> is not valid XML, and Excel has to repair the file to open
  // it - so the failure belonged here, not on the reader's machine.
  for (const sz of [NaN, Infinity, -Infinity, -5, 0, 1e9]) {
    assert.throws(
      () => writeSheetWithStyle({ font: { sz } }),
      (e) => e.code === "BAD_STYLE" && /font size/.test(e.message),
      "sz " + String(sz) + " must be refused",
    );
  }
});

test("writer: a usable font size is written exactly", () => {
  for (const sz of [1, 8, 11, 20, 409]) {
    const xml = stylesXml(writeSheetWithStyle({ font: { sz } }));
    assert.ok(xml.includes('<sz val="' + sz + '"/>'), "sz " + sz + " should appear in styles.xml");
  }
});

test("writer: no attribute in the package is ever NaN or Infinity", () => {
  // The general form of the bug: a number reaching an attribute without a
  // finiteness check produces a file that cannot be parsed.
  const sheet = sheetWith({
    columns: [{ min: 0, max: 0, width: 12 }],
    freeze: { row: 1, col: 1 },
    autoFilter: "A1:B10",
  });
  const bytes = writeWorkbookSync({ sheetNames: ["S"], sheets: [sheet] });
  const zip = new ZipArchive(bytes);
  for (const entry of zip.entries) {
    if (!/\.(xml|rels)$/.test(entry.name)) continue;
    const text = entry.text();
    assert.ok(!/="NaN"/.test(text), entry.name + " contains a NaN attribute");
    assert.ok(!/="Infinity"/.test(text), entry.name + " contains an Infinity attribute");
    assert.ok(!/="undefined"/.test(text), entry.name + " contains an undefined attribute");
  }
});

test("writer: a bad colour is still refused", () => {
  for (const color of ["not-a-color", "#AB", "#AABBCCDDEEFF"]) {
    assert.throws(() => writeSheetWithStyle({ font: { color } }), (e) => e.code === "BAD_STYLE", JSON.stringify(color));
  }
});

/* --------------------------------------------- sparse row length (found by fuzz) */

test("writer: a sparse row index beyond the sheet is refused, not walked", () => {
  // A row is a sparse array, so its length is whatever the caller set. Writing to
  // index 2^31-1 made length 2^31 and the naive loop walked two billion holes:
  // 38 seconds for one row, then it emitted a column reference Excel cannot read.
  // Same class as the SYLK/DIF coordinate bomb - bounded on the way in.
  const row = [];
  row[2 ** 31 - 1] = { t: "number", v: 1 };
  const started = Date.now();
  assert.throws(
    () => writeWorkbookSync({ sheetNames: ["S"], sheets: [sheetWith({ rows: [row], rowCount: 1 })] }),
    (e) => e.code === "RANGE_TOO_LARGE" && /columns a worksheet can address/.test(e.message),
  );
  assert.ok(Date.now() - started < 1000, "must refuse immediately, not scan");

  const row2 = [];
  row2[1e9] = { t: "number", v: 1 };
  assert.throws(
    () => writeWorkbookSync({ sheetNames: ["S"], sheets: [sheetWith({ rows: [row2], rowCount: 1 })] }),
    (e) => e.code === "RANGE_TOO_LARGE",
  );
});

test("writer: the last legal column index still works", () => {
  // 16383 is the highest 0-based column; the bound must not reject it.
  const row = [];
  row[16383] = { t: "string", v: "last" };
  const bytes = writeWorkbookSync({ sheetNames: ["S"], sheets: [sheetWith({ rows: [row], rowCount: 1 })] });
  const back = readWorkbookSync(bytes);
  assert.strictEqual(back.sheets[0].columnCount, 16384);
});

test("writer: an ordinary sparse row is unaffected by the bound", () => {
  const row = [];
  row[0] = { t: "number", v: 1 };
  row[25] = { t: "number", v: 2 };   // columns A and Z, 26 entries
  const bytes = writeWorkbookSync({ sheetNames: ["S"], sheets: [sheetWith({ rows: [row], rowCount: 1 })] });
  const back = readWorkbookSync(bytes);
  assert.strictEqual(back.sheets[0].columnCount, 26);
  assert.strictEqual(back.sheets[0].rows[0][25].v, 2);
});
