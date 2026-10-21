/**
 * A workbook written BY HAND in the OOXML that Excel itself produces.
 *
 * Every other fixture in this repository is produced by gridxcel's writer or by
 * openpyxl. That makes the reader's coverage circular: it only ever sees the
 * subset of OOXML this project already knows how to emit. A real workbook from
 * Excel carries constructs the writer never produces - phonetic runs (`rPh`),
 * rich-text runs, inline strings, boolean and error cells, cached formula
 * results, a shared-strings table with `uniqueCount`, and style/relationship
 * parts laid out the way Excel lays them out.
 *
 * This file builds such a package byte by byte and reads it. No gridxcel writer
 * is involved, so a defect in the writer cannot hide a defect in the reader.
 */
import test from "node:test";
import assert from "node:assert/strict";
import { readWorkbookSync } from "../src/read.js";
import { writeWorkbookSync } from "../src/write.js";
import { cellFormatCode } from "../src/read.js";
import { storedZip } from "../scripts/zip-surgery.mjs";

const RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
const WB = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="0"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Total">Sheet1!$B$2</definedName></definedNames></workbook>';
const WBRELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
const CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>';
const SHEET = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  + '<dimension ref="A1:F3"/>'
  + '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>'
  + '<cols><col min="1" max="1" width="12.5" customWidth="1"/></cols>'
  + '<sheetData>'
  + '<row r="1" spans="1:6"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>inline</t></is></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#DIV/0!</v></c><c r="E1" t="str"><f>A1&amp;B1</f><v>cached</v></c><c r="F1" s="1"><v>3.14159</v></c></row>'
  + '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>42</v></c><c r="C2"/></row>'
  + '<row r="3"><c r="A3"><f>SUM(B1:B2)</f><v>45</v></c></row>'
  + '</sheetData>'
  + '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
  + '<autoFilter ref="A1:F1"/>'
  + '<hyperlinks><hyperlink ref="F1" r:id="rId9" tooltip="tip"/></hyperlinks>'
  + '</worksheet>';
const SST = '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">'
  + '<si><t>plain shared</t></si>'
  + '<si><r><rPr><b/></rPr><t>rich </t></r><r><t>text</t></r><rPh sb="0" eb="1"><t>PHONETIC</t></rPh></si>'
  + '</sst>';
const STYLES = '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  + '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>'
  + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
  + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
  + '<borders count="1"><border/></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  + '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
  + '</styleSheet>';
const SRELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>';

function excelStyleWorkbook() {
  return storedZip([
    ["[Content_Types].xml", CT],
    ["_rels/.rels", RELS],
    ["xl/workbook.xml", WB],
    ["xl/_rels/workbook.xml.rels", WBRELS],
    ["xl/worksheets/sheet1.xml", SHEET],
    ["xl/sharedStrings.xml", SST],
    ["xl/styles.xml", STYLES],
    ["xl/worksheets/_rels/sheet1.xml.rels", SRELS],
  ]);
}

test("a hand-written Excel-shaped package reads without warnings", () => {
  const book = readWorkbookSync(excelStyleWorkbook());
  assert.deepEqual(book.sheetNames, ["Sheet1"]);
  assert.deepEqual(book.warnings || [], [], "no warning should be raised for a valid package");
});

test("every cell type Excel writes is decoded", () => {
  const sheet = readWorkbookSync(excelStyleWorkbook()).sheets[0];
  const row1 = sheet.rows[0];
  assert.equal(row1[0].t, "string", "shared string");
  assert.equal(row1[0].v, "plain shared");
  assert.equal(row1[1].t, "string", "inlineStr");
  assert.equal(row1[1].v, "inline");
  assert.equal(row1[2].t, "boolean");
  assert.equal(row1[2].v, true);
  assert.equal(row1[3].t, "error");
  assert.equal(row1[3].v, "#DIV/0!");
  assert.equal(row1[4].t, "string", "t=str keeps the cached value");
  assert.equal(row1[4].v, "cached");
  assert.equal(row1[4].f, "A1&B1", "and the formula text");
  assert.equal(row1[5].t, "number");
  assert.equal(row1[5].v, 3.14159);
});

test("rich text and phonetic runs render as the visible text only", () => {
  // A phonetic run (rPh) is a pronunciation hint, not content. Including it would
  // corrupt every Japanese and Chinese cell that carries one.
  const sheet = readWorkbookSync(excelStyleWorkbook()).sheets[0];
  assert.equal(sheet.rows[1][0].v, "rich text", "runs concatenate; rPh is excluded");
});

test("sheet-level structure survives: freeze, merges, filter, hyperlink, defined names", () => {
  const book = readWorkbookSync(excelStyleWorkbook());
  const sheet = book.sheets[0];
  assert.deepEqual(sheet.freeze, { row: 1, col: 1 });
  assert.deepEqual(sheet.merges, ["A1:B1"]);
  assert.equal(sheet.autoFilter, "A1:F1");
  assert.deepEqual(book.definedNames, [{ name: "Total", ref: "Sheet1!$B$2" }]);
  const link = sheet.rows[0][5].l;
  assert.ok(link, "the hyperlink should be attached to F1");
  assert.equal(link.target, "https://example.com");
  assert.equal(link.tooltip, "tip");
});

test("a numeric cell with a custom style keeps its format code", () => {
  // cell.z is the numeric id on read (documented); the code is reached through
  // cellFormatCode or book.numFmts. Asserting the code on cell.z would be testing
  // a shape the library never promised.
  const book = readWorkbookSync(excelStyleWorkbook());
  const cell = book.sheets[0].rows[0][5];
  assert.equal(cell.z, 164, "the numeric numFmtId");
  assert.equal(book.numFmts.get(164), "0.00%", "resolvable through book.numFmts");
  assert.equal(cellFormatCode(book, cell), "0.00%", "and through the helper");
});

test("the hand-written package round-trips through the writer unchanged in meaning", () => {
  const first = readWorkbookSync(excelStyleWorkbook());
  const rewritten = writeWorkbookSync(first);
  const second = readWorkbookSync(rewritten);
  const flat = (s) => s.rows.map((row) => row.map((c) => (c ? c.t + "=" + String(c.v) : "")).join("|")).join("/");
  assert.equal(flat(second.sheets[0]), flat(first.sheets[0]));
});
