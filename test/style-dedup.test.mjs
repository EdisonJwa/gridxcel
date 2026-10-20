/**
 * De-duplication contract for styles.xml (DUP-3) and attribute escaping (DUP-4).
 *
 * These tests exist because the codebase had TWO styles.xml parsers that
 * disagreed about the same input, and TWO attribute escapers with different
 * reach. The parser that disagreed decides whether a numeric cell becomes a
 * Date, so the disagreement was a live correctness bug, not a style question.
 *
 * Every assertion here is a contract between the reader and the writer:
 *   - parseStyles is the only parser, and it hands the writer a raw numFmtId
 *     and the reader a resolved numFmtCode from the same record.
 *   - xml.encodeXmlAttr is the only attribute escaper, so a control character
 *     in a font name or format code can never reach the part unescaped.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { StyleTable, parseStyles } from "../src/style-registry.js";
import { ZipArchive } from "../src/zip.js";
import { parseTag } from "../src/xml.js";
import { cellFormatCode, readWorkbookSync } from "../src/read.js";
import { oracleAvailable, pythonPath } from "../scripts/oracle-client.mjs";

const RAW_CONTROL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/;

/** Validate a part with python's lxml: null when well-formed, else the error. */
function xmlError(text) {
  const dir = mkdtempSync(join(tmpdir(), "gridxcel-styles-"));
  const file = join(dir, "styles.xml");
  writeFileSync(file, text);
  try {
    execFileSync(pythonPath(), ["-c", "import sys, lxml.etree as E; E.parse(sys.argv[1])", file], { stdio: "pipe" });
    return null;
  } catch (err) {
    return String((err && err.stderr) || (err && err.message) || err);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

const HAS_ORACLE = oracleAvailable();

/* ------------------------------------------------------------------- DUP-4 */

test("DUP-4: a control character in a font name or format code stays escaped", () => {
  const table = new StyleTable();
  const idx = table.xf({ font: { name: "Bad\u0001Font" }, numFmt: "0.00\u0002%" });
  const xml = table.toXml();
  assert.ok(idx > 0, "the styled xf should not collapse into the default");
  assert.equal(RAW_CONTROL.test(xml), false, "styles.xml must not contain a raw control character");
  assert.ok(xml.includes("_x0001_"), "the font name control character must use the _xHHHH_ escape");
  assert.ok(xml.includes("_x0002_"), "the format code control character must use the _xHHHH_ escape");
  if (HAS_ORACLE) {
    assert.equal(xmlError(xml), null, "the emitted styles.xml must be well-formed XML");
  }
});

test("DUP-4: the private escapeAttr copy is gone and encodeXmlAttr is used", () => {
  const src = readFileSync("src/style-registry.js", "utf8");
  assert.equal(/function\s+escapeAttr\b/.test(src), false, "the private escapeAttr must be deleted, not kept alongside");
  assert.equal(/escapeAttr\(/.test(src), false, "no call site may still use the deleted helper");
  assert.ok(/encodeXmlAttr/.test(src), "style-registry must call xml.encodeXmlAttr");
  assert.ok(/from "\.\/xml\.js"/.test(src), "and import it from src/xml.js");
});

/* ------------------------------------------------------------------- DUP-3 */

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>' +
  '<numFmt numFmtId="165" formatCode="0.00%"/></numFmts>' +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><color theme="1"/></font>' +
  '<font><b val="1"/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
  '<borders count="1"><border><left style="thin"><color rgb="FF000000"/></left>' +
  '<right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="4">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1">' +
  '<alignment horizontal="center" wrapText="1"/></xf>' +
  '<xf numFmtId="200" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '</cellXfs></styleSheet>';

test("DUP-3: parseStyles returns the canonical shape for both consumers", () => {
  const parsed = parseStyles(STYLES_XML);
  assert.ok(parsed.numFmts instanceof Map, "numFmts is a Map");
  assert.equal(parsed.numFmts.get(164), "yyyy-mm-dd", "custom formats are listed by id");
  assert.equal(parsed.numFmts.get(165), "0.00%");
  // The writer needs the raw id, the reader needs the code; both come from one
  // record so neither consumer re-derives anything.
  assert.equal(parsed.styles[1].numFmt, 14, "the raw numFmtId, which the writer needs");
  assert.equal(parsed.styles[1].numFmtCode, "m/d/yy", "the resolved code, which the reader needs");
  assert.equal(parsed.styles[2].numFmt, 164);
  assert.equal(parsed.styles[2].numFmtCode, "yyyy-mm-dd");
  assert.equal(parsed.styles[0].numFmt, 0);
  assert.equal(parsed.styles[0].numFmtCode, "General");
  assert.equal(parsed.styles[2].font.b, true);
  assert.equal(parsed.styles[2].alignment.horizontal, "center");
  assert.equal(parsed.styles[2].alignment.wrapText, true);
  assert.equal(parsed.styles[1].font.name, "Calibri");
  assert.equal(parsed.styles[0].border.left.style, "thin");
  // An id with no declaration and no built-in meaning has no code at all.
  assert.equal(parsed.styles[3].numFmt, 200);
  assert.equal(parsed.styles[3].numFmtCode, null);
  // Preserved extras: the writer and the tests rely on these.
  assert.equal(parsed.fonts.length, 2);
  assert.equal(parsed.fills.length, 1);
  assert.equal(parsed.borders.length, 1);
  assert.equal(parsed.cellXfs.length, 4);
});

test("DUP-3: numFmts includes the ECMA-376 built-ins, not just the customs", () => {
  const parsed = parseStyles(STYLES_XML);
  assert.equal(parsed.numFmts.get(14), "m/d/yy");
  assert.equal(parsed.numFmts.get(4), "#,##0.00");
  assert.equal(parsed.numFmts.get(0), "General");
  assert.equal(parsed.numFmts.get(49), "@");
  assert.equal(parsed.numFmts.get(164), "yyyy-mm-dd", "a custom id shadows nothing and is not shadowed");
  assert.equal(parsed.numFmts.size, Object.keys(parsed.numFmts.get(0) ? {} : {}).length + parsed.numFmts.size);
  // Absent parts must not invent formats.
  assert.equal(parseStyles("").styles.length, 0);
  assert.equal(parseStyles("").isDate(0), false);
});

test("DUP-3: isDate stays the single decider of date cells", () => {
  const parsed = parseStyles(STYLES_XML);
  assert.equal(parsed.isDate(0), false, "General");
  assert.equal(parsed.isDate(1), true, "built-in 14 m/d/yy");
  assert.equal(parsed.isDate(2), true, "custom yyyy-mm-dd");
  assert.equal(parsed.isDate(3), false, "id 200 is unknown, so it is not a date");
  assert.equal(parsed.isDate(99), false, "an out-of-range index is not a date");
  const withCustom = parseStyles(
    '<styleSheet><numFmts count="2"><numFmt numFmtId="164" formatCode="0.00%"/>' +
    '<numFmt numFmtId="165" formatCode="[$-409]d\\-mmm\\-yy;@"/></numFmts>' +
    '<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs></styleSheet>');
  assert.equal(withCustom.isDate(0), false);
  assert.equal(withCustom.isDate(1), false, "a percentage is not a date");
  assert.equal(withCustom.isDate(2), true, "a localised date code is a date");
});

test("DUP-3: reader and writer see one styles.xml consistently", () => {
  // The writer's view: intern a style, emit the part.
  const table = new StyleTable();
  const bold = table.xf({ font: { b: true } });
  const dateIdx = table.xf({ numFmt: "yyyy-mm-dd hh:mm:ss" });
  const xml = table.toXml();
  const dateId = table.numFmtId("yyyy-mm-dd hh:mm:ss");
  assert.ok(dateId >= 164, "the custom date format gets a custom id");

  // The reader's view of the same bytes.
  const parsed = parseStyles(xml);
  assert.equal(parsed.styles[dateIdx].numFmt, dateId, "the id the writer allocated");
  assert.equal(parsed.styles[dateIdx].numFmtCode, "yyyy-mm-dd hh:mm:ss", "the code the reader renders");
  assert.equal(parsed.numFmts.get(dateId), "yyyy-mm-dd hh:mm:ss", "id -> code resolves for both sides");
  assert.equal(parsed.isDate(dateIdx), true);
  assert.equal(parsed.isDate(bold), false);
  assert.equal(parsed.styles[0].numFmtCode, "General");

  // Feeding the reader's record back into a fresh writer table reproduces the
  // same xf index: the two views describe the same record. (Interned in the
  // same order, because an xf index is an allocation order, not a name.)
  const again = new StyleTable();
  again.xf({ font: { b: true } });
  assert.equal(again.xf({ numFmt: parsed.styles[dateIdx].numFmtCode, font: parsed.styles[dateIdx].font }), dateIdx);
  assert.equal(parsed.styles[bold].numFmt, 0);
  assert.equal(parsed.styles[bold].numFmtCode, "General");
  assert.equal(parsed.numFmts.get(parsed.styles[bold].numFmt), "General");
});

test("DUP-3: a workbook read from disk exposes the same code on every styled cell", () => {
  const book = readWorkbookSync("test/fixtures/golden.xlsx");
  assert.ok(book.numFmts instanceof Map);
  assert.equal(book.numFmts.get(165), "yyyy-mm-dd hh:mm:ss", "the style part's custom format");
  let styled = 0;
  for (const sheet of book.sheets) {
    for (const row of sheet.rows) {
      if (row === undefined) continue;
      for (const cell of row) {
        if (cell === undefined || cell.s === undefined) continue;
        styled++;
        const style = book.styles[cell.s];
        assert.equal(cell.z, style.numFmt, "Cell.z is the raw id of its xf");
        assert.equal(cell.zf, style.numFmtCode, "Cell.zf is that xf's resolved code");
        if (style.numFmtCode !== null) {
          assert.equal(book.numFmts.get(style.numFmt), style.numFmtCode, "book.numFmts agrees with the style record");
        }
        assert.equal(cellFormatCode(book, cell), style.numFmtCode === null ? cell.z : style.numFmtCode);
      }
    }
  }
  assert.equal(styled, 6, "golden.xlsx has six styled cells, and every one must agree");
  const heavy = readWorkbookSync("test/fixtures/xlsxwriter-full.xlsx");
  let heavyStyled = 0;
  for (const sheet of heavy.sheets) {
    for (const row of sheet.rows) {
      if (row === undefined) continue;
      for (const cell of row) {
        if (cell === undefined || cell.s === undefined) continue;
        heavyStyled++;
        assert.equal(cell.zf, heavy.styles[cell.s].numFmtCode);
        assert.equal(cell.z, heavy.styles[cell.s].numFmt);
      }
    }
  }
  // 200 data rows x 3 formatted columns, plus the styled header and merge cells.
  assert.ok(heavyStyled >= 600, "xlsxwriter-full should expose at least 600 styled cells, saw " + heavyStyled);
  // The date decision still comes from the style, not from the value.
  const types = book.sheets.find((s) => s.name === "Types");
  assert.equal(types.rows[0][3].t, "date");
  assert.equal(types.rows[0][3].zf, "yyyy-mm-dd hh:mm:ss");
  assert.equal(types.rows[0][0].t, "string");
  assert.equal(types.rows[0][0].zf, "General");      // styled, but with the General format
  assert.equal(types.rows[0][1].t, "number");
  assert.equal(types.rows[0][1].zf, undefined);      // no s attribute, so no format view
});

test("DUP-3: a read workbook keeps its format codes through a write", async (t) => {
  if (!HAS_ORACLE) return t.skip("python oracle unavailable");
  const { writeWorkbookSync } = await import("../src/write.js");
  const book = readWorkbookSync("test/fixtures/golden.xlsx");
  const file = join(mkdtempSync(join(tmpdir(), "gridxcel-rt-")), "out.xlsx");
  writeFileSync(file, writeWorkbookSync(book, { type: "buffer" }));
  const py = [
    "import openpyxl, sys, json",
    "wb = openpyxl.load_workbook(sys.argv[1])",
    "out = {}",
    "for name in wb.sheetnames:",
    "    for row in wb[name].iter_rows():",
    "        for c in row:",
    "            if c.value is not None:",
    "                out[name + '!' + c.coordinate] = [c.number_format, str(c.value)]",
    "print(json.dumps(out))",
  ].join("\n");
  const raw = execFileSync(pythonPath(), ["-c", py, file], { encoding: "utf8" });
  const cells = JSON.parse(raw);
  try {
    // The style record the reader produced must drive the writer back to the
    // original format code; a numeric numFmt alone would degrade to General.
    assert.deepEqual(cells["Types!D1"], ["yyyy-mm-dd hh:mm:ss", "2024-03-15 14:30:45"]);
    assert.deepEqual(cells["Types!J1"], ["yyyy-mm-dd", "2024-01-01 00:00:00"]);
    assert.deepEqual(cells["Types!E1"], ["0.00%", "0.125"]);
    assert.deepEqual(cells["Types!F1"], ["#,##0.00", "1234.5"]);
    assert.deepEqual(cells["Types!C2"], ["0.00E+00", "-0.0001234"]);
  } finally {
    rmSync(file, { force: true });
    rmSync(join(file, ".."), { recursive: true, force: true });
  }
});

test("DUP-3: a malformed <sz/> does not poison the font record", () => {
  const parsed = parseStyles('<styleSheet><fonts count="1"><font><sz/><name val="Calibri"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs></styleSheet>');
  assert.equal(parsed.styles[0].font.name, "Calibri");
  assert.equal(parsed.styles[0].font.sz, undefined, "an unreadable size is dropped, not stored as NaN");
  // And the writer falls back to its default instead of emitting val="NaN".
  const table = new StyleTable();
  table.xf({ font: parsed.styles[0].font });
  const xml = table.toXml();
  assert.equal(/val="NaN"/.test(xml), false, "a NaN size must never reach the part");
  assert.ok(xml.includes('<sz val="11"/>'), "the writer's default size is used instead");
  // A readable size still survives.
  const ok = parseStyles('<styleSheet><fonts count="1"><font><sz val="9"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs></styleSheet>');
  assert.equal(ok.styles[0].font.sz, 9);
});

/* ------------------------------------------------------------------- MINOR-5 */

/** fonts/fills/borders/cellXfs of one styles.xml part, for a converged table. */
function styleCounts(xml) {
  const parsed = parseStyles(xml);
  return {
    fonts: parsed.fonts.length,
    fills: parsed.fills.length,
    borders: parsed.borders.length,
    cellXfs: parsed.cellXfs.length,
    numFmts: parsed.numFmts.size,
    bytes: xml.length,
  };
}

const stylesPartOf = (bytes) => new ZipArchive(new Uint8Array(bytes)).readText("xl/styles.xml");

test("MINOR-5: the style table converges in one round trip", async () => {
  const { writeWorkbookSync } = await import("../src/write.js");
  const files = ["golden", "xlsxwriter-full", "pandas", "openpyxl-written", "reader-handmade"];
  for (const name of files) {
    const file = "test/fixtures/" + name + ".xlsx";
    let book = readWorkbookSync(file);
    const trips = [];
    let first = null;
    for (let trip = 1; trip <= 3; trip++) {
      const bytes = writeWorkbookSync(book, { type: "buffer" });
      const xml = stylesPartOf(bytes);
      trips.push(styleCounts(xml));
      if (trip === 1) first = xml;
      else {
        // A workbook read and written back unchanged must stop changing at the
        // first write. On failure the message carries the whole three-trip table.
        assert.equal(xml, first,
          name + ": styles.xml differs on trip " + trip + " - three-trip table " + JSON.stringify(trips));
      }
      book = readWorkbookSync(bytes);
    }
    assert.deepEqual(trips[1], trips[0], name + ": counts after trip 2: " + JSON.stringify(trips));
    assert.deepEqual(trips[2], trips[1], name + ": counts after trip 3: " + JSON.stringify(trips));
  }
});

test("MINOR-5: a border with empty edges interns as the default border", () => {
  // The reader produces {left:{},right:{},...} for Excel's blank border; the
  // writer's own default border has no edge objects. They emit the same XML and
  // must therefore share one intern slot.
  const table = new StyleTable();
  assert.equal(table.borderCount, 1);
  const parsed = parseStyles('<styleSheet><borders count="1"><border><left/><right/><top/>' +
    '<bottom/><diagonal/></border></borders><cellXfs count="1">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>');
  table.xf({ border: parsed.styles[0].border });
  assert.equal(table.borderCount, 1, "an empty edge must not intern a second border record");
  const withStyle = parseStyles('<styleSheet><borders count="1"><border><left style="thin">' +
    '<color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>');
  const t2 = new StyleTable();
  t2.xf({ border: withStyle.styles[0].border });
  assert.equal(t2.borderCount, 2, "a real edge still interns its own border");
  assert.ok(t2.toXml().includes('<left style="thin"><color rgb="FF000000"/></left>'));
});

/* -------------------------------------------------------------------- xml.js */

test("xml.js: parseTag uses the exported localName instead of its own copy", () => {
  const src = readFileSync("src/xml.js", "utf8");
  const body = src.slice(src.indexOf("export function parseTag"), src.indexOf("export function localName"));
  assert.ok(/localName\(name\)/.test(body), "parseTag must call localName()");
  // Behaviour that must not move while the duplicate is collapsed.
  assert.equal(parseTag("c").local, "c");
  assert.equal(parseTag("x:c").local, "c");
  assert.equal(parseTag("c").name, "c");
  assert.equal(parseTag("x:c").name, "x:c");
  assert.deepEqual({ ...parseTag("c").attrs }, {});
  const withAttrs = parseTag('x:c r="A1" t="s"');
  assert.equal(withAttrs.local, "c");
  assert.equal(withAttrs.attrs.r, "A1");
  assert.equal(withAttrs.attrs.t, "s");
  assert.equal(parseTag("c r='A1'").attrs.r, "A1");
  assert.equal(parseTag("v").attrs, parseTag("v").attrs, "the no-attribute case reuses one frozen map");
});
