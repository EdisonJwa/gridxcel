/**
 * Writer acceptance tests.
 *
 * The oracle is Python (openpyxl 3.1.5 / XlsxWriter 3.2.9 / zipfile / lxml),
 * never another gridxcel module: every value, merge, number format and part
 * name asserted here was observed by an independent implementation.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

import { WorkbookWriter, buildParts, numberLiteral } from "../src/workbook-writer.js";
import { StyleTable, parseStyles } from "../src/style-registry.js";
import { writeWorkbook, writeWorkbookSync, toBuffer } from "../src/write.js";
import { GridxcelError } from "../src/errors.js";

const PYTHON = process.env.GRIDXCEL_PYTHON ||
  "C:\\Users\\admin\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe";
const ROOT = fileURLToPath(new URL("..", import.meta.url));
const OUT = path.join(ROOT, "bench", "writer-eng");
const ORACLE = path.join(OUT, "oracle.py");
mkdirSync(OUT, { recursive: true });

// The oracle is embedded so the test file is self-contained: it is rewritten
// next to the generated workbooks on every run.
const ORACLE_SRC = String.raw`"""Independent oracle for the gridxcel writer.

Every fact this script prints comes from an implementation that is not
gridxcel: zipfile for the container, openpyxl for the workbook model and
XlsxWriter for the twin workbook used in the cross-writer comparison.
"""
import datetime
import json
import sys
import zipfile


def iso(value):
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, bool):
        return value
    return value


def cell_record(cell):
    return {"t": cell.data_type, "v": iso(cell.value), "nf": cell.number_format}


def sheet_dump(ws):
    cells = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cells[cell.coordinate] = cell_record(cell)
    cols = {}
    for key, dim in ws.column_dimensions.items():
        if dim.width is not None or dim.hidden:
            cols[key] = {"w": None if dim.width is None else round(float(dim.width), 6),
                         "hidden": bool(dim.hidden)}
    rows = {}
    for key, dim in ws.row_dimensions.items():
        if dim.height is not None or dim.hidden:
            rows[str(key)] = {"ht": None if dim.height is None else round(float(dim.height), 6),
                              "hidden": bool(dim.hidden)}
    tab = ws.sheet_properties.tabColor
    links = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink is not None:
                links[cell.coordinate] = cell.hyperlink.target
    return {
        "state": ws.sheet_state,
        "max_row": ws.max_row,
        "max_col": ws.max_column,
        "cells": cells,
        "merges": sorted(str(rng) for rng in ws.merged_cells.ranges),
        "freeze": ws.freeze_panes,
        "autofilter": None if ws.auto_filter.ref is None else str(ws.auto_filter.ref),
        "cols": cols,
        "rows": rows,
        "tab": None if tab is None else tab.rgb,
        "hyperlinks": links,
    }


def dump(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=False)
    book = {}
    for ws in wb.worksheets:
        book[ws.title] = sheet_dump(ws)
    names = {}
    for name, defined in wb.defined_names.items():
        names[name] = str(defined.value)
    return {
        "sheets": wb.sheetnames,
        "book": book,
        "names": names,
        "props": {"title": wb.properties.title, "creator": wb.properties.creator},
        "epoch": wb.epoch.year,
    }


def zipinfo(path):
    with zipfile.ZipFile(path) as z:
        return {"testzip": z.testzip(), "names": z.namelist(),
                "sizes": {i.filename: i.file_size for i in z.infolist()},
                "compressed": {i.filename: i.compress_size for i in z.infolist()},
                "methods": {i.filename: i.compress_type for i in z.infolist()}}


def stats(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    out = []
    for ws in wb.worksheets:
        first = None
        last = None
        count = 0
        for row in ws.iter_rows(values_only=True):
            count += 1
            if first is None:
                first = [iso(v) for v in row]
            last = [iso(v) for v in row]
        out.append({"title": ws.title, "rows": count, "max_row": ws.max_row,
                    "max_column": ws.max_column, "first": first, "last": last})
    wb.close()
    return {"sheets": out}


def mkxlsxwriter(path):
    """Write the same logical workbook the JS side writes, with XlsxWriter."""
    import xlsxwriter
    wb = xlsxwriter.Workbook(path)
    money = wb.add_format({"num_format": "#,##0.00"})
    ratio = wb.add_format({"num_format": "0.0000"})
    head = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E79",
                          "align": "center"})
    when = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})

    ws = wb.add_worksheet("Data")
    for i, title in enumerate(["Name", "Value", "When", "Ratio", "Flag", "Double", "Note"]):
        ws.write_string(0, i, title, head)
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 12.5)
    ws.set_column(2, 2, 20)
    ws.set_column(3, 3, 8.43, None, {"hidden": True})
    ws.set_column(4, 4, 10, None, {"level": 1})
    ws.set_row(0, 28)
    ws.set_row(4, None, None, {"hidden": True})
    ws.autofilter(0, 0, 500, 6)
    ws.freeze_panes(1, 0)
    special = ["\u65e5\u672c\u8a9e \U0001F600 \u00e9\u00e8", 'quote"and,comma',
               "line1\nline2", "  padded  "]
    for n in range(1, 501):
        row = n
        ws.write_string(row, 0, special[n - 1] if n <= 4 else "row-%d" % n)
        ws.write_number(row, 1, n * 1.5, money)
        ws.write_datetime(row, 2, datetime.datetime(2024, 1, 1, 12, 0, n % 60), when)
        ws.write_number(row, 3, n / 7.0, ratio)
        ws.write_boolean(row, 4, n % 2 == 0)
        ws.write_formula(row, 5, "=B%d*2" % (row + 1))
        ws.write_string(row, 6, "unicode \u2713 \u65e5\u672c\u8a9e \U0001F600")
    ws.merge_range(502, 0, 502, 2, "Merged title")

    layout = wb.add_worksheet("Layout")
    layout.merge_range(0, 0, 0, 2, "merged title")
    layout.merge_range(2, 0, 3, 1, "a3")
    layout.set_column(0, 0, 30)
    layout.set_column(1, 1, 12.5)
    layout.set_column(3, 3, 8.43, None, {"hidden": True})
    layout.set_row(0, 28)
    layout.set_row(4, None, None, {"hidden": True})
    layout.freeze_panes(1, 1)
    layout.set_tab_color("#FF0000")
    layout.write_url(6, 0, "https://example.com/gridxcel?x=1&y=2", None, "example link", "Example")
    layout.write_number(6, 2, 7)
    layout.write_boolean(7, 0, True)

    hidden = wb.add_worksheet("Hidden")
    hidden.write_string(0, 0, "secret")
    hidden.hide()

    wb.set_properties({"title": "gridxcel interop", "author": "writer-eng"})
    wb.define_name("Total", "=Data!$B$2:$B$501")
    wb.close()
    return {"ok": True}


COMMANDS = {"dump": dump, "zipinfo": zipinfo, "stats": stats, "mkxlsxwriter": mkxlsxwriter}


def main():
    # The Windows console default (cp949 here) cannot encode the unicode the
    # tests push through the writer, so the JSON channel is forced to UTF-8.
    sys.stdout.reconfigure(encoding="utf-8")
    command = sys.argv[1]
    result = COMMANDS[command](*sys.argv[2:])
    sys.stdout.write(json.dumps(result, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    main()`;
writeFileSync(ORACLE, ORACLE_SRC);

function havePython() {
  return existsSync(PYTHON);
}

function py(command, ...args) {
  const stdout = execFileSync(PYTHON, [ORACLE, command, ...args], {
    cwd: ROOT, encoding: "utf8", maxBuffer: 512 * 1024 * 1024,
  });
  return JSON.parse(stdout);
}

/* --------------------------------------------------------------- fixtures */

const HEADERS = ["Name", "Value", "When", "Ratio", "Flag", "Double", "Note"];
const SPECIAL = [
  "\u65e5\u672c\u8a9e \ud83d\ude00 \u00e9\u00e8",
  'quote"and,comma',
  "line1\nline2",
  "  padded  ",
];
const UNICODE_NOTE = "unicode \u2713 \u65e5\u672c\u8a9e \ud83d\ude00";
const HYPERLINK_URL = "https://example.com/gridxcel?x=1&y=2";

const headerStyle = {
  font: { b: true, color: "FFFFFFFFffff" .slice(0, 8) },
  fill: { pattern: "solid", fgColor: "FF1F4E79" },
  alignment: { horizontal: "center" },
};
const dataStyle = {
  font: { name: "Calibri", sz: 11 },
  border: { bottom: { style: "thin", color: "FFCCCCCC" } },
};

/** The workbook mirrored by oracle.py's mkxlsxwriter twin. */
function interopBook() {
  const rows = [];
  rows.push(HEADERS.map((h) => ({ v: h, t: "string", s: headerStyle })));
  for (let n = 1; n <= 500; n++) {
    const r = n + 1;
    rows.push([
      { v: n <= 4 ? SPECIAL[n - 1] : "row-" + n, s: dataStyle },
      { v: n * 1.5, s: dataStyle, z: "#,##0.00" },
      { v: new Date(2024, 0, 1, 12, 0, n % 60), t: "date", s: dataStyle, z: "yyyy-mm-dd hh:mm:ss" },
      { v: n / 7, s: dataStyle, z: "0.0000" },
      { v: n % 2 === 0, s: dataStyle },
      { f: "B" + r + "*2", v: n * 3, s: dataStyle },
      { v: UNICODE_NOTE, s: dataStyle },
    ]);
  }
  rows.push([]);                            // Excel row 502, deliberately empty
  rows.push([{ v: "Merged title" }]);       // Excel row 503
  const layout = [
    [{ v: "merged title" }],
    [],
    [{ v: "a3" }],
    [],
    [],
    [],
    [{ v: "example link" }, null, { v: 7 }],
    [{ v: true }],
  ];
  return {
    sheets: [
      {
        name: "Data",
        rows,
        originRow: 1,
        originCol: 0,
        columnCount: 7,
        rowCount: rows.length,
        merges: ["A503:C503"],
        columns: [
          { min: 0, max: 0, width: 28 },
          { min: 1, max: 1, width: 12.5 },
          { min: 2, max: 2, width: 20 },
          { min: 3, max: 3, width: 8.43, hidden: true },
          { min: 4, max: 4, width: 10, outlineLevel: 1 },
        ],
        rowInfo: new Map([[1, { index: 1, height: 28 }], [5, { index: 5, hidden: true }]]),
        freeze: { row: 1, col: 0 },
        autoFilter: "A1:G501",
      },
      {
        name: "Layout",
        rows: layout,
        originRow: 1,
        originCol: 0,
        columnCount: 3,
        rowCount: layout.length,
        merges: ["A1:C1", "A3:B4"],
        columns: [
          { min: 0, max: 0, width: 30 },
          { min: 1, max: 1, width: 12.5 },
          { min: 3, max: 3, width: 8.43, hidden: true },
        ],
        rowInfo: new Map([[1, { index: 1, height: 28 }], [5, { index: 5, hidden: true }]]),
        freeze: { row: 1, col: 1 },
        tabColor: "FFFF0000",
        hyperlinks: [{ ref: "A7", target: HYPERLINK_URL, tooltip: "Example" }],
      },
      {
        name: "Hidden",
        rows: [[{ v: "secret" }]],
        originRow: 1,
        originCol: 0,
        columnCount: 1,
        rowCount: 1,
        merges: [],
        columns: [],
        rowInfo: new Map(),
        hidden: true,
      },
    ],
    dateSystem: 1900,
    numFmts: new Map(),
    styles: [],
    props: { title: "gridxcel interop", creator: "writer-eng" },
    definedNames: [{ name: "Total", ref: "Data!$B$2:$B$501" }],
  };
}

function smallBook() {
  return {
    sheets: [{
      name: "S1",
      rows: [
        [{ v: "alpha" }, { v: 1 }, { v: true }],
        [{ v: 2.5 }, { v: new Date(2024, 2, 15, 14, 30, 45), t: "date" }, { v: "\u00e9t\u00e9" }],
      ],
      originRow: 1, originCol: 0, columnCount: 3, rowCount: 2,
      merges: [], columns: [], rowInfo: new Map(),
    }],
    dateSystem: 1900, numFmts: new Map(), styles: [],
    props: { title: "small", creator: "writer-eng" },
    definedNames: [],
  };
}

/* ------------------------------------------------------------------ tests */

test("OPC layout: exact part set, zip integrity, no theme part", { skip: !havePython() }, () => {
  const file = path.join(OUT, "layout.xlsx");
  writeWorkbookSync(interopBook(), { path: file, overwrite: true, deterministic: true });
  const info = py("zipinfo", file);
  assert.equal(info.testzip, null, "zipfile.testzip() must report no damage");
  assert.deepEqual(info.names.slice().sort(), [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/app.xml",
    "docProps/core.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "xl/workbook.xml",
    "xl/worksheets/_rels/sheet2.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/sheet3.xml",
  ].sort());
  assert.ok(!info.names.some((n) => n.indexOf("theme") >= 0), "no theme part");
  assert.ok(!info.names.some((n) => n.indexOf("sheet1.xml.rels") >= 0),
    "sheet1 has no hyperlinks, so it must have no rels part");
});

test("interop: openpyxl reads back every value, format, merge and name",
  { skip: !havePython() }, () => {
    const file = path.join(OUT, "interop.xlsx");
    writeWorkbookSync(interopBook(), { path: file, overwrite: true, deterministic: true });
    const d = py("dump", file);

    assert.deepEqual(d.sheets, ["Data", "Layout", "Hidden"]);
    assert.equal(d.epoch, 1899, "1900 date system");

    const data = d.book.Data;
    assert.equal(data.max_row, 503);
    assert.equal(data.max_col, 7);
    assert.deepEqual(data.merges, ["A503:C503"]);
    assert.equal(data.freeze, "A2");
    assert.equal(data.autofilter, "A1:G501");
    assert.deepEqual(data.rows["1"], { ht: 28.0, hidden: false });
    assert.deepEqual(data.rows["5"], { ht: null, hidden: true });
    assert.deepEqual(data.cols.A, { w: 28.0, hidden: false });
    assert.deepEqual(data.cols.B, { w: 12.5, hidden: false });
    assert.deepEqual(data.cols.C, { w: 20.0, hidden: false });
    assert.deepEqual(data.cols.D, { w: 8.43, hidden: true });
    assert.deepEqual(data.cols.E, { w: 10.0, hidden: false });

    const cells = data.cells;
    HEADERS.forEach((h, i) => {
      assert.equal(cells[String.fromCharCode(65 + i) + "1"].v, h);
    });
    // unicode, control characters and quoting survive byte-exactly
    assert.equal(cells.A2.v, SPECIAL[0]);
    assert.equal(cells.A3.v, SPECIAL[1]);
    assert.equal(cells.A4.v, SPECIAL[2]);
    assert.equal(cells.A5.v, SPECIAL[3]);
    assert.equal(cells.G501.v, UNICODE_NOTE);
    assert.equal(cells.A501.v, "row-500");
    // numbers keep IEEE-754 identity
    assert.equal(cells.B2.v, 1.5);
    assert.equal(cells.B501.v, 750);
    assert.equal(cells.B2.nf, "#,##0.00");
    assert.equal(cells.D2.v, 1 / 7);
    assert.equal(cells.D2.nf, "0.0000");
    // booleans
    assert.equal(cells.E2.v, false);
    assert.equal(cells.E3.v, true);
    assert.equal(cells.E2.t, "b");
    // dates carry the date number format and read back as datetimes
    assert.equal(cells.C2.v, "2024-01-01T12:00:01");
    assert.equal(cells.C2.nf, "yyyy-mm-dd hh:mm:ss");
    assert.equal(cells.C501.v, "2024-01-01T12:00:20");   // n = 500, n % 60 = 20
    // formulas keep their text
    assert.equal(cells.F2.v, "=B2*2");
    assert.equal(cells.F2.t, "f");
    assert.equal(cells.F501.v, "=B501*2");
    assert.equal(cells.A503.v, "Merged title");

    const layout = d.book.Layout;
    assert.deepEqual(layout.merges, ["A1:C1", "A3:B4"]);
    assert.equal(layout.freeze, "B2");
    assert.equal(layout.tab, "FFFF0000");
    assert.equal(layout.cells.A7.v, "example link");
    assert.equal(layout.hyperlinks.A7, HYPERLINK_URL);
    assert.equal(layout.cells.C7.v, 7);
    assert.equal(layout.cells.A8.v, true);
    assert.deepEqual(layout.rows["5"], { ht: null, hidden: true });
    assert.deepEqual(layout.cols.D, { w: 8.43, hidden: true });

    assert.equal(d.book.Hidden.state, "hidden");
    assert.equal(d.book.Hidden.cells.A1.v, "secret");

    assert.deepEqual(d.names, { Total: "Data!$B$2:$B$501" });
    assert.deepEqual(d.props, { title: "gridxcel interop", creator: "writer-eng" });
  });

test("interop: an XlsxWriter workbook with the same content is semantically identical",
  { skip: !havePython() }, () => {
    const mine = path.join(OUT, "interop.xlsx");
    const twin = path.join(OUT, "interop-xlsxwriter.xlsx");
    if (!existsSync(mine)) writeWorkbookSync(interopBook(), { path: mine, overwrite: true, deterministic: true });
    py("mkxlsxwriter", twin);
    const a = py("dump", mine);
    const b = py("dump", twin);
    assert.deepEqual(b.sheets, a.sheets, "sheet names and order");
    assert.deepEqual(b.epoch, a.epoch, "date system");
    assert.deepEqual(b.names, a.names, "defined names");
    assert.deepEqual(b.props, a.props, "document properties");
    // Cross-writer comparison of writer-independent facts only: two different
    // writers do not agree on record shapes or on the last ulp of a double.
    // XlsxWriter serialises with %.16g and loses it; gridxcel writes the
    // shortest spelling that parses back to the same double, which the
    // openpyxl test above proves value for value.
    const compareCells = (mine, twin, sheet) => {
      assert.deepEqual(Object.keys(twin).sort(), Object.keys(mine).sort(), "cell set of " + sheet);
      for (const ref of Object.keys(mine)) {
        const x = mine[ref];
        const y = twin[ref];
        assert.equal(y.t, x.t, ref + " type on " + sheet);
        assert.equal(y.nf, x.nf, ref + " number format on " + sheet);
        if (x.t === "n" && typeof x.v === "number" && typeof y.v === "number") {
          const tolerance = Math.abs(x.v) * 1e-14 + Number.MIN_VALUE;
          assert.ok(Math.abs(y.v - x.v) <= tolerance,
            ref + " on " + sheet + ": " + y.v + " vs " + x.v);
        } else {
          assert.deepEqual(y.v, x.v, ref + " value on " + sheet);
        }
      }
    };
    assert.equal(a.book.Data.cells.D2.v, 1 / 7, "gridxcel writes the exact double for 1/7");
    console.log("    1/7: gridxcel=" + a.book.Data.cells.D2.v + " xlsxwriter=" + b.book.Data.cells.D2.v);
    for (const name of a.sheets) {
      compareCells(a.book[name].cells, b.book[name].cells, name);
      assert.deepEqual(b.book[name].merges, a.book[name].merges, "merges of sheet " + name);
      // Column width *values* are not comparable across the two writers:
      // XlsxWriter treats set_column(w) as the Excel-UI character width and
      // rewrites it through its pixel conversion (28 -> 28.710938), while
      // gridxcel writes the OOXML width attribute verbatim. The exact width
      // gridxcel emits is asserted against openpyxl in the test above; here we
      // check that the same columns exist with the same visibility.
      const colsOf = (dump) => Object.fromEntries(Object.entries(dump)
        .map(([key, dim]) => [key, dim.hidden]));
      assert.deepEqual(colsOf(b.book[name].cols), colsOf(a.book[name].cols),
        "column visibility of sheet " + name);
      // Row records are compared as height/hidden semantics: the oracle only
      // reports rows that carry one of the two.
      assert.deepEqual(b.book[name].rows, a.book[name].rows, "row heights/hidden of sheet " + name);
      assert.deepEqual(b.book[name].freeze, a.book[name].freeze, "freeze of sheet " + name);
      assert.deepEqual(b.book[name].autofilter, a.book[name].autofilter, "filter of sheet " + name);
      assert.deepEqual(b.book[name].tab, a.book[name].tab, "tab colour of sheet " + name);
      assert.deepEqual(b.book[name].state, a.book[name].state, "state of sheet " + name);
      assert.deepEqual(b.book[name].hyperlinks, a.book[name].hyperlinks, "links of sheet " + name);
      assert.deepEqual(b.book[name].max_row, a.book[name].max_row, "max_row of sheet " + name);
      assert.deepEqual(b.book[name].max_col, a.book[name].max_col, "max_col of sheet " + name);
    }
  });

test("style dedup: 10000 cells sharing one style produce exactly one cellXfs entry",
  { skip: !havePython() }, () => {
    const shared = {
      font: { b: true, color: "FFFF0000" },
      fill: { pattern: "solid", fgColor: "FFFFFF00" },
      border: { top: { style: "thin", color: "FF000000" } },
      alignment: { horizontal: "right", wrapText: true },
    };
    const w = new WorkbookWriter({ deterministic: true });
    const sheet = w.addSheet("Dedup");
    let written = 0;
    for (let r = 1; r <= 1000; r++) {
      const row = [];
      for (let c = 0; c < 10; c++) row.push({ v: r * 10 + c, s: shared });
      sheet.addRow(r, row);
      written += row.length;
    }
    sheet.end();
    assert.equal(written, 10000);

    const sheetXml = w.stringParts().get("xl/worksheets/sheet1.xml");
    const stylesXml = w.stringParts().get("xl/styles.xml");
    const cellXfs = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
    assert.equal(cellXfs[1], "2", "one shared style + the mandatory default");
    assert.equal((cellXfs[2].match(/<xf /g) || []).length, 2);
    const styleAttrs = sheetXml.match(/ s="(\d+)"/g) || [];
    assert.equal(styleAttrs.length, 10000, "every cell carries a style");
    assert.ok(styleAttrs.every((a) => a === ' s="1"'), "all 10000 cells point at the same xf");
    assert.equal((sheetXml.match(/<c /g) || []).length, 10000);

    // Equal-but-distinct style objects must dedup just as aggressively.
    const w2 = new WorkbookWriter({ deterministic: true });
    const sheet2 = w2.addSheet("Dedup2");
    for (let r = 1; r <= 10000; r++) {
      sheet2.addRow(r, [{
        v: r,
        s: {
          font: { b: true, color: "FFFF0000" },
          fill: { pattern: "solid", fgColor: "FFFFFF00" },
          border: { top: { style: "thin", color: "FF000000" } },
          alignment: { horizontal: "right", wrapText: true },
        },
      }]);
    }
    sheet2.end();
    const cellXfs2 = w2.stringParts().get("xl/styles.xml").match(/<cellXfs count="(\d+)">/);
    assert.equal(cellXfs2[1], "2");

    const file = path.join(OUT, "dedup-10000.xlsx");
    writeFileSync(file, w.finish());
    const stats = py("stats", file);
    assert.equal(stats.sheets[0].title, "Dedup");
    assert.equal(stats.sheets[0].rows, 1000);
    assert.equal(stats.sheets[0].max_column, 10);
    assert.equal(stats.sheets[0].first[0], 10);
    assert.equal(stats.sheets[0].last[9], 10009);
    const size = py("zipinfo", file).sizes["xl/styles.xml"];
    assert.ok(size < 2048, "styles.xml stays tiny, got " + size + " bytes");
  });

test("deterministic: repeated writes are byte-identical", () => {
  const a = writeWorkbookSync(interopBook(), { deterministic: true });
  const b = writeWorkbookSync(interopBook(), { deterministic: true });
  assert.equal(a.length, b.length);
  assert.equal(Buffer.compare(Buffer.from(a), Buffer.from(b)), 0);

  const w1 = new WorkbookWriter({ deterministic: true });
  const s1 = w1.addSheet("S");
  s1.addRow(1, [{ v: 1 }, { v: "x" }]);
  s1.end();
  const w2 = new WorkbookWriter({ deterministic: true });
  const s2 = w2.addSheet("S");
  s2.addRow(1, [{ v: 1 }, { v: "x" }]);
  s2.end();
  assert.equal(Buffer.compare(Buffer.from(w1.finish()), Buffer.from(w2.finish())), 0);
});

test("streaming: 100k rows with bounded heap", { skip: !havePython() }, () => {
  const baseline = process.memoryUsage().heapUsed;
  let peak = 0;
  const sample = () => {
    const used = process.memoryUsage().heapUsed - baseline;
    if (used > peak) peak = used;
  };
  const w = new WorkbookWriter({ deterministic: true });
  const sheet = w.addSheet("Big");
  const shared = "shared string";
  for (let i = 1; i <= 100000; i++) {
    sheet.addRow(i, [i, (i % 1000) / 7, shared, i % 2 === 0]);
    if (i % 10000 === 0) sample();
  }
  sheet.end();
  assert.equal(sheet.rowCount, 100000);
  const sheetXml = w.stringParts().get("xl/worksheets/sheet1.xml");
  sample();
  assert.equal((sheetXml.match(/<row /g) || []).length, 100000);
  assert.equal((sheetXml.match(/<c /g) || []).length, 400000);
  const bytes = w.finish();
  sample();
  assert.ok(peak < 250 * 1024 * 1024, "peak heap delta was " + (peak / 1048576).toFixed(1) + " MB");

  const file = path.join(OUT, "stream-100k.xlsx");
  writeFileSync(file, bytes);
  const stats = py("stats", file);
  assert.equal(stats.sheets[0].rows, 100000, "python counts 100000 rows");
  assert.equal(stats.sheets[0].max_row, 100000);
  assert.equal(stats.sheets[0].max_column, 4);
  assert.equal(stats.sheets[0].first[0], 1);
  assert.equal(stats.sheets[0].last[0], 100000);
  assert.equal(stats.sheets[0].last[2], shared);
  assert.equal(stats.sheets[0].last[3], true);
  console.log("    streaming: " + (bytes.length / 1024).toFixed(1) + " kB zip, peak heap +" +
    (peak / 1048576).toFixed(1) + " MB");
});

test("inline strings: no sharedStrings part, values still exact", { skip: !havePython() }, () => {
  const book = smallBook();
  const file = path.join(OUT, "inline.xlsx");
  writeWorkbookSync(book, { path: file, overwrite: true, inlineStrings: true, deterministic: true });
  const info = py("zipinfo", file);
  assert.equal(info.testzip, null);
  assert.ok(!info.names.includes("xl/sharedStrings.xml"));
  const d = py("dump", file);
  assert.equal(d.book.S1.cells.A1.v, "alpha");
  assert.equal(d.book.S1.cells.C2.v, "\u00e9t\u00e9");
  assert.equal(d.book.S1.cells.A2.v, 2.5);
  // the wall clock of the Date survives exactly
  assert.equal(d.book.S1.cells.B2.v, "2024-03-15T14:30:45");
  assert.equal(d.book.S1.cells.B2.nf, "yyyy-mm-dd hh:mm:ss");

  // ...and the shared-string path stores each unique text once.
  const shared = path.join(OUT, "shared.xlsx");
  writeWorkbookSync(book, { path: shared, overwrite: true, deterministic: true });
  const sst = py("zipinfo", shared).sizes["xl/sharedStrings.xml"];
  assert.ok(sst > 0 && sst < 512, "sharedStrings.xml is small, got " + sst);
});

test("date system: 1904 workbook sets workbookPr date1904", { skip: !havePython() }, () => {
  const book = smallBook();
  book.dateSystem = 1904;
  const file = path.join(OUT, "date1904.xlsx");
  writeWorkbookSync(book, { path: file, overwrite: true, deterministic: true });
  const d = py("dump", file);
  assert.equal(d.epoch, 1904);
  const workbookXml = execFileSync(PYTHON, ["-c",
    "import sys,zipfile;print(zipfile.ZipFile(sys.argv[1]).read('xl/workbook.xml').decode())", file],
    { encoding: "utf8" });
  assert.match(workbookXml, /<workbookPr date1904="1"\/>/);
  assert.match(workbookXml, /<calcPr calcId="0" fullCalcOnLoad="1"\/>/);
});

test("recalc: false omits calcPr", () => {
  const parts = buildParts(smallBook(), { recalc: false });
  assert.ok(!parts.get("xl/workbook.xml").includes("calcPr"));
  const on = buildParts(smallBook());
  assert.ok(on.get("xl/workbook.xml").includes('<calcPr calcId="0" fullCalcOnLoad="1"/>'));
});

test("compressionLevel 0 stores parts instead of deflating", { skip: !havePython() }, () => {
  const stored = path.join(OUT, "stored.xlsx");
  const deflated = path.join(OUT, "deflated.xlsx");
  writeWorkbookSync(interopBook(), { path: stored, overwrite: true, compressionLevel: 0, deterministic: true });
  writeWorkbookSync(interopBook(), { path: deflated, overwrite: true, compressionLevel: 9, deterministic: true });
  const a = py("zipinfo", stored);
  const b = py("zipinfo", deflated);
  assert.equal(a.testzip, null);
  assert.equal(b.testzip, null);
  const sum = (info) => Object.values(info.compressed).reduce((x, y) => x + y, 0);
  assert.ok(sum(a) > sum(b) * 2, "stored (" + sum(a) + ") must be much larger than level 9 (" + sum(b) + ")");
  assert.ok(Object.values(a.methods).every((m) => m === 0), "level 0 must store every part");
  assert.ok(Object.values(b.methods).every((m) => m === 8), "level 9 must deflate every part");
  assert.deepEqual(py("dump", stored).book.Data.cells.F501, py("dump", deflated).book.Data.cells.F501);
});

test("return types follow the SheetJS type option", () => {
  const book = smallBook();
  assert.ok(writeWorkbookSync(book, { deterministic: true }) instanceof Uint8Array);
  assert.ok(Buffer.isBuffer(writeWorkbookSync(book, { type: "buffer", deterministic: true })));
  const array = writeWorkbookSync(book, { type: "array", deterministic: true });
  assert.ok(array instanceof ArrayBuffer);
  assert.equal(array.byteLength, writeWorkbookSync(book, { deterministic: true }).length);
  const b64 = writeWorkbookSync(book, { type: "base64", deterministic: true });
  assert.equal(typeof b64, "string");
  assert.equal(Buffer.from(b64, "base64").length, array.byteLength);
  const bin = writeWorkbookSync(book, { type: "binary", deterministic: true });
  assert.equal(bin.length, array.byteLength);
  assert.equal(bin.slice(0, 2), "PK");
  assert.equal(writeWorkbookSync(book, { type: "string", deterministic: true }), bin);
  assert.ok(toBuffer(book, { deterministic: true }) instanceof Uint8Array);
  assert.throws(() => writeWorkbookSync(book, { type: "nope" }), (e) => e.code === "BAD_TYPE");
});

test("overwrite guard: refuse an existing path until overwrite is set", async () => {
  const file = path.join(OUT, "guard.xlsx");
  writeWorkbookSync(smallBook(), { path: file, overwrite: true });
  assert.throws(() => writeWorkbookSync(smallBook(), { path: file }),
    (e) => e instanceof GridxcelError && e.code === "FILE_EXISTS");
  await assert.rejects(writeWorkbook(smallBook(), { path: file }),
    (e) => e instanceof GridxcelError && e.code === "FILE_EXISTS");
  const bytes = await writeWorkbook(smallBook(), { path: file, overwrite: true });
  assert.ok(bytes instanceof Uint8Array && bytes.length > 0);
});

test("values that XLSX cannot represent are rejected", () => {
  const w = new WorkbookWriter();
  const sheet = w.addSheet("S");
  assert.throws(() => sheet.addRow(1, [{ v: NaN }]), (e) => e.code === "BAD_VALUE");
  assert.throws(() => sheet.addRow(2, [{ v: Infinity }]), (e) => e.code === "BAD_VALUE");
  assert.throws(() => sheet.addRow(3, [{ v: -Infinity, t: "number" }]), (e) => e.code === "BAD_VALUE");
  assert.throws(() => sheet.addRow(4, [{ v: "2024-01-01", t: "date" }]), (e) => e.code === "BAD_VALUE");
  assert.throws(() => sheet.addRow(5, [{ v: "#NOPE", t: "error" }]), (e) => e.code === "BAD_VALUE");
  assert.throws(() => sheet.addRow(6, [{ v: 1, t: "wat" }]), (e) => e.code === "BAD_TYPE");
  assert.throws(() => sheet.addRow(7, [{ v: 1, s: -3 }]), (e) => e.code === "BAD_STYLE");
  assert.throws(() => sheet.addRow(8, [{ v: 1, z: {} }]), (e) => e.code === "BAD_STYLE");
  assert.throws(() => sheet.addRow(9, [{ v: "x", s: { fill: { fgColor: "not-a-colour" } } }]),
    (e) => e.code === "BAD_STYLE");
  // A numeric serial is a legal date input; one cell must survive the round trip.
  sheet.addRow(10, [{ v: 45292, t: "date" }]);
});

test("workbook structure errors carry stable codes", () => {
  const w = new WorkbookWriter();
  w.addSheet("S1");
  assert.throws(() => w.addSheet("S1"), (e) => e.code === "BAD_SHEET_NAME");
  assert.throws(() => w.addSheet("bad/name"), (e) => e.code === "BAD_SHEET_NAME");
  assert.throws(() => w.addSheet("x".repeat(32)), (e) => e.code === "BAD_SHEET_NAME");
  const s = w.addSheet("S2");
  assert.throws(() => s.addRow(0, []), (e) => e.code === "BAD_ROW");
  s.addRow(3, []);
  assert.throws(() => s.addRow(2, []), (e) => e.code === "ROW_ORDER");
  s.end();
  assert.throws(() => s.addRow(4, []), (e) => e.code === "SHEET_CLOSED");
  w.finish();
  assert.throws(() => w.addSheet("S3"), (e) => e.code === "WRITER_CLOSED");
});

test("number literals never use exponent notation", () => {
  assert.equal(numberLiteral(1), "1");
  assert.equal(numberLiteral(-0.5), "-0.5");
  assert.equal(numberLiteral(1e21), "1" + "0".repeat(21));
  assert.equal(numberLiteral(1.5e-7), "0.00000015");
  assert.equal(Number(numberLiteral(1e21)), 1e21);
  assert.equal(Number(numberLiteral(1.5e-7)), 1.5e-7);
  assert.equal(Number(numberLiteral(1e-7)), 1e-7);
  assert.equal(Number(numberLiteral(123456789012345678901234567890)), 123456789012345678901234567890);
});

test("buildParts exposes the raw XML for every part", () => {
  const parts = buildParts(interopBook());
  assert.ok(parts.get("[Content_Types].xml").startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'));
  const sheet = parts.get("xl/worksheets/sheet1.xml");
  assert.match(sheet, /<dimension ref="A1:G503"\/>/);
  assert.match(sheet, /<sheetViews><sheetView tabSelected="1" workbookViewId="0">/);
  assert.match(sheet, /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/);
  assert.match(sheet, /<row r="1" spans="1:7" ht="28" customHeight="1">/);
  assert.match(sheet, /<row r="5" spans="1:7" hidden="1">/);
  assert.ok(!sheet.includes('<row r="502"'), "an empty row with no row info is not emitted");
  const layout = parts.get("xl/worksheets/sheet2.xml");
  assert.match(layout, /<row r="5" hidden="1"\/>/, "row info alone still emits the row");
  assert.match(layout, /<tabColor rgb="FFFF0000"\/>/);
  assert.match(layout, /<hyperlink ref="A7" r:id="rId1" tooltip="Example"\/>/);
  assert.match(parts.get("xl/worksheets/_rels/sheet2.xml.rels"),
    /Target="https:\/\/example\.com\/gridxcel\?x=1&amp;y=2"/);
  assert.match(sheet, /<cols><col min="1" max="1" width="28" customWidth="1"\/>/);
  assert.match(sheet, /<autoFilter ref="A1:G501"\/>/);
  assert.match(sheet, /<mergeCells count="1"><mergeCell ref="A503:C503"\/><\/mergeCells>/);
  assert.match(parts.get("xl/worksheets/_rels/sheet2.xml.rels"), /TargetMode="External"/);
  assert.match(parts.get("xl/styles.xml"), /<numFmt numFmtId="164" formatCode="#,##0.00"\/>|formatCode="0.0000"/);
});

test("read-model round trip: numeric Cell.z and StyleObject.numFmt ids resolve to codes",
  { skip: !havePython() }, () => {
    // Shape a reader produces: cells carry numeric style indices and numeric
    // number-format ids, and the workbook carries the id -> code table.
    const book = {
      sheets: [{
        name: "Types",
        rows: [
          [{ v: "label" }, { v: 42 }, { v: true }],
          [
            { v: new Date(2024, 2, 15, 14, 30, 45), t: "date", s: 1, z: 166 },
            { v: 0.125, s: 2, z: 10 },
            { v: 1234.5, s: 3, z: 4 },
          ],
          [
            { v: new Date(2024, 3, 1), t: "date", s: 1 },
            { v: 0.5, s: 2 },
            { v: 7, s: 0, z: 0 },
          ],
          // A date cell whose source xf is the default one: the id (0) is
          // General, but the cell type still has to reach Excel as a date.
          [
            { v: new Date(2024, 4, 5, 6, 7, 8), t: "date", s: 0, z: 0 },
            { v: new Date(2024, 5, 6), t: "date", s: 0, z: 0 },
            { v: 1 },
          ],
        ],
        originRow: 1, originCol: 0, columnCount: 3, rowCount: 4,
        merges: [], columns: [], rowInfo: new Map(),
      }],
      dateSystem: 1900,
      numFmts: new Map([[166, "yyyy-mm-dd hh:mm:ss"]]),
      styles: [
        { numFmt: 0 },
        { numFmt: 166, font: { b: true } },
        { numFmt: 10 },
        { numFmt: 4 },
      ],
      props: { title: "roundtrip", creator: "writer-eng" },
      definedNames: [],
    };
    const file = path.join(OUT, "roundtrip-formats.xlsx");
    writeWorkbookSync(book, { path: file, overwrite: true, deterministic: true });
    const d = py("dump", file);
    const cells = d.book.Types.cells;
    // the custom format survives as itself, not as a default date format
    assert.equal(cells.A2.nf, "yyyy-mm-dd hh:mm:ss");
    assert.equal(cells.A2.v, "2024-03-15T14:30:45");
    assert.equal(cells.A3.nf, "yyyy-mm-dd hh:mm:ss");
    assert.equal(cells.A3.v, "2024-04-01T00:00:00");
    // built-in ids are reused, not re-registered as customs
    assert.equal(cells.B2.nf, "0.00%");
    assert.equal(cells.C2.nf, "#,##0.00");
    assert.equal(cells.B3.nf, "0.00%");
    assert.equal(cells.C3.nf, "General");
    assert.equal(cells.B2.v, 0.125);
    assert.equal(cells.A4.nf, "yyyy-mm-dd hh:mm:ss", "a date cell with a General xf is still a date");
    assert.equal(cells.A4.v, "2024-05-05T06:07:08");
    assert.equal(cells.B4.nf, "yyyy-mm-dd", "a midnight date gets the date-only default");
    assert.equal(cells.B4.v, "2024-06-06T00:00:00");
    assert.equal(cells.C4.nf, "General");
    assert.equal(cells.C4.v, 1);
    // only the one genuinely custom code is written to <numFmts>
    const parts = buildParts(book);
    const styles = parts.get("xl/styles.xml");
    const numFmts = styles.match(/<numFmts count="(\d+)"/);
    // the one custom code from the source book, plus the date-only default
    // invented for the midnight date cell in row 4
    assert.equal(numFmts[1], "2");
    assert.match(styles, /<numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"\/>/);
    assert.match(styles, /<numFmt numFmtId="165" formatCode="yyyy-mm-dd"\/>/);
    assert.ok(!/<numFmt numFmtId="10"/.test(styles), "built-in 10 must not be re-declared");
    assert.ok(!/<numFmt numFmtId="4"/.test(styles), "built-in 4 must not be re-declared");
  });

test("parseStyles reads the writer's own styles.xml back", () => {
  const parts = buildParts(interopBook());
  const parsed = parseStyles(parts.get("xl/styles.xml"));
  assert.ok(parsed.styles.length >= 3, "default + header + data styles");
  assert.equal(parsed.styles[0].numFmt, 0);
  const codes = Array.from(parsed.numFmts.values());
  assert.ok(codes.includes("0.0000"));
  assert.ok(codes.includes("yyyy-mm-dd hh:mm:ss"));
  const dateXf = parsed.styles.findIndex((s, i) => parsed.isDate(i));
  assert.ok(dateXf > 0, "the date format is classified as a date");
  assert.ok(!parsed.isDate(0), "General is not a date");
  assert.equal(parsed.styles[0].font.name, "Calibri");
  assert.equal(parsed.fonts[0].sz, 11);
  assert.equal(parsed.fills[0].pattern, "none");
  assert.equal(parsed.fills[1].pattern, "gray125");
});

test("StyleTable: built-ins are reused, customs start at 164 and dedup", () => {
  const t = new StyleTable();
  assert.equal(t.numFmtId("General"), 0);
  assert.equal(t.numFmtId(undefined), 0);
  assert.equal(t.numFmtId("#,##0.00"), 4);
  assert.equal(t.numFmtId("m/d/yy"), 14);
  assert.equal(t.numFmtId("0.0000"), 164);
  assert.equal(t.numFmtId("0.0000"), 164);
  assert.equal(t.numFmtId("yyyy-mm-dd"), 165);
  assert.equal(t.numFmtId("0.0000"), 164);
  assert.equal(t.xf({}), 0);
  assert.equal(t.xf({ numFmt: "General" }), 0);
  const bold = { font: { b: true } };
  const first = t.xf(bold);
  assert.ok(first > 0);
  assert.equal(t.xf(bold), first);
  assert.equal(t.xf({ font: { b: true } }), first, "value-equal styles collapse");
  assert.notEqual(t.xf({ font: { i: true } }), first);
  assert.equal(t.xfCount, 3, "default + bold + italic");
  assert.match(t.toXml(), /<cellStyles count="1">/);
});

test("golden round trip: read -> write keeps every value, format and merge (openpyxl-verified)",
  { skip: !havePython() }, async () => {
    let read = null;
    try {
      read = await import("../src/read.js");
    } catch {
      read = null;
    }
    if (read === null || typeof read.readWorkbookSync !== "function") {
      console.log("    src/read.js not available yet - skipped");
      return;
    }
    const golden = path.join(ROOT, "test", "fixtures", "golden.xlsx");
    const out = path.join(OUT, "golden-roundtrip.xlsx");
    const book = read.readWorkbookSync(golden, { cellDates: true });
    writeWorkbookSync(book, { path: out, overwrite: true, deterministic: true });
    const before = py("dump", golden);
    const after = py("dump", out);
    assert.deepEqual(after.sheets, before.sheets, "sheets survive the round trip");
    for (const name of before.sheets) {
      const src = before.book[name];
      const dst = after.book[name];
      assert.deepEqual(Object.keys(dst.cells).sort(), Object.keys(src.cells).sort(),
        "cell set of " + name);
      for (const ref of Object.keys(src.cells)) {
        assert.equal(dst.cells[ref].v, src.cells[ref].v,
          ref + " value on " + name + " (" + JSON.stringify(src.cells[ref]) + " -> " +
          JSON.stringify(dst.cells[ref]) + ")");
        assert.equal(dst.cells[ref].nf, src.cells[ref].nf, ref + " number format on " + name);
      }
      assert.deepEqual(dst.merges, src.merges, "merges of " + name);
      assert.equal(dst.freeze, src.freeze, "freeze of " + name);
      assert.deepEqual(dst.cols, src.cols, "columns of " + name);
      assert.deepEqual(dst.rows, src.rows, "rows of " + name);
      // openpyxl writes tab colours with an 00 alpha byte, the reader keeps
      // the RGB triple and gridxcel re-emits it with FF alpha. Only the RGB
      // part is visible in Excel, so that is what must survive.
      const rgb = (v) => (v === null || v === undefined ? null : String(v).slice(-6).toUpperCase());
      assert.equal(rgb(dst.tab), rgb(src.tab), "tab colour of " + name);
      assert.deepEqual(dst.hyperlinks, src.hyperlinks, "hyperlinks of " + name);
    }
    assert.equal(py("zipinfo", out).testzip, null, "the round-tripped package is intact");
  });

test("round trip through src/read.js when the reader has landed", async () => {
  let read = null;
  try {
    read = await import("../src/read.js");
  } catch {
    read = null;
  }
  const file = path.join(OUT, "interop.xlsx");
  if (!existsSync(file)) writeWorkbookSync(interopBook(), { path: file, overwrite: true, deterministic: true });
  if (read === null || typeof read.readWorkbookSync !== "function") {
    // The reader module is owned by another teammate and may not exist yet;
    // the python oracle above is the independent check in that case.
    console.log("    src/read.js not available yet - verified through openpyxl instead");
    return;
  }
  const book = read.readWorkbookSync(file, { cellDates: false });
  assert.deepEqual(book.sheets.map((s) => s.name), ["Data", "Layout", "Hidden"]);
  const data = book.sheets[0];
  assert.equal(data.merges.length, 1);
  assert.equal(data.merges[0], "A503:C503");
});