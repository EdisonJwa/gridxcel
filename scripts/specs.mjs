/**
 * Write-path specifications for scripts/interop.mjs.
 *
 * Every workbook here is described with the FROZEN data model from
 * docs/ARCHITECTURE.md / src/types.d.ts (Book = { sheets: SheetData[] },
 * Cell = { v, t, ... }, sparse rows indexed from originCol). The expected
 * values recorded next to each cell are computed from the same literals, so
 * the harness never asks a gridxcel module what the right answer is.
 */

/** Excel 1900-system serial for a JS Date, computed independently. */
export function serialOf(date, { local = false } = {}) {
  const epoch = local
    ? new Date(1899, 11, 30).getTime()
    : Date.UTC(1899, 11, 30);
  const ms = local
    ? new Date(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds()).getTime()
    : date.getTime();
  let serial = (ms - epoch) / 86400000;
  // Excel's 1900 system includes the non-existent 1900-02-29.
  if (serial >= 60) serial = serial;
  return serial;
}

export function colName(col) {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

class Spec {
  constructor(name, notes = []) {
    this.name = name;
    this.notes = notes;
    this.sheets = [];
    this._sheet = null;
  }

  sheet(name) {
    this._sheet = { name, rows: [], merges: [], expect: [], maxRow: 0, maxCol: 0, expectMerges: [] };
    this.sheets.push(this._sheet);
    return this;
  }

  /** col is 0-based; row is 1-based. Creates sparse rows as needed. */
  cell(row, col, value, { t, z, f, e, s: style } = {}) {
    const s = this._sheet;
    while (s.rows.length < row) s.rows.push([]);
    const line = s.rows[row - 1];
    while (line.length < col + 1) line.push(null);
    const cell = { v: value, t: t ?? inferType(value) };
    if (z !== undefined) cell.z = z;
    if (f !== undefined) cell.f = f;
    if (e !== undefined) cell.e = e;
    if (style !== undefined) cell.s = style;
    line[col] = cell;
    const ref = colName(col) + row;
    const exp = { ref, kind: cell.t, format: z === undefined ? "General" : String(z) };
    if (value instanceof Date) {
      exp.kind = "date";
      exp.serialUTC = serialOf(value);
      exp.serialLocal = serialOf(value, { local: true });
      exp.dateUTC = isoNoZone(value);
    } else {
      exp.value = value;
    }
    if (f !== undefined) exp.formula = f.replace(/^=/, "").replace(/\s+/g, "");
    s.expect.push(exp);
    s.maxRow = Math.max(s.maxRow, row);
    s.maxCol = Math.max(s.maxCol, col + 1);
    return this;
  }

  merge(range) {
    this._sheet.merges.push(range);
    this._sheet.expectMerges.push(range.toUpperCase().replace(/\$/g, ""));
    return this;
  }

  /** Book in the frozen shape plus the expectation record. */
  build() {
    const sheets = this.sheets.map((s) => ({
      name: s.name,
      rows: s.rows,
      originRow: 1,
      originCol: 0,
      columnCount: s.maxCol,
      rowCount: s.maxRow,
      merges: [...s.merges],
      columns: [],
      rowInfo: new Map(),
    }));
    const book = {
      sheets,
      dateSystem: 1900,
      numFmts: new Map(),
      styles: [],
      props: { title: "gridxcel interop spec: " + this.name },
      definedNames: [],
    };
    const expect = this.sheets.map((s) => ({
      name: s.name,
      rowCount: s.maxRow,
      columnCount: s.maxCol,
      merges: [...s.expectMerges].sort(),
      cells: s.expect,
    }));
    return { name: this.name, book, expect, notes: this.notes };
  }
}

function inferType(v) {
  if (v === null || v === undefined) return "empty";
  if (v instanceof Date) return "date";
  if (typeof v === "number") return "number";
  if (typeof v === "boolean") return "boolean";
  return "string";
}

function isoNoZone(d) {
  const p = (n, w = 2) => String(n).padStart(w, "0");
  return (
    d.getUTCFullYear() + "-" + p(d.getUTCMonth() + 1) + "-" + p(d.getUTCDate()) +
    "T" + p(d.getUTCHours()) + ":" + p(d.getUTCMinutes()) + ":" + p(d.getUTCSeconds())
  );
}

/* ------------------------------------------------------------- the specs */

export function writeSpecs() {
  const specs = [];

  // 1. Value/type coverage on the write path.
  {
    const s = new Spec("types", [
      "Numbers must round-trip as IEEE-754 doubles (compared with ===).",
      "The Date timezone convention is not pinned by the contract; the harness accepts UTC or local and reports which one was used.",
    ]);
    s.sheet("Types");
    const nums = [0, 1, -1, 0.1, 1 / 3, 1e-9, 1e15, -2.5e-7, 1234567890123, Math.PI];
    nums.forEach((n, i) => s.cell(1 + i, 0, n, { t: "number" }));
    const strs = ["hello", "", "caf\u00e9 \u4e2d\u6587 \ud83d\ude00", "x".repeat(500), "<&>\"'", " pad ", "line1\nline2", "tab\tsep"];
    strs.forEach((v, i) => s.cell(1 + i, 1, v, { t: "string" }));
    s.cell(1, 2, true, { t: "boolean" });
    s.cell(2, 2, false, { t: "boolean" });
    s.cell(3, 2, new Date(Date.UTC(2024, 0, 2, 3, 4, 5)), { t: "date", z: "yyyy-mm-dd hh:mm:ss" });
    s.cell(4, 2, new Date(Date.UTC(1999, 11, 31, 23, 59, 59)), { t: "date", z: "yyyy-mm-dd hh:mm:ss" });
    s.cell(5, 2, new Date(Date.UTC(1904, 0, 1, 0, 0, 0)), { t: "date", z: "yyyy-mm-dd" });
    s.cell(6, 2, "#DIV/0!", { t: "error", e: "#DIV/0!" });
    // Row 1000 forces the writer to emit a row index, not a dense block.
    s.cell(1000, 3, "last");
    s.sheet("Second Sheet");
    s.cell(1, 0, "second sheet value", { t: "string" });
    s.cell(2, 0, 42, { t: "number" });
    specs.push(s.build());
  }

  // 2. Number formats, merges, sheet geometry.
  {
    const s = new Spec("formats", [
      "Number-format codes are compared as text: python's cell.number_format must equal the spec code.",
    ]);
    s.sheet("Formats");
    const fmts = ["$#,##0.00", "0.00%", "yyyy-mm-dd hh:mm:ss", "0.000", "#,##0", "General", "0.00E+00", "@"];
    fmts.forEach((code, i) => s.cell(1 + i, 0, i === 2 ? new Date(Date.UTC(2024, 2, 4, 5, 6, 7)) : 1234.5678, {
      t: i === 2 ? "date" : "number",
      z: code,
    }));
    s.merge("A12:C12");
    s.cell(12, 0, "merged block", { t: "string" });
    s.merge("B13:B14");
    s.sheet("Empty");
    specs.push(s.build());
  }

  // 3. Formulas.
  {
    const s = new Spec("formulas", [
      "python/openpyxl reports formula text with a leading '='; gridxcel stores it without.",
    ]);
    s.sheet("Calc");
    s.cell(1, 0, 2, { t: "number" });
    s.cell(2, 0, 3, { t: "number" });
    s.cell(3, 0, 5, { t: "number", f: "A1+A2" });
    s.cell(4, 0, 0, { t: "number", f: "SUM(A1:A3)" });
    s.cell(5, 0, "a", { t: "string" });
    s.cell(6, 0, "a1", { t: "string", f: 'CONCATENATE(A5,"1")' });
    specs.push(s.build());
  }

  // 4. Cell styles: fonts, fills, borders, alignment.
  //
  // These records were only ever read back by gridxcel's own reader, so a write
  // that dropped <b/> or a fill colour round-tripped perfectly. openpyxl parses
  // the same styles.xml, so this spec puts every attribute we can emit in front of
  // an independent implementation.
  {
    const s = new Spec("styles", [
      "openpyxl reads fonts, fills, borders and alignment from the styles.xml we write.",
      "Alignment is only visible on a cell that uses the style, not on the xf record.",
    ]);
    s.sheet("Styled");
    s.cell(1, 0, 1, { t: "number", s: { font: { b: true, sz: 14, color: "#FF0000" } } });
    s.cell(2, 0, 2, { t: "number", s: { font: { i: true } } });
    s.cell(3, 0, 3, { t: "number", s: { font: { u: true } } });
    s.cell(4, 0, 4, { t: "number", s: { font: { strike: true } } });
    s.cell(5, 0, 5, { t: "number", s: { font: { name: "Arial", sz: 9 } } });
    s.cell(6, 0, 6, { t: "number", s: { fill: { pattern: "solid", fgColor: "#FFFF00" } } });
    s.cell(7, 0, 7, { t: "number", s: { border: { top: { style: "thin" }, bottom: { style: "double" } } } });
    s.cell(8, 0, 8, { t: "number", s: { alignment: { horizontal: "center", vertical: "top", wrapText: true } } });
    s.cell(9, 0, 9, { t: "number", s: { alignment: { horizontal: "right" } } });
    s.cell(10, 0, "styled text", { t: "string", s: { font: { b: true, i: true } } });
    specs.push(s.build());
  }

  return specs;
}

/**
 * Sparse AOA generator for the volume benchmarks: builds one row at a time so
 * the benchmark measures the writer, not the fixture builder.
 */
export function* volumeRows(rows, cols) {
  const epoch = Date.UTC(2020, 0, 1);
  for (let r = 0; r < rows; r++) {
    const line = new Array(cols);
    for (let c = 0; c < cols; c++) {
      const m = c % 5;
      if (m === 0) line[c] = { v: r * cols + c, t: "number" };
      else if (m === 1) line[c] = { v: (r * cols + c) / 7, t: "number" };
      else if (m === 2) line[c] = { v: "s" + r + "_" + c, t: "string" };
      else if (m === 3) line[c] = { v: r % 2 === 0, t: "boolean" };
      else line[c] = { v: new Date(epoch + r * 60000), t: "date", z: "yyyy-mm-dd hh:mm:ss" };
    }
    yield line;
  }
}
