/**
 * OPC part assembly and streaming worksheet emission.
 *
 * WorkbookWriter owns the whole package: the style table, the shared string
 * table and one SheetWriter per sheet. Rows are serialised to XML strings as
 * they are pushed, so a caller can write a million-row sheet without ever
 * holding a sheet-sized object graph; finish() deflates each part and hands
 * the lot to ZipWriter.
 *
 * The part layout is deliberately minimal - no theme, no drawings - which
 * both Excel and openpyxl accept and which keeps a small book under 3 kB.
 */

import { ZipWriter } from "./zip.js";
import { encodeXmlText, encodeXmlAttr } from "./xml.js";
import { XML_DECL, NS_MAIN, NS_REL, NS_PKG_REL, NS_CT, NS_CP, NS_DC, NS_DCTERMS, NS_DCMITYPE, NS_XSI, NS_EP, NS_VT } from "./constants.js";
import { GridxcelError } from "./errors.js";
import { encodeCell, encodeRange, decodeRange, boundsToRef, makeBounds, extendBounds } from "./address.js";
import { dateToSerial, BUILTIN_FORMATS } from "./numfmt.js";
import { StyleTable, normalizeColor, colorAttrs, colorElement } from "./style-registry.js";
// Shared interfaces: value coercion lives in
// json.js, sheet-name rules in api.js. Neither may be re-implemented here.
import { cellFromValue } from "./json.js";
import { freeSheetName, validateSheetName } from "./api.js";

// Namespace URIs and the XML declaration for the parts this module writes.
// Module-local on purpose: no other module reads them (style-registry keeps its
// own copies at the bottom of that file, because importing them from here would
// be a cycle - this module imports StyleTable from there).

/**
 * The grid a worksheet may address.
 *
 * These are Excel's own limits. They are enforced on input because a writer that
 * accepts an impossible coordinate emits a file the consumer has to repair, and
 * the failure appears on the other side of the wire rather than here.
 */
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/** Excel caps a column width at 255 characters; anything larger is clamped. */
const MAX_COL_WIDTH = 255;

const REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const CT_RELS = "application/vnd.openxmlformats-package.relationships+xml";
const CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SHARED_STRINGS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml";
const CT_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

const DEFAULT_ROW_HEIGHT = 15;
const DATE_CODE = "yyyy-mm-dd";
const DATETIME_CODE = "yyyy-mm-dd hh:mm:ss";
const EPOCH_ISO = "1980-01-01T00:00:00Z";

/** t values accepted on input that are not already the canonical spelling. */
const TYPE_ALIASES = {
  n: "number", s: "string", b: "boolean", d: "date", e: "error",
  z: "empty", str: "string", inlineStr: "string",
};

/** Every error literal XLSX can carry in a t="e" cell. */
const ERROR_LITERALS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"]);

const UTF8 = new TextEncoder();

/* --------------------------------------------------------------- scalars */

/**
 * Shortest decimal spelling that parses back to the identical double, with
 * exponent notation expanded because Excel writes plain decimals in <v>.
 */
export function numberLiteral(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e21) return String(n);
  const s = String(n);
  const e = s.indexOf("e");
  if (e < 0) return s;
  return expandExponential(s);
}

function expandExponential(s) {
  const e = s.indexOf("e");
  const mant = s.slice(0, e);
  const exp = parseInt(s.slice(e + 1), 10);
  let sign = "";
  let digits = mant;
  if (digits.charCodeAt(0) === 45) { sign = "-"; digits = digits.slice(1); }
  const dot = digits.indexOf(".");
  let pointPos = (dot < 0 ? digits.length : dot) + exp;
  if (dot >= 0) digits = digits.slice(0, dot) + digits.slice(dot + 1);
  if (pointPos <= 0) return sign + "0." + "0".repeat(-pointPos) + digits;
  if (pointPos >= digits.length) return sign + digits + "0".repeat(pointPos - digits.length);
  return sign + digits.slice(0, pointPos) + "." + digits.slice(pointPos);
}

function toW3CDTF(ms) {
  return new Date(ms).toISOString().replace(/\.\d{3}Z$/, "Z");
}

function badValue(ref, what, v) {
  return new GridxcelError(
    "cannot write " + what + " at " + ref + ": " + (typeof v === "number" ? String(v) : JSON.stringify(v)) +
    " is not representable in XLSX", "BAD_VALUE");
}

function finite(n, ref) {
  if (typeof n !== "number" || !Number.isFinite(n)) throw badValue(ref, "number", n);
  return n;
}

function coerceNumber(v, ref) {
  if (typeof v === "number") return finite(v, ref);
  if (typeof v === "bigint") {
    const n = Number(v);
    if (!Number.isSafeInteger(n)) throw badValue(ref, "bigint", v);
    return n;
  }
  if (typeof v === "string" && v.trim() !== "") {
    const n = Number(v);
    if (Number.isFinite(n)) return n;
  }
  throw badValue(ref, "number", v);
}

/* -------------------------------------------------------- shared strings */

const PRESERVE_RE = /^\s|\s$/;

/** <t> child, preserving significant leading/trailing whitespace. */
function textElement(text) {
  const preserve = PRESERVE_RE.test(text) ? ' xml:space="preserve"' : "";
  return "<t" + preserve + ">" + encodeXmlText(text) + "</t>";
}

function plainSi(text) {
  return "<si>" + textElement(text) + "</si>";
}

function richSi(runs) {
  let s = "<si>";
  for (const run of runs) {
    if (run === undefined || run === null) continue;
    const text = run.v === undefined ? "" : String(run.v);
    const props = [];
    if (run.b) props.push("<b/>");
    if (run.i) props.push("<i/>");
    if (run.u) props.push("<u/>");
    if (run.sz !== undefined) props.push('<sz val="' + run.sz + '"/>');
    if (run.color !== undefined && run.color !== null) props.push(colorElement(normalizeColor(run.color)));
    s += "<r>";
    if (props.length) s += "<rPr>" + props.join("") + "</rPr>";
    s += "<t" + (PRESERVE_RE.test(text) ? ' xml:space="preserve"' : "") + ">" + encodeXmlText(text) + "</t></r>";
  }
  return s + "</si>";
}

/**
 * First-seen-order shared string table. The dedup key is the finished <si>
 * element, so equal text and equal rich-text runs collapse to one entry.
 */
class SharedStrings {
  constructor() {
    this.items = [];
    this.index = new Map();
    this.count = 0;
  }
  add(siXml) {
    this.count++;
    let i = this.index.get(siXml);
    if (i === undefined) {
      i = this.items.length;
      this.items.push(siXml);
      this.index.set(siXml, i);
    }
    return i;
  }
  get uniqueCount() { return this.items.length; }
  toXml() {
    if (this.items.length === 0) return null;
    return XML_DECL + '<sst xmlns="' + NS_MAIN + '" count="' + this.count +
      '" uniqueCount="' + this.items.length + '">' + this.items.join("") + "</sst>";
  }
}

/* ------------------------------------------------------------- sheet prep */

function normalizeMerges(merges) {
  if (merges === undefined || merges === null) return [];
  const list = Array.isArray(merges) ? merges : Object.keys(merges);
  const out = [];
  for (const m of list) {
    if (m === undefined || m === null) continue;
    const ref = String(m);
    const range = decodeRange(ref);          // throws on garbage
    if (range.s.r === range.e.r && range.s.c === range.e.c) continue;   // a single cell is not a merge
    out.push(ref);
  }
  return out;
}

/**
 * Normalise the `columns` option.
 *
 * Two spellings are documented and both must work:
 *
 *   [{ width: 8 }, { width: 24 }]        positional: entry i is column i
 *   [{ min: 2, max: 4, width: 20 }]      explicit 0-based span
 *
 * The positional form used to be SILENTLY DROPPED, because an entry with no
 * min/index fell through the index check - so the documented example in COOKBOOK
 * set two widths and got neither, with no error. An entry that cannot be placed
 * is now a mistake the caller hears about rather than a no-op.
 */
function normalizeColumns(columns) {
  if (columns === undefined || columns === null) return [];
  const out = [];
  const push = (c, position) => {
    if (c === undefined || c === null) return;
    if (typeof c === "number") {
      if (!Number.isInteger(c) || c < 0 || c >= MAX_COLS) {
        throw new GridxcelError(
          "column index " + c + " is outside the sheet (0.." + (MAX_COLS - 1) + ")",
          "RANGE_TOO_LARGE", { column: c });
      }
      out.push({ min: c, max: c });
      return;
    }
    // Explicit index, else the entry's own position in the array.
    const start = c.min !== undefined ? c.min : (c.index !== undefined ? c.index : position);
    if (start === undefined) {
      // Silently dropping the entry means a caller sets a width and gets a
      // default column with no signal. An entry with nothing to act on is a
      // mistake worth naming.
      throw new GridxcelError(
        "a column option has no min/index, so it does not say which column it applies to",
        "BAD_VALUE", { column: c });
    }
    const min = start;
    const max = c.max !== undefined ? c.max : start;
    let width = c.width;
    if (width !== undefined && width !== null) {
      // A width must be a finite, positive number of characters. NaN and
      // Infinity are unrepresentable in the attribute and a negative width is
      // meaningless; silently emitting nothing left the caller with a default
      // column and no sign that the request was ignored.
      const n = Number(width);
      if (!Number.isFinite(n) || n <= 0) {
        throw new GridxcelError(
          "column " + (typeof c.min === "number" ? c.min + 1 : "?") + " has width " + String(width) +
            "; a column width must be a finite positive number of characters",
          "BAD_VALUE", { width });
      }
      width = Math.min(n, MAX_COL_WIDTH);
    }
    // A column index must address the sheet, not past its edge.
    const lo = Math.min(min, max);
    const hi = Math.max(min, max);
    if (!Number.isInteger(lo) || !Number.isInteger(hi) || lo < 0 || hi >= MAX_COLS) {
      throw new GridxcelError(
        "column range " + lo + ".." + hi + " is outside the sheet (0.." + (MAX_COLS - 1) + ")",
        "RANGE_TOO_LARGE", { min: lo, max: hi });
    }
    out.push({
      min: lo, max: hi, width, hidden: c.hidden,
      outlineLevel: c.outlineLevel, style: c.style, bestFit: c.bestFit, customWidth: c.customWidth,
    });
  };
  if (Array.isArray(columns)) columns.forEach((c, i) => push(c, i));
  else if (typeof columns === "object") push(columns, undefined);
  return out;
}

function toRowInfoMap(rowInfo) {
  if (rowInfo === undefined || rowInfo === null) return new Map();
  if (rowInfo instanceof Map) {
    const out = new Map();
    for (const [k, v] of rowInfo) out.set(Number(k), v);
    return out;
  }
  const out = new Map();
  for (const [k, v] of Object.entries(rowInfo)) out.set(Number(k), v);
  return out;
}

/**
 * Autofilter: `true` means "the used range", a string must be a real A1 range.
 *
 * The ref was previously written through verbatim, so `autoFilter: "not a range"`
 * produced `<autoFilter ref="not a range"/>` - invalid XML that Excel repairs.
 */
function normalizeAutoFilter(value) {
  if (value === undefined || value === null || value === false) return null;
  if (value === true) return true;
  const text = String(value);
  try {
    // decodeRange returns { s, e } (start and end cells), not a bounds object.
    const range = decodeRange(text);
    return encodeRange(range.s, range.e);
  } catch {
    throw new GridxcelError(
      "autoFilter " + JSON.stringify(text) + " is not a valid A1 range (expected something like \"A1:D10\")",
      "BAD_RANGE", { autoFilter: text });
  }
}

function normalizeFreeze(freeze) {
  if (!freeze) return null;
  const row = freeze.row | 0;
  const col = freeze.col | 0;
  if (row <= 0 && col <= 0) return null;
  // The pane's topLeftCell is derived from these, so an absurd split emitted a
  // cell reference past the end of the grid ("CFDGSXM1000000001") and produced a
  // file Excel would repair. Refuse instead.
  if (row >= MAX_ROWS || col >= MAX_COLS) {
    throw new GridxcelError(
      "freeze pane at row " + row + ", column " + col + " is outside the sheet (max " +
        MAX_ROWS + " x " + MAX_COLS + ")",
      "RANGE_TOO_LARGE", { row, col });
  }
  return { row: Math.max(0, row), col: Math.max(0, col) };
}

function normalizeHyperlinks(links) {
  if (links === undefined || links === null) return [];
  const out = [];
  const one = (ref, val) => {
    if (val === undefined || val === null) { out.push({ ref }); return; }
    if (typeof val === "string") {
      // "Sheet2!A1" is a location; anything else is an external target.
      if (val.indexOf("://") < 0 && val.indexOf("!") >= 0) out.push({ ref, location: val });
      else out.push({ ref, target: val });
      return;
    }
    out.push({
      ref,
      target: val.target !== undefined ? val.target : val.url,
      location: val.location, tooltip: val.tooltip, display: val.display,
    });
  };
  if (Array.isArray(links)) {
    for (const item of links) {
      if (typeof item === "string") { one(item, undefined); continue; }
      if (item === undefined || item === null) continue;
      const ref = item.ref !== undefined ? item.ref : item.cell;
      const val = item.target !== undefined || item.url !== undefined || item.location !== undefined
        ? item : undefined;
      one(ref, val);
    }
  } else if (typeof links === "object") {
    for (const [ref, val] of Object.entries(links)) one(ref, val);
  }
  for (const h of out) {
    if (h.ref === undefined || h.ref === null || h.ref === "") {
      throw new GridxcelError("hyperlink needs a ref (A1 address)", "BAD_HYPERLINK");
    }
    decodeRange(String(h.ref));   // validate
  }
  return out;
}

/* ------------------------------------------------------------ sheet writer */

/**
 * One worksheet. addRow() serialises immediately, so the only state that
 * grows with the sheet is the array of finished row strings.
 */
export class SheetWriter {
  constructor(book, name, opts = {}) {
    this.book = book;
    this.name = name;
    this.index = book.sheets.length;
    this.partName = "xl/worksheets/sheet" + (this.index + 1) + ".xml";
    this.relsPartName = "xl/worksheets/_rels/sheet" + (this.index + 1) + ".xml.rels";
    this._rows = [];
    this._rowNums = [];
    this._count = 0;
    this._lastRow = 0;
    this._closed = false;
    this._finalized = false;
    this._sheetData = null;
    this._bounds = makeBounds();
    this._merges = normalizeMerges(opts.merges);
    this._columns = normalizeColumns(opts.columns);
    this._rowInfo = toRowInfoMap(opts.rowInfo);
    this._freeze = normalizeFreeze(opts.freeze);
    this._autoFilter = normalizeAutoFilter(opts.autoFilter);
    this._tabColor = opts.tabColor === undefined || opts.tabColor === null ? null : normalizeColor(opts.tabColor);
    this._hidden = opts.hidden === true;
    this._hyperlinks = normalizeHyperlinks(opts.hyperlinks);
    this._originCol = opts.originCol === undefined || opts.originCol === null ? 0 : opts.originCol | 0;
    this._dims = opts.dims ? String(opts.dims) : null;
    this._defaultRowHeight = opts.defaultRowHeight;
    this._defaultColWidth = opts.defaultColWidth;
    this._inlineStrings = opts.inlineStrings === true ? true
      : opts.inlineStrings === false ? false : book.inlineStrings;
    this._styleObjects = opts.styleObjects || null;   // old xf index -> StyleObject
    this._styleIndex = opts.styleIndex || null;       // old xf index -> new cellXfs index
    this._numFmtCodes = opts.numFmtCodes || null;     // read numFmtId -> format code
    this._styleCache = new Map();
  }

  get rowCount() { return this._count; }
  get hidden() { return this._hidden; }

  /** Append one row. rowIndex is 1-based; cells are indexed from originCol. */
  addRow(rowIndex, cells) {
    if (this._closed || this._finalized) {
      throw new GridxcelError("sheet " + JSON.stringify(this.name) + " is already ended", "SHEET_CLOSED");
    }
    const r = rowIndex === undefined || rowIndex === null ? this._lastRow + 1 : rowIndex;
    if (!Number.isInteger(r) || r < 1) {
      throw new GridxcelError("row index must be a positive integer, got " + rowIndex, "BAD_ROW");
    }
    if (r <= this._lastRow) {
      throw new GridxcelError("rows must be written in ascending order: row " + r + " after row " + this._lastRow, "ROW_ORDER");
    }
    this._lastRow = r;
    this._count++;
    const info = this._rowInfo.get(r);
    let parts = null;
    let minC = -1;
    let maxC = -1;
    if (cells !== undefined && cells !== null) {
      const base = this._originCol;
      // Bounds and A1 refs are 0-based (address.js convention); only the row
      // element's r attribute is 1-based.
      const r0 = r - 1;
      for (let j = 0; j < cells.length; j++) {
        const cell = cells[j];
        if (cell === undefined || cell === null) continue;
        const col = base + j;
        const xml = this._cellXml(cell, r0, col);
        if (xml === null) continue;
        if (parts === null) parts = [];
        parts.push(xml);
        if (minC < 0 || col < minC) minC = col;
        if (col > maxC) maxC = col;
        extendBounds(this._bounds, r0, col);
      }
    }
    if (parts === null && info === undefined) return;
    this._rows.push(this._rowXml(r, parts, minC, maxC, info));
    this._rowNums.push(r);
  }

  /** Freeze the sheet; further addRow calls throw. Safe to call twice. */
  end() {
    this._closed = true;
    return this;
  }

  /* ------------------------------------------------------------ emission */

  _cellXml(input, r, c) {
    const ref = encodeCell(r, c);
    // One coercion contract for the whole library: json.cellFromValue. The
    // writer used to infer the type here and answer "empty" for anything it did
    // not recognise, which dropped the caller's cell without a word; the
    // canonical coercion throws BAD_CELL instead, so the same value now lands
    // the same way through aoa_to_sheet, normalizeCell and this writer.
    const cell = cellFromValue(input, { date1904: this.book.date1904 === true });
    if (cell === undefined) return null;
    // A bare Date is the one value whose format code the writer keeps: the
    // canonical cell stamps "yyyy-mm-dd", while the code that fits the value -
    // date vs date-time - is derivable from the serial about to be emitted.
    if (input instanceof Date) delete cell.z;
    // json.cellFromValue always returns a cell carrying a t (it throws BAD_CELL
    // rather than returning a typeless one), so no undefined/null guard is
    // needed here - F15 removed the one that could never fail.
    let t = cell.t;
    if (TYPE_ALIASES[t] !== undefined) t = TYPE_ALIASES[t];
    const v = cell.v;
    const f = cell.f;
    const z = cell.z;
    const s = cell.s;
    const e = cell.e;
    const serial = cell.serial;
    const runs = cell.r;
    let fXml = "";
    if (f !== undefined && f !== null && f !== "") {
      let text = String(f);
      if (text.charCodeAt(0) === 61) text = text.slice(1);   // drop the leading '='
      if (text.length === 0) throw new GridxcelError("empty formula at " + ref, "BAD_FORMULA");
      fXml = "<f>" + encodeXmlText(text) + "</f>";
    }
    switch (t) {
      case "empty":
        if (fXml === "") return null;
        return this._cellTag(ref, this._resolveStyle(s, z, "number", undefined), "", fXml);
      case "number": {
        if (v === undefined || v === null) {
          if (fXml === "") throw badValue(ref, "number", v);
          return this._cellTag(ref, this._resolveStyle(s, z, "number", undefined), "", fXml);
        }
        const n = coerceNumber(v, ref);
        return this._cellTag(ref, this._resolveStyle(s, z, "number", n), "", fXml + "<v>" + numberLiteral(n) + "</v>");
      }
      case "date": {
        const ser = this._serial(v, serial, ref);
        return this._cellTag(ref, this._resolveStyle(s, z, "date", ser), "", fXml + "<v>" + numberLiteral(ser) + "</v>");
      }
      case "boolean": {
        const b = v === true || v === 1 || v === "1" || v === "TRUE" || v === "true";
        return this._cellTag(ref, this._resolveStyle(s, z, "boolean", b ? 1 : 0), ' t="b"', fXml + "<v>" + (b ? "1" : "0") + "</v>");
      }
      case "string": {
        if (v === undefined || v === null) {
          // A formula whose result is a string but whose cached value is absent.
          if (fXml === "") return null;
          return this._cellTag(ref, this._resolveStyle(s, z, "string", undefined), ' t="str"', fXml);
        }
        const text = String(v);
        const style = this._resolveStyle(s, z, "string", 0);
        if (fXml !== "") return this._cellTag(ref, style, ' t="str"', fXml + "<v>" + encodeXmlText(text) + "</v>");
        if (Array.isArray(runs) && runs.length > 0) {
          return this._cellTag(ref, style, ' t="s"', "<v>" + this.book.sharedStrings.add(richSi(runs)) + "</v>");
        }
        if (this._inlineStrings) {
          return this._cellTag(ref, style, ' t="inlineStr"', "<is>" + textElement(text) + "</is>");
        }
        return this._cellTag(ref, style, ' t="s"', "<v>" + this.book.sharedStrings.add(plainSi(text)) + "</v>");
      }
      case "error": {
        const lit = String(v === undefined || v === null ? e : v);
        if (!ERROR_LITERALS.has(lit)) {
          throw new GridxcelError("unknown error literal " + JSON.stringify(lit) + " at " + ref +
            "; expected one of " + Array.from(ERROR_LITERALS).join(" "), "BAD_VALUE");
        }
        return this._cellTag(ref, this._resolveStyle(s, z, "error", undefined), ' t="e"', fXml + "<v>" + encodeXmlText(lit) + "</v>");
      }
      default:
        throw new GridxcelError("unknown cell type " + JSON.stringify(t) + " at " + ref, "BAD_TYPE");
    }
  }

  _cellTag(ref, styleIdx, tAttr, inner) {
    let s = '<c r="' + ref + '"';
    if (styleIdx > 0) s += ' s="' + styleIdx + '"';
    if (tAttr !== "") s += tAttr;
    return inner === "" ? s + "/>" : s + ">" + inner + "</c>";
  }

  _serial(v, serial, ref) {
    if (typeof serial === "number") return finite(serial, ref);
    if (v instanceof Date) return finite(dateToSerial(v, this.book.date1904), ref);
    if (typeof v === "number") return finite(v, ref);
    throw new GridxcelError("date cell at " + ref + " needs a Date or a numeric serial, got " + typeof v, "BAD_VALUE");
  }

  /**
   * Cell.s (style object, or an index into the source workbook's table) plus
   * Cell.z (number-format code or id) -> cellXfs index.
   */
  _resolveStyle(s, z, t, value) {
    const table = this.book.styleTable;
    // Cell.z is a format code or a numFmtId; both end up as a code.
    const code = this._formatCode(z);
    const hasZ = code !== undefined && code !== "";
    if (s === undefined || s === null) {
      if (!hasZ) return t === "date" ? table.xf({ numFmt: defaultDateCode(value) }) : 0;
      return table.xf({ numFmt: code });
    }
    if (typeof s === "object") {
      return hasZ ? table.xf({ ...s, numFmt: code }) : table.xf(s);
    }
    if (!Number.isInteger(s) || s < 0) {
      throw new GridxcelError("invalid style index " + s, "BAD_STYLE");
    }
    const fallback = () => (t === "date" ? table.xf({ numFmt: defaultDateCode(value) }) : 0);
    if (this._styleObjects === null) return hasZ ? table.xf({ numFmt: code }) : fallback();
    const base = this._styleObjects[s];
    if (base === undefined) return hasZ ? table.xf({ numFmt: code }) : fallback();
    if (!hasZ || (t === "date" && code === "General")) {
      // No usable format id (or a General one on a cell that the reader typed
      // as a date): prefer the seeded style, and if that resolved to nothing
      // keep the cell rendering as a date rather than as a raw serial.
      const mapped = this._styleIndex[s];
      const idx = mapped === undefined ? 0 : mapped;
      return idx === 0 ? fallback() : idx;
    }
    const key = s + "\u0000" + code;
    let idx = this._styleCache.get(key);
    if (idx === undefined) {
      idx = table.xf({ ...base, numFmt: code });
      this._styleCache.set(key, idx);
    }
    return idx;
  }

  /** Cell.z, which may be a code string or a numFmtId, -> a format code. */
  _formatCode(z) {
    if (z === undefined || z === null || z === "") return undefined;
    if (typeof z !== "number") return z;
    return codeForNumFmtId(z, this._numFmtCodes);
  }

  _mappedStyleIndex(n) {
    if (typeof n !== "number" || this._styleIndex === null) return 0;
    const mapped = this._styleIndex[n];
    return mapped === undefined ? 0 : mapped;
  }

  _rowXml(r, parts, minC, maxC, info) {
    let s = '<row r="' + r + '"';
    if (minC >= 0) s += ' spans="' + (minC + 1) + ":" + (maxC + 1) + '"';
    if (info !== undefined && info !== null) {
      if (info.height !== undefined && info.height !== null) s += ' ht="' + info.height + '" customHeight="1"';
      if (info.hidden) s += ' hidden="1"';
      if (info.outlineLevel) s += ' outlineLevel="' + info.outlineLevel + '"';
      if (info.style !== undefined && info.style !== null) {
        s += ' s="' + this._mappedStyleIndex(info.style) + '" customFormat="1"';
      }
    }
    if (parts === null || parts.length === 0) return s + "/>";
    return s + ">" + parts.join("") + "</row>";
  }

  _finalize() {
    if (this._finalized) return;
    this._finalized = true;
    this._closed = true;
    for (const m of this._merges) {
      const range = decodeRange(m);
      extendBounds(this._bounds, range.s.r, range.s.c);
      extendBounds(this._bounds, range.e.r, range.e.c);
    }
    for (const h of this._hyperlinks) {
      const range = decodeRange(String(h.ref));
      extendBounds(this._bounds, range.s.r, range.s.c);
      extendBounds(this._bounds, range.e.r, range.e.c);
    }
  }

  _hasRow(r) {
    const a = this._rowNums;
    let lo = 0;
    let hi = a.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (a[mid] === r) return true;
      if (a[mid] < r) lo = mid + 1;
      else hi = mid - 1;
    }
    return false;
  }

  _sheetDataXml() {
    if (this._sheetData !== null) return this._sheetData;
    const pending = [];
    for (const r of this._rowInfo.keys()) if (!this._hasRow(r)) pending.push(r);
    if (pending.length === 0) {
      this._sheetData = this._rows.join("");
      return this._sheetData;
    }
    // Row info for rows that never received a cell still has to reach the
    // file, and sheetData rows must stay in ascending r order.
    pending.sort((a, b) => a - b);
    const merged = [];
    let i = 0;
    for (const r of pending) {
      while (i < this._rowNums.length && this._rowNums[i] < r) merged.push(this._rows[i++]);
      if (i < this._rowNums.length && this._rowNums[i] === r) { i++; continue; }
      merged.push(this._rowXml(r, null, -1, -1, this._rowInfo.get(r)));
    }
    while (i < this._rows.length) merged.push(this._rows[i++]);
    this._sheetData = merged.join("");
    return this._sheetData;
  }

  _dimensionRef() {
    if (this._dims !== null) return this._dims;
    if (this._bounds.maxR < 0) return this._merges.length > 0 ? this._merges[0] : "A1";
    return boundsToRef(this._bounds);
  }

  _sheetViewsXml() {
    let s = "<sheetViews><sheetView";
    if (this.index === 0) s += ' tabSelected="1"';
    s += ' workbookViewId="0"';
    const f = this._freeze;
    if (f === null) return s + "/></sheetViews>";
    s += ">";
    let pane = "topLeft";
    if (f.row > 0 && f.col > 0) pane = "bottomRight";
    else if (f.row > 0) pane = "bottomLeft";
    else if (f.col > 0) pane = "topRight";
    const topLeft = encodeCell(f.row, f.col);
    s += "<pane";
    if (f.col > 0) s += ' xSplit="' + f.col + '"';
    if (f.row > 0) s += ' ySplit="' + f.row + '"';
    s += ' topLeftCell="' + topLeft + '" activePane="' + pane + '" state="frozen"/>';
    s += '<selection pane="' + pane + '" activeCell="' + topLeft + '" sqref="' + topLeft + '"/>';
    return s + "</sheetView></sheetViews>";
  }

  _sheetFormatPrXml() {
    let levelRow = 0;
    let levelCol = 0;
    for (const info of this._rowInfo.values()) {
      if (info && info.outlineLevel > levelRow) levelRow = info.outlineLevel;
    }
    for (const col of this._columns) {
      if (col.outlineLevel > levelCol) levelCol = col.outlineLevel;
    }
    const height = this._defaultRowHeight === undefined ? DEFAULT_ROW_HEIGHT : this._defaultRowHeight;
    let s = '<sheetFormatPr defaultRowHeight="' + height + '"';
    if (this._defaultColWidth !== undefined) s += ' defaultColWidth="' + this._defaultColWidth + '"';
    if (levelRow > 0) s += ' outlineLevelRow="' + levelRow + '"';
    if (levelCol > 0) s += ' outlineLevelCol="' + levelCol + '"';
    return s + "/>";
  }

  _colsXml() {
    if (this._columns.length === 0) return "";
    let s = "<cols>";
    for (const col of this._columns) {
      s += '<col min="' + (col.min + 1) + '" max="' + (col.max + 1) + '"';
      if (col.width !== undefined && col.width !== null) {
        s += ' width="' + col.width + '" customWidth="1"';
      }
      if (col.hidden) s += ' hidden="1"';
      if (col.outlineLevel) s += ' outlineLevel="' + col.outlineLevel + '"';
      if (col.bestFit) s += ' bestFit="1"';
      if (col.style !== undefined && col.style !== null) s += ' style="' + this._mappedStyleIndex(col.style) + '"';
      s += "/>";
    }
    return s + "</cols>";
  }

  _hyperlinksXml() {
    let s = "<hyperlinks>";
    let rid = 0;
    for (const h of this._hyperlinks) {
      s += '<hyperlink ref="' + encodeXmlAttr(String(h.ref)) + '"';
      if (h.target) {
        rid++;
        s += ' r:id="rId' + rid + '"';
      }
      if (h.location) s += ' location="' + encodeXmlAttr(String(h.location)) + '"';
      if (h.tooltip) s += ' tooltip="' + encodeXmlAttr(String(h.tooltip)) + '"';
      if (h.display) s += ' display="' + encodeXmlAttr(String(h.display)) + '"';
      s += "/>";
    }
    return s + "</hyperlinks>";
  }

  /** The worksheet part, or null when the sheet holds nothing at all. */
  toXml() {
    this._finalize();
    let s = XML_DECL + '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">';
    if (this._tabColor !== null) s += "<sheetPr><tabColor" + colorAttrs(this._tabColor) + "/></sheetPr>";
    s += '<dimension ref="' + this._dimensionRef() + '"/>';
    s += this._sheetViewsXml();
    s += this._sheetFormatPrXml();
    s += this._colsXml();
    s += "<sheetData>" + this._sheetDataXml() + "</sheetData>";
    const filter = this._autoFilter === true ? this._dimensionRef() : this._autoFilter;
    if (filter !== null) s += '<autoFilter ref="' + filter + '"/>';
    if (this._merges.length > 0) {
      s += '<mergeCells count="' + this._merges.length + '">';
      for (const m of this._merges) s += '<mergeCell ref="' + encodeXmlAttr(m) + '"/>';
      s += "</mergeCells>";
    }
    if (this._hyperlinks.length > 0) s += this._hyperlinksXml();
    s += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
    return s + "</worksheet>";
  }

  /** sheetN.xml.rels, or null when there is nothing to relate. */
  relsXml() {
    let hasExternal = false;
    for (const h of this._hyperlinks) if (h.target) { hasExternal = true; break; }
    if (!hasExternal) return null;
    let s = XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">';
    let rid = 0;
    for (const h of this._hyperlinks) {
      if (!h.target) continue;
      rid++;
      s += '<Relationship Id="rId' + rid + '" Type="' + REL_HYPERLINK + '" Target="' +
        encodeXmlAttr(String(h.target)) + '" TargetMode="External"/>';
    }
    return s + "</Relationships>";
  }
}

function defaultDateCode(serial) {
  if (typeof serial !== "number") return DATE_CODE;
  return Math.abs(serial - Math.round(serial)) > 1e-9 ? DATETIME_CODE : DATE_CODE;
}

/**
 * numFmtId from a read workbook -> the format code this writer re-emits.
 *
 * A read cell carries the numeric id (Cell.z) and a read style carries the
 * numeric id in StyleObject.numFmt; without this resolution the writer would
 * have to invent a code and the format would degrade to General (or to the
 * default date format) on a read -> write round trip. Custom ids resolve
 * through the workbook's numFmts map, built-ins through ECMA-376 so they map
 * back onto their own built-in id.
 */
function codeForNumFmtId(id, numFmtCodes) {
  if (typeof id !== "number") return id;          // already a code
  if (id === 0) return "General";
  const custom = numFmtCodes === null || numFmtCodes === undefined ? undefined : numFmtCodes.get(id);
  if (custom !== undefined) return custom;
  return BUILTIN_FORMATS[id];                     // undefined when unknown
}

/* --------------------------------------------------------- workbook writer */

/**
 * Builds a complete .xlsx package.
 *
 * opts: date1904/dateSystem, props, definedNames, recalc, inlineStrings,
 *       compressionLevel, deterministic, activeTab, defaultRowHeight.
 */
export class WorkbookWriter {
  constructor(opts = {}) {
    this.opts = opts;
    this.date1904 = opts.date1904 === true || opts.dateSystem === 1904;
    this.styleTable = opts.styleTable instanceof StyleTable ? opts.styleTable : new StyleTable();
    this.sharedStrings = new SharedStrings();
    this.inlineStrings = opts.inlineStrings === true;
    this.recalc = opts.recalc !== false;
    this.props = opts.props || {};
    this.definedNames = opts.definedNames || [];
    this.activeTab = opts.activeTab === undefined ? 0 : opts.activeTab;
    this.deterministic = opts.deterministic === true;
    this.compressionLevel = opts.compressionLevel === undefined ? 6 : opts.compressionLevel;
    this.application = opts.application || "gridxcel";
    this.sheets = [];
    this._names = new Set();
    this._strings = null;
    this._bytes = null;
  }

  addSheet(name, opts = {}) {
    if (this._strings !== null || this._bytes !== null) {
      throw new GridxcelError("cannot add a sheet after the package has been built", "WRITER_CLOSED");
    }
    // Scan for a free name rather than guessing from the sheet count: the count
    // says nothing about which names are taken, so "Sheet2" + one existing sheet
    // produced "Sheet2" and threw BAD_SHEET_NAME on a call that should succeed.
    const resolved = name === undefined || name === null ? freeSheetName(this._names) : String(name);
    validateSheetName(resolved, this._names);
    const sheet = new SheetWriter(this, resolved, opts);
    this._names.add(resolved);
    this.sheets.push(sheet);
    return sheet;
  }

  /** Raw XML for every part, keyed by part name. */
  stringParts() {
    if (this._strings !== null) return this._strings;
    for (const sheet of this.sheets) sheet._finalize();
    const sharedStrings = this.sharedStrings.toXml();
    const map = new Map();
    map.set("[Content_Types].xml", this._contentTypesXml());
    map.set("_rels/.rels", this._rootRelsXml());
    map.set("docProps/app.xml", this._appXml());
    map.set("docProps/core.xml", this._coreXml());
    map.set("xl/workbook.xml", this._workbookXml());
    map.set("xl/_rels/workbook.xml.rels", this._workbookRelsXml());
    map.set("xl/styles.xml", this.styleTable.toXml());
    if (sharedStrings !== null) map.set("xl/sharedStrings.xml", sharedStrings);
    for (const sheet of this.sheets) {
      map.set(sheet.partName, sheet.toXml());
      const rels = sheet.relsXml();
      if (rels !== null) map.set(sheet.relsPartName, rels);
    }
    this._strings = map;
    return map;
  }

  /** stringParts() encoded to UTF-8. Cached: encoding a big sheet is not free. */
  parts() {
    if (this._encoded !== undefined && this._encoded !== null) return this._encoded;
    const out = new Map();
    for (const [name, text] of this.stringParts()) out.set(name, UTF8.encode(text));
    this._encoded = out;
    return out;
  }

  /** Deflate every part and assemble the zip. Idempotent. */
  finish() {
    if (this._bytes !== null) return this._bytes;
    const zip = new ZipWriter({ compressionLevel: this.compressionLevel, deterministic: this.deterministic });
    for (const [name, text] of this.stringParts()) zip.add(name, text);
    this._bytes = zip.build();
    return this._bytes;
  }

  /* ------------------------------------------------------------- parts */

  _contentTypesXml() {
    let s = XML_DECL + '<Types xmlns="' + NS_CT + '">';
    s += '<Default Extension="rels" ContentType="' + CT_RELS + '"/>';
    s += '<Default Extension="xml" ContentType="application/xml"/>';
    s += '<Override PartName="/xl/workbook.xml" ContentType="' + CT_WORKBOOK + '"/>';
    for (const sheet of this.sheets) {
      s += '<Override PartName="/' + sheet.partName + '" ContentType="' + CT_WORKSHEET + '"/>';
    }
    s += '<Override PartName="/xl/styles.xml" ContentType="' + CT_STYLES + '"/>';
    if (this.sharedStrings.uniqueCount > 0) {
      s += '<Override PartName="/xl/sharedStrings.xml" ContentType="' + CT_SHARED_STRINGS + '"/>';
    }
    s += '<Override PartName="/docProps/core.xml" ContentType="' + CT_CORE + '"/>';
    s += '<Override PartName="/docProps/app.xml" ContentType="' + CT_APP + '"/>';
    return s + "</Types>";
  }

  _rootRelsXml() {
    let s = XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">';
    s += '<Relationship Id="rId1" Type="' + REL_OFFICE_DOC + '" Target="xl/workbook.xml"/>';
    s += '<Relationship Id="rId2" Type="' + REL_CORE + '" Target="docProps/core.xml"/>';
    s += '<Relationship Id="rId3" Type="' + REL_APP + '" Target="docProps/app.xml"/>';
    return s + "</Relationships>";
  }

  _workbookRelsXml() {
    let s = XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">';
    this.sheets.forEach((sheet, i) => {
      s += '<Relationship Id="rId' + (i + 1) + '" Type="' + REL_WORKSHEET + '" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
    });
    let rid = this.sheets.length + 1;
    s += '<Relationship Id="rId' + rid + '" Type="' + REL_STYLES + '" Target="styles.xml"/>';
    rid++;
    if (this.sharedStrings.uniqueCount > 0) {
      s += '<Relationship Id="rId' + rid + '" Type="' + REL_SHARED_STRINGS + '" Target="sharedStrings.xml"/>';
    }
    return s + "</Relationships>";
  }

  _workbookXml() {
    let s = XML_DECL + '<workbook xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">';
    if (this.date1904) s += '<workbookPr date1904="1"/>';
    s += '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="18000" activeTab="' +
      this.activeTab + '"/></bookViews>';
    s += "<sheets>";
    this.sheets.forEach((sheet, i) => {
      s += '<sheet name="' + encodeXmlAttr(sheet.name) + '" sheetId="' + (i + 1) + '"' +
        (sheet.hidden ? ' state="hidden"' : "") + ' r:id="rId' + (i + 1) + '"/>';
    });
    s += "</sheets>";
    if (this.definedNames.length > 0) {
      s += "<definedNames>";
      for (const dn of this.definedNames) {
        if (dn === undefined || dn === null || !dn.name || !dn.ref) {
          throw new GridxcelError("defined name needs both name and ref", "BAD_DEFINED_NAME");
        }
        s += '<definedName name="' + encodeXmlAttr(String(dn.name)) + '"';
        if (dn.sheet !== undefined && dn.sheet !== null) s += ' localSheetId="' + dn.sheet + '"';
        if (dn.hidden) s += ' hidden="1"';
        if (dn.comment) s += ' comment="' + encodeXmlAttr(String(dn.comment)) + '"';
        s += ">" + encodeXmlText(String(dn.ref)) + "</definedName>";
      }
      s += "</definedNames>";
    }
    if (this.recalc) s += '<calcPr calcId="0" fullCalcOnLoad="1"/>';
    return s + "</workbook>";
  }

  _coreXml() {
    const p = this.props;
    let s = XML_DECL + '<cp:coreProperties xmlns:cp="' + NS_CP + '" xmlns:dc="' + NS_DC +
      '" xmlns:dcterms="' + NS_DCTERMS + '" xmlns:dcmitype="' + NS_DCMITYPE + '" xmlns:xsi="' + NS_XSI + '">';
    if (p.title) s += "<dc:title>" + encodeXmlText(String(p.title)) + "</dc:title>";
    if (p.subject) s += "<dc:subject>" + encodeXmlText(String(p.subject)) + "</dc:subject>";
    if (p.creator) s += "<dc:creator>" + encodeXmlText(String(p.creator)) + "</dc:creator>";
    if (p.keywords) s += "<cp:keywords>" + encodeXmlText(String(p.keywords)) + "</cp:keywords>";
    if (p.description) s += "<dc:description>" + encodeXmlText(String(p.description)) + "</dc:description>";
    if (p.lastModifiedBy) s += "<cp:lastModifiedBy>" + encodeXmlText(String(p.lastModifiedBy)) + "</cp:lastModifiedBy>";
    if (p.category) s += "<cp:category>" + encodeXmlText(String(p.category)) + "</cp:category>";
    const created = this._stamp(p.created);
    const modified = this._stamp(p.modified === undefined ? p.created : p.modified);
    s += '<dcterms:created xsi:type="dcterms:W3CDTF">' + created + "</dcterms:created>";
    s += '<dcterms:modified xsi:type="dcterms:W3CDTF">' + modified + "</dcterms:modified>";
    return s + "</cp:coreProperties>";
  }

  _stamp(value) {
    if (value instanceof Date) return toW3CDTF(value.getTime());
    if (typeof value === "number") return toW3CDTF(value);
    if (typeof value === "string" && value !== "") return value;
    // A deterministic package must not embed the wall clock.
    if (this.deterministic) return EPOCH_ISO;
    return toW3CDTF(Date.now());
  }

  _appXml() {
    const names = this.sheets.map((sheet) => sheet.name);
    let s = XML_DECL + '<Properties xmlns="' + NS_EP + '" xmlns:vt="' + NS_VT + '">';
    s += "<Application>" + encodeXmlText(String(this.props.application || this.application)) + "</Application>";
    s += "<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>";
    s += '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr>' +
      "</vt:variant><vt:variant><vt:i4>" + names.length + "</vt:i4></vt:variant></vt:vector></HeadingPairs>";
    s += '<TitlesOfParts><vt:vector size="' + names.length + '" baseType="lpstr">' +
      names.map((n) => "<vt:lpstr>" + encodeXmlText(n) + "</vt:lpstr>").join("") + "</vt:vector></TitlesOfParts>";
    if (this.props.company) s += "<Company>" + encodeXmlText(String(this.props.company)) + "</Company>";
    s += "<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>" +
      "<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>";
    return s + "</Properties>";
  }
}

/* --------------------------------------------------------------- book API */

/**
 * Configure a WorkbookWriter from a read-model Book and push every row into
 * it. The single code path behind buildParts() and writeWorkbook().
 */
export function createBookWriter(book, opts = {}) {
  if (book === undefined || book === null) {
    throw new GridxcelError("buildParts needs a workbook", "BAD_BOOK");
  }
  const sheets = Array.isArray(book) ? book : book.sheets;
  if (!Array.isArray(sheets)) {
    throw new GridxcelError("buildParts needs book.sheets to be an array", "BAD_BOOK");
  }
  const writer = new WorkbookWriter({
    date1904: book.dateSystem === 1904,
    props: book.props,
    definedNames: book.definedNames,
    recalc: opts.recalc,
    inlineStrings: opts.inlineStrings,
    deterministic: opts.deterministic,
    compressionLevel: opts.compressionLevel,
    activeTab: opts.activeTab,
    defaultRowHeight: opts.defaultRowHeight,
  });
  // Seed the style table so numeric Cell.s indices from a read workbook keep
  // working after dedup renumbers the cellXfs table.
  const numFmtCodes = new Map();
  if (book.numFmts && typeof book.numFmts.forEach === "function") {
    book.numFmts.forEach((code, id) => numFmtCodes.set(Number(id), code));
  }
  const sourceStyles = Array.isArray(book.styles) ? book.styles : [];
  const styleObjects = [];
  const styleIndex = [];
  for (const style of sourceStyles) {
    const source = style === undefined || style === null ? {} : style;
    const numFmt = codeForNumFmtId(source.numFmt, numFmtCodes);
    const object = numFmt === source.numFmt ? source : { ...source, numFmt };
    styleObjects.push(object);
    styleIndex.push(writer.styleTable.xf(object));
  }
  for (const sheet of sheets) {
    if (sheet === undefined || sheet === null) continue;
    const originRow = sheet.originRow === undefined || sheet.originRow === null ? 1 : sheet.originRow;
    const originCol = sheet.originCol === undefined || sheet.originCol === null ? 0 : sheet.originCol;
    const target = writer.addSheet(sheet.name, {
      columns: sheet.columns,
      rowInfo: sheet.rowInfo,
      freeze: sheet.freeze,
      autoFilter: sheet.autoFilter,
      merges: sheet.merges,
      tabColor: sheet.tabColor,
      hidden: sheet.hidden === true,
      hyperlinks: sheet.hyperlinks,
      dims: sheet.dims,
      originCol,
      styleObjects,
      styleIndex,
      numFmtCodes,
      inlineStrings: opts.inlineStrings,
    });
    const rows = sheet.rows;
    // Rows the reader could not keep in the dense array are still part of the
    // sheet and must survive a rewrite. Merge them with the dense rows in
    // ascending order, which is what SheetWriter.addRow requires.
    const sparse = sheet.sparseRows instanceof Map ? sheet.sparseRows : null;
    if (Array.isArray(rows)) {
      let nextSparse = null;
      if (sparse !== null) {
        nextSparse = [...sparse.keys()].sort((a, b) => a - b);
      }
      let si = 0;
      for (let i = 0; i < rows.length; i++) {
        while (nextSparse !== null && si < nextSparse.length && nextSparse[si] < i) {
          const cells = sparse.get(nextSparse[si]);
          if (cells !== undefined && cells !== null) target.addRow(originRow + nextSparse[si], cells);
          si++;
        }
        const row = rows[i];
        if (row === undefined || row === null) continue;
        target.addRow(originRow + i, row);
      }
      while (nextSparse !== null && si < nextSparse.length) {
        const cells = sparse.get(nextSparse[si]);
        if (cells !== undefined && cells !== null) target.addRow(originRow + nextSparse[si], cells);
        si++;
      }
    }
    target.end();
  }
  return writer;
}

/**
 * Turn a read-model Book into package parts without touching the filesystem.
 * Returns the raw XML strings so tests and the streaming API share one code
 * path.
 */
export function buildParts(book, opts = {}) {
  return createBookWriter(book, opts).stringParts();
}

