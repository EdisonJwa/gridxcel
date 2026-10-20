/**
 * Sheet <-> plain data: arrays of arrays, arrays of objects, and the value
 * shaping rules shared by csv.js and compat.js.
 *
 * Ownership: api-eng (docs/ARCHITECTURE.md). Imports address.js and numfmt.js
 * only - never the zip layer - so a partially built tree still loads.
 *
 * Cell construction, frozen by the data model:
 *   string   -> { t:"string",  v }
 *   number   -> { t:"number",  v }
 *   boolean  -> { t:"boolean", v }
 *   Date     -> { t:"date", v: Date, serial: dateToSerial(v), z: dateNF }
 *   error    -> { t:"error", v: null, e: "#NUM!" }  (NaN / +-Infinity)
 *   nullish  -> no cell at all, or { t:"empty" } when sheetStubs is set
 *   object   -> a Cell ({v|f|t}) passes through; anything else throws BAD_CELL
 *
 * The serial always travels with the Date so writer-eng never has to rebuild a
 * calendar, and z carries the format code so a written file still looks like a
 * date. Every produced cell also carries its A1 address, computed through
 * src/address.js, matching what readWorkbook hands back.
 *
 * Options understood here and in csv.js:
 *   raw        sheetToJSON: A/B/C keys and no header row; elsewhere: no
 *              formatted text (Cell.w is ignored, numbers use their value)
 *   rawNumbers keep numeric values numeric even when Cell.w exists
 *   defval     value used for cells that are missing inside the range
 *   blankrows  keep rows with no content (sheetToJSON: false for object
 *              output, true for header:1; sheetToAOA: true - a dense grid
 *              keeps its rows)
 *   header     sheetToJSON: 1 for arrays of arrays, "A" for column-letter
 *              keys, string[] for names, n > 1 for a 1-based header row. Only
 *              the absent and n > 1 forms consume a header row; SheetJS treats
 *              the first row as data for the other three.
 *              aoaToSheet: string[] prepended above the data
 *   range      "A1:C5" / {s,e} / a 1-based first row (SheetJS style)
 *   sheetStubs emit { t:"empty" } instead of dropping nullish values
 *   dateNF     date format code, default "yyyy-mm-dd"
 *   cellDates  false emits { t:"number", v: serial } for Date inputs
 *   skipHidden drop rows/columns flagged hidden (rowInfo / columns)
 *   sheet|name name for a produced sheet
 *
 * DELIBERATE DEVIATION: with header:1 SheetJS leaves array holes for missing
 * cells; we emit null there (or defval when supplied), the same dense shape
 * sheetToAOA produces. A hole carries no information but silently changes
 * length and iteration semantics for every consumer, so the grid is always
 * dense. This is the one place the SheetJS output is not byte-identical.
 */

import { GridxcelError } from "./errors.js";
import { encodeCell, encodeCol, encodeRange, toRange, decodeCell } from "./address.js";
import { MAX_COLUMNS, DEFAULT_SHEET_NAME } from "./constants.js";
import { dateToSerial, builtinFormat, formatValue } from "./numfmt.js";

export const DEFAULT_DATE_NF = "yyyy-mm-dd";

const A1_RE = /^[A-Za-z]{1,3}[0-9]{1,7}$/;
const KEY_CACHE = new Map();

/* ------------------------------------------------------------ value -> cell */

function inferType(v) {
  if (v === null || v === undefined) return "empty";
  if (typeof v === "number") return "number";
  if (typeof v === "boolean") return "boolean";
  if (v instanceof Date) return "date";
  return "string";
}

/** An empty sheet with every documented field present. */
export function blankSheet(name = DEFAULT_SHEET_NAME) {
  return {
    name,
    rows: [],
    originRow: 1,
    originCol: 0,
    columnCount: 0,
    rowCount: 0,
    merges: [],
    columns: [],
    rowInfo: new Map(),
  };
}

/**
 * Build a date cell. t stays "date" with a Date value by default; cellDates
 * false produces the SheetJS-shaped { t:"number", v: serial } instead.
 */
export function dateCell(date, opts = {}) {
  const dateNF = opts.dateNF || DEFAULT_DATE_NF;
  const serial = dateToSerial(date, opts.date1904 === true);
  if (opts.cellDates === false) return { v: serial, t: "number", z: dateNF };
  return { v: date, t: "date", serial, z: dateNF };
}

/**
 * Build a Cell from a plain JS value. undefined means "no cell here".
 * Rich objects ({v,t,f,z,s,w,c}) pass through with their fields preserved.
 *
 * A plain object or an array is NOT a cell and is not coerced: it throws
 * BAD_CELL. This is the agreed resolution of DUP-1. json.js used to
 * JSON.stringify such a value, which wrote a cell containing text the caller
 * never supplied - a silent wrong answer, and the worst kind of divergence
 * because api.normalizeCell threw BAD_CELL for the very same input. The two
 * now agree on throwing, and nothing is invented. api.normalizeCell delegates
 * here, so this is the only implementation.
 */
export function cellFromValue(v, opts = {}) {
  if (v === null || v === undefined) {
    return opts.sheetStubs ? { v: null, t: "empty" } : undefined;
  }
  const tt = typeof v;
  if (tt === "string") return { v, t: "string" };
  if (tt === "boolean") return { v, t: "boolean" };
  if (tt === "number") {
    if (Number.isFinite(v)) return { v, t: "number" };
    // Excel has no NaN/Infinity: the model's honest answer is an error cell.
    return { v: null, t: "error", e: "#NUM!" };
  }
  if (v instanceof Date) return dateCell(v, opts);
  if (tt === "object") {
    if (!Array.isArray(v) && ("v" in v || "f" in v || "t" in v)) {
      const cell = { ...v };
      if (cell.t === undefined) cell.t = inferType(cell.v);
      if (cell.t === "date" && cell.v instanceof Date && cell.serial === undefined) {
        cell.serial = dateToSerial(cell.v, opts.date1904 === true);
      }
      if (cell.v === undefined) cell.v = null;
      return cell;
    }
    // Not a Cell: an array, or an object without v/f/t. See the doc comment.
    throw new GridxcelError("cell object must have a v property", "BAD_CELL", {
      input: Array.isArray(v) ? "array" : Object.keys(v).slice(0, 8),
    });
  }
  throw new GridxcelError("unsupported cell value of type " + tt, "BAD_CELL");
}

/* ------------------------------------------------------------ sheet -> cell */

function originRowOf(sheet) {
  return sheet.originRow === undefined ? 1 : sheet.originRow;
}
function originColOf(sheet) {
  return sheet.originCol === undefined ? 0 : sheet.originCol;
}

/**
 * The row object holding row r, or undefined.
 * r is a 0-based sheet row index, the same convention as A1 ranges
 * ("A1" -> r 0) and as sheetRange()/resolveRange(); originRow is 1-based.
 */
export function rowAt(sheet, r) {
  const i = r - (originRowOf(sheet) - 1);
  return i < 0 ? undefined : sheet.rows[i];
}

/** The cell at absolute row r / column c, or undefined. */
export function cellAt(sheet, r, c) {
  const row = rowAt(sheet, r);
  if (!row) return undefined;
  const j = c - originColOf(sheet);
  return j < 0 ? undefined : row[j];
}

/** True when a cell counts as content (a formula with no cached value does). */
function hasContent(cell) {
  if (!cell || cell.t === "empty") return false;
  if (cell.v === undefined || cell.v === null) return cell.f !== undefined;
  return true;
}

/** Used range of a sheet, derived from the cells that actually hold content. */
export function sheetRange(sheet) {
  const rows = sheet.rows || [];
  const base = originRowOf(sheet) - 1;
  const col0 = originColOf(sheet);
  let minR = Infinity;
  let minC = Infinity;
  let maxR = -1;
  let maxC = -1;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    for (let j = 0; j < row.length; j++) {
      if (!hasContent(row[j])) continue;
      const r = base + i;
      const c = col0 + j;
      if (r < minR) minR = r;
      if (c < minC) minC = c;
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
  }
  if (maxR < 0) return { s: { r: 0, c: 0 }, e: { r: -1, c: -1 }, empty: true };
  return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC }, empty: false };
}

/**
 * Resolve opts.range against the used range.
 * A bare number is the 1-based first row and runs to the last used row, which
 * is how SheetJS spells "start at row N".
 */
export function resolveRange(sheet, opts = {}) {
  const used = sheetRange(sheet);
  const r = opts.range;
  if (r === undefined || r === null) return used;
  if (typeof r === "number") {
    return { s: { r: Math.max(0, Math.floor(r) - 1), c: used.s.c }, e: { r: used.e.r, c: used.e.c }, empty: used.empty };
  }
  const parsed = toRange(r);
  if (!parsed || !parsed.s || !parsed.e) {
    throw new GridxcelError("range must be \"A1:B2\", {s,e} or a row number", "BAD_RANGE");
  }
  return { s: parsed.s, e: parsed.e, empty: parsed.e.r < parsed.s.r };
}

/* ------------------------------------------------------------- hidden flags */

/** True when row r (1-based) is flagged hidden in sheet.rowInfo. */
export function isRowHidden(sheet, r) {
  const info = sheet.rowInfo;
  if (!info) return false;
  let rec;
  if (info instanceof Map) rec = info.get(r + 1);
  else rec = info[r + 1] !== undefined ? info[r + 1] : info[r];
  return !!(rec && rec.hidden);
}

/** True when column c (0-based) falls inside a hidden entry of sheet.columns. */
export function isColHidden(sheet, c) {
  const cols = sheet.columns;
  if (!Array.isArray(cols)) return false;
  for (const col of cols) {
    if (!col || !col.hidden) continue;
    const single = col.index === undefined ? col.min : col.index;
    const min = col.min === undefined ? single : col.min;
    const max = col.max === undefined ? (col.index === undefined ? min : col.index) : col.max;
    if (min === undefined) continue;
    if (c >= min && c <= (max === undefined ? min : max)) return true;
  }
  return false;
}

/* -------------------------------------------------------------- cell -> value */

/**
 * The JavaScript value a consumer sees for a cell.
 * Cell.w (pre-rendered text) wins unless rawNumbers is set, mirroring SheetJS:
 * formatting is display, the value is data.
 */
export function cellValue(cell, opts = {}) {
  if (!cell || cell.t === "empty") return undefined;
  if (cell.t === "error") return cell.e !== undefined ? cell.e : cell.v === undefined ? undefined : cell.v;
  if (cell.v === undefined || cell.v === null) return undefined;
  if (opts.rawNumbers !== true && typeof cell.w === "string") return cell.w;
  return cell.v;
}

/**
 * The display text of a cell, used by CSV output and compat.format_cell.
 * Numbers use String(v) rather than Excel's 11-significant-digit General
 * format: a CSV field is data, and String(v) survives a parse unchanged.
 * An explicit format code (or Cell.w) still wins when one is present.
 */
export function cellText(cell, opts = {}) {
  if (!cell || cell.t === "empty") return "";
  const t = cell.t;
  const v = cell.v;
  if (opts.raw !== true && typeof cell.w === "string") return cell.w;
  if (t === "error") return cell.e !== undefined ? String(cell.e) : v === undefined || v === null ? "" : String(v);
  if (v === undefined || v === null) return "";
  if (t === "boolean" || typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (v instanceof Date) {
    return formatValue(v, dateCode(cell, opts), opts);
  }
  if (t === "date" || typeof v === "number") {
    const code = numberFormatCode(cell.z);
    if (t === "date") {
      const serial = typeof v === "number" ? v : dateToSerial(v, opts.date1904 === true);
      return formatValue(serial, dateCode(cell, opts), opts);
    }
    if (code && opts.raw !== true) return formatValue(v, code, opts);
    return String(v);
  }
  return String(v);
}

/**
 * The date format code for a cell. An explicit opts.dateNF wins (the caller
 * asked for that rendering), then the cell's own format, then the default -
 * so a sheet read from a real file still shows the format it was saved with.
 */
function dateCode(cell, opts) {
  return opts.dateNF || numberFormatCode(cell.z) || DEFAULT_DATE_NF;
}

/** Cell.z is an id (read) or a code (write): both spellings are accepted. */
function numberFormatCode(z) {
  if (z === undefined || z === null) return undefined;
  if (typeof z === "string") {
    if (!/^[0-9]+$/.test(z)) return z;
    const id = Number(z);
    if (id === 0) return undefined;
    return builtinFormat(id);
  }
  if (typeof z === "number") {
    if (z === 0) return undefined;
    return builtinFormat(z);
  }
  return undefined;
}

/* ------------------------------------------------------------ sheet shapes */

function hasA1Keys(obj) {
  for (const k of Object.keys(obj)) {
    if (k.charAt(0) === "!") continue;
    if (A1_RE.test(k)) return true;
  }
  return false;
}

/**
 * Resolve sheet-shaped input into a gridxcel Sheet.
 *
 * Deliberately not a re-export of api.normalizeSheet (DUP-5). They answer
 * different questions: api.normalizeSheet builds a Sheet out of raw caller
 * values for the write path, while this recognises input that is already
 * sheet-shaped - a gridxcel Sheet, a sparse or dense SheetJS worksheet
 * ({!ref/A1: {...}} or {!data}), an AOA, or a Book ({sheets}, resolved through
 * opts.sheet) - and adapts it for the read/convert paths. The SheetJS
 * spellings are why it exists and cannot move into api.js, and merging the two
 * would either drop that compatibility or make the core depend on compat
 * shapes. What they genuinely share, the value -> Cell rule, is not repeated
 * here: it is json.js's cellFromValue, which api.normalizeCell delegates to.
 */
export function normalizeSheet(sheet, opts = {}) {
  if (sheet === null || sheet === undefined) return blankSheet();
  if (Array.isArray(sheet)) return aoaToSheet(sheet, opts);
  if (Array.isArray(sheet.rows)) return sheet;
  if (Array.isArray(sheet.sheets)) {
    const picked = pickFromBook(sheet, opts);
    return normalizeSheet(picked, opts);
  }
  if (sheet["!ref"] !== undefined || sheet["!data"] !== undefined || hasA1Keys(sheet)) {
    return worksheetToSheet(sheet);
  }
  throw new GridxcelError(
    "not a Sheet: expected { rows } or a SheetJS worksheet with !ref",
    "BAD_SHEET",
    { keys: Object.keys(sheet).slice(0, 8) },
  );
}

function pickFromBook(book, opts = {}) {
  const sel = opts.sheet === undefined ? opts.sheetName : opts.sheet;
  if (typeof sel === "number") return book.sheets[sel];
  if (typeof sel === "string") {
    const found = book.sheets.find((s) => s && s.name === sel);
    if (!found) throw new GridxcelError("no sheet named " + JSON.stringify(sel), "NO_SHEET");
    return found;
  }
  return book.sheets[0];
}

/** Convert a SheetJS worksheet (sparse or !data dense) into a gridxcel Sheet. */
export function worksheetToSheet(ws) {
  const cells = new Map();
  const colWidth = MAX_COLUMNS;
  let minR = Infinity;
  let minC = Infinity;
  let maxR = -1;
  let maxC = -1;
  const put = (r, c, cell) => {
    if (!cell) return;
    cells.set(r * colWidth + c, cell);
    if (r < minR) minR = r;
    if (c < minC) minC = c;
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  };

  const dense = Array.isArray(ws["!data"]) ? ws["!data"] : null;
  if (dense) {
    for (let r = 0; r < dense.length; r++) {
      const row = dense[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) put(r, c, fromWorksheetCell(row[c]));
    }
  } else {
    for (const k of Object.keys(ws)) {
      if (k.charAt(0) === "!" || !A1_RE.test(k)) continue;
      const addr = KEY_CACHE.get(k) || (KEY_CACHE.set(k, decodeCell(k)), KEY_CACHE.get(k));
      put(addr.r, addr.c, fromWorksheetCell(ws[k]));
    }
  }

  const ref = ws["!ref"] === undefined ? null : toRange(ws["!ref"]);
  if (ref) {
    if (ref.s.r < minR) minR = ref.s.r;
    if (ref.s.c < minC) minC = ref.s.c;
    if (ref.e.r > maxR) maxR = ref.e.r;
    if (ref.e.c > maxC) maxC = ref.e.c;
  }
  if (maxR < 0) return blankSheet();

  const originRow = minR + 1;
  const originCol = minC;
  const rows = [];
  for (let r = minR; r <= maxR; r++) {
    const row = [];
    for (let c = minC; c <= maxC; c++) {
      const cell = cells.get(r * colWidth + c);
      if (cell) row[c - minC] = cell;
    }
    rows.push(row);
  }
  const sheet = blankSheet();
  sheet.rows = rows;
  sheet.originRow = originRow;
  sheet.originCol = originCol;
  sheet.rowCount = maxR + 1;
  sheet.columnCount = maxC + 1;
  if (Array.isArray(ws["!merges"])) {
    sheet.merges = ws["!merges"].map((m) => encodeRange(m.s, m.e));
  }
  if (Array.isArray(ws["!cols"])) {
    sheet.columns = ws["!cols"].map((c, i) => ({
      index: i,
      min: c && c.min !== undefined ? c.min : i,
      max: c && c.max !== undefined ? c.max : i,
      width: c ? c.width : undefined,
      hidden: c ? !!c.hidden : false,
    }));
  }
  if (Array.isArray(ws["!rows"])) {
    for (let i = 0; i < ws["!rows"].length; i++) {
      const r = ws["!rows"][i];
      if (r) sheet.rowInfo.set(originRow + i, { index: originRow + i, height: r.hpt, hidden: !!r.hidden });
    }
  }
  return sheet;
}

function fromWorksheetCell(c) {
  if (c === null || c === undefined) return undefined;
  if (typeof c !== "object") return cellFromValue(c);
  const t = c.t;
  if (t === "z") return { v: null, t: "empty" };
  if (t === "d") {
    const d = c.v instanceof Date ? c.v : new Date(c.v);
    const cell = dateCell(d, { dateNF: typeof c.z === "string" ? c.z : undefined });
    if (c.f !== undefined) cell.f = c.f;
    return cell;
  }
  const out = { v: c.v === undefined ? null : c.v, t: mapSheetJsType(t, c.v) };
  if (c.f !== undefined) out.f = c.f;
  if (c.z !== undefined) out.z = c.z;
  if (c.w !== undefined) out.w = c.w;
  if (c.c !== undefined) out.c = c.c;
  if (out.t === "boolean" && typeof out.v !== "boolean") out.v = String(out.v).toUpperCase() === "TRUE";
  return out;
}

function mapSheetJsType(t, v) {
  if (t === "s" || t === "str") return "string";
  if (t === "n") return "number";
  if (t === "b") return "boolean";
  if (t === "e") return "error";
  if (t === "date" || t === "d") return "date";
  if (t === "empty") return "empty";
  return inferType(v);
}

/* --------------------------------------------------------------- AOA <-> sheet */

/**
 * Array of arrays -> Sheet. Rows are sparse; opts.header prepends a row.
 * Values obey cellFromValue, so Date, boolean, number and rich cells all land
 * in their documented shape.
 */
export function aoaToSheet(rows, opts = {}) {
  if (!Array.isArray(rows)) throw new GridxcelError("aoaToSheet expects an array of rows", "BAD_SHEET");
  const sheet = blankSheet(opts.sheet || opts.name || DEFAULT_SHEET_NAME);
  let width = 0;
  const addRow = (row) => {
    const out = [];
    if (row !== null && row !== undefined) {
      if (!Array.isArray(row)) throw new GridxcelError("each AOA row must be an array", "BAD_SHEET");
      const r = sheet.originRow - 1 + sheet.rows.length;
      for (let c = 0; c < row.length; c++) {
        const cell = cellFromValue(row[c], opts);
        if (!cell) continue;
        if (cell.a === undefined) cell.a = encodeCell(r, c);
        out[c] = cell;
      }
    }
    if (out.length > width) width = out.length;
    sheet.rows.push(out);
  };
  if (Array.isArray(opts.header)) addRow(opts.header);
  for (const row of rows) addRow(row);
  sheet.rowCount = sheet.rows.length;
  sheet.columnCount = width;
  return sheet;
}

/**
 * Largest row span the dense converters will materialise.
 *
 * A worksheet is untrusted input: a single <row r="40000000"> makes the reader
 * report a forty-million-row extent, and a dense conversion of that range
 * allocates gigabytes from a one-kilobyte file. The default caps the dense
 * output at 262,144 rows (about 16 MiB), and callers can pass an explicit
 * `range` or raise `maxRows` when they genuinely want more. The failure is an
 * actionable error rather than a process abort.
 */
const MAX_DENSE_ROWS = 262144;


/**
 * Refuse a dense conversion whose row span is untrustworthy.
 *
 * A worksheet is untrusted input: a single <row r="40000000"> makes the reader
 * report a forty-million-row extent, and densifying that range allocates
 * gigabytes from a one-kilobyte file. Callers who pass an explicit `range` are
 * taken at their word; callers who inherited the sheet's own extent are not.
 *
 * @param {object} ws normalised sheet
 * @param {object} range resolved range
 * @param {object} opts options; `range` and `maxRows` are honoured
 * @returns {number} the span that will actually be materialised
 */
export function assertDenseRange(ws, range, opts = {}) {
  const requested = range.e.r - range.s.r + 1;
  const declared = typeof ws.rowCount === "number" && ws.rowCount > 0
    ? ws.rowCount - (ws.originRow || 1) + 1
    : 0;
  const span = opts.range === undefined ? Math.max(requested, declared) : requested;
  const limit = Number.isInteger(opts.maxRows) && opts.maxRows > 0 ? opts.maxRows : MAX_DENSE_ROWS;
  if (span > limit) {
    throw new GridxcelError(
      "refusing to materialise " + span + " dense rows (limit " + limit + "); " +
        "the sheet declares " + span + " rows and this is untrusted input - " +
        "pass a bounded range (for example { range: \"A1:Z1000\" }), raise maxRows, " +
        "or stream it with eachRow()",
      "RANGE_TOO_LARGE",
      { span, limit, lastRow: range.e.r + 1 },
    );
  }
  return span;
}

/** Sheet -> dense array of arrays. Missing cells become defval ?? null. */
export function sheetToAOA(sheet, opts = {}) {
  const ws = normalizeSheet(sheet, opts);
  const range = resolveRange(ws, opts);
  const out = [];
  if (range.e.r < range.s.r) return out;
  assertDenseRange(ws, range, opts);
  const blankrows = opts.blankrows !== false; // dense grid: keep rows by default
  const skipHidden = opts.skipHidden === true;
  for (let r = range.s.r; r <= range.e.r; r++) {
    if (skipHidden && isRowHidden(ws, r)) continue;
    const row = [];
    let has = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (skipHidden && isColHidden(ws, c)) continue;
      const v = cellValue(cellAt(ws, r, c), opts);
      if (v === undefined) row.push(opts.defval === undefined ? null : opts.defval);
      else {
        row.push(v);
        has = true;
      }
    }
    if (!has && !blankrows) continue;
    out.push(row);
  }
  return out;
}

/* -------------------------------------------------------------- JSON <-> sheet */

/** Header key for column c: the cell's text, or "" when the header is empty. */
function headerKey(ws, row, c) {
  const cell = cellAt(ws, row, c);
  if (!cell || !hasContent(cell)) return "";
  return cellText(cell, { raw: true });
}

/**
 * SheetJS appends _1, _2... to repeated labels so a duplicate header cell
 * cannot silently overwrite the column before it (S, h, e, e, t, J, S becomes
 * S, h, e, e_1, t, J, S_1, which is what SheetJS produces for the same sheet).
 */
function disambiguate(keys) {
  const seen = new Set();
  return keys.map((key) => {
    if (key === "") return key;
    let candidate = key;
    let n = 1;
    while (seen.has(candidate)) candidate = key + "_" + n++;
    seen.add(candidate);
    return candidate;
  });
}

/**
 * Sheet -> objects, or arrays of arrays when header is 1. The header spellings
 * follow SheetJS exactly:
 *   (absent)       the first row of the range is the header row, data follows
 *   header: 1      array of arrays; EVERY row of the range is data
 *   header: "A"    keys are A, B, C...; every row of the range is data
 *                  (raw:true is the older spelling of the same thing)
 *   header: [...]  keys are these names; every row of the range is data. A
 *                  caller-supplied list is used verbatim, as SheetJS does, and
 *                  columns past its end fall back to the sheet's header text
 *   header: n      n > 1: row n (1-based) is the header row, data follows
 *
 * Keys read from the sheet are disambiguated, blank header cells produce no
 * key, and blankrows defaults to true for arrays (a dense grid keeps its rows)
 * and false for objects.
 */
export function sheetToJSON(sheet, opts = {}) {
  const ws = normalizeSheet(sheet, opts);
  const range = resolveRange(ws, opts);
  const out = [];
  if (range.e.r < range.s.r) return out;

  const skipHidden = opts.skipHidden === true;
  const cols = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    if (skipHidden && isColHidden(ws, c)) continue;
    cols.push(c);
  }

  const header = opts.header;
  const asArrays = header === 1;
  const blankrows = opts.blankrows === undefined ? asArrays : opts.blankrows === true;

  let keys = null;
  let headerRow = range.s.r;
  let firstData = range.s.r;
  if (asArrays) {
    firstData = range.s.r;
  } else if (header === "A" || opts.raw === true) {
    keys = cols.map((c) => encodeCol(c));
  } else if (Array.isArray(header)) {
    keys = cols.map((c, i) => {
      const name = header[i];
      if (name === undefined || name === null) return headerKey(ws, headerRow, c);
      return String(name);
    });
  } else {
    if (typeof header === "number" && header > 1) headerRow = Math.floor(header) - 1;
    keys = disambiguate(cols.map((c) => headerKey(ws, headerRow, c)));
    firstData = headerRow + 1;
  }

  for (let r = firstData; r <= range.e.r; r++) {
    if (skipHidden && isRowHidden(ws, r)) continue;
    if (asArrays) {
      const row = [];
      let has = false;
      for (const c of cols) {
        const v = cellValue(cellAt(ws, r, c), opts);
        if (v === undefined) row.push(opts.defval === undefined ? null : opts.defval);
        else {
          row.push(v);
          has = true;
        }
      }
      if (!has && !blankrows) continue;
      out.push(row);
      continue;
    }
    const row = {};
    let has = false;
    for (let i = 0; i < cols.length; i++) {
      const key = keys[i];
      if (key === undefined || key === "") continue;
      const v = cellValue(cellAt(ws, r, cols[i]), opts);
      if (v === undefined) {
        if (opts.defval !== undefined) row[key] = opts.defval;
        continue;
      }
      row[key] = v;
      has = true;
    }
    if (!has && !blankrows) continue;
    out.push(row);
  }
  return out;
}

/**
 * Array of objects -> Sheet. Keys keep first-seen order across every row, so
 * the header row and the objects the reader gets back stay in a stable order.
 * opts.header replaces/extends the key list, opts.skipHeader suppresses the
 * header row itself. An AOA input is passed straight to aoaToSheet.
 */
export function jsonToSheet(rows, opts = {}) {
  let list = rows;
  if (list === null || list === undefined) list = [];
  if (!Array.isArray(list)) {
    if (typeof list === "object") list = [list];
    else throw new GridxcelError("jsonToSheet expects an array of objects", "BAD_SHEET");
  }
  if (list.length && Array.isArray(list[0])) return aoaToSheet(list, opts);

  const keys = [];
  const seen = new Set();
  const addKey = (k) => {
    const key = String(k);
    if (!seen.has(key)) {
      seen.add(key);
      keys.push(key);
    }
  };
  if (Array.isArray(opts.header)) for (const k of opts.header) addKey(k);
  for (const row of list) {
    if (row === null || typeof row !== "object" || Array.isArray(row)) continue;
    for (const k of Object.keys(row)) addKey(k);
  }

  const aoa = [];
  if (opts.skipHeader !== true) aoa.push(keys.slice());
  for (const row of list) {
    const src = row !== null && typeof row === "object" && !Array.isArray(row) ? row : {};
    aoa.push(keys.map((k) => src[k]));
  }
  // opts.header means "key order" here, not "prepend this row", so it must not
  // reach aoaToSheet where the same name means the opposite.
  const sheet = aoaToSheet(aoa, { ...opts, header: undefined });
  sheet.name = opts.sheet || opts.name || sheet.name;
  return sheet;
}
