/**
 * CSV / DSV: RFC 4180 plus the extensions every real spreadsheet pipeline
 * needs (BOM, lone CR, ragged rows, embedded newlines, delimiters, type
 * inference) and the SheetJS-compatible option names.
 *
 * Ownership: api-eng (docs/ARCHITECTURE.md). Depends on json.js for the sheet
 * and cell model and address.js for addressing; numfmt.js is reached through
 * json.js, never directly.
 *
 * Options
 *   delimiter|FS   field separator, default "," (";" and "\t" are the usual
 *                  alternatives)
 *   quote          quote character, default '"'
 *   newline        "\n" (default) or "\r\n"; lineEnding is an explicit alias
 *   lineEnding     terminator for every row, including the last
 *   bom            prepend U+FEFF on output, strip it on input (always)
 *   header         string[] row written above the data / prepended to a parse
 *   alwaysQuote    quote every field (alias forceQuotes)
 *   escapeFormulae DEFAULT TRUE: a text field whose FIRST character is one of
 *                  = + - @ TAB CR LF gets a leading apostrophe, so Excel and
 *                  Sheets open it as text instead of executing it (the CSV
 *                  injection guard). TAB, CR and LF are included because
 *                  spreadsheet applications strip leading whitespace before
 *                  deciding that a cell is a formula, so "\t=1+1" is evaluated
 *                  exactly like "=1+1". Only text cells are escaped - numbers
 *                  and dates keep their exact shape, so a numeric -2.5 stays
 *                  numeric and is not quoted. csvToSheet does NOT strip the
 *                  apostrophe again; the guard is one-way by design, exactly
 *                  like Excel's own CSV output.
 *   dateNF         date format for Date cells, default "yyyy-mm-dd"
 *   quotedStrings  false (default): a field's type comes from its content, as
 *                  SheetJS and pandas do. true: a quoted field is taken
 *                  literally, and sheetToCSV quotes text that would otherwise
 *                  be re-typed, so csvToSheet(sheetToCSV(s), {quotedStrings:
 *                  true}) round-trips strings that look like numbers or dates.
 *   raw            no type inference on parse, no formatting on output
 *   blankrows      output: emit rows with no content (default true);
 *                  parse: keep blank lines (default true)
 *   skipHidden     drop rows/columns flagged hidden
 *   range, defval, sheetStubs, cellDates, sheet|name  as in json.js
 */

import { GridxcelError, NotImplementedError } from "./errors.js";
import { encodeCell } from "./address.js";
import { DEFAULT_SHEET_NAME } from "./constants.js";
import {
  assertDenseRange,
  DEFAULT_DATE_NF,
  blankSheet,
  cellAt,
  cellFromValue,
  cellText,
  dateCell,
  isColHidden,
  isRowHidden,
  normalizeSheet,
  resolveRange,
} from "./json.js";
import { emptyBook } from "./api.js";

const NUM_RE = /^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$/;
const BOOL_RE = /^(?:true|false)$/i;
const ISO_RE = /^([0-9]{4})-([0-9]{2})-([0-9]{2})(?:[T ]([0-9]{2}):([0-9]{2})(?::([0-9]{2})(?:\.([0-9]{1,9}))?)?(Z|[+-][0-9]{2}:?[0-9]{2})?)?$/;

/* ------------------------------------------------------------------ helpers */

function dsvDelimiter(opts) {
  const d = opts.delimiter === undefined ? opts.FS : opts.delimiter;
  if (d === undefined || d === null) return ",";
  const s = String(d);
  if (s.length === 0) throw new GridxcelError("delimiter must not be empty", "BAD_OPTION");
  return s;
}

function quoteChar(opts) {
  const q = opts.quote === undefined || opts.quote === null ? '"' : String(opts.quote);
  if (q.length !== 1) throw new GridxcelError("quote must be a single character", "BAD_OPTION");
  return q;
}

function lineEndingOf(opts) {
  if (opts.lineEnding !== undefined) return String(opts.lineEnding);
  if (opts.newline !== undefined) return String(opts.newline);
  return "\n";
}

function isTextCell(cell) {
  if (!cell) return false;
  if (cell.t === "string") return true;
  return cell.t === undefined && typeof cell.v === "string";
}

/** Text that a content-based parser would hand back as a non-string. */
function wouldBeRetyped(text, opts) {
  if (text === "") return false;
  if (BOOL_RE.test(text)) return true;
  if (NUM_RE.test(text)) return true;
  return parseDate(text, opts) !== null;
}

/* ------------------------------------------------------------- date parsing */

function validDate(y, m, d, h, mi, s, ms) {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  if (h > 23 || mi > 59 || s > 59) return null;
  const date = new Date(y, m - 1, d, h, mi, s, ms);
  // Reject impossible days the way Excel does (2024-02-31 is text, not a date).
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return null;
  return date;
}

/** ISO-8601 date or date-time, with or without a zone. */
function parseISO(text) {
  const m = ISO_RE.exec(text);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (m[4] === undefined) return validDate(y, mo, d, 0, 0, 0, 0);
  const ms = m[7] === undefined ? 0 : Number(m[7].padEnd(3, "0").slice(0, 3));
  const date = validDate(y, mo, d, Number(m[4]), Number(m[5]), Number(m[6] || 0), ms);
  if (!date) return null;
  if (m[8] !== undefined && m[8] !== "") {
    // An explicit zone means an instant: let Date normalise it, then convert
    // the instant back into local wall-clock time, which is what Excel stores.
    const instant = new Date(text.replace(" ", "T"));
    return Number.isNaN(instant.getTime()) ? date : instant;
  }
  return date;
}

/**
 * Compile a number-format code into a value parser, so a CSV written with
 * dateNF "mm/dd/yy" reads back as dates. Only the date/time tokens are
 * interpreted; everything else is a literal. "m" after an hour token (or
 * before seconds) is minutes, the one genuinely ambiguous rule.
 */
function compileDatePattern(nf) {
  if (typeof nf !== "string" || nf === "") return null;
  let re = "^";
  const parts = [];
  let i = 0;
  while (i < nf.length) {
    const ch = nf.charAt(i);
    if (ch === '"') {
      const end = nf.indexOf('"', i + 1);
      re += escapeRe(nf.slice(i + 1, end < 0 ? nf.length : end));
      i = end < 0 ? nf.length : end + 1;
      continue;
    }
    if (ch === "\\") {
      re += escapeRe(nf.charAt(i + 1));
      i += 2;
      continue;
    }
    const lower = ch.toLowerCase();
    if (lower === "y" || lower === "m" || lower === "d" || lower === "h" || lower === "s") {
      let j = i;
      while (j < nf.length && nf.charAt(j).toLowerCase() === lower) j++;
      const width = j - i;
      let kind = lower;
      if (lower === "m") {
        const prev = parts.length ? parts[parts.length - 1] : "";
        kind = prev === "h" || prev === "n" ? "n" : "m";
      }
      parts.push(kind);
      if (lower === "y") re += width >= 4 ? "([0-9]{4})" : "([0-9]{2})";
      else re += width >= 2 ? "([0-9]{2})" : "([0-9]{1,2})";
      i = j;
      continue;
    }
    if ((lower === "a" || lower === "p") && nf.slice(i, i + 5).toUpperCase() === "AM/PM") {
      parts.push("ap");
      re += "([AaPp][Mm])";
      i += 5;
      continue;
    }
    re += escapeRe(ch);
    i++;
  }
  if (!parts.some((p) => p === "y" || p === "m" || p === "d")) return null;
  let regex;
  try {
    regex = new RegExp(re + "$");
  } catch {
    return null;
  }
  return { regex, parts };
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const PATTERN_CACHE = new Map();

function parseWithPattern(text, nf) {
  let compiled = PATTERN_CACHE.get(nf);
  if (compiled === undefined) {
    compiled = compileDatePattern(nf);
    PATTERN_CACHE.set(nf, compiled);
  }
  if (!compiled) return null;
  const m = compiled.regex.exec(text);
  if (!m) return null;
  let y = 1970;
  let mo = 1;
  let d = 1;
  let h = 0;
  let mi = 0;
  let s = 0;
  let pm = false;
  for (let i = 0; i < compiled.parts.length; i++) {
    const v = Number(m[i + 1]);
    switch (compiled.parts[i]) {
      case "y": y = m[i + 1].length === 2 ? (v < 70 ? 2000 + v : 1900 + v) : v; break;
      case "m": mo = v; break;
      case "d": d = v; break;
      case "h": h = v; break;
      case "n": mi = v; break;
      case "s": s = v; break;
      case "ap": pm = /p/i.test(m[i + 1]); break;
      default: break;
    }
  }
  if (pm && h < 12) h += 12;
  return validDate(y, mo, d, h, mi, s, 0);
}

/**
 * Try every accepted date spelling: ISO (always) and the dateNF pattern.
 * Returns a Date or null; anything ambiguous stays text.
 */
function parseDate(text, opts) {
  if (typeof text !== "string" || text.length === 0 || text.length > 40) return null;
  const iso = parseISO(text);
  if (iso) return iso;
  const nf = opts.dateNF || DEFAULT_DATE_NF;
  return parseWithPattern(text, nf);
}

/* ------------------------------------------------------------------ writing */

/**
 * Everything a rendered field needs, resolved once per call.
 *
 * The caller cannot change an option mid-render, so re-deriving (and
 * re-validating) the quote character for every cell was a call per field on a
 * path that runs once per cell of the export. Building the plan also validates
 * the options up front, so an unusable quote is reported even when the sheet
 * turns out to have nothing to render.
 */
function renderPlan(opts) {
  return {
    delimiter: dsvDelimiter(opts),
    quote: quoteChar(opts),
    always: opts.alwaysQuote === true || opts.forceQuotes === true,
    quotedStrings: opts.quotedStrings === true,
    escapeFormulae: opts.escapeFormulae !== false,
  };
}

function quoteField(text, cell, opts, plan) {
  let need = plan.always;
  if (!need && text !== "") {
    need = text.includes(plan.delimiter) || text.includes(plan.quote) ||
      text.includes("\n") || text.includes("\r");
  }
  if (!need && plan.quotedStrings && isTextCell(cell) && wouldBeRetyped(text, opts)) need = true;
  if (!need) return text;
  const quote = plan.quote;
  return quote + text.split(quote).join(quote + quote) + quote;
}

/**
 * Characters that make a spreadsheet application treat a text field as a
 * formula. Covers OWASP's list: = + - @ TAB CR LF.
 *
 * The guard applies to text cells only: a numeric -2.5 is a number, not a
 * formula, and is written unchanged.
 *
 * A leading SPACE is deliberately excluded. Excel does not evaluate " =1+1" as
 * a formula - the space makes it text - so escaping it would add an apostrophe
 * to a very common and entirely harmless shape of data (" hello"), changing the
 * value to protect against nothing.
 */
const FORMULA_LEAD = /^[=+\-@\t\r\n]/;

function fieldFor(cell, opts, plan) {
  if (!cell) return "";
  let text = cellText(cell, opts);
  if (plan.escapeFormulae && isTextCell(cell) && FORMULA_LEAD.test(text)) text = "'" + text;
  return quoteField(text, cell, opts, plan);
}

function renderRow(cells, opts, plan) {
  const out = [];
  for (const cell of cells) out.push(fieldFor(cell, opts, plan));
  return out.join(plan.delimiter);
}

/**
 * Sheet -> CSV text. Rows are terminated, including the last one, which is
 * what RFC 4180 recommends and what Python's csv module produces.
 */
export function sheetToCSV(sheet, opts = {}) {
  // Styled CSV (HTML-flavoured text) is not a thing here: say so loudly rather
  // than hand back plain text the caller believes is styled.
  if (opts.style_text === true) {
    throw new NotImplementedError("styled CSV output is not supported; call sheetToCSV without style_text");
  }
  const ws = normalizeSheet(sheet, opts);
  const plan = renderPlan(opts);
  const eol = lineEndingOf(opts);
  const range = resolveRange(ws, opts);
  // Same bound as the JSON converters: a dense render must not be driven by a
  // row count that came out of an untrusted file.
  assertDenseRange(ws, range, opts);
  const lines = [];
  for (const line of csvLines(ws, range, opts, plan)) lines.push(line);
  const bom = opts.bom ? "\ufeff" : "";
  if (!lines.length) return bom;
  return bom + lines.join(eol) + eol;
}

/**
 * The rows of a sheet as CSV text lines, one at a time.
 *
 * Both sheetToCSV and createCSVStream walk this, so the string form and the
 * streaming form cannot disagree about quoting, formula escaping or blank-row
 * handling - the class of bug this codebase keeps finding in duplicated
 * converters.
 */
function* csvLines(ws, range, opts, plan) {
  const blankrows = opts.blankrows !== false;
  const skipHidden = opts.skipHidden === true;
  if (range.e.r < range.s.r) return;
  if (Array.isArray(opts.header)) {
    yield renderRow(opts.header.map((h) => cellFromValue(String(h), opts)), opts, plan);
  }
  for (let r = range.s.r; r <= range.e.r; r++) {
    if (skipHidden && isRowHidden(ws, r)) continue;
    const cells = [];
    let has = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (skipHidden && isColHidden(ws, c)) continue;
      const cell = cellAt(ws, r, c);
      if (cell && (cell.t !== "empty" && (cell.v !== undefined && cell.v !== null || cell.f !== undefined))) has = true;
      cells.push(cell);
    }
    if (!has && !blankrows) continue;
    yield renderRow(cells, opts, plan);
  }
}

/**
 * Sheet -> an async iterable of CSV text chunks, so a large sheet never becomes
 * one string.
 *
 * sheetToCSV is O(output) in retained memory: 184.2 MiB for a 200,000 x 10
 * sheet, because it builds every line and then joins them. This yields the same
 * bytes in chunks, holding one row at a time, so it can be piped to a file, a
 * socket or a Response without that cost.
 *
 * Chunks are emitted on row boundaries rather than byte boundaries: a quoted
 * field may legally contain a newline, and splitting there would force a
 * consumer to reassemble before it could parse.
 *
 * @param {object|Array} sheet a Sheet, or anything normalizeSheet accepts
 * @param {object} [opts] the sheetToCSV options, plus `bom`
 * @param {object} [streamOpts] `{ chunkRows?: number }` - rows per chunk. Default 512.
 * @returns {AsyncGenerator<string>}
 */
export async function* createCSVStream(sheet, opts = {}, streamOpts = {}) {
  if (opts.style_text === true) {
    throw new NotImplementedError("styled CSV output is not supported; call createCSVStream without style_text");
  }
  const ws = normalizeSheet(sheet, opts);
  const plan = renderPlan(opts);
  const eol = lineEndingOf(opts);
  const range = resolveRange(ws, opts);
  assertDenseRange(ws, range, opts);
  const perChunk = Math.max(1, Number(streamOpts.chunkRows) || 512);
  let buf = opts.bom ? "\ufeff" : "";
  let pending = 0;
  for (const line of csvLines(ws, range, opts, plan)) {
    buf += pending === 0 ? line : eol + line;
    pending++;
    if (pending >= perChunk) {
      yield buf + eol;
      buf = "";
      pending = 0;
    }
  }
  // A trailing terminator on the last row, matching sheetToCSV and RFC 4180.
  if (pending > 0) yield buf + eol;
  else if (buf.length) yield buf;
}

/**
 * The second argument of the explicit-separator forms.
 *
 * The documented signature is (value, delimiter, opts?), and the companion
 * option-object form has to keep working because every other function in this
 * module takes options there and callers already pass them. Accepting both is
 * one typeof check and removes a trap: as a plain alias, sheetToDSV(sheet, ";")
 * silently ignored the ";" and emitted commas.
 */
function dsvOptions(delimiter, opts) {
  if (typeof delimiter === "string") return { ...(opts || {}), delimiter };
  if (delimiter === undefined || delimiter === null) return opts || {};
  return delimiter;
}

/** Sheet -> DSV text with an explicit separator: sheetToDSV(sheet, ";", opts?). */
export function sheetToDSV(sheet, delimiter, opts) {
  return sheetToCSV(sheet, dsvOptions(delimiter, opts));
}

/* ------------------------------------------------------------------ parsing */

/**
 * RFC 4180 parser with the usual real-world tolerance:
 *   - \r\n, \n and lone \r all end a record
 *   - "" inside a quoted field is one quote
 *   - a quote that does not start a field is literal text
 *   - text after a closing quote is kept (lenient, never silent data loss)
 *   - an unterminated quote runs to end of input
 * Records are arrays of { text, quoted } so the caller can decide whether
 * quoting carries meaning.
 */
export function parseDSV(text, opts = {}) {
  const delimiter = dsvDelimiter(opts);
  const quote = quoteChar(opts);
  const records = [];
  let record = [];
  let field = "";
  let quoted = false;
  let closed = false;
  let inQuotes = false;
  let i = 0;
  const n = text.length;

  const endField = () => {
    record.push({ text: field, quoted });
    field = "";
    quoted = false;
    closed = false;
  };

  while (i < n) {
    const ch = text.charAt(i);
    if (inQuotes) {
      if (ch === quote) {
        if (text.charAt(i + 1) === quote) {
          field += quote;
          i += 2;
          continue;
        }
        inQuotes = false;
        closed = true;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === quote && field === "" && !closed) {
      inQuotes = true;
      quoted = true;
      i++;
      continue;
    }
    if (text.startsWith(delimiter, i)) {
      endField();
      i += delimiter.length;
      continue;
    }
    if (ch === "\r" || ch === "\n") {
      endField();
      records.push(record);
      record = [];
      i += ch === "\r" && text.charAt(i + 1) === "\n" ? 2 : 1;
      continue;
    }
    field += ch;
    i++;
  }
  // A trailing line terminator already closed the last record; only push a
  // pending one when it holds something.
  if (field.length > 0 || quoted || closed || record.length > 0) {
    endField();
    records.push(record);
  }
  return records;
}

function fieldToCell(f, opts) {
  const text = f.text;
  if (text === "") return { v: "", t: "string" };
  if (opts.raw === true) return { v: text, t: "string" };
  // A quoted field is literal text when the caller asks for it; by default the
  // content decides, which is what SheetJS and pandas do.
  if (opts.quotedStrings === true && f.quoted) return { v: text, t: "string" };
  if (BOOL_RE.test(text)) return { v: text.toUpperCase() === "TRUE", t: "boolean" };
  if (NUM_RE.test(text)) {
    const num = Number(text);
    if (Number.isFinite(num)) return { v: num, t: "number" };
  }
  const date = parseDate(text, opts);
  if (date) return dateCell(date, { dateNF: opts.dateNF || DEFAULT_DATE_NF, date1904: opts.date1904 });
  return { v: text, t: "string" };
}

/**
 * CSV text -> Sheet. Ragged rows are padded to the widest row with empty
 * strings, and empty fields stay empty strings (never null) - a CSV field
 * that exists is a value, even when it is blank.
 */
export function csvToSheet(text, opts = {}) {
  if (typeof text !== "string") {
    throw new GridxcelError("csvToSheet expects a string", "BAD_INPUT", { type: typeof text });
  }
  const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const records = parseDSV(src, opts);
  const sheet = blankSheet(opts.sheet || opts.name || DEFAULT_SHEET_NAME);
  const rows = sheet.rows;

  if (Array.isArray(opts.header)) {
    rows.push(opts.header.map((h) => {
      const cell = cellFromValue(String(h), opts);
      return cell || { v: "", t: "string" };
    }));
  }
  for (const rec of records) {
    if (opts.blankrows === false && rec.every((f) => f.text === "")) continue;
    rows.push(rec.map((f) => fieldToCell(f, opts)));
  }

  let width = 0;
  for (const row of rows) if (row.length > width) width = row.length;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < width; c++) {
      const cell = row[c];
      if (cell === undefined) row[c] = { v: "", t: "string", a: encodeCell(r, c) };
      else if (cell.a === undefined) cell.a = encodeCell(r, c);
    }
  }
  sheet.rowCount = rows.length;
  sheet.columnCount = width;
  return sheet;
}

/** DSV text -> Sheet with an explicit separator: dsvToSheet(text, ";", opts?). */
export function dsvToSheet(text, delimiter, opts) {
  return csvToSheet(text, dsvOptions(delimiter, opts));
}

/* -------------------------------------------------------------- book helpers */

function pickSheet(book, opts = {}) {
  if (book === null || book === undefined) return undefined;
  if (Array.isArray(book.sheets)) {
    const sel = opts.sheet === undefined ? opts.sheetName : opts.sheet;
    if (typeof sel === "number") return book.sheets[sel];
    if (typeof sel === "string") {
      const found = book.sheets.find((s) => s && s.name === sel);
      if (!found) throw new GridxcelError("no sheet named " + JSON.stringify(sel), "NO_SHEET");
      return found;
    }
    return book.sheets[0];
  }
  if (Array.isArray(book.SheetNames) && book.Sheets) {
    const name = opts.sheet === undefined ? book.SheetNames[0] : opts.sheet;
    return book.Sheets[name];
  }
  return book;
}

/** Book (or Sheet) -> CSV text. Uses the first sheet unless opts.sheet says. */
export function toCSV(book, opts = {}) {
  return sheetToCSV(pickSheet(book, opts), opts);
}

/** CSV text -> Book holding one sheet, ready for writeWorkbook. */
export function fromCSV(text, opts = {}) {
  // The canonical shape, not a copy of it: a Book literal written out here is
  // how the reader and api versions drifted apart in the first place (F8/F22).
  const book = emptyBook();
  book.sheets = [csvToSheet(text, opts)];
  if (opts.date1904) book.dateSystem = 1904;
  return book;
}
