/**
 * DIF: Data Interchange Format, the plain-text table format Lotus 1-2-3
 * popularised and Excel still exports.
 *
 * The grammar, per the 1981 DIF technical specification:
 *
 *   - A DIF file has a HEADER section and a DATA section.
 *   - A header chunk is THREE lines: an all-caps identifier, a "n,n" pair, and a
 *     quoted string. VALUE_TABLE is `TABLE / 0,1 / "EXCEL"`.
 *   - A data chunk is TWO lines: a "type,n" pair, then a keyword or a quoted
 *     string.
 *
 * The type in the pair decides what the second line is:
 *
 *   -1  directive. The second line is BOT (begin tuple = a new row) or EOD (end
 *       of data). Anything else is ignored.
 *    0  numeric. The value is the SECOND NUMBER OF THE PAIR, not the second
 *       line; the second line is a validity keyword: V (valid), NA, ERROR, TRUE
 *       or FALSE.
 *    1  string. The second line is the string in double quotes.
 *
 * Reading and writing live together so they cannot drift, the same rule the
 * other formats follow.
 */

import { GridxcelError } from "./errors.js";
import { blankSheet, cellFromValue } from "./json.js";
import { DEFAULT_SHEET_NAME } from "./constants.js";

/**
 * The largest cell a DIF file may address. The predecessor formats carry row and
 * column counts as plain integers, so an 85-byte file can claim a billion of
 * either; the reader refuses rather than allocating.
 */
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/** DIF is line-based; accept CRLF, LF or a lone CR. */
const LINE_END = /\r\n|\n|\r/;

/**
 * One data chunk: a "type,n" line and the line that follows it.
 *
 * Returns `{ directive }` for BOT/EOD, `{ cell }` for a value, or null when the
 * chunk carries nothing (an unterminated file, or an unknown directive).
 */
function difChunk(pairLine, nextLine) {
  const comma = pairLine.indexOf(",");
  if (comma < 0) return null;
  const type = Number(pairLine.slice(0, comma));
  const second = pairLine.slice(comma + 1).trim();
  const next = nextLine === undefined ? "" : nextLine.trim();

  if (type === -1) {
    if (next === "BOT") return { directive: "row" };
    if (next === "EOD") return { directive: "end" };
    return null;
  }
  if (type === 0) {
    // The number is in the pair; the second line says whether it is valid.
    const n = Number(second);
    if (next === "TRUE") return { cell: cellFromValue(true) };
    if (next === "FALSE") return { cell: cellFromValue(false) };
    if (next === "NA" || next === "ERROR") return { cell: { v: "#N/A", t: "error", e: "#N/A" } };
    return Number.isFinite(n) ? { cell: cellFromValue(n) } : null;
  }
  if (type === 1) {
    let s = next;
    if (s.length >= 2 && s.charCodeAt(0) === 34 && s.charCodeAt(s.length - 1) === 34) {
      s = s.slice(1, -1).replace(/""/g, '"');
    }
    // An empty STRING is a value, not an absent cell: a spreadsheet distinguishes
    // "" from an empty cell, and the writer emits 1,0 / "" to hold a column.
    // Returning null here turned every empty string into a hole on round trip.
    return { cell: cellFromValue(s) };
  }
  return null;
}

/**
 * DIF text -> Sheet.
 *
 * @param {string} text the file contents
 * @param {object} [opts] `{ name?: string, sheet?: string }`
 * @returns {object} Sheet
 */
export function difToSheet(text, opts = {}) {
  if (typeof text !== "string") {
    throw new GridxcelError("difToSheet expects a string", "BAD_INPUT", { type: typeof text });
  }
  const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const lines = src.split(LINE_END);
  if (lines.length < 2 || lines[0].trim().toUpperCase() !== "TABLE") {
    throw new GridxcelError("not a DIF file: the first line must be TABLE", "BAD_DIF", { first: (lines[0] || "").slice(0, 40) });
  }

  const sheet = blankSheet(opts.name || opts.sheet || DEFAULT_SHEET_NAME);
  let maxCol = -1;

  // Headers are THREE lines, data is TWO. Walk the header section first, then
  // switch to pair-walking at DATA. Treating header chunks as pairs is what
  // misaligned every row in the first attempt.
  let at = 0;
  for (; at < lines.length; at += 3) {
    const id = lines[at].trim().toUpperCase();
    if (id === "DATA") { at += 3; break; }
    // VECTORS declares columns, TUPLES declares rows. Both are hints, but a
    // caller sizes work from them, so an absurd claim is refused rather than
    // reported.
    if (id === "VECTORS" || id === "TUPLES") {
      const declared = Number((lines[at + 1] || "").split(",")[1]);
      const limit = id === "VECTORS" ? MAX_COLS : MAX_ROWS;
      if (Number.isFinite(declared) && declared > limit) {
        throw new GridxcelError(
          "DIF header declares " + declared + " " + (id === "VECTORS" ? "columns" : "rows") +
            ", beyond the supported extent of " + limit + "; refusing to allocate it",
          "RANGE_TOO_LARGE", { [id === "VECTORS" ? "columns" : "rows"]: declared });
      }
    }
  }
  let row = null;
  for (; at + 1 < lines.length; at += 2) {
    const pair = lines[at].trim();
    if (pair === "") { at -= 1; continue; }
    const chunk = difChunk(pair, lines[at + 1]);
    if (!chunk) continue;
    if (chunk.directive === "end") break;
    if (chunk.directive === "row") { row = []; sheet.rows.push(row); continue; }
    if (row === null) { row = []; sheet.rows.push(row); }
    if (!chunk.cell) { row.push(undefined); continue; }
    row.push(chunk.cell);
    if (row.length - 1 > maxCol) maxCol = row.length - 1;
  }
  // Trailing blank rows carry no information and would inflate rowCount.
  while (sheet.rows.length && sheet.rows[sheet.rows.length - 1].every((c) => c === undefined)) sheet.rows.pop();
  sheet.rowCount = sheet.rows.length;
  sheet.columnCount = maxCol + 1;
  return sheet;
}

/**
 * Sheet -> DIF text.
 *
 * @param {object} sheet a Sheet, or anything normalizeSheet accepts
 * @param {object} [opts] `{ eol?: string }`
 * @returns {string}
 */
/**
 * Format a DIF string value, refusing one this parser could not read back.
 *
 * A DIF chunk is two LINES. A value containing a newline would split into extra
 * lines, shift every following chunk, and be read back as garbage rather than
 * raising - so the writer refuses instead. DIF has no escape for it.
 */
function difString(value, row, col) {
  const text = value === undefined || value === null ? "" : String(value);
  if (text.indexOf("\n") >= 0 || text.indexOf("\r") >= 0) {
    throw new GridxcelError(
      "cannot write a DIF cell containing a line break (row " + row + ", column " + col +
        "); DIF chunks are line-based and the format has no escape for it - " +
        "write CSV instead, or strip the newline",
      "UNREPRESENTABLE", { row, col });
  }
  return '"' + text.replace(/"/g, '""') + '"';
}

export function sheetToDIF(sheet, opts = {}) {
  const eol = typeof opts.eol === "string" ? opts.eol : "\r\n";
  const rows = sheet.rows || [];
  let cols = 0;
  for (const row of rows) if (row.length > cols) cols = row.length;

  const out = [
    "TABLE", "0,1", '"EXCEL"',
    "VECTORS", "0," + cols, '""',
    "TUPLES", "0," + rows.length, '""',
    "DATA", "0,0", '""',
  ];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    // A row starts with the directive chunk -1,0 / BOT.
    out.push("-1,0", "BOT");
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell === undefined || cell === null || cell.t === "empty") {
        // DIF has no empty cell: a blank string keeps the column position, which
        // is what a spreadsheet exports for a hole.
        out.push("1,0", '""');
        continue;
      }
      if (cell.t === "boolean") { out.push("0,0", cell.v ? "TRUE" : "FALSE"); continue; }
      if (cell.t === "error") { out.push("0,0", "NA"); continue; }
      if (cell.t === "number") { out.push("0," + cell.v, "V"); continue; }
      if (cell.t === "date") {
        const serial = cell.serial !== undefined ? cell.serial : cell.v;
        out.push("0," + serial, "V");
        continue;
      }
      out.push("1,0", difString(cell.v, r + 1, c + 1));
    }
  }
  out.push("-1,0", "EOD");
  return out.join(eol) + eol;
}
