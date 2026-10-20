/**
 * SYLK: the "Symbolic Link" format Microsoft Excel used before XLSX.
 *
 * A SYLK file is a sequence of records, each one a semicolon-separated list of
 * fields ending in a record terminator. The first field is a record id; the rest
 * are letter-prefixed values:
 *
 *   ID;P                     identification, P = production
 *   B;Y3;X4                  bounds: rows and columns
 *   C;Y1;X2;K42              cell at (row 1, column 2), K = numeric value
 *   E                        end
 *
 * Coordinates are 1-based and, unusually, `Y` comes before `X`. A record that
 * omits Y or X addresses the same row or column as the previous one, which is
 * how real files keep their size down.
 *
 * Only what a spreadsheet actually uses is read here: cells with values (K), and
 * the text and formula fields (E, F) that go with them. Style records are parsed
 * far enough to skip them safely and are not applied, because a style table that
 * silently half-applies is worse than one that is honestly not read.
 *
 * Reading and writing are both here so the two cannot drift, the same rule the
 * other formats follow.
 */

import { GridxcelError } from "./errors.js";
import { blankSheet, cellFromValue, dateCell } from "./json.js";
import { DEFAULT_SHEET_NAME } from "./constants.js";

/** Record terminators. SYLK is line-based; Excel writes CRLF. */
const RECORD_END = /\r\n|\n|\r/;

/**
 * Split one record into its fields, honouring a quoted string that contains a
 * semicolon. A quoted value runs to the closing quote; `""` inside it is a
 * literal quote, which is the same convention as CSV.
 */
function splitRecord(record) {
  const out = [];
  let i = 0;
  const n = record.length;
  while (i <= n) {
    const start = i;
    let field = "";
    // A field's first character is its id letter.
    while (i < n && record.charCodeAt(i) !== 59) {
      const c = record.charCodeAt(i);
      if (c === 34) {
        // Quoted section: copy until the matching quote, doubling to escape.
        i++;
        for (;;) {
          if (i >= n) throw new GridxcelError("unterminated quoted field in SYLK record " + JSON.stringify(record), "BAD_SYMBOL");
          const q = record.charCodeAt(i);
          if (q === 34) {
            if (record.charCodeAt(i + 1) === 34) { field += '"'; i += 2; continue; }
            i++;
            break;
          }
          field += record[i];
          i++;
        }
        continue;
      }
      field += record[i];
      i++;
    }
    out.push(field);
    if (i >= n) break;
    i++; // skip the ';'
    if (i > start && i >= n) { out.push(""); break; }
  }
  return out;
}

/** Decode a `K` field: a number, a quoted string, or one of the logic words. */
function cellFromK(raw) {
  if (raw === undefined || raw === "") return undefined;
  const trimmed = raw.trim();
  if (trimmed === "TRUE") return cellFromValue(true);
  if (trimmed === "FALSE") return cellFromValue(false);
  if (trimmed === "NA" || trimmed === "#N/A") return { v: "#N/A", t: "error", e: "#N/A" };
  // A leading zero would be a string ("007"); anything that round-trips is numeric.
  const asNumber = Number(trimmed);
  if (trimmed !== "" && Number.isFinite(asNumber) && !(trimmed.length > 1 && trimmed.charCodeAt(0) === 48 && trimmed.indexOf(".") < 0)) {
    return cellFromValue(asNumber);
  }
  return cellFromValue(raw);
}

/**
 * SYLK text -> Sheet.
 *
 * @param {string} text the file contents
 * @param {object} [opts] `{ name?: string }`
 * @returns {object} Sheet
 */
export function sylkToSheet(text, opts = {}) {
  if (typeof text !== "string") {
    throw new GridxcelError("sylkToSheet expects a string", "BAD_INPUT", { type: typeof text });
  }
  const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const records = src.split(RECORD_END).filter((r) => r.length > 0);

  if (!records.length || splitRecord(records[0])[0].toUpperCase() !== "ID") {
    throw new GridxcelError("not a SYLK file: the first record must be ID", "BAD_SYMBOL", { first: records[0] === undefined ? "" : records[0].slice(0, 40) });
  }

  const sheet = blankSheet(opts.name || opts.sheet || DEFAULT_SHEET_NAME);
  /** Sparse rows keyed by 1-based row number, each a sparse array by 0-based column. */
  const byRow = new Map();
  let lastY = 0;
  let lastX = 0;
  let rowCount = 0;
  let maxCol = -1;

  for (const record of records) {
    const fields = splitRecord(record);
    const id = fields[0].toUpperCase();
    if (id === "C") {
      let y = lastY;
      let x = lastX;
      let k;
      for (let i = 1; i < fields.length; i++) {
        const f = fields[i];
        if (f.length === 0) continue;
        const tag = f[0].toUpperCase();
        const value = f.slice(1);
        if (tag === "Y") y = parseInt(value, 10);
        else if (tag === "X") x = parseInt(value, 10);
        else if (tag === "K") k = value;
      }
      if (!(y > 0) || !(x > 0)) {
        throw new GridxcelError("SYLK cell record has no usable coordinates " + JSON.stringify(record), "BAD_SYMBOL");
      }
      lastY = y;
      lastX = x;
      const cell = cellFromK(k);
      if (cell === undefined) continue;
      let row = byRow.get(y);
      if (row === undefined) { row = []; byRow.set(y, row); }
      row[x - 1] = cell;
      if (y > rowCount) rowCount = y;
      if (x - 1 > maxCol) maxCol = x - 1;
    } else if (id === "B") {
      // Bounds only: a hint, not a promise. The cells decide what is present.
      for (let i = 1; i < fields.length; i++) {
        const f = fields[i];
        if (!f) continue;
        if (f[0].toUpperCase() === "Y") rowCount = Math.max(rowCount, parseInt(f.slice(1), 10) || 0);
      }
    }
  }

  // Materialise: rows must be contiguous from originRow, so fill the gaps.
  for (let r = 1; r <= rowCount; r++) sheet.rows.push(byRow.get(r) || []);
  sheet.rowCount = rowCount;
  sheet.columnCount = maxCol + 1;
  return sheet;
}

/**
 * Quote a field for SYLK if it contains a separator, a quote or a line break.
 */
function q(value) {
  const s = String(value);
  if (s.indexOf(";") < 0 && s.indexOf('"') < 0 && s.indexOf("\r") < 0 && s.indexOf("\n") < 0) return s;
  return '"' + s.replace(/"/g, '""') + '"';
}

/**
 * Sheet -> SYLK text.
 *
 * @param {object} sheet a Sheet, or anything normalizeSheet accepts
 * @param {object} [opts] `{ eol?: string }`
 * @returns {string}
 */
export function sheetToSYLK(sheet, opts = {}) {
  const eol = typeof opts.eol === "string" ? opts.eol : "\r\n";
  const rows = sheet.rows || [];
  const out = ["ID;P"];
  let maxCol = 0;
  for (const row of rows) {
    if (row.length > maxCol) maxCol = row.length;
  }
  if (rows.length && maxCol) out.push("B;Y" + rows.length + ";X" + maxCol);

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell === undefined || cell === null || cell.t === "empty") continue;
      const y = r + 1;
      const x = c + 1;
      let k;
      if (cell.t === "boolean") k = cell.v ? "TRUE" : "FALSE";
      else if (cell.t === "error") k = q(cell.e || cell.v || "#N/A");
      else if (cell.t === "date") {
        // SYLK has no date type: write the serial, which is what the cell holds.
        k = String(cell.serial !== undefined ? cell.serial : cell.v);
      } else if (cell.t === "number") k = String(cell.v);
      else k = q(cell.v === undefined || cell.v === null ? "" : cell.v);
      if (k === "") continue;
      out.push("C;Y" + y + ";X" + x + ";K" + k);
    }
  }
  out.push("E");
  return out.join(eol) + eol;
}
