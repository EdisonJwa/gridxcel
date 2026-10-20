/**
 * gridxcel - a zero-dependency, streaming-first XLSX reader and writer.
 *
 *   import { read, write, eachRow, sheetToJSON } from "gridxcel";
 *
 *   const book = await read("report.xlsx", { cellDates: true });
 *   await write(book, { path: "copy.xlsx" });
 *   for await (const { sheet, r, cells } of eachRow("huge.xlsx")) { ... }
 *
 * Every module is imported statically. A lazy loader used to sit in front of
 * ./read.js, but the file was imported statically here anyway, so the loader
 * could only ever succeed - the MODULE_UNAVAILABLE path it promised was
 * unreachable, and node:module kept the whole entry point out of browser
 * bundles. Static imports are what make the browser IIFE build possible.
 */
import {
  freeSheetName,
  readBook, writeBook, normalizeSheet, normalizeBook, normalizeCell,
  mergeBooks, renameSheet, countCells, validateSheetName, VERSION,
} from "./api.js";
import { DEFAULT_SHEET_NAME } from "./constants.js";
import {
  aoaToSheet as aoaToSheetImpl,
  jsonToSheet as jsonToSheetImpl,
  sheetToAOA as sheetToAOAImpl,
  sheetToJSON as sheetToJSONImpl,
} from "./json.js";
import { csvToSheet as csvToSheetImpl, sheetToCSV as sheetToCSVImpl } from "./csv.js";
// One home for the display-text rule (MINOR-3): format_cell is a SheetJS method,
// so gridxcel/compat owns its semantics and this surface calls that function
// rather than keeping a second copy of the idea - the two copies disagreed once
// already (DUP-2). The dependency points this way on purpose: the reverse would
// make the migration entry point import the zip layer, which is what compat.js
// promises it does not do. The end state is json.js exporting it next to
// cellText, at which point this import disappears.
import { utils as compatUtils } from "./compat.js";
import { GridxcelError, BadZipError, FormatError, SchemaError, NotImplementedError, StrictWarningError, PlatformError } from "./errors.js";
import { encodeCol, decodeCol, encodeCell, decodeCell, encodeRange, decodeRange, toRange, splitSheetRef } from "./address.js";
import { isDateFormat, formatNumber, formatValue, serialToDate, dateToSerial, generalNumber, BUILTIN_FORMATS } from "./numfmt.js";
import { XmlParser, parseTag, encodeXmlText, encodeXmlAttr, decodeEntities } from "./xml.js";
import { ZipArchive, ZipWriter, createZip, ZipError } from "./zip.js";
import { crc32, Crc32 } from "./crc32.js";
import { WorkbookWriter, SheetWriter, buildParts, createBookWriter, numberLiteral } from "./workbook-writer.js";
import { StyleTable, parseStyles, normalizeColor } from "./style-registry.js";
import { readWorkbookSync, eachRow as eachRowRaw, cellFormatCode } from "./read.js";
import { writeWorkbook, writeWorkbookSync, toBuffer } from "./write.js";

export { GridxcelError, BadZipError, FormatError, SchemaError, NotImplementedError, StrictWarningError, PlatformError };
export { encodeCol, decodeCol, encodeCell, decodeCell, encodeRange, decodeRange, toRange, splitSheetRef };
export { isDateFormat, formatNumber, formatValue, serialToDate, dateToSerial, generalNumber, BUILTIN_FORMATS };
export { XmlParser, parseTag, encodeXmlText, encodeXmlAttr, decodeEntities };
export { ZipArchive, ZipWriter, createZip, ZipError, crc32, Crc32 };
export { VERSION, normalizeSheet, normalizeBook, normalizeCell, mergeBooks, renameSheet, countCells };

/* Streaming and low-level surfaces that the documentation promises. */
export { WorkbookWriter, SheetWriter, buildParts, createBookWriter, numberLiteral };
export { StyleTable, parseStyles, normalizeColor };
export { writeWorkbook, writeWorkbookSync, toBuffer, cellFormatCode };

/**
 * Read a workbook from bytes, a path, a stream, or an async iterable.
 *
 * Non-fatal problems are reported two ways: an explicit options.onWarning
 * callback, and a `warnings` array attached to the returned Book. The array is
 * always present (empty when the file was clean) so callers can inspect
 * problems without opting into a callback. The same array is passed to
 * onWarning, which keeps the two views from diverging.
 */
export function read(input, options) {
  const collected = [];
  const userSink = options && options.onWarning;
  const merged = userSink
    ? { ...options, onWarning: (w) => { collected.push(w); userSink(w); } }
    : { ...(options || {}), onWarning: (w) => { collected.push(w); } };
  return readBook(input, merged).then((book) => {
    if (book && typeof book === "object" && !Array.isArray(book.warnings)) book.warnings = collected;
    return book;
  });
}

/**
 * Read synchronously from bytes. Works wherever the synchronous deflate
 * backend does (Node; see platform.js), not just where require-of-ESM does.
 * Mirrors read(): non-fatal problems end up on the returned Book's warnings
 * array as well as in options.onWarning.
 * @param {Uint8Array|ArrayBuffer} bytes
 */
export function readSync(bytes, options) {
  const collected = [];
  const userSink = options && options.onWarning;
  const merged = userSink
    ? { ...options, onWarning: (w) => { collected.push(w); userSink(w); } }
    : { ...(options || {}), onWarning: (w) => { collected.push(w); } };
  const book = readWorkbookSync(bytes, merged);
  if (book && typeof book === "object" && !Array.isArray(book.warnings)) book.warnings = collected;
  return book;
}

/** Write a workbook to bytes or, with options.path, to disk. */
export const write = writeBook;

/** Write a workbook to a file path. */
export function writeFile(book, path, options = {}) {
  return writeBook(book, { ...options, path });
}

/** Read a workbook from a file path. */
export function readFile(path, options = {}) {
  return readBook(path, options);
}

/**
 * Stream rows out of a workbook without materialising a sheet.
 * Yields { sheet, r, cells } with cells sparse and indexed from column 0.
 */
export async function* eachRow(input, options = {}) {
  yield* eachRowRaw(input, options);
}

/**
 * Read only the sheet names, without parsing any worksheet.
 * The reader keeps this metadata on Book.sheetNames when sheets:false is used,
 * so we prefer it and fall back to the parsed sheets.
 */
export async function sheetNames(input) {
  const book = await readBook(input, { sheets: false });
  if (Array.isArray(book.sheetNames) && book.sheetNames.length) return book.sheetNames.slice();
  return book.sheets.map((s) => s.name);
}

/** Read the first sheet as an array of JSON objects. */
export async function readJSON(input, options = {}) {
  const book = await readBook(input, options);
  const sheet = book.sheets[0];
  return sheet ? sheetToJSONImpl(sheet, options) : [];
}

/** Write raw rows straight to xlsx bytes. */
export function writeRows(rows, options = {}) {
  return writeBook(normalizeBook(rows, options), options);
}

/**
 * Convert an array of arrays into a Sheet. This is the canonical converter, so
 * the native entry point, `utils.aoa_to_sheet` and `gridxcel/compat` produce
 * the same cells - A1 addresses and date serials included.
 */
export function aoaToSheet(rows, options) {
  return aoaToSheetImpl(rows, options);
}

/**
 * The converters below stay promise-based on this entry point, which is what
 * docs/API.md documents. The implementation is the same synchronous function
 * that `gridxcel/compat` exposes under the SheetJS name, so which import a
 * caller uses cannot change the result.
 */
export async function sheetToAOA(sheet, options) {
  return sheetToAOAImpl(sheet, options);
}

/** Convert a Sheet into JSON objects. */
export async function sheetToJSON(sheet, options) {
  return sheetToJSONImpl(sheet, options);
}

/** Convert JSON objects into a Sheet. */
export async function jsonToSheet(rows, options) {
  return jsonToSheetImpl(rows, options);
}

/** Convert a Sheet into CSV text. */
export async function sheetToCSV(sheet, options) {
  return sheetToCSVImpl(sheet, options);
}

/** Parse CSV text into a Sheet. */
export async function csvToSheet(text, options) {
  return csvToSheetImpl(text, options);
}

/**
 * The SheetJS-shaped helper namespace in its native flavour.
 *
 * `gridxcel/compat` exports the same method names and is the surface a
 * migration should use: synchronous and SheetJS-shaped throughout. This one
 * keeps the promise-based converters the root API documents, and exactly two
 * members differ on purpose - both pinned by test/api-eng-dedup.test.mjs so
 * they cannot drift again:
 *   - `book_new()` returns the native Workbook (`{ sheets }`); the compat book
 *     also carries the SheetJS `SheetNames`/`Sheets` views.
 *   - the converters return Promises here, values there.
 * Every other member delegates to the same canonical function compat calls.
 */
export const utils = {
  aoa_to_sheet: (rows, opts) => aoaToSheetImpl(rows, opts),
  json_to_sheet: async (rows, opts) => jsonToSheetImpl(rows, opts),
  sheet_to_json: async (sheet, opts) => sheetToJSONImpl(sheet, opts),
  sheet_to_csv: async (sheet, opts) => sheetToCSVImpl(sheet, opts),
  sheet_to_aoa: async (sheet, opts) => sheetToAOAImpl(sheet, opts),
  book_new: () => normalizeBook(null),
  sheet_add_aoa: (sheet, rows, opts) => {
    const origin = (opts && opts.origin) || { r: 0, c: 0 };
    rows.forEach((row, ri) => {
      const target = (sheet.rows[origin.r + ri] = sheet.rows[origin.r + ri] || []);
      row.forEach((v, ci) => { target[origin.c + ci] = normalizeCell(v); });
      if (target.length > sheet.columnCount) sheet.columnCount = target.length;
    });
    sheet.rowCount = sheet.rows.length;
    return sheet;
  },
  book_append_sheet: (book, sheet, name) => {
    // An empty or missing name means "pick one for me", which is what SheetJS's
    // falsy-name rule does; compat has the same rule so the two never disagree.
    const taken = book.sheets.map((s) => s.name);
    const explicit = name === undefined || name === null || name === "" ? undefined : String(name);
    const resolved = validateSheetName(explicit === undefined ? freeSheetName(taken) : explicit, taken);
    const added = normalizeSheet(sheet, resolved);
    added.name = resolved;
    book.sheets.push(added);
    return book;
  },
  decode_range: decodeRange,
  encode_cell: (cell, c) => (c === undefined ? encodeCell(cell.r, cell.c) : encodeCell(cell, c)),
  encode_range: (a, b) => {
    if (b) return encodeRange(a, b);
    if (typeof a === "string") {
      const range = decodeRange(a);
      return encodeRange(range.s, range.e);
    }
    return encodeRange(a.s, a.e);
  },
  format_cell: (cell, v, opts) => compatUtils.format_cell(cell, v, opts),
};

export default {
  VERSION, read, readSync, write, writeFile, readFile, eachRow, sheetNames, readJSON, writeRows,
  aoaToSheet, sheetToAOA, sheetToJSON, jsonToSheet, sheetToCSV, csvToSheet, utils,
  normalizeSheet, normalizeBook, normalizeCell, mergeBooks, renameSheet, countCells,
  GridxcelError, BadZipError, ZipError, FormatError, SchemaError, NotImplementedError, StrictWarningError, PlatformError,
};
