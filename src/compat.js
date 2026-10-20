/**
 * SheetJS-shaped compatibility surface: change one import and keep going.
 *
 *   import { readFile, utils } from "gridxcel/compat.js";
 *   const wb = readFile("report.xlsx");
 *   const rows = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
 *
 * Ownership: api-eng (docs/ARCHITECTURE.md). The reader and the writer are
 * imported STATICALLY: this entry point exists to be a drop-in for SheetJS, and
 * SheetJS is synchronous, so `readFile("x.xlsx")` has to work on the line after
 * the import. A dynamic import cannot give that without a top-level await, and a
 * top-level await here makes the library impossible to ship as an IIFE
 * (<script> tag) bundle - esbuild rejects it outright. There is no cycle:
 * read.js and write.js do not reach this module, and index.js imports both
 * statically already.
 *
 * The cost is that importing gridxcel/compat pulls the reader and the writer
 * (about 344 kB of src/) even if only utils.sheet_to_csv is used. That is the
 * right trade for a migration surface; a partial tree that lacks read.js or
 * write.js can no longer load this module at all, which is why available() is
 * now a settled report rather than a capability probe. preload() and available()
 * are kept for compatibility with published docs and earlier versions; both are
 * truthful no-ops now (see their JSDoc).
 *
 * Returned workbooks are gridxcel Books. They also carry SheetNames and Sheets
 * accessors, so the classic SheetJS dance above works unchanged; utils.* accept
 * either a gridxcel Sheet or a SheetJS worksheet ({!ref, A1: {...}}).
 *
 * write()/writeFile() accept a gridxcel Book, a SheetJS workbook
 * ({SheetNames, Sheets}), a bare worksheet, or an array of arrays.
 *
 * Deliberate differences from SheetJS, documented rather than silent:
 *   - the SheetJS-named entry points are synchronous and use the *Sync build
 *     of the reader/writer; readAsync/writeAsync/readFileAsync/writeFileAsync
 *     use the streaming implementations and accept paths and streams
 *   - write() with no opts.type returns a Uint8Array; SheetJS returns a binary
 *     string. type "array" is a Uint8Array, "buffer" a Node Buffer, "binary" a
 *     latin1 string, "base64" base64 text, "string" text for text formats
 *   - type "file" is not a thing here: call writeFile()
 *   - writeFile() overwrites by default (as SheetJS does); pass
 *     overwrite:false to be refused
 *   - bookType "xlsb", "xls", "ods", "html", "prn" and encrypted files throw
 *     NotImplementedError instead of producing something subtly wrong
 */

import { readWorkbook, readWorkbookSync } from "./read.js";
import { writeWorkbook, writeWorkbookSync } from "./write.js";
import { GridxcelError, NotImplementedError } from "./errors.js";
import { decodeCell, decodeRange, encodeCell, encodeCol, encodeRange, toRange } from "./address.js";
import {
  aoaToSheet,
  blankSheet,
  cellAt,
  cellText,
  jsonToSheet,
  normalizeSheet,
  sheetRange,
  sheetToAOA,
  sheetToJSON,
} from "./json.js";
import { csvToSheet, sheetToCSV, toCSV } from "./csv.js";
import { sheetToSYLK, sylkToSheet } from "./sylk.js";
import { difToSheet, sheetToDIF } from "./dif.js";
import {
  emptyBook, freeSheetName, validateSheetName } from "./api.js";
import { DEFAULT_SHEET_NAME } from "./constants.js";
import {
  asBuffer,
  decodeBase64,
  decodeLatin1,
  encodeBase64,
  encodeLatin1,
  readFileNodeSync,
  writeFileNodeSync,
} from "./platform.js";
import { SIG_LOCAL } from "./zip.js";

/* --------------------------------------------------------- peer modules */

/*
 * read.js and write.js are static imports (see the module header), so by the
 * time anything here runs both peers are loaded. What used to be a probe -
 * "can this build reach the reader?" - is now a fact, and the two functions
 * below report it. They stay exported because docs/API.md, docs/MIGRATION.md
 * and compat.d.ts publish them.

/**
 * Which peer modules this build can reach. With the peers imported statically
 * the answer is settled before this module finishes evaluating: read:true and
 * write:true, or this module would not have loaded. Kept exported, and kept
 * truthful, for the docs and callers that check it.
 */
export function available() {
  return {
    read: true,
    write: true,
    readError: null,
    writeError: null,
  };
}

/**
 * Resolve once the peers are usable. They already are: this resolves on the
 * next turn and returns available(). Kept so existing callers and the published
 * examples keep working; new code does not need it.
 */
export async function preload() {
  return available();
}

/* ----------------------------------------------------------- book plumbing */

/**
 * Add the SheetJS views to a Book. Getters, not stored copies, so the two
 * spellings can never drift apart.
 */
export function attachViews(book) {
  if (!book || typeof book !== "object") return book;
  if (!Object.getOwnPropertyDescriptor(book, "SheetNames")) {
    Object.defineProperty(book, "SheetNames", {
      enumerable: true,
      configurable: true,
      get() {
        return (book.sheets || []).map((s) => s.name);
      },
    });
  }
  if (!Object.getOwnPropertyDescriptor(book, "Sheets")) {
    Object.defineProperty(book, "Sheets", {
      enumerable: true,
      configurable: true,
      get() {
        const out = {};
        for (const s of book.sheets || []) out[s.name] = s;
        return out;
      },
    });
  }
  return book;
}

/** Accept every workbook-ish input SheetJS users hand to write(). */
export function normalizeBookInput(input) {
  if (input === null || input === undefined) return attachViews(emptyBook());
  if (Array.isArray(input.sheets)) return attachViews(input);
  if (Array.isArray(input.SheetNames) && input.Sheets) {
    const book = emptyBook();
    book.sheets = input.SheetNames.map((name, i) => {
      const sheet = normalizeSheet(input.Sheets[name]);
      return { ...sheet, name: name || sheet.name || "Sheet" + (i + 1) };
    });
    if (input.dateSystem === 1904 || (input.WBProps && input.WBProps.date1904)) book.dateSystem = 1904;
    if (input.Props) book.props = input.Props;
    if (input.Workbook && Array.isArray(input.Workbook.Names)) {
      book.definedNames = input.Workbook.Names.map((n) => ({ name: n.Name, ref: n.Ref, sheet: n.Sheet }));
    }
    return attachViews(book);
  }
  const sheet = normalizeSheet(input);
  const book = emptyBook();
  book.sheets = [sheet];
  return attachViews(book);
}

function bookTypeOf(opts) {
  const t = opts.bookType || opts.book || "xlsx";
  return String(t).toLowerCase();
}

function stripType(opts) {
  const copy = { ...opts };
  delete copy.type;
  delete copy.bookType;
  delete copy.book;
  return copy;
}

/** Options for the writer: it does not need to shape or save anything. */
function outputOptions(opts) {
  const copy = { ...opts };
  delete copy.type;
  delete copy.path;
  delete copy.overwrite;
  return copy;
}

/* -------------------------------------------------------------- read/write */

/**
 * "Is this the start of a zip?" - the same signature the container reader
 * recognises (zip.SIG_LOCAL), not a second copy of the magic bytes. Two bytes
 * are enough to tell a zip from CSV text, and the signature is compared
 * little-endian because that is how it sits in the file.
 */
function looksLikeZip(bytes) {
  return bytes.length > 1 && bytes[0] === (SIG_LOCAL & 0xff) && bytes[1] === ((SIG_LOCAL >>> 8) & 0xff);
}

function decodeText(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

/**
 * Sniff which text format this is.
 *
 * A SYLK file begins with the `ID` record; a DIF file begins with `TABLE`.
 * CSV has no signature, so it is the fallback, and that fallback is not free: a
 * CSV whose first field happens to be "ID" or "TABLE" would be misread. The sniff
 * therefore requires the full signature ("ID;", "TABLE\n"), not just the word,
 * and `{ bookType: "csv" }` overrides it when a file is genuinely ambiguous.
 *
 * Sniffing rather than requiring bookType matters for migration: SheetJS picks
 * the format from the content, and so does every caller who has been handing it
 * a path ending in .slk.
 */
function sniffTextFormat(text, opts) {
  if (opts.bookType === "csv" || opts.bookType === "txt") return "csv";
  if (opts.bookType !== undefined && opts.bookType !== null) return String(opts.bookType);
  // A BOM is not part of the format; comparing against it made BOM-prefixed SYLK
  // fall through to CSV. Both parsers strip it themselves.
  const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const head = src.length > 512 ? src.slice(0, 512) : src;
  // SYLK must begin with the ID record. A leading blank line is tolerated by some
  // writers, so skip leading whitespace before deciding.
  const fromFirstToken = head.replace(/^(?:[ \t]*[\r\n])+[ \t]*/, "");
  if (/^ID;/.test(fromFirstToken)) return "sylk";
  if (/^TABLE\r?\n/.test(fromFirstToken)) return "dif";
  return "csv";
}

function readTextBook(text, opts) {
  const book = emptyBook();
  if (opts.date1904) book.dateSystem = 1904;
  const format = sniffTextFormat(text, opts);
  if (format === "sylk") {
    book.sheets = [sylkToSheet(text, opts)];
  } else if (format === "dif") {
    book.sheets = [difToSheet(text, opts)];
  } else if (format === "csv" || format === "txt") {
    book.sheets = [csvToSheet(text, opts)];
  } else {
    throw new GridxcelError("cannot read " + JSON.stringify(format) + " from text; supported text formats are csv, txt, sylk and dif", "BAD_TYPE", { bookType: format });
  }
  book.sheetNames = book.sheets.map((s) => s.name);
  return attachViews(book);
}

/**
 * Bytes or text -> Book. A string, or bytes that are not a zip, is read as
 * CSV/DSV: SheetJS users hand CSV to read() constantly and a hard failure
 * there is the classic migration trap.
 */
function toBytes(data, opts) {
  const type = opts.type;
  if (data === null || data === undefined) throw new GridxcelError("read() needs data", "BAD_INPUT");
  if (typeof data === "string") {
    if (type === "base64") return decodeBase64(data);
    if (type === "binary") return decodeLatin1(data);
    if (type === "string" || type === undefined) {
      return data.startsWith("PK") ? decodeLatin1(data) : null;
    }
    throw new GridxcelError("read() cannot interpret a string as " + JSON.stringify(type), "BAD_TYPE");
  }
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (Array.isArray(data)) return Uint8Array.from(data);
  throw new GridxcelError("read() expects bytes, base64/binary text or CSV text", "BAD_INPUT", { type: typeof data });
}

/** Read a workbook. Synchronous, SheetJS-style; await also works. */
export function read(data, opts = {}) {
  const bytes = toBytes(data, opts);
  if (bytes === null) return readTextBook(String(data), opts);
  if (!looksLikeZip(bytes)) return readTextBook(decodeText(bytes), opts);
  return attachViews(readWorkbookSync(bytes, stripType(opts)));
}

/** Read a workbook through the async/streaming reader. */
export async function readAsync(data, opts = {}) {
  const bytes = toBytes(data, opts);
  if (bytes === null) return readTextBook(String(data), opts);
  if (!looksLikeZip(bytes)) return readTextBook(decodeText(bytes), opts);
  return attachViews(await readWorkbook(bytes, stripType(opts)));
}

/**
 * Read a workbook from a path. Synchronous; use readFileAsync for streams.
 * The file read is platform.readFileNodeSync - the same seam the writer uses,
 * which raises PlatformError [NOT_SUPPORTED] in a browser.
 */
export function readFile(path, opts = {}) {
  const bytes = readFileNodeSync(path);
  return read(bytes, { ...opts, type: "buffer" });
}

/** Read a workbook from a path through the async reader (paths, streams). */
export async function readFileAsync(path, opts = {}) {
  return attachViews(await readWorkbook(path, stripType(opts)));
}

function shapeBytes(bytes, opts) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const type = opts.type;
  if (type === undefined || type === "array") return u8;
  if (type === "buffer") return asBuffer(u8);
  if (type === "base64") return encodeBase64(u8);
  if (type === "binary" || type === "string") return encodeLatin1(u8);
  if (type === "file") throw new NotImplementedError("write type \"file\" is not supported; call writeFile(wb, path)");
  throw new GridxcelError("unknown write type " + JSON.stringify(type), "BAD_TYPE");
}

function shapeText(text, opts) {
  const type = opts.type;
  if (type === undefined || type === "string") return text;
  const bytes = new TextEncoder().encode(text);
  if (type === "buffer") return asBuffer(bytes);
  if (type === "array") return bytes;
  if (type === "base64") return encodeBase64(bytes);
  if (type === "binary") return encodeLatin1(bytes);
  throw new GridxcelError("unknown write type " + JSON.stringify(type), "BAD_TYPE");
}

/** Workbook -> xlsx bytes (or csv/txt/json text) shaped by opts.type. */
export function write(wb, opts = {}) {
  const type = bookTypeOf(opts);
  const book = normalizeBookInput(wb);
  if (type === "csv" || type === "txt") {
    const text = toCSV(book, type === "txt"
      ? { ...opts, delimiter: "\t", escapeFormulae: false, bom: opts.bom === undefined ? true : opts.bom }
      : opts);
    return shapeText(text, opts);
  }
  if (type === "sylk") {
    return shapeText(sheetToSYLK(book.sheets[0] || blankSheet(), opts), opts);
  }
  if (type === "dif") {
    return shapeText(sheetToDIF(book.sheets[0] || blankSheet(), opts), opts);
  }
  if (type === "json") {
    const sheet = book.sheets[0];
    const rows = sheet ? sheetToJSON(sheet, opts) : [];
    return shapeText(JSON.stringify(rows, null, opts.indent === undefined ? 2 : opts.indent), opts);
  }
  if (type === "xlsx" || type === "xlsm") {
    return shapeBytes(writeWorkbookSync(book, outputOptions(opts)), opts);
  }
  throw new NotImplementedError(
    "compat.write does not support bookType " + JSON.stringify(type) + " (xlsx, xlsm, csv, txt, json, sylk, dif)",
  );
}

/** Workbook -> xlsx bytes through the streaming writer (honours opts.path). */
export async function writeAsync(wb, opts = {}) {
  const type = bookTypeOf(opts);
  if (type !== "xlsx" && type !== "xlsm") return write(wb, opts);
  return writeWorkbook(normalizeBookInput(wb), opts);
}

/**
 * Workbook -> file. Overwrites unless overwrite:false, like SheetJS.
 *
 * Writing goes through platform.writeFileNodeSync - the library's one
 * synchronous fs call - with the "wx" flag, so the no-overwrite guard is atomic
 * instead of an existsSync race. This module has no node:fs import of its own.
 */
export function writeFile(wb, path, opts = {}) {
  if (typeof path !== "string" || path === "") throw new GridxcelError("writeFile needs a path", "BAD_INPUT");
  const flag = opts.overwrite === false ? "wx" : "w";
  const type = bookTypeOf(opts);
  const bytes = type === "csv" || type === "txt" || type === "json"
    ? new TextEncoder().encode(write(wb, { ...opts, type: "string" }))
    : write(wb, { ...opts, type: "array" });
  try {
    writeFileNodeSync(path, bytes, { flag });
  } catch (err) {
    if (err && err.code === "EEXIST") {
      throw new GridxcelError("refusing to overwrite " + path + ": pass overwrite: true to replace it", "FILE_EXISTS");
    }
    throw err;
  }
  return undefined;
}

/** Workbook -> file through the streaming writer. */
export async function writeFileAsync(wb, path, opts = {}) {
  const type = bookTypeOf(opts);
  if (type !== "xlsx" && type !== "xlsm") return writeFile(wb, path, opts);
  // writeWorkbook owns the path handling (lazy node:fs/promises) and the
  // overwrite guard. SheetJS's writeFile overwrites by default.
  await writeWorkbook(normalizeBookInput(wb), {
    ...opts,
    path,
    type: undefined,
    overwrite: opts.overwrite !== false,
  });
  return undefined;
}

/* ------------------------------------------------------------ SheetJS views */

function sheetJsType(cell) {
  if (cell.t === "string") return "s";
  if (cell.t === "number") return "n";
  if (cell.t === "boolean") return "b";
  if (cell.t === "date") return cell.v instanceof Date ? "d" : "n";
  if (cell.t === "error") return "e";
  return "s";
}

/** gridxcel Sheet -> SheetJS worksheet ({!ref, A1: {v,t,z,w,f}}). */
export function toWorksheet(sheet) {
  const ws = normalizeSheet(sheet);
  const range = sheetRange(ws);
  const out = {};
  if (!range.empty) {
    out["!ref"] = encodeRange(range.s, range.e);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = cellAt(ws, r, c);
        if (!cell || cell.t === "empty") continue;
        const copy = { v: cell.v, t: sheetJsType(cell) };
        if (cell.z !== undefined) copy.z = cell.z;
        if (cell.w !== undefined) copy.w = cell.w;
        if (cell.f !== undefined) copy.f = cell.f;
        if (cell.c !== undefined) copy.c = cell.c;
        out[encodeCell(r, c)] = copy;
      }
    }
  }
  if (Array.isArray(ws.merges) && ws.merges.length) {
    out["!merges"] = ws.merges.map((ref) => {
      const parsed = toRange(ref);
      return { s: parsed.s, e: parsed.e };
    });
  }
  return out;
}

/** gridxcel Book -> SheetJS workbook ({SheetNames, Sheets, Props}). */
export function toSheetJSBook(book) {
  const source = Array.isArray(book.sheets) ? book : normalizeBookInput(book);
  const Sheets = {};
  const SheetNames = [];
  for (const s of source.sheets) {
    Sheets[s.name] = toWorksheet(s);
    SheetNames.push(s.name);
  }
  return { SheetNames, Sheets, Props: source.props || {} };
}

/** Existing sheet names, for either workbook shape. */
function sheetNamesOf(wb) {
  if (Array.isArray(wb.sheets)) return wb.sheets.map((s) => s.name);
  return Array.isArray(wb.SheetNames) ? wb.SheetNames.slice() : [];
}

function resolveSheet(wb, which) {
  if (typeof which === "number") {
    if (Array.isArray(wb.sheets)) return wb.sheets[which];
    const name = Array.isArray(wb.SheetNames) ? wb.SheetNames[which] : undefined;
    return name === undefined ? undefined : wb.Sheets[name];
  }
  if (typeof which === "string") {
    if (Array.isArray(wb.sheets)) return wb.sheets.find((s) => s && s.name === which);
    return wb.Sheets ? wb.Sheets[which] : undefined;
  }
  if (Array.isArray(wb.sheets)) return wb.sheets[0];
  const name = Array.isArray(wb.SheetNames) ? wb.SheetNames[0] : undefined;
  return name === undefined ? undefined : wb.Sheets[name];
}

/* ------------------------------------------------------------------- utils */

/**
 * The SheetJS-shaped helper namespace: synchronous, SheetJS-shaped, and the
 * surface a migration should use.
 *
 * Its native twin (`utils` from the package root) delegates every member to
 * the same canonical function, so a result cannot depend on which import
 * produced it. Two members differ on purpose, and test/api-eng-dedup.test.mjs
 * asserts both differences so they cannot drift again:
 *   - `book_new()` here returns the SheetJS shape (SheetNames/Sheets views over
 *     a native book); the root one returns the plain native Workbook.
 *   - every converter here returns a value synchronously, as SheetJS does; the
 *     root namespace returns Promises, which docs/API.md documents.
 */
export const utils = {
  /** Array of arrays -> Sheet. */
  aoa_to_sheet: (rows, opts) => aoaToSheet(rows, opts),
  /** Array of objects -> Sheet, keys in first-seen order. */
  json_to_sheet: (rows, opts) => jsonToSheet(rows, opts),
  /** Sheet -> array of objects (SheetJS option names honoured). */
  sheet_to_json: (sheet, opts) => sheetToJSON(sheet, opts),
  /** Sheet -> CSV text. */
  sheet_to_csv: (sheet, opts) => sheetToCSV(sheet, opts),
  /** Sheet -> tab text with a BOM, as SheetJS emits for .txt. */
  sheet_to_txt: (sheet, opts) => "\ufeff" + sheetToCSV(sheet, { ...opts, delimiter: "\t", escapeFormulae: false }),
  /** Sheet -> dense array of arrays (gridxcel extension). */
  sheet_to_aoa: (sheet, opts) => sheetToAOA(sheet, opts),
  /**
   * A workbook with no sheets, carrying the SheetJS SheetNames/Sheets views.
   * Deliberately different from the root `utils.book_new`, which returns the
   * plain native Workbook (see the namespace comment above).
   */
  book_new: () => attachViews(emptyBook()),
  /**
   * Append a sheet and return its name, as SheetJS does: the first free
   * "SheetN" when no name is given, the canonical sheet-name rules (api.js
   * validateSheetName) applied to the rest, and a duplicate name rejected here
   * rather than at write time.
   */
  book_append_sheet: (wb, ws, name) => {
    const taken = sheetNamesOf(wb);
    const explicit = name === undefined || name === null || name === "" ? undefined : String(name);
    const resolved = validateSheetName(explicit === undefined ? freeSheetName(taken) : explicit, taken);
    const sheet = normalizeSheet(ws);
    sheet.name = resolved;
    if (Array.isArray(wb.sheets)) wb.sheets.push(sheet);
    else {
      if (!Array.isArray(wb.SheetNames)) wb.SheetNames = [];
      if (!wb.Sheets) wb.Sheets = {};
      wb.SheetNames.push(resolved);
      wb.Sheets[resolved] = sheet;
    }
    return resolved;
  },
  /** SheetJS visibility: 0 visible, 1 hidden, 2 very hidden. */
  book_set_sheet_visibility: (wb, which, visibility) => {
    const sheet = resolveSheet(wb, which);
    if (!sheet) throw new GridxcelError("no such sheet: " + JSON.stringify(which), "NO_SHEET");
    const vis = Number(visibility) || 0;
    sheet.hidden = vis > 0;
    sheet.visibility = vis;
    return wb;
  },
  /** "A1:B2" or {s,e} -> {s:{r,c}, e:{r,c}}. */
  decode_range: (ref) => {
    if (ref === null || ref === undefined) throw new GridxcelError("decode_range needs a range", "BAD_RANGE");
    if (typeof ref === "number") return toRange(ref);
    return typeof ref === "string" ? decodeRange(ref) : ref;
  },
  /** {r,c} -> "A1" (a two-argument (r, c) form is accepted too). */
  encode_cell: (cell, c) => (c === undefined ? encodeCell(cell.r, cell.c) : encodeCell(cell, c)),
  /** {s,e} -> "A1:B2"; also accepts (s, e) and a range string. */
  encode_range: (a, b) => {
    if (b) return encodeRange(a, b);
    if (typeof a === "string") {
      const parsed = decodeRange(a);
      return encodeRange(parsed.s, parsed.e);
    }
    return encodeRange(a.s, a.e);
  },
  /** "A1" -> {r,c}. */
  decode_cell: (ref) => decodeCell(ref),
  /** 0 -> "A". */
  encode_col: (c) => encodeCol(c),
  /**
   * The display text of a cell (Cell.w wins, exactly like SheetJS). The root
   * `utils.format_cell` is the same function body over the same canonical
   * json.cellText, so both surfaces answer identically.
   */
  format_cell: (cell, v, opts) => {
    if (!cell || cell.t === "z" || cell.t === "empty") return "";
    if (cell.w !== undefined) return cell.w;
    const value = v === undefined || v === null ? cell.v : v;
    return cellText({ ...cell, v: value, w: undefined }, { ...(opts || {}), rawNumbers: true });
  },
};

/** const XLSX = XLSX_like; keeps existing call sites compiling unchanged. */
export const XLSX_like = { read, readFile, write, writeFile, utils, version: "0.1.0" };

export default XLSX_like;
