/**
 * The public surface: normalisation of caller input, dispatch to the reader or
 * writer, and the small conveniences that make the library pleasant to use.
 *
 * Peer modules are loaded lazily and cached. That keeps two useful properties:
 * the CSV/JSON helpers are usable without ever touching the zip layer, and a
 * partially built tree still loads instead of exploding at import time.
 */

import { GridxcelError } from "./errors.js";
import { decodeRange, encodeRange, encodeCell, decodeCell } from "./address.js";
import { MAX_SHEET_NAME_LENGTH, DEFAULT_SHEET_NAME } from "./constants.js";
// Statically imported, not lazily: json.js owns the canonical value -> Cell
// coercion (cellFromValue) and normalising a value should not need an await.
// json.js touches neither the zip layer nor the filesystem, so the laziness
// that keeps CSV/JSON usable in a partial tree is unaffected.
import { cellFromValue } from "./json.js";
// Path writes go through the platform seam: a literal import("node:fs/promises")
// is resolved by the bundler at build time and would fail the browser bundle even
// though only the CSV/JSON output paths reach it.
import { writeFileNode } from "./platform.js";

export const VERSION = "0.1.0";

const lazy = new Map();

/**
 * Peer modules, as literal thunks.
 *
 * The specifier has to be visible to the bundler. A dynamic import() with a
 * variable argument cannot be rewritten, so it survives into the browser bundle
 * as a runtime import with no module loader behind it - verified in the Web-only
 * sandbox as "A dynamic import callback was not specified". A table of literal
 * thunks keeps the laziness (the module is evaluated on first use, not at import
 * time) and lets a bundler inline each peer.
 */
const LAZY_PEERS = {
  "./csv.js": () => import("./csv.js"),
  "./json.js": () => import("./json.js"),
  "./read.js": () => import("./read.js"),
  "./write.js": () => import("./write.js"),
};

async function load(spec, name) {
  const key = spec + "#" + name;
  if (lazy.has(key)) return lazy.get(key);
  const peer = LAZY_PEERS[spec];
  if (peer === undefined) {
    throw new GridxcelError("internal module " + spec + " is not a known peer module", "MODULE_UNAVAILABLE", { spec, name });
  }
  let mod;
  try {
    mod = await peer();
  } catch (cause) {
    throw new GridxcelError(
      "internal module " + spec + " is not available yet (missing export " + name + ")",
      "MODULE_UNAVAILABLE",
      { spec, name, cause: cause && cause.message },
    );
  }
  if (typeof mod[name] !== "function") {
    throw new GridxcelError(spec + " does not export " + name + "()", "MODULE_INCOMPLETE", { spec, name });
  }
  lazy.set(key, mod[name]);
  return mod[name];
}

/* ------------------------------------------------------------ normalisation */

/**
 * Coerce a cell input into the internal Cell shape.
 *
 * json.js owns this rule (cellFromValue) and this delegates to it, so the two
 * entry points cannot disagree about what a cell is - which is exactly what
 * DUP-1 was. The agreed handling of a plain object or array is documented
 * there: it throws BAD_CELL rather than inventing a stringified value.
 */
export function normalizeCell(input) {
  return cellFromValue(input);
}

/** Coerce a raw 2-D array (or array of objects) into a Sheet. */
export function normalizeSheet(input, name = DEFAULT_SHEET_NAME) {
  if (input === null || input === undefined) {
    return { name, rows: [], originRow: 1, originCol: 0, columnCount: 0, rowCount: 0, merges: [], columns: [], rowInfo: new Map() };
  }
  if (typeof input === "object" && !Array.isArray(input) && Array.isArray(input.rows)) return { ...input, name: input.name || name };
  if (!Array.isArray(input)) throw new GridxcelError("sheet must be an array of rows or a sheet object", "BAD_SHEET");
  const rows = input.map((row) => {
    if (row === null || row === undefined) return [];
    if (Array.isArray(row)) return row.map(normalizeCell);
    if (typeof row === "object") {
      // Array-of-objects form: keys are column letters or header names.
      const out = [];
      for (const [k, v] of Object.entries(row)) {
        const col = /^[A-Za-z]{1,3}$/.test(k) ? decodeCell(k + "1").c : Number(k);
        if (Number.isInteger(col) && col >= 0) out[col] = normalizeCell(v);
      }
      return out;
    }
    return [normalizeCell(row)];
  });
  let columnCount = 0;
  for (const r of rows) if (r.length > columnCount) columnCount = r.length;
  return { name, rows, originRow: 1, originCol: 0, columnCount, rowCount: rows.length, merges: [], columns: [], rowInfo: new Map() };
}

/* ------------------------------------------------------------ sheet names */

/** Characters Excel forbids anywhere in a sheet name. */
const ILLEGAL_SHEET_NAME_CHARS = /[\\\/?*\[\]:]/;

/** Names Excel keeps for its own use, compared case-insensitively. */
const RESERVED_SHEET_NAMES = new Set(["history"]);

/**
 * The first free default sheet name: Sheet1, Sheet2, ... skipping any already in
 * use. Comparison is case-insensitive, matching validateSheetName's uniqueness
 * rule, so "sheet1" is not offered when "Sheet1" exists.
 *
 * This lived in two places and was missing from a third. The third is the one
 * that mattered: WorkbookWriter.addSheet named an unnamed sheet
 * "Sheet" + (sheets.length + 1), which collides as soon as the caller has used
 * that exact name, so the same call that succeeded through either utils surface
 * threw BAD_SHEET_NAME through the writer.
 *
 * @param {Iterable<string>} taken names already in the workbook
 * @returns {string} a name that validateSheetName will accept
 */
export function freeSheetName(taken) {
  const list = Array.from(taken, (t) => String(t).toLowerCase());
  const prefix = DEFAULT_SHEET_NAME.replace(/[0-9]+$/, "");
  let n = 1;
  while (list.includes((prefix + n).toLowerCase())) n++;
  return prefix + n;
}

/**
 * Throw GridxcelError / BAD_SHEET_NAME when a sheet name is not usable in a
 * workbook, so every entry point rejects the same names at the same moment.
 *
 * @param {string} name
 * @param {Set<string>|string[]|Map} [taken] names already in the workbook, for
 *   the uniqueness rule. Comparison is case-insensitive, as Excel's is.
 * @returns {string} the name, so callers can write `name = validateSheetName(name, taken)`
 * @throws {GridxcelError} code BAD_SHEET_NAME
 */
export function validateSheetName(name, taken) {
  if (typeof name !== "string") {
    throw new GridxcelError("sheet name must be a string, got " + typeof name, "BAD_SHEET_NAME");
  }
  if (name.length === 0) {
    throw new GridxcelError("sheet name must not be empty", "BAD_SHEET_NAME");
  }
  if (name.length > MAX_SHEET_NAME_LENGTH) {
    throw new GridxcelError(
      "sheet name must be 1.." + MAX_SHEET_NAME_LENGTH + " characters, got " + name.length,
      "BAD_SHEET_NAME",
      { name },
    );
  }
  if (ILLEGAL_SHEET_NAME_CHARS.test(name)) {
    throw new GridxcelError(
      "sheet name " + JSON.stringify(name) + " may not contain \\ / ? * [ ] :",
      "BAD_SHEET_NAME",
      { name },
    );
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new GridxcelError(
      "sheet name " + JSON.stringify(name) + " may not start or end with an apostrophe",
      "BAD_SHEET_NAME",
      { name },
    );
  }
  if (RESERVED_SHEET_NAMES.has(name.toLowerCase())) {
    throw new GridxcelError("sheet name " + JSON.stringify(name) + " is reserved by Excel", "BAD_SHEET_NAME", { name });
  }
  if (takenIncludes(taken, name)) {
    throw new GridxcelError("duplicate sheet name " + JSON.stringify(name), "BAD_SHEET_NAME", { name });
  }
  return name;
}

/** Case-insensitive membership over the accepted `taken` shapes: Set, Map or array. */
function takenIncludes(taken, name) {
  if (!taken) return false;
  const names = Array.isArray(taken) ? taken : typeof taken.keys === "function" ? taken.keys() : taken;
  if (typeof names[Symbol.iterator] !== "function") return false;
  const lower = name.toLowerCase();
  for (const n of names) if (String(n).toLowerCase() === lower) return true;
  return false;
}

/**
 * The canonical empty workbook.
 *
 * The same seven-field literal used to be written out in seven places (here,
 * compat, read and csv). Every new Book field then needed seven edits, and a
 * missed one made the Book's shape depend on which entry point produced it.
 * `props` becomes Book.props; callers that need another date system set
 * `book.dateSystem` after the call.
 */
export function emptyBook(props) {
  return {
    sheets: [],
    dateSystem: 1900,
    numFmts: new Map(),
    styles: [],
    props: props || {},
    definedNames: [],
  };
}

/** Build a Book from the many shapes callers pass to write helpers. */
export function normalizeBook(input, options = {}) {
  const base = emptyBook();
  if (options.date1904) base.dateSystem = 1904;
  if (input === null || input === undefined) return base;

  if (Array.isArray(input)) {
    // Either an array of sheets or a single 2-D array.
    if (input.length && Array.isArray(input[0]) && !Array.isArray(input[0][0])) {
      const name = options.sheet || DEFAULT_SHEET_NAME;
      base.sheets.push(normalizeSheet(input, name));
    } else {
      input.forEach((s, i) => base.sheets.push(normalizeSheet(s, "Sheet" + (i + 1))));
    }
    return base;
  }

  if (typeof input === "object") {
    if (Array.isArray(input.sheets)) {
      base.sheets = input.sheets.map((s, i) => normalizeSheet(s, "Sheet" + (i + 1)));
      if (input.dateSystem === 1904) base.dateSystem = 1904;
      if (input.props) base.props = input.props;
      if (input.definedNames) base.definedNames = input.definedNames;
      if (input.numFmts) base.numFmts = input.numFmts;
      if (input.styles) base.styles = input.styles;
      return base;
    }
    // SheetJS-shaped: { SheetNames, Sheets: { name: sheet } }
    if (Array.isArray(input.SheetNames) && input.Sheets) {
      base.sheets = input.SheetNames.map((n) => normalizeSheet(input.Sheets[n], n));
      return base;
    }
    // Single sheet object
    if (Array.isArray(input.rows)) {
      base.sheets = [normalizeSheet(input, input.name || options.sheet || DEFAULT_SHEET_NAME)];
      return base;
    }
  }
  throw new GridxcelError("cannot interpret the workbook input", "BAD_BOOK");
}

/* ------------------------------------------------------------------ reading */

/**
 * Read a workbook.
 * @param {Uint8Array|ArrayBuffer|string|ReadableStream|AsyncIterable} input
 * @param {object} [options]
 * @returns {Promise<object>} Book
 */
export async function readBook(input, options) {
  const fn = await load("./read.js", "readWorkbook");
  return fn(input, options);
}

/* ------------------------------------------------------------------ writing */

/**
 * Write a workbook.
 * options.path writes to disk; options.bookType selects xlsx/csv/json.
 */
export async function writeBook(book, options = {}) {
  const type = (options.bookType || "xlsx").toLowerCase();
  if (type === "csv") {
    const sheetToCSV = await load("./csv.js", "sheetToCSV");
    const sheet = normalizeBook(book, options).sheets[0];
    const text = sheetToCSV(sheet, options);
    if (options.path) {
      await writeFileNode(options.path, text, "utf8");
      return undefined;
    }
    return text;
  }
  if (type === "json") {
    const fn = await load("./json.js", "sheetToJSON");
    const sheet = normalizeBook(book, options).sheets[0];
    const rows = fn(sheet, options);
    const text = JSON.stringify(rows, null, options.indent === undefined ? 2 : options.indent);
    if (options.path) {
      await writeFileNode(options.path, text, "utf8");
      return undefined;
    }
    return text;
  }
  if (type !== "xlsx" && type !== "xlsm") {
    throw new GridxcelError("unsupported bookType " + JSON.stringify(type) + " (xlsx, csv, json)", "BAD_TYPE");
  }
  const fn = await load("./write.js", "writeWorkbook");
  return fn(normalizeBook(book, options), options);
}

/* ------------------------------------------------------------------ helpers */

/** Merge two or more books; sheet names are preserved, later sheets win. */
export function mergeBooks(...books) {
  const out = emptyBook();
  for (const b of books) {
    const nb = normalizeBook(b);
    out.sheets.push(...nb.sheets);
    if (nb.dateSystem === 1904) out.dateSystem = 1904;
  }
  return out;
}

/** Rename a sheet, rejecting duplicates and illegal names the way Excel does. */
export function renameSheet(book, from, to) {
  const sheet = book.sheets.find((s) => s.name === from);
  if (!sheet) throw new GridxcelError("no sheet named " + JSON.stringify(from), "NO_SHEET");
  // Every format rule comes from the shared validator, so an illegal name fails
  // here instead of surfacing later when the workbook is written.
  const name = validateSheetName(to);
  const others = book.sheets.filter((s) => s !== sheet);
  // A clash keeps its own code: API.md documents DUPLICATE_SHEET for rename.
  if (others.some((s) => String(s.name).toLowerCase() === name.toLowerCase())) {
    throw new GridxcelError("sheet " + name + " already exists", "DUPLICATE_SHEET");
  }
  sheet.name = name;
  return book;
}

/** Count non-empty cells across a book (useful in tests and diagnostics). */
export function countCells(book) {
  let n = 0;
  for (const s of book.sheets) for (const row of s.rows) for (const c of row) if (c && c.t && c.t !== "empty") n++;
  return n;
}

export { decodeRange, encodeRange, encodeCell, decodeCell };
