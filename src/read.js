/**
 * The read path: bytes (or a path / stream / async iterable) in, Book out.
 *
 * Layout of the pipeline:
 *   opc.js            package plumbing + workbook metadata (small parts)
 *   read.js           orchestration: styles, shared strings, comments,
 *                     sheet selection, range/strict/warning policy
 *   sheet-reader.js   worksheet XML -> Sheet or a bounded-memory row stream
 *
 * Two entry points share every helper: readWorkbook is async and accepts
 * anything streamable, readWorkbookSync takes bytes or a path. Worksheet parts
 * are never inflated wholesale: a STORED part is handed to the scanner as the
 * raw byte view the zip already holds, a DEFLATED part is piped through
 * entry.stream().
 */

import { ZipArchive } from "./zip.js";
import { FormatError, GridxcelError, StrictWarningError } from "./errors.js";
// Filesystem access is Node-only and lives behind the platform seam, so this
// module carries no node:fs/node:url import of its own: a browser bundle can
// load it, and the file helpers raise PlatformError [NOT_SUPPORTED] there.
import { fileUrlToPath, readFileNode, readFileNodeSync } from "./platform.js";
import { XmlParser, localName } from "./xml.js";
import { BUILTIN_FORMATS } from "./numfmt.js";
import { toRange } from "./address.js";
import { getPart, parseRels, prewarmMeta, readWorkbookMeta, resolveTarget, relsPathOf } from "./opc.js";
import {
  createRowStream, parseSharedStrings, parseSharedStringsSync, parseSheetXml, parseSheetXmlSync,
} from "./sheet-reader.js";
// THE styles parser. This module used to carry its own copy, which drifted from
// the writer's on numFmt (id vs resolved code) and on whether numFmts covered
// the built-ins. Both consumers now read one record.
import { parseStyles } from "./style-registry.js";
import { MAX_SHEETS } from "./constants.js";

const utf8 = new TextEncoder();

/* ------------------------------------------------------------------ warnings */

/** Routes non-fatal problems to onWarning, or throws in strict mode. */
class WarningSink {
  constructor(opts) {
    this.strict = opts.strict === true;
    this.sink = typeof opts.onWarning === "function" ? opts.onWarning : undefined;
    this.part = undefined;
    this.count = 0;
  }

  warn(code, message, extra) {
    const w = { code, message };
    const part = extra && extra.part !== undefined ? extra.part : this.part;
    if (part) w.part = part;
    if (extra) {
      if (extra.row !== undefined) w.row = extra.row;
      if (extra.col !== undefined) w.col = extra.col;
    }
    this.count++;
    if (this.strict) throw new StrictWarningError(w);
    if (this.sink) this.sink(w);
  }
}

/* ------------------------------------------------------------------- options */

function normalizeOptions(options) {
  const o = options === undefined || options === null ? {} : options;
  const opts = {
    sheets: o.sheets === undefined ? true : o.sheets,
    sheetStubs: o.sheetStubs === true,
    cellDates: o.cellDates === true,
    cellText: o.cellText === true,
    cellFormula: o.cellFormula !== false,
    cellComments: o.cellComments === true,
    range: o.range,
    strict: o.strict === true,
    onWarning: typeof o.onWarning === "function" ? o.onWarning : undefined,
    keepParts: o.keepParts === true,
    bounds: boundsOf(o.range),
  };
  return opts;
}

/** {s:{r,c},e:{r,c}} or "A1:B10" -> inclusive bound box, or null. */
function boundsOf(range) {
  if (range === undefined || range === null) return null;
  const r = toRange(range);
  return {
    minR: Math.min(r.s.r, r.e.r),
    minC: Math.min(r.s.c, r.e.c),
    maxR: Math.max(r.s.r, r.e.r),
    maxC: Math.max(r.s.c, r.e.c),
  };
}

/* --------------------------------------------------------------------- input */

/** Filesystem failures become typed errors; the raw error is kept as cause. */
function fileError(path, err) {
  const map = {
    ENOENT: "FILE_NOT_FOUND",
    EACCES: "FILE_ACCESS_DENIED",
    EPERM: "FILE_ACCESS_DENIED",
    EISDIR: "FILE_IS_DIRECTORY",
  };
  const code = map[err && err.code] || "FILE_READ_FAILED";
  const wrapped = new GridxcelError("cannot read " + JSON.stringify(String(path)) + ": " + (err && err.message), code, { path: String(path) });
  wrapped.cause = err;
  return wrapped;
}

async function readPathBytes(path) {
  try {
    return await readFileNode(path);
  } catch (err) {
    // A PlatformError already says that a filesystem is unavailable; wrapping it
    // as FILE_READ_FAILED would bury the actionable part of the message.
    if (err instanceof GridxcelError) throw err;
    throw fileError(path, err);
  }
}

function readPathBytesSync(path) {
  try {
    return readFileNodeSync(path);
  } catch (err) {
    if (err instanceof GridxcelError) throw err;
    throw fileError(path, err);
  }
}

async function toBytes(input) {
  // Non-file URLs keep their string form; a file: URL becomes a path on Node and
  // raises the platform's typed refusal elsewhere.
  if (input instanceof URL) input = fileUrlToPath(input);
  if (typeof input === "string") return readPathBytes(input);
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  if (typeof Blob !== "undefined" && input instanceof Blob) return new Uint8Array(await input.arrayBuffer());
  if (input && typeof input.getReader === "function") return collect(input);
  if (input && (input[Symbol.asyncIterator] || input[Symbol.iterator])) return collect(input);
  throw new GridxcelError("unsupported workbook input", "BAD_INPUT");
}

function toBytesSync(input) {
  if (input instanceof URL) input = fileUrlToPath(input);
  if (typeof input === "string") return readPathBytesSync(input);
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  throw new GridxcelError("readWorkbookSync needs bytes or a file path, not a stream", "BAD_INPUT");
}

/** Concatenate a byte stream / async iterable into one buffer. */
async function collect(source) {
  const parts = [];
  let total = 0;
  for await (const chunk of source) {
    const bytes = typeof chunk === "string"
      ? utf8.encode(chunk)
      : chunk instanceof Uint8Array ? chunk
        : chunk instanceof ArrayBuffer ? new Uint8Array(chunk)
          : new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength);
    parts.push(bytes);
    total += bytes.length;
  }
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/* -------------------------------------------------------------------- styles */

function colorFrom(attrs) {
  if (attrs.rgb) return attrs.rgb;
  const out = {};
  if (attrs.theme !== undefined) out.theme = parseInt(attrs.theme, 10);
  if (attrs.tint !== undefined) out.tint = parseFloat(attrs.tint);
  if (attrs.indexed !== undefined) out.indexed = parseInt(attrs.indexed, 10);
  if (attrs.auto !== undefined) out.auto = true;
  return Object.keys(out).length ? out : undefined;
}

/**
 * The sheet reader consumes styles through a tiny method interface
 * (numFmtId/formatOf/isDate per xf index), because it looks a style up for
 * every cell. This adapts the canonical parseStyles result to it; nothing
 * here re-parses the part or re-resolves an id, so the reader cannot drift
 * from the writer's view of the same styles.xml.
 */
function readerStyles(parsed) {
  const styles = parsed.styles;
  return {
    numFmts: parsed.numFmts,
    styles,
    numFmtId: (i) => (styles[i] === undefined ? undefined : styles[i].numFmt),
    code: (i) => (styles[i] === undefined ? undefined : styles[i].numFmtCode),
    formatOf: (i) => {
      const style = styles[i];
      return style === undefined || style.numFmtCode === null ? undefined : style.numFmtCode;
    },
    isDate: parsed.isDate,
  };
}

const NO_STYLES = {
  numFmts: new Map(),
  styles: [],
  numFmtId: () => undefined,
  code: () => undefined,
  formatOf: () => undefined,
  isDate: () => false,
};

function loadStyles(zip, meta, sink) {
  const path = meta.pathOf.styles;
  if (!path) return NO_STYLES;
  const entry = getPart(zip, path);
  if (!entry) return NO_STYLES;
  const previous = sink.part;
  sink.part = path;
  try {
    return readerStyles(parseStyles(entry.text()));
  } catch (err) {
    sink.warn("BAD_STYLES", "styles could not be read: " + err.message);
    return NO_STYLES;
  } finally {
    sink.part = previous;
  }
}

/**
 * Resolved number-format code for a cell that came out of this reader.
 * Cells carry ids (s = xf index, z = numFmtId); writers need the code.
 * @param {object} book a Book returned by readWorkbook/readWorkbookSync
 * @param {{s?: number, z?: number}} cell
 * @returns {string|number|undefined}
 */
export function cellFormatCode(book, cell) {
  if (!cell) return undefined;
  if (cell.s !== undefined && book && Array.isArray(book.styles)) {
    const style = book.styles[cell.s];
    if (style) {
      // A style record carries both views: prefer the resolved code, fall back
      // to the id, which book.numFmts resolves (built-ins included).
      if (style.numFmtCode !== undefined && style.numFmtCode !== null) return style.numFmtCode;
      if (style.numFmt !== undefined && book.numFmts instanceof Map) {
        const hit = book.numFmts.get(style.numFmt);
        if (hit !== undefined) return hit;
      }
      if (style.numFmt !== undefined) return style.numFmt;
    }
  }
  if (cell.z !== undefined && book && book.numFmts instanceof Map) {
    const hit = book.numFmts.get(cell.z);
    if (hit !== undefined) return hit;
    if (BUILTIN_FORMATS[cell.z] !== undefined) return BUILTIN_FORMATS[cell.z];
  }
  return cell.z;
}

/* ------------------------------------------------------------ shared strings */

async function loadSharedStrings(zip, meta, sink) {
  const path = meta.pathOf.sharedStrings;
  if (!path) return null;                                   // inline strings only
  const entry = getPart(zip, path);
  if (!entry) return null;
  const previous = sink.part;
  sink.part = path;
  try {
    return await parseSharedStrings(chunksFor(entry, sink), (code, message) => sink.warn(code, message));
  } catch (err) {
    if (err instanceof StrictWarningError) throw err;
    sink.warn("BAD_SHARED_STRINGS", "shared strings could not be read: " + err.message);
    return null;
  } finally {
    sink.part = previous;
  }
}

function loadSharedStringsSync(zip, meta, sink) {
  const path = meta.pathOf.sharedStrings;
  if (!path) return null;
  const entry = getPart(zip, path);
  if (!entry) return null;
  const previous = sink.part;
  sink.part = path;
  try {
    return parseSharedStringsSync(entry.bytes(), (code, message) => sink.warn(code, message));
  } catch (err) {
    if (err instanceof StrictWarningError) throw err;
    sink.warn("BAD_SHARED_STRINGS", "shared strings could not be read: " + err.message);
    return null;
  } finally {
    sink.part = previous;
  }
}

/* ------------------------------------------------------------------ comments */

/** Comments part of a worksheet: relationship first, then the usual name. */
function commentsPathFor(zip, sheetPath, sheetRels) {
  for (const rel of sheetRels.values()) {
    if (rel.kind === "comments" && rel.mode === "Internal") return resolveTarget(sheetPath, rel.target);
  }
  const m = /(?:^|\/)sheet([0-9]+)\.xml$/.exec(sheetPath);
  if (m) {
    const guess = "/xl/comments" + m[1] + ".xml";
    if (getPart(zip, guess)) return guess;
  }
  return undefined;
}

/** comments1.xml -> Map<"A1", text> with rich text runs concatenated. */
export function parseCommentsXml(xml) {
  const out = new Map();
  let cur = null;
  const parser = new XmlParser({
    onOpen(name, attrs) {
      if (localName(name) === "comment") cur = { ref: attrs.ref, text: "" };
    },
    onClose(name) {
      if (localName(name) === "comment") {
        if (cur && cur.ref) out.set(cur.ref, cur.text);
        cur = null;
      }
    },
    onText(text, node) {
      if (cur && node && node.local === "t") cur.text += text;
    },
  });
  if (xml) { parser.write(xml); parser.end(); }
  return out;
}

/**
 * Warm the per-sheet parts the synchronous loaders read: the sheet's
 * relationships part and, when it names one, its comments part.
 *
 * Same reason as opc.prewarmMeta(): the Web decompressor is asynchronous, so on
 * a host without node:zlib the async entry points cache the small parts first
 * and the shared synchronous loaders (loadStyles, relsFor, loadComments) then
 * find them in the entry cache. Worksheets and sharedStrings are deliberately
 * NOT warmed - they stay on the chunked/streaming path.
 */
async function prewarmSheetParts(zip, meta) {
  const sheets = (meta && meta.pathOf && meta.pathOf.sheets) || [];
  for (const sheetPath of sheets) {
    if (!sheetPath) continue;
    const relsEntry = getPart(zip, relsPathOf(sheetPath));
    let sheetRels = new Map();
    if (relsEntry) {
      try {
        sheetRels = parseRels(await relsEntry.textAsync());
      } catch {
        sheetRels = new Map();     // relsFor() reports it with its usual warning
      }
    }
    const commentsPath = commentsPathFor(zip, sheetPath, sheetRels);
    const commentsEntry = commentsPath ? getPart(zip, commentsPath) : undefined;
    if (!commentsEntry) continue;
    try {
      await commentsEntry.textAsync();
    } catch {
      // loadComments() reports the failure when the caller asked for comments.
    }
  }
}

function loadComments(zip, sheetPath, sheetRels, sink) {
  const path = commentsPathFor(zip, sheetPath, sheetRels);
  if (!path) return null;
  const entry = getPart(zip, path);
  if (!entry) return null;
  const previous = sink.part;
  sink.part = path;
  try {
    return parseCommentsXml(entry.text());
  } catch (err) {
    if (err instanceof StrictWarningError) throw err;
    sink.warn("BAD_COMMENTS", "comments could not be read: " + err.message);
    return null;
  } finally {
    sink.part = previous;
  }
}

/* ------------------------------------------------------------------ zip glue */

/**
 * Open the container. Bytes that are not a zip at all are a caller error, not
 * a damaged workbook: returning an empty Book here would turn a typo or a
 * mislabelled file into an apparent success.
 */
function openContainer(bytes) {
  try {
    return new ZipArchive(bytes);
  } catch (err) {
    const e = new FormatError("input is not a zip container (" + err.message + ")", "NOT_A_ZIP", { size: bytes.length });
    e.cause = err;
    throw e;
  }
}

/**
 * A zip that carries no readable workbook part is not a workbook either.
 * Anything past this point is a workbook that may still be partly damaged,
 * which is the tolerant path (warnings, not throws).
 */
function requireWorkbook(meta) {
  if (!meta.workbookPart) {
    throw new FormatError("package contains no workbook part (no officeDocument relationship and no sheet main part)", "NOT_A_WORKBOOK");
  }
  const read = meta.workbookRead || { elements: 0, hasRoot: false };
  if (read.elements === 0) {
    throw new FormatError("workbook part " + meta.workbookPart + " is missing, empty or not XML", "NOT_A_WORKBOOK", { part: meta.workbookPart });
  }
  if (!read.hasRoot) {
    throw new FormatError("workbook part " + meta.workbookPart + " has no workbook element", "NOT_A_WORKBOOK", { part: meta.workbookPart });
  }
}

/**
 * Bytes to feed the row scanner for one part. STORED parts stay zero-copy
 * (the scanner slices them itself); DEFLATED parts are inflated live.
 */
function chunksFor(entry, sink) {
  if (entry.method === 0) return entry.bytes();
  if (entry.method === 8) return entry.stream();
  sink.warn("UNSUPPORTED_METHOD", "unsupported compression method " + entry.method + " for " + entry.name, { part: entry.name });
  return null;
}

function relsFor(zip, partPath, sink) {
  const path = relsPathOf(partPath);
  const entry = getPart(zip, path);
  if (!entry) return new Map();
  try {
    return parseRels(entry.text());
  } catch (err) {
    sink.warn("BAD_XML", "relationships could not be read: " + err.message, { part: path });
    return new Map();
  }
}

/* ------------------------------------------------------------ sheet selection */

function selectSheets(declared, wanted, sink) {
  if (wanted === true || wanted === undefined) return declared;
  if (wanted === false) return [];
  if (!Array.isArray(wanted)) {
    sink.warn("NO_SHEET", "sheets option must be true, false or an array");
    return declared;
  }
  const picked = [];
  for (const item of wanted) {
    if (typeof item === "number") {
      const hit = declared[item];
      if (hit) picked.push(hit);
      else sink.warn("NO_SHEET", "no sheet at index " + item);
    } else if (typeof item === "string") {
      const hit = declared.find((s) => s.name === item);
      if (hit) picked.push(hit);
      else sink.warn("NO_SHEET", "no sheet named " + JSON.stringify(item));
    } else {
      sink.warn("NO_SHEET", "sheet selector must be a name or an index");
    }
  }
  return picked;
}

/* ------------------------------------------------------------------ assembly */

function emptyBook() {
  return {
    sheets: [],
    sheetNames: [],
    dateSystem: 1900,
    numFmts: new Map(),
    styles: [],
    props: {},
    definedNames: [],
  };
}

function readerOptionsFor(metaSheet, opts, shared, styles, zip, sink, date1904) {
  const sheetPath = metaSheet.path;
  const sheetRels = relsFor(zip, sheetPath, sink);
  const o = {
    name: metaSheet.name,
    hidden: metaSheet.hidden === true,
    sharedStrings: shared,
    styles,
    date1904,
    bounds: opts.bounds,
    cellDates: opts.cellDates,
    cellText: opts.cellText,
    cellFormula: opts.cellFormula,
    sheetStubs: opts.sheetStubs,
    warn: (code, message, extra) => sink.warn(code, message, extra),
    relTarget: (rid) => {
      const rel = sheetRels.get(rid);
      if (!rel) return undefined;
      if (rel.mode !== "Internal") return { target: rel.target, mode: rel.mode };
      return { target: resolveTarget(sheetPath, rel.target), mode: rel.mode };
    },
  };
  if (opts.cellComments) {
    const comments = loadComments(zip, sheetPath, sheetRels, sink);
    if (comments) o.comments = comments;
  }
  return o;
}

function declaredSheets(meta, sink) {
  if (!meta.sheets.length) {
    sink.warn("NO_SHEETS", "workbook declares no sheets");
    return [];
  }
  // A workbook is untrusted input, and the sheet list is a number inside it.
  // A 794 kB file declaring 100,000 <sheet> entries cost 5.9 s and 321 MB
  // before this cap; the parts beyond it are skipped with a warning rather
  // than refused, so the sheets that were named are still readable.
  if (meta.sheets.length > MAX_SHEETS) {
    sink.warn(
      "TOO_MANY_SHEETS",
      "workbook declares " + meta.sheets.length + " sheets; only the first " + MAX_SHEETS + " are read",
    );
    return meta.sheets.slice(0, MAX_SHEETS);
  }
  return meta.sheets;
}

function keepRawParts(zip, book) {
  const parts = new Map();
  for (const entry of zip.entries) {
    if (entry.isDirectory) continue;
    try {
      parts.set(entry.name, entry.bytes());
    } catch {
      // An unreadable part is already reported when the reader needs it.
    }
  }
  book.parts = parts;
}

/**
 * The async twin: keepParts must not depend on a synchronous decompressor, so
 * the async entry points collect through bytesAsync() (node:zlib when present,
 * DecompressionStream when not).
 */
async function keepRawPartsAsync(zip, book) {
  const parts = new Map();
  for (const entry of zip.entries) {
    if (entry.isDirectory) continue;
    try {
      parts.set(entry.name, await entry.bytesAsync());
    } catch {
      // Same contract as the sync twin.
    }
  }
  book.parts = parts;
}

function finishBook(book, zip, opts, sink) {
  if (opts.keepParts) keepRawParts(zip, book);
  return book;
}

async function finishBookAsync(book, zip, opts) {
  if (opts.keepParts) await keepRawPartsAsync(zip, book);
  return book;
}

async function assemble(zip, meta, opts, sink) {
  for (const w of meta.warnings) sink.warn(w.code, w.message, w);
  const date1904 = meta.date1904 === true;
  const styles = loadStyles(zip, meta, sink);
  const book = emptyBook();
  book.dateSystem = date1904 ? 1904 : 1900;
  book.numFmts = styles.numFmts;
  book.styles = styles.styles;
  book.props = meta.props;
  book.definedNames = meta.definedNames;
  book.sheetNames = meta.sheets.map((s) => s.name);
  const declared = declaredSheets(meta, sink);
  if (opts.sheets !== false && declared.length) {
    const shared = await loadSharedStrings(zip, meta, sink);
    for (const metaSheet of selectSheets(declared, opts.sheets, sink)) {
      if (metaSheet.kind && metaSheet.kind !== "worksheet") {
        sink.warn("UNSUPPORTED_SHEET_TYPE", "sheet " + JSON.stringify(metaSheet.name) + " is a " + metaSheet.kind, { part: metaSheet.path });
        continue;
      }
      const entry = getPart(zip, metaSheet.path);
      if (!entry) {
        sink.warn("MISSING_PART", "worksheet part is missing", { part: metaSheet.path });
        continue;
      }
      sink.part = metaSheet.path;
      const chunks = chunksFor(entry, sink);
      if (!chunks) continue;
      const oh = readerOptionsFor(metaSheet, opts, shared, styles, zip, sink, date1904);
      try {
        book.sheets.push(await parseSheetXml(chunks, oh));
      } catch (err) {
        if (err instanceof StrictWarningError) throw err;
        sink.warn("BAD_SHEET", "worksheet could not be parsed: " + err.message);
      }
    }
  }
  return finishBookAsync(book, zip, opts);
}

function assembleSync(zip, meta, opts, sink) {
  for (const w of meta.warnings) sink.warn(w.code, w.message, w);
  const date1904 = meta.date1904 === true;
  const styles = loadStyles(zip, meta, sink);
  const book = emptyBook();
  book.dateSystem = date1904 ? 1904 : 1900;
  book.numFmts = styles.numFmts;
  book.styles = styles.styles;
  book.props = meta.props;
  book.definedNames = meta.definedNames;
  book.sheetNames = meta.sheets.map((s) => s.name);
  const declared = declaredSheets(meta, sink);
  if (opts.sheets !== false && declared.length) {
    const shared = loadSharedStringsSync(zip, meta, sink);
    for (const metaSheet of selectSheets(declared, opts.sheets, sink)) {
      if (metaSheet.kind && metaSheet.kind !== "worksheet") {
        sink.warn("UNSUPPORTED_SHEET_TYPE", "sheet " + JSON.stringify(metaSheet.name) + " is a " + metaSheet.kind, { part: metaSheet.path });
        continue;
      }
      const entry = getPart(zip, metaSheet.path);
      if (!entry) {
        sink.warn("MISSING_PART", "worksheet part is missing", { part: metaSheet.path });
        continue;
      }
      sink.part = metaSheet.path;
      let bytes;
      try {
        bytes = entry.bytes();
      } catch (err) {
        if (err instanceof StrictWarningError) throw err;
        sink.warn("BAD_SHEET", "worksheet could not be inflated: " + err.message);
        continue;
      }
      if (entry.method !== 0 && entry.method !== 8) {
        sink.warn("UNSUPPORTED_METHOD", "unsupported compression method " + entry.method, { part: entry.name });
        continue;
      }
      const oh = readerOptionsFor(metaSheet, opts, shared, styles, zip, sink, date1904);
      try {
        book.sheets.push(parseSheetXmlSync(bytes, oh));
      } catch (err) {
        if (err instanceof StrictWarningError) throw err;
        sink.warn("BAD_SHEET", "worksheet could not be parsed: " + err.message);
      }
    }
  }
  return finishBook(book, zip, opts, sink);
}

/* -------------------------------------------------------------------- public */

/**
 * Read a workbook from bytes, a file path, a ReadableStream or an async
 * iterable of chunks.
 * @param {Uint8Array|ArrayBuffer|string|ReadableStream|AsyncIterable} input
 * @param {object} [options] ReadOptions
 * @returns {Promise<object>} Book
 */
export async function readWorkbook(input, options = {}) {
  const opts = normalizeOptions(options);
  const sink = new WarningSink(opts);
  const bytes = await toBytes(input);
  const zip = openContainer(bytes);
  // Cache the small parts first. The metadata sweep and the style, relationship
  // and comment loaders are synchronous by contract (they are shared with
  // readWorkbookSync), so a host with only an async decompressor needs those
  // parts warm before the sync code reaches for them.
  await prewarmMeta(zip);
  const meta = readWorkbookMeta(zip);
  requireWorkbook(meta);
  await prewarmSheetParts(zip, meta);
  return assemble(zip, meta, opts, sink);
}

/**
 * Synchronous read from bytes or a file path (Streams cannot be read
 * synchronously, so readWorkbookSync rejects them).
 * @returns {object} Book
 */
export function readWorkbookSync(input, options = {}) {
  const opts = normalizeOptions(options);
  const sink = new WarningSink(opts);
  const bytes = toBytesSync(input);
  const zip = openContainer(bytes);
  const meta = readWorkbookMeta(zip);
  requireWorkbook(meta);
  return assembleSync(zip, meta, opts, sink);
}

/**
 * Stream every row of every selected sheet as { sheet, r, cells } without ever
 * materialising a sheet. cells is a sparse array indexed by absolute column
 * (A == 0) so the generator can hand each row over as soon as it is complete;
 * sheetStubs therefore fills gaps from column A, not from the sheet's first
 * populated column (which is only known once the whole part is read).
 *
 * Every derived cell field the batch reader produces is produced here as well
 * (t, serial, z, zf, f, w, c) with one documented exception: cell.l, the
 * hyperlink, is batch-only, because <hyperlinks> follows <sheetData> and
 * honouring it while streaming would require retaining rows.
 *
 * @param {Uint8Array|ArrayBuffer|string|ReadableStream|AsyncIterable} input
 * @param {object} [options] ReadOptions
 * @returns {AsyncGenerator<{sheet: string, r: number, cells: object[]}>}
 */
export async function* eachRow(input, options = {}) {
  const opts = normalizeOptions(options);
  const sink = new WarningSink(opts);
  const bytes = await toBytes(input);
  const zip = openContainer(bytes);
  // Same warm-up as readWorkbook: the metadata sweep, styles and sheet
  // relationships loaders are synchronous and shared with the sync reader.
  await prewarmMeta(zip);
  const meta = readWorkbookMeta(zip);
  requireWorkbook(meta);
  await prewarmSheetParts(zip, meta);
  for (const w of meta.warnings) sink.warn(w.code, w.message, w);
  const date1904 = meta.date1904 === true;
  const styles = loadStyles(zip, meta, sink);
  const declared = declaredSheets(meta, sink);
  if (opts.sheets === false || !declared.length) return;
  const shared = await loadSharedStrings(zip, meta, sink);
  for (const metaSheet of selectSheets(declared, opts.sheets, sink)) {
    if (metaSheet.kind && metaSheet.kind !== "worksheet") {
      sink.warn("UNSUPPORTED_SHEET_TYPE", "sheet " + JSON.stringify(metaSheet.name) + " is a " + metaSheet.kind, { part: metaSheet.path });
      continue;
    }
    const entry = getPart(zip, metaSheet.path);
    if (!entry) {
      sink.warn("MISSING_PART", "worksheet part is missing", { part: metaSheet.path });
      continue;
    }
    sink.part = metaSheet.path;
    const chunks = chunksFor(entry, sink);
    if (!chunks) continue;
    const oh = readerOptionsFor(metaSheet, opts, shared, styles, zip, sink, date1904);
    oh.collectMeta = false;                 // no sheet-level retention on this path
    // Comments are attached here too; hyperlinks cannot be, because the
    // <hyperlinks> element follows <sheetData> and a single-pass stream must
    // not buffer rows to wait for it.
    oh.relTarget = undefined;
    const stream = createRowStream(chunks, oh);
    for await (const row of stream) {
      yield { sheet: metaSheet.name, r: row.r, cells: row.cells };
    }
  }
}
