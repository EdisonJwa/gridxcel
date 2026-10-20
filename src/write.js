/**
 * Write entry points: writeWorkbook (async), writeWorkbookSync and toBuffer.
 *
 * All three funnel through createBookWriter so a book written from disk, from
 * a buffer or streamed by hand produces byte-identical parts.
 *
 * The module itself is host-agnostic: no Node import is evaluated at load time.
 * options.path is the only Node-only feature - the filesystem calls come from
 * ./platform.js, which reaches node:fs through a dynamic import and raises
 * PlatformError [NOT_SUPPORTED] in a browser. Everything else (building the
 * package, every options.type shape) works with the Web platform alone.
 */

import { GridxcelError, NotImplementedError } from "./errors.js";
import { createBookWriter } from "./workbook-writer.js";
import { asBuffer, encodeBase64, encodeLatin1, writeFileNode, writeFileNodeSync } from "./platform.js";

/** SheetJS-style output shaping. Anything else returns the raw bytes. */
function shape(bytes, type) {
  switch (type) {
    case undefined:
    case null:
    case "":
    case "uint8array":
    case "bytes":
      return bytes;
    case "buffer":
      // A Node Buffer when the host has one (unchanged), the bytes themselves
      // in a browser, where Buffer does not exist. See platform.asBuffer.
      return asBuffer(bytes);
    case "array":
      // Return a real ArrayBuffer with no slack or offset.
      if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) return bytes.buffer;
      return bytes.slice().buffer;
    case "base64":
      // btoa over a latin1 string: see platform.encodeBase64 for why that beat
      // a Buffer dependency and a hand-rolled base64 table.
      return encodeBase64(bytes);
    case "binary":
    case "string":
      // Both spellings are a latin1 "binary string", which is what SheetJS
      // produces for spreadsheet formats; one char per byte, exactly what
      // Buffer#toString("latin1") returned.
      return encodeLatin1(bytes);
    default:
      throw new GridxcelError('unsupported write type "' + type +
        '": expected buffer, array, base64, binary or string', "BAD_TYPE");
  }
}

function level(options) {
  const value = options.compressionLevel === undefined ? 6 : options.compressionLevel;
  if (!Number.isInteger(value) || value < 0 || value > 9) {
    throw new GridxcelError("compressionLevel must be an integer 0..9, got " + options.compressionLevel, "BAD_OPTION");
  }
  return value;
}

function buildBytes(book, options) {
  if (options.bookType !== undefined && options.bookType !== null && options.bookType !== "xlsx") {
    throw new NotImplementedError("writeWorkbook only emits xlsx; asked for " + options.bookType +
      " - use the csv/json helpers for that book type");
  }
  const writer = createBookWriter(book, {
    compressionLevel: level(options),
    deterministic: options.deterministic === true,
    recalc: options.recalc,
    inlineStrings: options.inlineStrings,
    activeTab: options.activeTab,
    defaultRowHeight: options.defaultRowHeight,
  });
  return writer.finish();
}

/** Write a book to disk, honouring the atomic no-overwrite default. */
function persist(path, bytes, overwrite) {
  // "wx" makes the existence check atomic instead of a stat race.
  return { path, bytes, flag: overwrite === true ? "w" : "wx" };
}

function fileError(err, path, overwrite) {
  if (err && err.code === "EEXIST") {
    return new GridxcelError("refusing to overwrite " + path + ": pass overwrite: true to replace it", "FILE_EXISTS");
  }
  return err;
}

/**
 * Build the package in memory. The return shape follows options.type:
 * "buffer" -> Buffer on Node, Uint8Array elsewhere; "array" -> ArrayBuffer;
 * "base64" -> base64 string; "binary"/"string" -> latin1 binary string;
 * anything else -> Uint8Array.
 */
export function toBuffer(book, options = {}) {
  return buildBytes(book, options);
}

/**
 * Synchronous write. Returns the bytes (or the shaped value when options.type
 * is set); options.path also writes the file.
 */
export function writeWorkbookSync(book, options = {}) {
  const bytes = buildBytes(book, options);
  const path = options.path;
  if (path !== undefined && path !== null) {
    const target = persist(path, bytes, options.overwrite);
    try {
      writeFileNodeSync(target.path, target.bytes, { flag: target.flag });
    } catch (err) {
      throw fileError(err, path, options.overwrite);
    }
  }
  return options.type === undefined || options.type === null || options.type === ""
    ? bytes
    : shape(bytes, options.type);
}

/**
 * Async write. Returns the bytes (or the shaped value when options.type is
 * set); options.path also writes the file (Node only, via a lazy
 * node:fs/promises import).
 */
export async function writeWorkbook(book, options = {}) {
  const bytes = buildBytes(book, options);
  const path = options.path;
  if (path !== undefined && path !== null) {
    const target = persist(path, bytes, options.overwrite);
    try {
      await writeFileNode(target.path, target.bytes, { flag: target.flag });
    } catch (err) {
      throw fileError(err, path, options.overwrite);
    }
  }
  return options.type === undefined || options.type === null || options.type === ""
    ? bytes
    : shape(bytes, options.type);
}
