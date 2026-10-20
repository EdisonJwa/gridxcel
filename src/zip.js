/**
 * Minimal, correct ZIP (PKZIP APPNOTE 6.3.x) reader + writer.
 *
 * Scope: what OPC/.xlsx/.ods actually use - STORE and DEFLATE, Zip64 when a
 * part exceeds 4 GiB, UTF-8 names (flag bit 11), data descriptors on read.
 *
 * Dependency-free and host-agnostic: the raw deflate primitives come from
 * ./platform.js, which is the only module allowed to reach for node:zlib. This
 * file therefore has no static Node import, so a browser bundler can load it.
 *
 * The async half of the API (readWorkbook, eachRow, ZipEntry.stream,
 * ZipEntry.bytesAsync/textAsync) works on every platform. The synchronous half
 * (bytes, text, ZipEntry.bytes() through readWorkbookSync, ZipWriter.add's
 * deflate step) needs a synchronous codec, which only Node provides; outside
 * Node those raise PlatformError [NOT_SUPPORTED] instead of a ReferenceError,
 * and ZipWriter stores its parts rather than failing (a stored zip is valid).
 */

import { crc32, Crc32 } from "./crc32.js";
import { BadZipError, GridxcelError } from "./errors.js";
import {
  asBuffer,
  deflateBackend,
  deflateBackendSync,
  hasSyncDeflate,
  nodeStreams,
  readFileNode,
  webBytes,
  webInflateRawStream,
} from "./platform.js";

export const SIG_LOCAL = 0x04034b50;
export const SIG_CENTRAL = 0x02014b50;
export const SIG_EOCD = 0x06054b50;
export const SIG_EOCD64 = 0x06064b50;
export const SIG_EOCD64_LOC = 0x07064b50;
export const SIG_DATA_DESC = 0x08074b50;

export const METHOD_STORE = 0;
export const METHOD_DEFLATE = 8;

/**
 * Zip-level failures reuse the shared taxonomy so that a caller can always
 * catch container problems with a single `instanceof GridxcelError` check.
 * The extra `offset` field points at the byte that could not be parsed.
 */
export class ZipError extends BadZipError {
  constructor(message, code = "ZIP_ERROR", offset) {
    super(message, code, offset === undefined ? undefined : { offset });
    this.name = "ZipError";
    if (offset !== undefined) this.offset = offset;
  }
}

/**
 * Map a node:zlib failure onto the shared taxonomy.
 *
 * Inflate is driven straight from attacker-controlled bytes, so it fails with
 * a plain Error carrying a Z_* code. Letting that through would break the
 * promise that every failure raised by this module is a GridxcelError and would
 * make a corrupt part an unhandled crash for a caller that catches by type. The
 * original error stays reachable as `cause`, which is what still tells a
 * truncated stream (Z_BUF_ERROR) from a corrupt one (Z_DATA_ERROR).
 */
function inflateError(cause, name) {
  const zlibCode = cause && cause.code ? String(cause.code) : "Z_UNKNOWN";
  const detail = cause && cause.message ? cause.message : String(cause);
  const err = new ZipError(
    "could not inflate " + name + ": " + detail + " (" + zlibCode + ")",
    "INFLATE_FAILED",
  );
  err.details = { part: name, zlibCode };
  err.cause = cause;
  return err;
}

const utf8Decoder = new TextDecoder("utf-8");
const utf8Encoder = new TextEncoder();

function decodeName(buf, start, len, flags) {
  const bytes = buf.subarray(start, start + len);
  if (flags & 0x800) return utf8Decoder.decode(bytes);
  let ascii = true;
  for (let i = 0; i < bytes.length; i++) if (bytes[i] > 0x7f) { ascii = false; break; }
  if (ascii) {
    let s = "";
    for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return s;
  }
  return utf8Decoder.decode(bytes);
}

/** Walk a zip extra field. Returns Map<id, {off, len}>. */
function parseExtra(buf, start, len) {
  const out = new Map();
  let p = start;
  const end = start + len;
  while (p + 4 <= end) {
    const id = buf[p] | (buf[p + 1] << 8);
    const size = buf[p + 2] | (buf[p + 3] << 8);
    if (p + 4 + size > end) break;
    out.set(id, { off: p + 4, len: size });
    p += 4 + size;
  }
  return out;
}

function readU64(buf, off) {
  const lo = (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
  const hi = (buf[off + 4] | (buf[off + 5] << 8) | (buf[off + 6] << 16) | (buf[off + 7] << 24)) >>> 0;
  return hi * 4294967296 + lo;
}

export class ZipEntry {
  constructor(archive, index, raw) {
    this.archive = archive;
    this.index = index;
    this.name = raw.name;
    this.flags = raw.flags;
    this.method = raw.method;
    this.crc = raw.crc;
    this.compressedSize = raw.compressedSize;
    this.size = raw.size;
    this.localOffset = raw.localOffset;
    this.dosTime = raw.dosTime;
    this.dosDate = raw.dosDate;
    this.extra = raw.extra;
    this._data = null;
  }

  get isDirectory() { return this.name.endsWith("/"); }

  /** Byte range of the entry payload inside the archive. */
  _range() {
    const buf = this.archive.buf;
    const off = this.localOffset;
    if (buf[off] !== 0x50 || buf[off + 1] !== 0x4b || buf[off + 2] !== 0x03 || buf[off + 3] !== 0x04) {
      throw new ZipError("bad local file header for " + this.name, "BAD_LOCAL_HEADER", off);
    }
    const nameLen = buf[off + 26] | (buf[off + 27] << 8);
    const extraLen = buf[off + 28] | (buf[off + 29] << 8);
    const start = off + 30 + nameLen + extraLen;
    return buf.subarray(start, start + this.compressedSize);
  }

  /** Decompress this entry (cached). */
  bytes() {
    if (this._data) return this._data;
    const raw = this._range();
    let out;
    if (this.method === METHOD_STORE) out = raw;
    else if (this.method === METHOD_DEFLATE) {
      out = this.compressedSize === 0 && this.size === 0 ? new Uint8Array(0) : this._inflate(raw);
    } else {
      throw new ZipError("unsupported compression method " + this.method + " for " + this.name, "UNSUPPORTED_METHOD");
    }
    this._data = out;
    if (this.archive.verifyCrc !== false) this._verify(out);
    return out;
  }

  /**
   * Inflate the raw payload, with any zlib failure mapped onto the taxonomy.
   *
   * Synchronous, therefore Node-only: deflateBackendSync() raises PlatformError
   * [NOT_SUPPORTED] on a host without node:zlib, and that error is deliberately
   * allowed through unchanged - it names the async alternative, whereas
   * wrapping it as INFLATE_FAILED would suggest a corrupt archive.
   */
  _inflate(raw) {
    const backend = deflateBackendSync();
    try {
      return new Uint8Array(backend.inflateRawSync(raw));
    } catch (cause) {
      throw inflateError(cause, this.name);
    }
  }

  /** Async twin of _inflate(): node:zlib when present, DecompressionStream when not. */
  async _inflateAsync(raw) {
    const backend = await deflateBackend();
    try {
      return new Uint8Array(await backend.inflateRawBytes(raw));
    } catch (cause) {
      throw inflateError(cause, this.name);
    }
  }

  /**
   * Compare the decompressed bytes against the CRC-32 the archive recorded.
   * A mismatch means the file is corrupt: every value read from this part is
   * suspect, so failing loudly is the only safe response.
   */
  _verify(out) {
    // Some writers legitimately record 0 for a streamed part; skip only then.
    if (this.crc === 0 && out.length > 0) return;
    const actual = crc32(out);
    if (actual !== this.crc) {
      throw new ZipError(
        "CRC mismatch in " + this.name + ": expected 0x" + this.crc.toString(16) +
          ", computed 0x" + actual.toString(16) + " - the archive is corrupt",
        "CRC_MISMATCH",
        { name: this.name, expected: this.crc, actual },
      );
    }
  }

  /** Decode as UTF-8 text. Synchronous: Node-only, like bytes(). */
  text() { return utf8Decoder.decode(this.bytes()); }

  /**
   * Async twins of bytes()/text() for hosts without a synchronous codec.
   *
   * They exist because the Web decompressor is asynchronous by construction, and
   * a browser still has to be able to read the small XML parts (content types,
   * workbook, relationships, styles) that are stored deflated. On Node they use
   * the same zlib code as the synchronous pair, so there is no second
   * decompression path to keep in step; the result is cached in the same slot,
   * so mixing the two spellings costs one inflate.
   */
  async bytesAsync() {
    if (this._data) return this._data;
    const raw = this._range();
    let out;
    if (this.method === METHOD_STORE) out = raw;
    else if (this.method === METHOD_DEFLATE) {
      out = this.compressedSize === 0 && this.size === 0 ? new Uint8Array(0) : await this._inflateAsync(raw);
    } else {
      throw new ZipError("unsupported compression method " + this.method + " for " + this.name, "UNSUPPORTED_METHOD");
    }
    this._data = out;
    if (this.archive.verifyCrc !== false) this._verify(out);
    return out;
  }

  /** Decode as UTF-8 text, without needing the synchronous backend. */
  async textAsync() { return utf8Decoder.decode(await this.bytesAsync()); }

  /**
   * Live inflate stream over this entry, without buffering the whole part.
   *
   * Returns a Node Readable on Node (unchanged) and a Web ReadableStream
   * everywhere else. Both are async-iterable, which is the only property the
   * consumers rely on, and both check the CRC-32 as the bytes flow.
   */
  stream() {
    const raw = this._range();
    if (this.method !== METHOD_STORE && this.method !== METHOD_DEFLATE) {
      throw new ZipError("unsupported compression method " + this.method + " for " + this.name, "UNSUPPORTED_METHOD");
    }
    const streams = nodeStreams();
    if (streams !== null) {
      const src = streams.Readable.from([raw]);
      if (this.method === METHOD_STORE) return this._verifying(src);
      return this._verifying(this._inflating(src, streams));
    }
    // No node:stream: a browser, a worker or a Deno process.
    if (this.method === METHOD_STORE) return this._verifyingWeb(webBytes(raw));
    return this._verifyingWeb(this._inflatingWeb(webInflateRawStream(raw)));
  }

  /**
   * Inflate a live payload, re-emitting zlib failures as ZipError.
   *
   * `pipe` forwards data but not errors, so the mapping has to sit on the
   * inflate side: the wrapper transform is the stream the caller sees. A clean
   * end of the payload closes the inflate stream normally and changes nothing.
   */
  _inflating(src, streams) {
    const entry = this;
    const inflate = deflateBackendSync().createInflateRaw();
    const out = new streams.Transform({
      transform(chunk, _enc, cb) { cb(null, chunk); },
    });
    inflate.on("error", (cause) => out.destroy(inflateError(cause, entry.name)));
    out.on("close", () => { inflate.destroy(); src.destroy(); });
    return src.pipe(inflate).pipe(out);
  }

  /**
   * Wrap a part stream so the bytes flowing through it are CRC-checked, and the
   * stream errors at the first sign of corruption. Streaming readers get the
   * same guarantee as buffered ones without buffering the part.
   */
  _verifying(source) {
    if (this.archive.verifyCrc === false) return source;
    const entry = this;
    const hash = new Crc32();
    let seen = 0;
    const check = new (nodeStreams().Transform)({
      transform(chunk, _enc, cb) {
        hash.update(chunk);
        seen += chunk.length;
        cb(null, chunk);
      },
      flush(cb) {
        if (entry.crc === 0 && seen > 0) return cb();
        const actual = hash.digest();
        if (actual !== entry.crc) {
          return cb(new ZipError(
            "CRC mismatch in " + entry.name + ": expected 0x" + entry.crc.toString(16) +
              ", computed 0x" + actual.toString(16) + " - the archive is corrupt",
            "CRC_MISMATCH",
            { name: entry.name, expected: entry.crc, actual },
          ));
        }
        cb();
      },
    });
    source.on("error", (e) => check.destroy(e));
    return source.pipe(check);
  }

  /**
   * Map a Web DecompressionStream failure onto the zip taxonomy.
   *
   * The Node path does this inside _inflating() by listening on the inflate
   * transform; a DecompressionStream reports through the stream itself, so the
   * mapping has to sit on the reader. Without it a corrupt part would reject
   * with a bare TypeError and break the promise that every failure raised by
   * this module is a GridxcelError (the same contract MED-4 restored for
   * node:zlib).
   */
  _inflatingWeb(source) {
    const entry = this;
    const reader = source.getReader();
    return new ReadableStream({
      async pull(controller) {
        try {
          const { done, value } = await reader.read();
          if (done) controller.close();
          else controller.enqueue(value);
        } catch (cause) {
          controller.error(inflateError(cause, entry.name));
        }
      },
      cancel(reason) {
        return reader.cancel(reason);
      },
    });
  }

  /**
   * The Web-stream twin of _verifying(): same CRC contract, no Node transform.
   *
   * It pumps the source by hand rather than using pipeThrough(new
   * TransformStream(...)) so that the only Web globals this module needs are
   * ReadableStream and the compression streams. TransformStream is standard, but
   * it is not universal - a Web-only sandbox without it would otherwise fail
   * here for a reason that has nothing to do with the archive.
   */
  _verifyingWeb(source) {
    if (this.archive.verifyCrc === false) return source;
    const entry = this;
    const reader = source.getReader();
    const hash = new Crc32();
    let seen = 0;
    return new ReadableStream({
      async pull(controller) {
        let result;
        try {
          result = await reader.read();
        } catch (cause) {
          // The source is either already mapped (an inflate failure) or a plain
          // byte stream that cannot fail this way; never let a bare TypeError out.
          controller.error(cause instanceof GridxcelError ? cause : inflateError(cause, entry.name));
          return;
        }
        if (result.done) {
          if (!(entry.crc === 0 && seen > 0)) {   // a streamed part may record no CRC
            const actual = hash.digest();
            if (actual !== entry.crc) {
              controller.error(new ZipError(
                "CRC mismatch in " + entry.name + ": expected 0x" + entry.crc.toString(16) +
                  ", computed 0x" + actual.toString(16) + " - the archive is corrupt",
                "CRC_MISMATCH",
                { name: entry.name, expected: entry.crc, actual },
              ));
              return;
            }
          }
          controller.close();
          return;
        }
        const bytes = result.value instanceof Uint8Array ? result.value : new Uint8Array(result.value);
        hash.update(bytes);
        seen += bytes.length;
        controller.enqueue(bytes);
      },
      cancel(reason) {
        return reader.cancel(reason);
      },
    });
  }
}

export class ZipArchive {
  /**
   * @param {Uint8Array} input
   * @param {{verifyCrc?: boolean}} [opts] verifyCrc (default true) checks every
   *   decompressed part against the CRC-32 recorded in the central directory.
   *   The cost is roughly 1.2 GB/s of payload, i.e. noise next to inflate, and
   *   it is the only thing standing between a flipped bit and a silently wrong
   *   cell value. Pass false only when the caller checks integrity itself.
   */
  constructor(input, opts = {}) {
    const buf = input instanceof Uint8Array ? input : new Uint8Array(input);
    this.buf = buf;
    this.verifyCrc = opts.verifyCrc !== false;
    /** @type {ZipEntry[]} */
    this.entries = [];
    /** @type {Map<string, ZipEntry>} */
    this.byName = new Map();
    this.comment = "";
    this._readDirectory();
  }

  /**
   * Read an archive from a filesystem path. Node-only: the filesystem call
   * comes from ./platform.js, so this module keeps no Node import of its own
   * and a browser bundler never has to resolve one.
   */
  static async open(path) {
    return new ZipArchive(await readFileNode(path));
  }

  get size() { return this.buf.length; }

  _readDirectory() {
    const buf = this.buf;
    const len = buf.length;
    if (len < 22) throw new ZipError("not a zip archive: too short", "NOT_ZIP");
    const minPos = Math.max(0, len - 65557);
    let eocd = -1;
    for (let p = len - 22; p >= minPos; p--) {
      if (buf[p] === 0x50 && buf[p + 1] === 0x4b && buf[p + 2] === 0x05 && buf[p + 3] === 0x06) {
        const cLen = buf[p + 20] | (buf[p + 21] << 8);
        if (p + 22 + cLen === len) { eocd = p; break; }
        if (eocd < 0) eocd = p;
      }
    }
    if (eocd < 0) throw new ZipError("not a zip archive: no end-of-central-directory record", "NOT_ZIP");

    let count = buf[eocd + 10] | (buf[eocd + 11] << 8);
    let cdSize = (buf[eocd + 12] | (buf[eocd + 13] << 8) | (buf[eocd + 14] << 16) | (buf[eocd + 15] << 24)) >>> 0;
    let cdOffset = (buf[eocd + 16] | (buf[eocd + 17] << 8) | (buf[eocd + 18] << 16) | (buf[eocd + 19] << 24)) >>> 0;
    const cmtLen = buf[eocd + 20] | (buf[eocd + 21] << 8);
    this.comment = utf8Decoder.decode(buf.subarray(eocd + 22, eocd + 22 + cmtLen));

    const loc = eocd - 20;
    if (loc >= 0 && buf[loc] === 0x50 && buf[loc + 1] === 0x4b && buf[loc + 2] === 0x06 && buf[loc + 3] === 0x07) {
      const z64off = readU64(buf, loc + 8);
      if (z64off + 56 <= len && buf[z64off] === 0x50 && buf[z64off + 1] === 0x4b &&
          buf[z64off + 2] === 0x06 && buf[z64off + 3] === 0x06) {
        count = readU64(buf, z64off + 32);
        cdSize = readU64(buf, z64off + 40);
        cdOffset = readU64(buf, z64off + 48);
      }
    }

    let p = cdOffset;
    const cdEnd = cdSize > 0 ? Math.min(len, cdOffset + cdSize) : len;
    while (p + 46 <= cdEnd && buf[p] === 0x50 && buf[p + 1] === 0x4b && buf[p + 2] === 0x01 && buf[p + 3] === 0x02) {
      const flags = buf[p + 8] | (buf[p + 9] << 8);
      const method = buf[p + 10] | (buf[p + 11] << 8);
      const dosTime = buf[p + 12] | (buf[p + 13] << 8);
      const dosDate = buf[p + 14] | (buf[p + 15] << 8);
      const crc = (buf[p + 16] | (buf[p + 17] << 8) | (buf[p + 18] << 16) | (buf[p + 19] << 24)) >>> 0;
      let csize = (buf[p + 20] | (buf[p + 21] << 8) | (buf[p + 22] << 16) | (buf[p + 23] << 24)) >>> 0;
      let usize = (buf[p + 24] | (buf[p + 25] << 8) | (buf[p + 26] << 16) | (buf[p + 27] << 24)) >>> 0;
      const nameLen = buf[p + 28] | (buf[p + 29] << 8);
      const extraLen = buf[p + 30] | (buf[p + 31] << 8);
      const partCmtLen = buf[p + 32] | (buf[p + 33] << 8);
      let localOffset = (buf[p + 42] | (buf[p + 43] << 8) | (buf[p + 44] << 16) | (buf[p + 45] << 24)) >>> 0;
      const name = decodeName(buf, p + 46, nameLen, flags);
      const extra = parseExtra(buf, p + 46 + nameLen, extraLen);

      const z64 = extra.get(0x0001);
      if (z64) {
        let q = z64.off;
        if (usize === 0xffffffff) { usize = readU64(buf, q); q += 8; }
        if (csize === 0xffffffff) { csize = readU64(buf, q); q += 8; }
        if (localOffset === 0xffffffff) { localOffset = readU64(buf, q); q += 8; }
      }

      const entry = new ZipEntry(this, this.entries.length, {
        name, flags, method, crc, compressedSize: csize, size: usize,
        localOffset, dosTime, dosDate, extra,
      });
      this.entries.push(entry);
      if (!this.byName.has(name)) this.byName.set(name, entry);
      p += 46 + nameLen + extraLen + partCmtLen;
    }

    if (count !== 0 && this.entries.length !== count) {
      this.entryCountMismatch = { declared: count, actual: this.entries.length };
    }
  }

  get(name) { return this.byName.get(name); }
  has(name) { return this.byName.has(name); }
  read(name) { const e = this.byName.get(name); return e ? e.bytes() : undefined; }
  readText(name) { const e = this.byName.get(name); return e ? e.text() : undefined; }
  names() { return this.entries.map((e) => e.name); }
  [Symbol.iterator]() { return this.entries[Symbol.iterator](); }
}

function dosDateTime(date) {
  const y = Math.max(1980, date.getFullYear());
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1);
  const d = ((y - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time: time & 0xffff, date: d & 0xffff };
}

/**
 * Build a zip archive from a list of parts.
 *
 * Sizes and CRCs are computed before each local header is emitted, so no data
 * descriptors and no seekable sink are required - the same code path drives a
 * Buffer (Node/browser) and a live HTTP response body (streaming write).
 */
export class ZipWriter {
  constructor(opts = {}) {
    this.version = opts.version === undefined ? 20 : opts.version;
    this.deterministic = opts.deterministic === true;
    this.compressionLevel = opts.compressionLevel === undefined ? 6 : opts.compressionLevel;
    this.files = [];
    this._parts = [];
    this._len = 0;
  }

  /**
   * @param {string} name
   * @param {Uint8Array|string} data
   * @param {{compress?: boolean, level?: number, date?: Date}} [opts]
   */
  add(name, data, opts = {}) {
    const raw = typeof data === "string" ? utf8Encoder.encode(data) : data;
    const level = opts.level === undefined ? this.compressionLevel : opts.level;
    const compress = opts.compress === undefined ? raw.length > 0 : opts.compress;
    let method = METHOD_STORE;
    let body = raw;
    // Deflate when a synchronous codec exists (Node), otherwise store. Storing
    // is a valid zip and keeps the synchronous writer API working in a browser;
    // the alternative - requiring an async assembler - would be a change to the
    // WorkbookWriter contract, which this module does not own.
    if (compress && level > 0 && hasSyncDeflate()) {
      const deflated = deflateBackendSync().deflateRawSync(raw, { level });
      if (deflated.length < raw.length) { body = deflated; method = METHOD_DEFLATE; }
    }
    const when = this.deterministic ? new Date(Date.UTC(1980, 0, 1)) : (opts.date || new Date());
    const dtd = dosDateTime(when);
    const crc = crc32(raw);
    const offset = this._len;
    const nameBytes = utf8Encoder.encode(name);
    const needsZip64 = raw.length >= 0xffffffff || body.length >= 0xffffffff || offset >= 0xffffffff;

    let extra = new Uint8Array(0);
    if (needsZip64) {
      extra = new Uint8Array(20);
      const dv = new DataView(extra.buffer);
      dv.setUint16(0, 0x0001, true);
      dv.setUint16(2, 16, true);
      dv.setBigUint64(4, BigInt(raw.length), true);
      dv.setBigUint64(12, BigInt(body.length), true);
    }

    const head = new Uint8Array(30 + nameBytes.length + extra.length);
    const hv = new DataView(head.buffer);
    hv.setUint32(0, SIG_LOCAL, true);
    hv.setUint16(4, this.version, true);
    hv.setUint16(6, 0x800, true);
    hv.setUint16(8, method, true);
    hv.setUint16(10, dtd.time, true);
    hv.setUint16(12, dtd.date, true);
    hv.setUint32(14, crc, true);
    hv.setUint32(18, needsZip64 ? 0xffffffff : body.length, true);
    hv.setUint32(22, needsZip64 ? 0xffffffff : raw.length, true);
    hv.setUint16(26, nameBytes.length, true);
    hv.setUint16(28, extra.length, true);
    head.set(nameBytes, 30);
    if (extra.length) head.set(extra, 30 + nameBytes.length);

    this.files.push({
      name, nameBytes, method, crc, csize: body.length, usize: raw.length,
      offset, time: dtd.time, date: dtd.date, isDir: name.endsWith("/"),
    });
    this._parts.push(head, body);
    this._len += head.length + body.length;
    return this;
  }

  build() {
    const cdParts = [];
    let cdSize = 0;
    const cdStart = this._len;
    for (const f of this.files) {
      const needsZip64 = f.usize >= 0xffffffff || f.csize >= 0xffffffff || f.offset >= 0xffffffff;
      let extra = new Uint8Array(0);
      let elen = 0;
      if (needsZip64) {
        const body = (f.usize >= 0xffffffff ? 8 : 0) + (f.csize >= 0xffffffff ? 8 : 0) + (f.offset >= 0xffffffff ? 8 : 0);
        extra = new Uint8Array(4 + body);
        const dv = new DataView(extra.buffer);
        dv.setUint16(0, 0x0001, true);
        dv.setUint16(2, body, true);
        let q = 4;
        if (f.usize >= 0xffffffff) { dv.setBigUint64(q, BigInt(f.usize), true); q += 8; }
        if (f.csize >= 0xffffffff) { dv.setBigUint64(q, BigInt(f.csize), true); q += 8; }
        if (f.offset >= 0xffffffff) { dv.setBigUint64(q, BigInt(f.offset), true); q += 8; }
        elen = extra.length;
      }
      const cd = new Uint8Array(46 + f.nameBytes.length + elen);
      const cv = new DataView(cd.buffer);
      cv.setUint32(0, SIG_CENTRAL, true);
      cv.setUint16(4, 0x031e, true);
      cv.setUint16(6, this.version, true);
      cv.setUint16(8, 0x800, true);
      cv.setUint16(10, f.method, true);
      cv.setUint16(12, f.time, true);
      cv.setUint16(14, f.date, true);
      cv.setUint32(16, f.crc, true);
      cv.setUint32(20, f.csize >= 0xffffffff ? 0xffffffff : f.csize, true);
      cv.setUint32(24, f.usize >= 0xffffffff ? 0xffffffff : f.usize, true);
      cv.setUint16(28, f.nameBytes.length, true);
      cv.setUint16(30, elen, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, f.isDir ? 0x41ed0010 : 0o100644 << 16, true);
      cv.setUint32(42, f.offset >= 0xffffffff ? 0xffffffff : f.offset, true);
      cd.set(f.nameBytes, 46);
      if (elen) cd.set(extra, 46 + f.nameBytes.length);
      cdParts.push(cd);
      cdSize += cd.length;
    }

    const count = this.files.length;
    const needsZip64 = count >= 0xffff || cdSize >= 0xffffffff || cdStart >= 0xffffffff;
    const tailLen = needsZip64 ? 56 + 20 + 22 : 22;
    const tail = new Uint8Array(tailLen);
    const tv = new DataView(tail.buffer);
    let tp = 0;
    if (needsZip64) {
      tv.setUint32(tp, SIG_EOCD64, true);
      tv.setBigUint64(tp + 4, 44n, true);
      tv.setUint16(tp + 12, 45, true);
      tv.setUint16(tp + 14, 45, true);
      tv.setUint32(tp + 16, 0, true);
      tv.setUint32(tp + 20, 0, true);
      tv.setBigUint64(tp + 24, BigInt(count), true);
      tv.setBigUint64(tp + 32, BigInt(count), true);
      tv.setBigUint64(tp + 40, BigInt(cdSize), true);
      tv.setBigUint64(tp + 48, BigInt(cdStart), true);
      tp += 56;
      tv.setUint32(tp, SIG_EOCD64_LOC, true);
      tv.setUint32(tp + 4, 0, true);
      tv.setBigUint64(tp + 8, BigInt(cdStart + cdSize), true);
      tv.setUint32(tp + 16, 1, true);
      tp += 20;
      tv.setUint32(tp, SIG_EOCD, true);
      tv.setUint16(tp + 4, 0, true);
      tv.setUint16(tp + 6, 0, true);
      tv.setUint16(tp + 8, 0xffff, true);
      tv.setUint16(tp + 10, 0xffff, true);
      tv.setUint32(tp + 12, 0xffffffff, true);
      tv.setUint32(tp + 16, 0xffffffff, true);
      tv.setUint16(tp + 20, 0, true);
    } else {
      tv.setUint32(tp, SIG_EOCD, true);
      tv.setUint16(tp + 4, 0, true);
      tv.setUint16(tp + 6, 0, true);
      tv.setUint16(tp + 8, count, true);
      tv.setUint16(tp + 10, count, true);
      tv.setUint32(tp + 12, cdSize, true);
      tv.setUint32(tp + 16, cdStart, true);
      tv.setUint16(tp + 20, 0, true);
    }

    const total = this._len + cdSize + tail.length;
    const out = new Uint8Array(total);
    let o = 0;
    for (const p of this._parts) { out.set(p, o); o += p.length; }
    for (const p of cdParts) { out.set(p, o); o += p.length; }
    out.set(tail, o);
    return out;
  }

  /** The archive as a Node Buffer when the host has one, else a Uint8Array. */
  toBuffer() { return asBuffer(this.build()); }
}

/** Build an archive from [name, data] pairs or a plain object. */
export function createZip(files, opts = {}) {
  const w = new ZipWriter(opts);
  const it = files instanceof Map || Array.isArray(files) ? files : Object.entries(files);
  for (const [name, data] of it) w.add(name, data, opts);
  return w.build();
}
