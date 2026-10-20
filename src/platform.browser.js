/**
 * platform.browser.js - the Web-platform baseline for the byte-level core.
 *
 * This file is what a browser bundle gets. package.json maps
 *
 *   "browser": { "./src/platform.js": "./src/platform.browser.js" }
 *
 * so a bundler targeting the browser (esbuild, Vite, webpack, Rollup all honour
 * the browser field) swaps this file in for src/platform.js. That is why this
 * file must contain no node: reference of any kind, not even a dynamic one:
 * esbuild resolves LITERAL dynamic imports at build time and fails the bundle
 * on an unresolvable specifier, and a computed one would be invisible to every
 * bundler and therefore unverifiable.
 *
 * STRUCTURE
 *   platform.browser.js  the baseline: Web APIs only. src/platform.js imports
 *                        it and adds Node capabilities on top, so the shared
 *                        codecs and stream helpers exist exactly once.
 *   platform.js          Node: node:zlib, node:stream, node:fs, node:url.
 *
 * WHAT A HOST HAS TO PROVIDE
 *   CompressionStream and DecompressionStream accepting "deflate-raw", plus
 *   ReadableStream/TransformStream, TextEncoder/TextDecoder and btoa/atob.
 *   Everything degrades to a typed PlatformError [NOT_SUPPORTED] when a
 *   capability is missing - never a ReferenceError, never wrong bytes.
 *
 * WHAT IS NODE-ONLY
 *   Anything synchronous. The Web compression API is asynchronous by
 *   construction (it returns a stream), so inflateRawSync/deflateRawSync cannot
 *   exist here; the synchronous entry points (readWorkbookSync,
 *   writeWorkbookSync, ZipArchive.bytes(), ZipEntry.text()) raise
 *   PlatformError [NOT_SUPPORTED] and name the async alternative. Filesystem
 *   access is Node-only for the obvious reason.
 */

import { PlatformError } from "./errors.js";

// Re-exported so this module keeps its documented surface: callers importing
// from the platform module directly still get it.
export { PlatformError };


function syncRequired(what) {
  return new PlatformError(
    what + " needs a synchronous deflate backend, which only Node provides. The Web " +
      "CompressionStream API is asynchronous only, so use an async entry point instead - " +
      "read/readWorkbook, write/writeWorkbook, eachRow or ZipEntry.stream() - which work " +
      "on every platform.");
}

/* --------------------------------------------------------------- detection */

/** False here by definition; src/platform.js overrides this to true. */
export function isNodePlatform() {
  return false;
}

/** No synchronous codec exists on the Web platform. */
export function hasSyncDeflate() {
  return false;
}

/** @throws {PlatformError} always, in this file. */
export function deflateBackendSync() {
  throw syncRequired("this call");
}

/** The Node stream classes; a browser has none. */
export function nodeStreams() {
  return null;
}

/* ---------------------------------------------------------- Web primitives */

/** A one-chunk ReadableStream over a byte array. */
export function webBytes(bytes) {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(bytes);
      controller.close();
    },
  });
}

/** Read a Web ReadableStream to the end. */
export async function collectStream(stream) {
  const chunks = [];
  let total = 0;
  const reader = stream.getReader();
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    const chunk = value instanceof Uint8Array ? value : new Uint8Array(value);
    chunks.push(chunk);
    total += chunk.length;
  }
  if (chunks.length === 1) return chunks[0];
  const out = new Uint8Array(total);
  let at = 0;
  for (const chunk of chunks) { out.set(chunk, at); at += chunk.length; }
  return out;
}

/**
 * Feature-detect the format rather than assume it: a host that rejects
 * "deflate-raw" would otherwise hand back a backend that exists and throws on
 * first use. Constructing both streams is cheap and does exactly that check.
 */
function assertRawFormat(format) {
  try {
    new CompressionStream(format);
    new DecompressionStream(format);
  } catch (cause) {
    throw new PlatformError(
      "this host does not support the Web " + format + " compression format, so only Node " +
        "(node:zlib) can deflate raw streams here: " + (cause && cause.message ? cause.message : cause));
  }
}

/** True when the async Web backend is usable here (deflate-raw verified). */
export function hasWebDeflate() {
  if (typeof DecompressionStream !== "function" || typeof CompressionStream !== "function") return false;
  try {
    assertRawFormat("deflate-raw");
    return true;
  } catch {
    return false;
  }
}

/**
 * @returns {{kind: "web",
 *   inflateRawBytes: (bytes: Uint8Array) => Promise<Uint8Array>,
 *   deflateRawBytes: (bytes: Uint8Array, opts?: {level?: number}) => Promise<Uint8Array>}}
 */
export function webDeflateBackend() {
  if (typeof DecompressionStream !== "function" || typeof CompressionStream !== "function") {
    throw new PlatformError(
      "this host provides neither node:zlib nor the Web CompressionStream API, so no " +
        "deflate backend is available; a modern browser, a worker or Node 18+ provides one.");
  }
  assertRawFormat("deflate-raw");
  return {
    kind: "web",
    inflateRawBytes: (bytes) => collectStream(webBytes(bytes).pipeThrough(new DecompressionStream("deflate-raw"))),
    deflateRawBytes: (bytes) => collectStream(webBytes(bytes).pipeThrough(new CompressionStream("deflate-raw"))),
  };
}

/**
 * Inflate a raw payload into a Web ReadableStream without buffering it. Used by
 * ZipEntry.stream() when there is no node:zlib; the caller maps failures onto
 * the zip error taxonomy.
 */
export function webInflateRawStream(bytes) {
  if (typeof DecompressionStream !== "function") {
    throw new PlatformError(
      "streaming inflate needs either node:zlib or the Web DecompressionStream API, and this " +
        "host has neither.");
  }
  assertRawFormat("deflate-raw");
  return webBytes(bytes).pipeThrough(new DecompressionStream("deflate-raw"));
}

/**
 * The backend for everything that can await: on the Web platform this is the
 * CompressionStream pair. src/platform.js overrides it to prefer node:zlib.
 */
export async function deflateBackend() {
  return webDeflateBackend();
}

/* ------------------------------------------------------------ text codecs */

/**
 * base64 WITHOUT Buffer.
 *
 * Choice: btoa over a latin1 string, in 3-byte-aligned blocks. btoa is the
 * base64 primitive every target guarantees (Node 16+, browsers, workers, Deno),
 * it needs no dependency, and the alternatives are worse here: Buffer would
 * reintroduce the Node coupling this module exists to avoid, and a hand-rolled
 * 64-character table is around a hundred lines that have to be trusted for
 * every byte. The 3-byte alignment is what makes concatenating per-block output
 * valid base64 - without it an interior block's padding would corrupt the
 * stream.
 */
export function encodeBase64(bytes) {
  if (typeof btoa !== "function") {
    throw new PlatformError("base64 output needs btoa, which this host does not provide.");
  }
  if (bytes.length === 0) return "";
  const BLOCK = 0x8000 * 3;           // multiple of 3: every block encodes independently
  let out = "";
  for (let i = 0; i < bytes.length; i += BLOCK) {
    out += btoa(encodeLatin1(bytes.subarray(i, Math.min(i + BLOCK, bytes.length))));
  }
  return out;
}

/**
 * latin1 ("binary string") WITHOUT Buffer: one char per byte, which is exactly
 * what Buffer#toString("latin1") produced and what SheetJS means by "binary".
 * Chunked because String.fromCharCode.apply has an argument limit.
 */
export function encodeLatin1(bytes) {
  const CHUNK = 0x8000;
  if (bytes.length <= CHUNK) return String.fromCharCode.apply(null, bytes);
  let out = "";
  for (let i = 0; i < bytes.length; i += CHUNK) {
    out += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(i + CHUNK, bytes.length)));
  }
  return out;
}

/**
 * A Node Buffer when the host has one, otherwise the bytes unchanged. The
 * SheetJS "buffer" output type means a Node Buffer; in a browser that type
 * cannot exist and a Uint8Array is the closest thing a consumer can use.
 * src/platform.js overrides this to return a real Buffer.
 */
export function asBuffer(bytes) {
  return bytes;
}

/**
 * base64 -> bytes, the inverse of encodeBase64. `atob` is the primitive every
 * target provides (Node 16+, browsers, workers, Deno); it yields a latin1
 * string, so one charCodeAt per character is the whole decoder. Throws
 * PlatformError rather than a bare URIError when the host or the input is bad,
 * so the failure stays inside the library's error taxonomy.
 */
export function decodeBase64(text) {
  if (typeof atob !== "function") {
    throw new PlatformError("base64 input needs atob, which this host does not provide.");
  }
  let binary;
  try {
    binary = atob(String(text));
  } catch (cause) {
    throw new PlatformError("input is not valid base64: " + (cause && cause.message ? cause.message : cause));
  }
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

/**
 * latin1 ("binary string") -> bytes, the inverse of encodeLatin1: one char per
 * byte, masked to 8 bits so a character above U+00FF lands on the byte
 * Buffer#from(str, "latin1") would have produced instead of a NaN.
 */
export function decodeLatin1(text) {
  const s = String(text);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

/* ------------------------------------------------------------- filesystem */

function noFilesystem(what) {
  return new PlatformError(
    what + " needs a filesystem, which a browser does not have; write to memory instead " +
      "(writeWorkbook without options.path returns the bytes) or run in Node.");
}

/** Node-only. @throws {PlatformError} always, in this file. */
export async function readFileNode(path) {
  throw noFilesystem("reading " + JSON.stringify(String(path)));
}

/** Node-only. @throws {PlatformError} always, in this file. */
export function readFileNodeSync(path) {
  throw noFilesystem("reading " + JSON.stringify(String(path)));
}

/** Node-only. @throws {PlatformError} always, in this file. */
export async function writeFileNode(path) {
  throw noFilesystem("writing to " + JSON.stringify(String(path)));
}

/** Node-only. @throws {PlatformError} always, in this file. */
export function writeFileNodeSync(path) {
  throw noFilesystem("writing to " + JSON.stringify(String(path)));
}

/**
 * A file: URL names a local file, so converting one is Node-only work. Any
 * other URL keeps its string form, which the caller then hands to the
 * filesystem layer and gets the same typed refusal for.
 */
export function fileUrlToPath(url) {
  const href = String(url);
  if (!href.startsWith("file:")) return href;
  throw noFilesystem("reading " + href);
}
