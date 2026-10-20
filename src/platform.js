/**
 * platform.js - the Node platform for the byte-level core.
 *
 * STRUCTURE
 *   src/platform.browser.js is the Web baseline: it holds every host-neutral
 *   piece (the CompressionStream backend, the stream helpers, the base64/latin1
 *   codecs, PlatformError) and references no node: module at all.
 *
 *   This file imports that baseline, re-exports it, and overrides the handful of
 *   capabilities Node improves on: node:zlib for synchronous and faster
 *   compression, node:stream for the streaming API, the Node Buffer for the
 *   SheetJS "buffer" output type, and node:fs/node:url for the filesystem
 *   helpers. Both files export the same surface, which the test suite asserts.
 *
 *   A bundler targeting the browser swaps this file for the baseline through
 *   package.json:
 *
 *     "browser": { "./src/platform.js": "./src/platform.browser.js" }
 *
 *   so the node: imports below are never resolved, let alone evaluated, in a
 *   browser bundle. That is why they can be plain static imports here, and why
 *   no computed "node:" + "zlib" specifier exists anywhere: a computed specifier
 *   is invisible to bundlers, so nothing can verify it.
 *
 * WHY THE SYNC PATH IS NODE-ONLY
 *   The Web compression API is asynchronous by construction (it returns a
 *   stream), so a host without node:zlib has no synchronous codec. On such a host
 *   the synchronous entry points raise PlatformError [NOT_SUPPORTED] naming the
 *   async alternative; they never throw a ReferenceError from a missing import
 *   and never return wrong bytes.
 *
 * A NEW PLATFORM
 *   Provide either a synchronous deflate pair (then everything works, including
 *   the sync APIs) or the Web streams pair (then the async and streaming APIs
 *   work and ZipWriter stores parts instead of deflating them - a stored zip is
 *   valid).
 */

import { deflateRawSync, inflateRawSync, createInflateRaw, createDeflateRaw } from "node:zlib";
import { Readable, Transform } from "node:stream";
import { readFileSync, writeFileSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath as nodeFileURLToPath } from "node:url";

// Everything the baseline provides: PlatformError, the Web stream helpers, the
// base64/latin1 codecs and the capability probes this file does not override.
export * from "./platform.browser.js";

/* --------------------------------------------------------------- detection */

/** True: this file only ever runs in Node. */
export function isNodePlatform() {
  return true;
}

/** Node always has a synchronous codec. */
export function hasSyncDeflate() {
  return true;
}

/** The Node stream classes, for the streaming API. */
export function nodeStreams() {
  return { Readable, Transform };
}

/* ------------------------------------------------------- deflate backends */

/**
 * The synchronous backend.
 *
 * @returns {{kind: "node", inflateRawSync: Function, deflateRawSync: Function,
 *            createInflateRaw: Function, createDeflateRaw: Function}}
 */
export function deflateBackendSync() {
  return {
    kind: "node",
    inflateRawSync,
    deflateRawSync,
    createInflateRaw,
    createDeflateRaw,
  };
}

/**
 * The backend for everything that can await. Node always wins here: zlib is
 * faster than the Web streams and it is already loaded.
 *
 * @returns {Promise<{kind: "node", inflateRawSync: Function, deflateRawSync: Function,
 *   inflateRawBytes: Function, deflateRawBytes: Function,
 *   createInflateRaw: Function, createDeflateRaw: Function}>}
 */
export async function deflateBackend() {
  return {
    kind: "node",
    inflateRawSync,
    deflateRawSync,
    createInflateRaw,
    createDeflateRaw,
    inflateRawBytes: async (bytes) => new Uint8Array(inflateRawSync(bytes)),
    deflateRawBytes: async (bytes) => new Uint8Array(deflateRawSync(bytes)),
  };
}

/* ------------------------------------------------------------ text codecs */

/**
 * A real Node Buffer for the SheetJS "buffer" output type. The baseline returns
 * the bytes unchanged; this override is what keeps type:"buffer" identical on
 * Node.
 */
export function asBuffer(bytes) {
  return Buffer.from(bytes);
}

/* ------------------------------------------------------------- filesystem */

/** Read a whole file. Node-only, like every filesystem helper. */
export async function readFileNode(path) {
  return new Uint8Array(await readFile(path));
}

/** Read a whole file synchronously. Node-only. */
export function readFileNodeSync(path) {
  return new Uint8Array(readFileSync(path));
}

/**
 * A path write. Errors keep their fs code (EEXIST and friends), so write.js can
 * still turn EEXIST into its own FILE_EXISTS error.
 */
export async function writeFileNode(path, bytes, opts) {
  return writeFile(path, bytes, opts);
}

/** The synchronous twin used by writeWorkbookSync({ path }). */
export function writeFileNodeSync(path, bytes, opts) {
  return writeFileSync(path, bytes, opts);
}

/** file: URL -> filesystem path. */
export function fileUrlToPath(url) {
  return nodeFileURLToPath(url);
}
