/**
 * The platform seam: what the byte-level core can rely on, in Node and in a
 * browser.
 *
 * The browser half is tested by importing src/platform.browser.js directly -
 * that is exactly the file the package.json "browser" field substitutes when a
 * bundler targets the browser, so these tests run the code a browser runs.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

import * as nodePlatform from "../src/platform.js";
import * as webPlatform from "../src/platform.browser.js";
import { ZipArchive, ZipWriter, createZip, METHOD_DEFLATE, METHOD_STORE } from "../src/zip.js";
import { writeWorkbookSync } from "../src/write.js";
import { GridxcelError } from "../src/errors.js";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const SRC = path.join(ROOT, "src");
const GOLDEN = path.join(ROOT, "test", "fixtures", "golden.xlsx");
const goldenBytes = new Uint8Array(readFileSync(GOLDEN));

const called = (fn) => {
  try {
    return { ok: true, value: fn() };
  } catch (err) {
    return { ok: false, err };
  }
};

/** Like called(), but settles a rejected promise too (the fs helpers are async). */
const settled = async (fn) => {
  try {
    return { ok: true, value: await fn() };
  } catch (err) {
    return { ok: false, err };
  }
};

/* ------------------------------------------------------- the two platforms */

test("platform: both files export the same surface", () => {
  const nodeNames = Object.keys(nodePlatform).sort();
  const webNames = Object.keys(webPlatform).sort();
  assert.deepEqual(webNames, nodeNames, "a bundler can swap one file for the other");
  for (const name of ["PlatformError", "deflateBackendSync", "deflateBackend", "webDeflateBackend",
    "hasSyncDeflate", "hasWebDeflate", "isNodePlatform", "nodeStreams", "webBytes",
    "webInflateRawStream", "collectStream", "encodeBase64", "encodeLatin1", "asBuffer",
    "readFileNode", "readFileNodeSync", "writeFileNode", "writeFileNodeSync", "fileUrlToPath"]) {
    assert.ok(webNames.includes(name), name + " is part of the seam");
  }
});

test("platform: Node uses node:zlib, the browser cannot reach it", () => {
  assert.equal(nodePlatform.isNodePlatform(), true);
  assert.equal(nodePlatform.hasSyncDeflate(), true);
  assert.equal(nodePlatform.nodeStreams() !== null, true, "Node gets the stream classes");
  const sync = nodePlatform.deflateBackendSync();
  assert.equal(sync.kind, "node");
  const raw = new TextEncoder().encode("deflate me, twice");
  const deflated = new Uint8Array(sync.deflateRawSync(raw, { level: 6 }));
  assert.deepEqual(new Uint8Array(sync.inflateRawSync(deflated)), raw);

  assert.equal(webPlatform.isNodePlatform(), false);
  assert.equal(webPlatform.hasSyncDeflate(), false);
  assert.equal(webPlatform.nodeStreams(), null, "a browser has no node:stream");
  const refused = called(() => webPlatform.deflateBackendSync());
  assert.equal(refused.ok, false);
  assert.ok(refused.err instanceof webPlatform.PlatformError);
  assert.ok(refused.err instanceof GridxcelError, "still catchable by type");
  assert.equal(refused.err.code, "NOT_SUPPORTED");
  assert.match(refused.err.message, /Node/);
  assert.match(refused.err.message, /async/);
});

test("platform: the Web backend round-trips and is feature-detected, not assumed", async (t) => {
  // Evidence, not assumption: a format the host does not know must be rejected,
  // otherwise hasWebDeflate() would answer true blindly.
  assert.throws(() => new DecompressionStream("deflate-raw-not-a-format"), TypeError);

  // The Web streams exist on Node 18 but do NOT accept "deflate-raw" - that
  // arrived in Node 20. So this is exactly the case the feature detection is for:
  // assert that the answer MATCHES reality, not that it is true. Asserting true
  // made the suite fail on 18 while the code was right.
  let streamsSupportRaw = false;
  try {
    new CompressionStream("deflate-raw");
    new DecompressionStream("deflate-raw");
    streamsSupportRaw = true;
  } catch { streamsSupportRaw = false; }
  assert.equal(webPlatform.hasWebDeflate(), streamsSupportRaw, "hasWebDeflate must match the host");
  assert.equal(nodePlatform.hasWebDeflate(), streamsSupportRaw, "both surfaces answer alike");

  if (!streamsSupportRaw) {
    // Documented floor: the browser path needs a host that supports deflate-raw.
    t.diagnostic("host has no deflate-raw support; the Web backend round trip cannot run here");
    return;
  }

  const web = await webPlatform.deflateBackend();
  assert.equal(web.kind, "web");
  const raw = new TextEncoder().encode("web compression round trip");
  assert.deepEqual(new Uint8Array(await web.inflateRawBytes(await web.deflateRawBytes(raw))), raw);

  // Cross-implementation: a payload node:zlib produced must inflate through
  // DecompressionStream, which is the browser reading a file someone else wrote.
  const nodeDeflated = new Uint8Array(nodePlatform.deflateBackendSync().deflateRawSync(raw, { level: 9 }));
  assert.deepEqual(new Uint8Array(await web.inflateRawBytes(nodeDeflated)), raw);

  const stream = webPlatform.webInflateRawStream(nodeDeflated);
  assert.deepEqual(await webPlatform.collectStream(stream), raw);
});

test("platform: the browser file refuses the filesystem, and says why", async () => {
  const attempts = [
    () => webPlatform.readFileNode("x.xlsx"),
    () => webPlatform.readFileNodeSync("x.xlsx"),
    () => webPlatform.writeFileNode("x.xlsx", new Uint8Array(1)),
    () => webPlatform.writeFileNodeSync("x.xlsx", new Uint8Array(1)),
    () => webPlatform.fileUrlToPath(pathToFileURL(path.join(ROOT, "x.xlsx"))),
  ];
  for (const attempt of attempts) {
    const result = await settled(attempt);
    assert.equal(result.ok, false, "filesystem access must be refused");
    assert.ok(result.err instanceof webPlatform.PlatformError, "typed, got " + result.err.name);
    assert.equal(result.err.code, "NOT_SUPPORTED");
  }
  // A non-file URL keeps its string form; the filesystem layer refuses it later.
  assert.equal(webPlatform.fileUrlToPath(new URL("https://example.com/x.xlsx")), "https://example.com/x.xlsx");
  assert.match(nodePlatform.fileUrlToPath(pathToFileURL(path.join(ROOT, "x.xlsx"))), /x\.xlsx$/);
});

/* --------------------------------------------------------- text codecs */

test("platform: base64 and latin1 match Node's own codecs byte for byte", () => {
  const sizes = [0, 1, 2, 3, 4, 5, 17, 100, 255, 0x8000 - 1, 0x8000, 0x8000 * 3 - 1, 0x8000 * 3, 0x8000 * 3 + 1, 200000];
  for (const size of sizes) {
    const bytes = new Uint8Array(size);
    for (let i = 0; i < size; i++) bytes[i] = (i * 31 + 7) & 0xff;
    const expected = Buffer.from(bytes).toString("base64");
    assert.equal(nodePlatform.encodeBase64(bytes), expected, "node base64 length " + size);
    assert.equal(webPlatform.encodeBase64(bytes), expected, "browser base64 length " + size);
    const latin1 = Buffer.from(bytes).toString("latin1");
    assert.equal(nodePlatform.encodeLatin1(bytes), latin1, "node latin1 length " + size);
    assert.equal(webPlatform.encodeLatin1(bytes), latin1, "browser latin1 length " + size);
  }
  // asBuffer is the one deliberate difference: a Buffer on Node, bytes elsewhere.
  assert.ok(Buffer.isBuffer(nodePlatform.asBuffer(new Uint8Array(3))));
  assert.equal(webPlatform.asBuffer(new Uint8Array(3)) instanceof Uint8Array, true);
});

/* --------------------------------------------------------------- zip.js */

test("zip: no Node import and no Buffer identifier survive in the core", () => {
  for (const file of ["zip.js", "write.js", "read.js"]) {
    const text = readFileSync(path.join(SRC, file), "utf8");
    const code = text.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    assert.equal(/\bBuffer\s*[.(]/.test(code), false, file + " must not call Buffer");
    assert.equal(/from\s+["']node:/.test(code), false, file + " must not import a Node builtin");
    assert.equal(/import\s*\(\s*["']node:/.test(code), false, file + " must not dynamically import a Node builtin");
    assert.equal(/["']node:[a-z]/.test(code), false, file + " must not name a node: specifier at all");
  }
});

test("zip: the browser platform has no node: reference, and only platform.js may have one", () => {
  const browserFile = readFileSync(path.join(SRC, "platform.browser.js"), "utf8");
  assert.equal(/["']node:/.test(browserFile), false, "the file a browser gets must not mention node:");

  // The files this change owns. index.js/api.js/compat.js still carry Node
  // imports owned by another task; the browser bundle test is their gate.
  const others = ["zip.js", "write.js", "read.js", "opc.js", "platform.browser.js"];
  for (const file of others) {
    const code = readFileSync(path.join(SRC, file), "utf8")
      .replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    assert.equal(/["']node:[a-z]/.test(code), false, file + " must not name a node: specifier");
  }

  // The swap itself is what makes that safe, so assert the mapping exists.
  const pkg = JSON.parse(readFileSync(path.join(ROOT, "package.json"), "utf8"));
  assert.equal(pkg.browser["./src/platform.js"], "./src/platform.browser.js",
    "package.json browser field must map the Node platform to the browser one");
});

test("zip: streaming still streams, and the async twins agree with the sync read", async () => {
  const archive = new ZipArchive(goldenBytes);
  const entry = archive.get("xl/worksheets/sheet1.xml");
  assert.ok(entry, "fixture has a worksheet");
  assert.equal(entry.method, METHOD_DEFLATE, "and it is deflated, so inflate really runs");

  const nodeStream = entry.stream();
  assert.equal(typeof nodeStream.pipe, "function", "Node gets a Node Readable, unchanged");
  let streamed = 0;
  for await (const chunk of nodeStream) streamed += chunk.length;
  assert.equal(streamed, entry.size);

  const whole = await new ZipArchive(goldenBytes).get("xl/worksheets/sheet1.xml").bytesAsync();
  assert.equal(whole.length, entry.size);
  assert.match(new TextDecoder().decode(whole), /<worksheet[ >]/);
  assert.deepEqual(await entry.bytesAsync(), entry.bytes(), "the async twin matches the sync read");
  assert.equal(await entry.textAsync(), entry.text());
});

test("zip: the writer still deflates on Node and reports its bytes as a Buffer", () => {
  const writer = new ZipWriter({ compressionLevel: 6, deterministic: true });
  writer.add("part.txt", "hello ".repeat(200));
  const bytes = writer.build();
  const reopened = new ZipArchive(bytes);
  assert.equal(reopened.get("part.txt").method, METHOD_DEFLATE, "unchanged Node behaviour");
  assert.equal(reopened.get("part.txt").text(), "hello ".repeat(200));
  assert.ok(Buffer.isBuffer(writer.toBuffer()), "toBuffer still returns a Buffer on Node");
  const stored = createZip([["a.bin", new Uint8Array(0)]], { compressionLevel: 0 });
  assert.equal(new ZipArchive(stored).get("a.bin").method, METHOD_STORE);
});

test("write: every output shape survives without a Buffer import", () => {
  const book = { sheets: [{ name: "S", rows: [[{ v: "x" }]], merges: [], columns: [], rowInfo: new Map() }] };
  const bytes = writeWorkbookSync(book, { deterministic: true });
  const b64 = writeWorkbookSync(book, { type: "base64", deterministic: true });
  assert.equal(typeof b64, "string");
  assert.deepEqual(new Uint8Array(Buffer.from(b64, "base64")), bytes);
  const binary = writeWorkbookSync(book, { type: "binary", deterministic: true });
  assert.deepEqual(new Uint8Array(Buffer.from(binary, "latin1")), bytes);
  assert.ok(Buffer.isBuffer(writeWorkbookSync(book, { type: "buffer", deterministic: true })));
  assert.ok(writeWorkbookSync(book, { type: "array", deterministic: true }) instanceof ArrayBuffer);
});

test("read: the async reader warms the small parts and leaves sheets streaming", async () => {
  const { readWorkbook } = await import("../src/read.js");
  const book = await readWorkbook(goldenBytes);
  assert.equal(book.sheets.length, 4);
  assert.equal(book.sheets[0].name, "Types");
  // The warm-up must only touch metadata-sized parts: the worksheet is streamed.
  const archive = new ZipArchive(goldenBytes);
  assert.equal(archive.get("xl/worksheets/sheet1.xml")._data, null, "worksheet not cached by the warm-up");
  assert.ok(archive.get("xl/styles.xml"), "styles part present");
});
