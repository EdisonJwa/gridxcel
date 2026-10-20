/**
 * Inflate failures must stay inside the typed error contract (MED-4 / LOW-6).
 *
 * A corrupt deflate stream used to surface as node:zlib's own Error, so a caller
 * doing `catch (e) { if (e instanceof GridxcelError) return 400; throw e; }`
 * crashed on a malformed upload instead of rejecting it. Every case below threw
 * a raw Z_DATA_ERROR or Z_BUF_ERROR before the fix; they are kept because the
 * failure mode is attacker-controlled bytes reaching a path that a service
 * cannot catch by type.
 *
 * The contract pinned here:
 *   - buffered and streaming inflate both raise ZipError [INFLATE_FAILED];
 *   - the message names the part and the original zlib error stays on `cause`,
 *     which is what still separates corruption (Z_DATA_ERROR) from truncation
 *     (Z_BUF_ERROR);
 *   - a valid part still reads as data, and CRC_MISMATCH is unchanged.
 */

import { test } from "node:test";
import assert from "node:assert";
import { createZip, ZipArchive, ZipWriter } from "../src/zip.js";
import { readWorkbook, readWorkbookSync, eachRow } from "../src/read.js";
import { GridxcelError } from "../src/errors.js";

const CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  + '<Default Extension="xml" ContentType="application/xml"/></Types>';
const RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
const WB = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>';
const WBRELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
const SHEET = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  + '<sheetData><row r="1"><c r="A1"><v>777</v></c></row></sheetData></worksheet>';
const PART = "xl/worksheets/sheet1.xml";
const STORED = "stored.bin";

/** A small, fully valid workbook whose worksheet part is deflated. */
function baseArchive() {
  return createZip([
    ["[Content_Types].xml", CT],
    ["_rels/.rels", RELS],
    ["xl/workbook.xml", WB],
    ["xl/_rels/workbook.xml.rels", WBRELS],
    [PART, SHEET],
  ], { compressionLevel: 9 });
}

/** Locate a local file header by name and report the fields the tests patch. */
function localHeader(bytes, name) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let i = 0; i + 30 <= bytes.length; i++) {
    if (bytes[i] !== 0x50 || bytes[i + 1] !== 0x4b || bytes[i + 2] !== 0x03 || bytes[i + 3] !== 0x04) continue;
    const nameLen = dv.getUint16(i + 26, true);
    if (nameLen !== name.length) continue;
    let same = true;
    for (let k = 0; k < name.length; k++) if (bytes[i + 30 + k] !== name.charCodeAt(k)) { same = false; break; }
    if (!same) continue;
    const extraLen = dv.getUint16(i + 28, true);
    return {
      off: i,
      nameLen,
      extraLen,
      dataStart: i + 30 + nameLen + extraLen,
      compressedSize: dv.getUint32(i + 18, true),
    };
  }
  throw new Error("no local header for " + name);
}

/** Set the declared compressed size in both the local header and the directory. */
function patchCompressedSize(bytes, name, size) {
  const out = Uint8Array.from(bytes);
  const dv = new DataView(out.buffer, out.byteOffset, out.byteLength);
  dv.setUint32(localHeader(out, name).off + 18, size, true);
  for (let i = 0; i + 46 <= out.length; i++) {
    if (out[i] !== 0x50 || out[i + 1] !== 0x4b || out[i + 2] !== 0x01 || out[i + 3] !== 0x02) continue;
    const nameLen = dv.getUint16(i + 28, true);
    if (nameLen !== name.length) continue;
    let same = true;
    for (let k = 0; k < name.length; k++) if (out[i + 46 + k] !== name.charCodeAt(k)) { same = false; break; }
    if (same) { dv.setUint32(i + 20, size, true); break; }
  }
  return out;
}

const BASE = baseArchive();
const LOC = localHeader(BASE, PART);
const MID = LOC.dataStart + (LOC.compressedSize >> 1);
const LAST = LOC.dataStart + LOC.compressedSize - 1;

/**
 * The five corruption variants the security review reproduced (MED-4), plus the
 * bogus local-header lengths from LOW-6. Each returns a fresh archive.
 */
const CORRUPTIONS = {
  "whole deflate stream zeroed": () => {
    const c = Uint8Array.from(BASE);
    c.fill(0, LOC.dataStart, LOC.dataStart + LOC.compressedSize);
    return c;
  },
  "second half zeroed": () => {
    const c = Uint8Array.from(BASE);
    c.fill(0, MID, LOC.dataStart + LOC.compressedSize);
    return c;
  },
  "first byte flipped": () => {
    const c = Uint8Array.from(BASE);
    c[LOC.dataStart] ^= 0xff;
    return c;
  },
  "last byte flipped": () => {
    const c = Uint8Array.from(BASE);
    c[LAST] ^= 0xff;
    return c;
  },
  "stream cut mid-block with the declared size adjusted": () =>
    patchCompressedSize(BASE, PART, LOC.compressedSize >> 1),
  "local header name and extra length 0xFFFF": () => {
    const c = Uint8Array.from(BASE);
    const dv = new DataView(c.buffer, c.byteOffset, c.byteLength);
    dv.setUint16(LOC.off + 26, 0xffff, true);
    dv.setUint16(LOC.off + 28, 0xffff, true);
    return c;
  },
};

/** Drain a stream, resolving with the number of bytes it produced. */
function drain(stream) {
  return new Promise((resolve, reject) => {
    let seen = 0;
    stream.on("data", (chunk) => { seen += chunk.length; });
    stream.on("end", () => resolve(seen));
    stream.on("error", reject);
  });
}

/** Assert the shape every inflate failure must have, and return true for assert. */
function isInflateFailure(err, label) {
  assert.ok(err instanceof GridxcelError, label + ": " + err.name + " [" + (err.code || "-") + "] is not a GridxcelError");
  assert.equal(err.code, "INFLATE_FAILED", label + ": code is " + err.code);
  assert.ok(err.message.includes(PART), label + ": the message must name the part, got " + JSON.stringify(err.message));
  assert.equal(err.details && err.details.part, PART, label + ": details.part must name the part");
  assert.ok(err.cause, label + ": the original zlib error must stay reachable as cause");
  assert.match(String(err.cause.code), /^Z_/, label + ": cause should be the zlib error, got " + (err.cause && err.cause.code));
  return true;
}

test("a corrupt deflate stream raises a typed error when the part is buffered", () => {
  for (const [label, make] of Object.entries(CORRUPTIONS)) {
    const entry = new ZipArchive(make()).get(PART);
    assert.throws(() => entry.bytes(), (err) => isInflateFailure(err, label), label);
  }
});

test("a corrupt deflate stream raises a typed error when the part is streamed", async () => {
  for (const [label, make] of Object.entries(CORRUPTIONS)) {
    const entry = new ZipArchive(make()).get(PART);
    await assert.rejects(drain(entry.stream()), (err) => isInflateFailure(err, label), label);
  }
});

test("cause distinguishes a truncated stream from a corrupt one", () => {
  const zlibCause = (bytes) => {
    const entry = new ZipArchive(bytes).get(PART);
    try { entry.bytes(); } catch (err) { return err.cause.code; }
    return null;
  };
  // Corruption flips bytes into a stream zlib can see is invalid.
  assert.equal(zlibCause(CORRUPTIONS["whole deflate stream zeroed"]()), "Z_DATA_ERROR");
  // Truncation ends the stream early, which zlib reports as an incomplete input.
  assert.equal(zlibCause(CORRUPTIONS["stream cut mid-block with the declared size adjusted"]()), "Z_BUF_ERROR");
  // The 0xFFFF local header makes the payload slice empty, i.e. truncated too.
  assert.equal(zlibCause(CORRUPTIONS["local header name and extra length 0xFFFF"]()), "Z_BUF_ERROR");
});

test("a valid deflate part is still returned as data on both paths", async () => {
  const archive = new ZipArchive(BASE);
  assert.equal(archive.readText(PART), SHEET, "buffered inflate must still return the part");

  const expected = new TextEncoder().encode(SHEET).length;
  assert.equal(await drain(new ZipArchive(BASE).get(PART).stream()), expected, "streamed inflate must still return the part");

  // And nothing above the entry level turned into an error either.
  const book = await readWorkbook(BASE);
  assert.equal(book.sheets.length, 1, "the workbook still reads");
  assert.equal(book.sheets[0].rows[0][0].v, 777, "the cell value still reads");
  let rows = 0;
  for await (const row of eachRow(BASE)) rows++;
  assert.ok(rows > 0, "the streaming reader still yields rows");
});

test("CRC_MISMATCH behaviour is unchanged, on both paths and with verifyCrc off", async () => {
  const w = new ZipWriter({ deterministic: true });
  w.add("[Content_Types].xml", CT);
  w.add("_rels/.rels", RELS);
  w.add("xl/workbook.xml", WB);
  w.add("xl/_rels/workbook.xml.rels", WBRELS);
  w.add(PART, SHEET, { compress: false });
  w.add(STORED, "the quick brown fox");
  const clean = w.build();

  // Flip a payload byte: a stored part needs no inflate, so this is the pure
  // CRC path and must behave exactly as it did before the inflate fix.
  const at = localHeader(clean, STORED).dataStart;
  const damaged = Uint8Array.from(clean);
  damaged[at] ^= 0x01;

  const buffered = new ZipArchive(damaged).get(STORED);
  assert.throws(() => buffered.bytes(), (err) => {
    assert.ok(err instanceof GridxcelError, "CRC failure must be typed");
    assert.equal(err.code, "CRC_MISMATCH");
    assert.ok(err.message.includes(STORED), "the CRC message still names the part");
    return true;
  }, "buffered CRC mismatch");

  await assert.rejects(drain(new ZipArchive(damaged).get(STORED).stream()), (err) => {
    assert.ok(err instanceof GridxcelError, "streamed CRC failure must be typed");
    assert.equal(err.code, "CRC_MISMATCH");
    return true;
  }, "streamed CRC mismatch");

  // The documented opt-out still returns the bytes as they are, damage included.
  const unchecked = new ZipArchive(damaged, { verifyCrc: false });
  const raw = unchecked.read(STORED);
  assert.equal(raw.length, 19, "opt-out returns the stored bytes");
  assert.notEqual(new TextDecoder().decode(raw), "the quick brown fox",
    "opt-out must hand back the damaged bytes rather than repair or reject them");
});

test("the reader never lets a raw zlib error escape", async () => {
  const corruptSheet = CORRUPTIONS["whole deflate stream zeroed"]();

  // eachRow streams the part, so the failure reaches the caller directly.
  await assert.rejects(async () => { for await (const row of eachRow(corruptSheet)) { void row; } },
    (err) => isInflateFailure(err, "eachRow"), "eachRow");

  // A part the reader loads eagerly fails the whole read rather than being
  // silently dropped, and the failure is typed.
  const corruptCT = (() => {
    const c = Uint8Array.from(BASE);
    const loc = localHeader(c, "[Content_Types].xml");
    c.fill(0, loc.dataStart, loc.dataStart + loc.compressedSize);
    return c;
  })();
  await assert.rejects(readWorkbook(corruptCT), (err) => {
    assert.ok(err instanceof GridxcelError, "readWorkbook threw " + err.name + " [" + (err.code || "-") + "]");
    return true;
  }, "readWorkbook on an eagerly-read part");

  // The synchronous reader may report the damage as a warning instead of
  // throwing; either way, no raw Error may escape.
  const warnings = [];
  try {
    readWorkbookSync(corruptSheet, { onWarning: (w) => warnings.push(w) });
  } catch (err) {
    assert.ok(err instanceof GridxcelError, "readWorkbookSync threw " + err.name + " [" + (err.code || "-") + "]");
    return;
  }
  assert.ok(warnings.some((w) => w.code === "BAD_SHEET"),
    "readWorkbookSync must report the part it could not inflate, got " + JSON.stringify(warnings.map((w) => w.code)));
});
