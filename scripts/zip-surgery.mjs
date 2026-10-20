/**
 * Hand-rolled zip byte surgery for the robustness suite.
 *
 * These helpers deliberately do NOT use src/zip.js: the inputs that test the
 * library must not be produced by the library itself. Everything here is a
 * direct read/write of the ZIP container layout (local file headers, the
 * end-of-central-directory record) plus CRC-32 for the hand-built archive.
 */
/* ------------------------------------------------------ byte-level helpers */

const u8 = (...bytes) => Uint8Array.from(bytes);
const ascii = (s) => new TextEncoder().encode(s);

/** Locate the local file header + data range of one zip entry, by hand. */
export function locateEntry(buf, name) {
  const want = ascii(name);
  for (let i = 0; i + 30 <= buf.length; i++) {
    if (buf[i] !== 0x50 || buf[i + 1] !== 0x4b || buf[i + 2] !== 0x03 || buf[i + 3] !== 0x04) continue;
    const method = buf[i + 8] | (buf[i + 9] << 8);
    const compSize = buf[i + 18] | (buf[i + 19] << 8) | (buf[i + 20] << 16) | (buf[i + 21] << 24);
    const nameLen = buf[i + 26] | (buf[i + 27] << 8);
    const extraLen = buf[i + 28] | (buf[i + 29] << 8);
    const nameStart = i + 30;
    if (nameLen !== want.length) continue;
    let same = true;
    for (let k = 0; k < want.length; k++) if (buf[nameStart + k] !== want[k]) { same = false; break; }
    if (!same) continue;
    const dataStart = nameStart + nameLen + extraLen;
    return { headerStart: i, dataStart, dataEnd: dataStart + (compSize >>> 0), method };
  }
  return null;
}

/** Index of the end-of-central-directory record, or -1. */
export function findEocd(buf) {
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf[i] === 0x50 && buf[i + 1] === 0x4b && buf[i + 2] === 0x05 && buf[i + 3] === 0x06) return i;
  }
  return -1;
}

export function flipByte(buf, offset) {
  const out = buf.slice();
  out[offset] = out[offset] ^ 0xff;
  return out;
}

/* A stored-only zip built here, so the "valid zip, not a workbook" case does
   not depend on any gridxcel code. */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(bytes) {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function storedZip(entries) {
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const [name, text] of entries) {
    const nameBytes = ascii(name);
    const data = typeof text === "string" ? ascii(text) : text;
    const crc = crc32(data);
    const lh = new Uint8Array(30 + nameBytes.length);
    const dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(8, 0, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, data.length, true);
    dv.setUint32(22, data.length, true);
    dv.setUint16(26, nameBytes.length, true);
    lh.set(nameBytes, 30);
    chunks.push(lh, data);
    const ch = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(10, 0, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(42, offset, true);
    ch.set(nameBytes, 46);
    central.push(ch);
    offset += lh.length + data.length;
  }
  const centralSize = central.reduce((n, c) => n + c.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  const total = chunks.reduce((n, c) => n + c.length, 0) + centralSize + 22;
  const out = new Uint8Array(total);
  let p = 0;
  for (const c of [...chunks, ...central, eocd]) { out.set(c, p); p += c.length; }
  return out;
}

export { u8, ascii };
