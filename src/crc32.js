/**
 * CRC-32 (IEEE 802.3), slice-by-8, table built lazily.
 * The exact polynomial every ZIP reader in existence expects.
 */

const T0 = new Int32Array(256);
const T1 = new Int32Array(256);
const T2 = new Int32Array(256);
const T3 = new Int32Array(256);
const T4 = new Int32Array(256);
const T5 = new Int32Array(256);
const T6 = new Int32Array(256);
const T7 = new Int32Array(256);

for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  T0[i] = c;
}
for (let i = 0; i < 256; i++) {
  T1[i] = (T0[i] >>> 8) ^ T0[T0[i] & 0xff];
  T2[i] = (T1[i] >>> 8) ^ T0[T1[i] & 0xff];
  T3[i] = (T2[i] >>> 8) ^ T0[T2[i] & 0xff];
  T4[i] = (T3[i] >>> 8) ^ T0[T3[i] & 0xff];
  T5[i] = (T4[i] >>> 8) ^ T0[T4[i] & 0xff];
  T6[i] = (T5[i] >>> 8) ^ T0[T5[i] & 0xff];
  T7[i] = (T6[i] >>> 8) ^ T0[T6[i] & 0xff];
}

/** Incremental CRC-32 over a sequence of buffers. */
export class Crc32 {
  constructor(seed = 0) {
    this.c = ~seed;
  }
  /** @param {Uint8Array} buf @returns {this} */
  update(buf) {
    let c = this.c;
    const n = buf.length;
    let i = 0;
    const n8 = n - 7;
    for (; i < n8; i += 8) {
      const v = c ^ (buf[i] | (buf[i + 1] << 8) | (buf[i + 2] << 16) | (buf[i + 3] << 24));
      c = T7[v & 0xff] ^ T6[(v >>> 8) & 0xff] ^ T5[(v >>> 16) & 0xff] ^ T4[(v >>> 24) & 0xff] ^
          T3[buf[i + 4]] ^ T2[buf[i + 5]] ^ T1[buf[i + 6]] ^ T0[buf[i + 7]];
    }
    for (; i < n; i++) c = (c >>> 8) ^ T0[(c ^ buf[i]) & 0xff];
    this.c = c;
    return this;
  }
  /** @returns {number} unsigned 32-bit CRC */
  digest() {
    return (~this.c) >>> 0;
  }
  reset() {
    this.c = -1;
    return this;
  }
}

/**
 * One-shot CRC-32 of a buffer (or a list of buffers, treated as one stream).
 * @param {Uint8Array|Uint8Array[]} input
 * @param {number} [seed]
 * @returns {number}
 */
export function crc32(input, seed = 0) {
  const h = new Crc32(seed);
  if (Array.isArray(input)) for (const b of input) h.update(b);
  else h.update(input);
  return h.digest();
}

/** Precomputed table, exported for reuse by other modules. */
export const CRC_TABLE = T0;
