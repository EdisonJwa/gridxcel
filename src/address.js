/**
 * A1 addressing: conversion, parsing, range maths, and comparison.
 * Pure string/number work - no allocation-heavy regex on the hot path.
 */

import { GridxcelError } from "./errors.js";
import { MAX_COLUMNS } from "./constants.js";

const A = 65;
const Z = 90;

/** 0 -> "A", 25 -> "Z", 26 -> "AA", 16383 -> "XFD". */
export function encodeCol(c) {
  if (c < 0 || !Number.isFinite(c)) throw new GridxcelError("invalid column index " + c, "BAD_ADDRESS");
  let s = "";
  c = c | 0;
  while (c >= 0) {
    s = String.fromCharCode(A + (c % 26)) + s;
    c = ((c / 26) | 0) - 1;
  }
  return s;
}

/** "A" -> 0, "Z" -> 25, "AA" -> 26. Case-insensitive. */
export function decodeCol(letters) {
  let c = 0;
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i);
    const v = code >= 97 ? code - 97 : code - A;
    if (v < 0 || v > 25) throw new GridxcelError("invalid column letters " + JSON.stringify(letters), "BAD_ADDRESS");
    c = c * 26 + v + 1;
  }
  return c - 1;
}

/** {r:0,c:0} -> "A1". */
export function encodeCell(r, c) {
  return encodeCol(c) + (r + 1);
}

/** "A1" -> {r:0,c:0}; absolute markers are ignored. */
export function decodeCell(a1) {
  let i = 0;
  if (a1.charCodeAt(0) === 36) i = 1;
  const start = i;
  while (i < a1.length) {
    const code = a1.charCodeAt(i);
    if ((code >= 65 && code <= 90) || (code >= 97 && code <= 122)) i++;
    else break;
  }
  if (i === start) throw new GridxcelError("invalid cell address " + JSON.stringify(a1), "BAD_ADDRESS");
  const c = decodeCol(a1.slice(start, i));
  if (a1.charCodeAt(i) === 36) i++;
  const numStart = i;
  while (i < a1.length && a1.charCodeAt(i) >= 48 && a1.charCodeAt(i) <= 57) i++;
  if (i === numStart) throw new GridxcelError("invalid cell address " + JSON.stringify(a1), "BAD_ADDRESS");
  const r = parseInt(a1.slice(numStart, i), 10) - 1;
  if (r < 0) throw new GridxcelError("invalid row in " + JSON.stringify(a1), "BAD_ADDRESS");
  // Reject trailing junk: "A1B" is not a cell, and silently ignoring the tail
  // hides caller bugs.
  if (i !== a1.length) throw new GridxcelError("invalid cell address " + JSON.stringify(a1), "BAD_ADDRESS");
  return { r, c };
}

/** "A1:B10" -> {s:{r,c}, e:{r,c}}. A single cell yields s === e. */
export function decodeRange(ref) {
  const bang = ref.lastIndexOf("!");
  if (bang >= 0) ref = ref.slice(bang + 1);
  const colon = ref.indexOf(":");
  if (colon < 0) {
    const c = decodeCell(ref);
    return { s: c, e: { r: c.r, c: c.c } };
  }
  return { s: decodeCell(ref.slice(0, colon)), e: decodeCell(ref.slice(colon + 1)) };
}

export function encodeRange(s, e) {
  return encodeCell(s.r, s.c) + ":" + encodeCell(e.r, e.c);
}

/** Split "Sheet1!A1:B2" (or quoted sheet name) into [sheetName|null, rangeText]. */
export function splitSheetRef(ref) {
  const bang = ref.lastIndexOf("!");
  if (bang < 0) return [null, ref];
  let name = ref.slice(0, bang);
  if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
  return [name, ref.slice(bang + 1)];
}

/** Normalise any accepted range spelling into {s:{r,c}, e:{r,c}}. */
export function toRange(range) {
  if (!range) return undefined;
  if (typeof range === "string") return decodeRange(range);
  if (typeof range === "number") return decodeRange("A" + range + ":" + encodeCol(MAX_COLUMNS - 1) + range);
  return range;
}

/** Grow a tracked bounding box; maxR/maxC are inclusive. */
export function makeBounds() {
  return { minR: Infinity, minC: Infinity, maxR: -1, maxC: -1 };
}
export function extendBounds(b, r, c) {
  if (r < b.minR) b.minR = r;
  if (c < b.minC) b.minC = c;
  if (r > b.maxR) b.maxR = r;
  if (c > b.maxC) b.maxC = c;
  return b;
}
export function boundsToRef(b) {
  if (b.maxR < 0) return "A1";
  return encodeRange({ r: b.minR, c: b.minC }, { r: b.maxR, c: b.maxC });
}
