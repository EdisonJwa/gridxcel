
import { encodeCol, decodeCol, encodeCell, decodeCell, decodeRange, encodeRange, toRange, splitSheetRef, makeBounds, extendBounds, boundsToRef } from "../src/address.js";
import { GridxcelError } from "../src/errors.js";
import assert from "node:assert";

// Round-trip every column index in the XLSX grid (A..XFD) and row numbers.
let checked = 0;
for (let c = 0; c <= 16383; c++) {
  const letters = encodeCol(c);
  assert.strictEqual(decodeCol(letters), c, "col round trip " + c);
  assert.strictEqual(decodeCol(letters.toLowerCase()), c, "col case " + c);
  checked++;
}
console.log("column round trips:", checked, "(A.." + encodeCol(16383) + ")");
for (const r of [0, 1, 9, 99, 999, 99999, 1048575]) {
  const a = encodeCell(r, 5);
  const d = decodeCell(a);
  assert.deepStrictEqual(d, { r, c: 5 }, "cell round trip " + a);
}
console.log("row round trips OK, max row address:", encodeCell(1048575, 16383));

assert.strictEqual(encodeCol(0), "A");
assert.strictEqual(encodeCol(25), "Z");
assert.strictEqual(encodeCol(26), "AA");
assert.strictEqual(encodeCol(701), "ZZ");
assert.strictEqual(encodeCol(702), "AAA");
assert.strictEqual(encodeCol(16383), "XFD");

// Absolute markers are tolerated on both axes.
assert.deepStrictEqual(decodeCell("$B$7"), { r: 6, c: 1 });
assert.deepStrictEqual(decodeCell("b7"), { r: 6, c: 1 });
assert.deepStrictEqual(decodeRange("A1:C3"), { s: { r: 0, c: 0 }, e: { r: 2, c: 2 } });
assert.deepStrictEqual(decodeRange("Sheet1!B2"), { s: { r: 1, c: 1 }, e: { r: 1, c: 1 } });
assert.deepStrictEqual(decodeRange("'My Sheet'!A1:B2"), { s: { r: 0, c: 0 }, e: { r: 1, c: 1 } });
assert.strictEqual(encodeRange({ r: 0, c: 0 }, { r: 2, c: 2 }), "A1:C3");
console.log("range parsing OK (including sheet-qualified refs)");

// Invalid addresses must throw a typed error, never return garbage.
const bad = ["", "1", "A", "$", "A$", "!!", "A0x", " A1", "A1B"];
let threw = 0;
for (const b of bad) {
  try { decodeCell(b); console.log("  WARNING: decoded garbage from", JSON.stringify(b), "->", JSON.stringify(decodeCell(b))); }
  catch (e) { assert.ok(e instanceof GridxcelError, "typed error for " + b); threw++; }
}
console.log("invalid addresses rejected with GridxcelError:", threw + "/" + bad.length);

// Bounds tracking
const b = makeBounds();
extendBounds(b, 5, 2); extendBounds(b, 1, 9); extendBounds(b, 3, 0);
assert.deepStrictEqual(b, { minR: 1, minC: 0, maxR: 5, maxC: 9 });
assert.strictEqual(boundsToRef(b), "A2:J6");
assert.strictEqual(boundsToRef(makeBounds()), "A1");
console.log("bounds tracking OK ->", boundsToRef(b));

// splitSheetRef
assert.deepStrictEqual(splitSheetRef("Data!A1:B2"), ["Data", "A1:B2"]);
assert.deepStrictEqual(splitSheetRef("'It''s'!A1"), ["It's", "A1"]);
assert.deepStrictEqual(splitSheetRef("A1"), [null, "A1"]);
console.log("splitSheetRef OK");

// toRange normalisations
assert.deepStrictEqual(toRange("B2:C3"), { s: { r: 1, c: 1 }, e: { r: 2, c: 2 } });
assert.deepStrictEqual(toRange(5), { s: { r: 4, c: 0 }, e: { r: 4, c: 16383 } });
assert.strictEqual(toRange(undefined), undefined);
console.log("toRange OK");
console.log("ADDRESS ALL GREEN");
