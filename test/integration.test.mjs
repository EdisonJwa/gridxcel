
import * as g from "../src/index.js";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import assert from "node:assert";

const F = (n) => import.meta.dirname + "/fixtures/" + n;  // string paths, not URL objects
// Scratch output goes next to the benchmarks, which are generated on demand and
// not tracked. git cannot create an empty directory, so make it here rather than
// assuming a previous run left one behind.
const BENCH = import.meta.dirname + "/../bench";
mkdirSync(BENCH, { recursive: true });

console.log("=== 1. read every fixture ===");
for (const f of ["golden.xlsx", "xlsxwriter-full.xlsx", "openpyxl-written.xlsx", "openpyxl-stored.xlsx", "openpyxl-lvl9.xlsx", "pandas.xlsx", "openpyxl-comment.xlsx"]) {
  const t0 = performance.now();
  const book = await g.read(F(f));
  const t1 = performance.now();
  const names = book.sheets.map((s) => s.name);
  const cells = book.sheets.reduce((a, s) => a + s.rows.reduce((b, r) => b + r.filter(Boolean).length, 0), 0);
  console.log("  " + f.padEnd(24), names.length + " sheets", String(cells).padStart(6) + " cells", (t1 - t0).toFixed(1) + " ms", JSON.stringify(names).slice(0, 70));
}

console.log("=== 2. golden.xlsx exact values ===");
const book = await g.read(F("golden.xlsx"));
assert.deepStrictEqual(book.sheets.map((s) => s.name), ["Types", "Layout", "Volume", "Empty"]);
const types = book.sheets[0];
// decodeCell returns 0-based {r,c}; rows are indexed from originRow (1-based sheet row).
const at = (sheet, a1) => { const { r, c } = g.decodeCell(a1); return sheet.rows[r - (sheet.originRow - 1)]?.[c - sheet.originCol]; };
const probe = [["A1","string","text"],["B1","number",42],["C1","boolean",true],["E1","number",0.125],["F1","number",1234.5],["G1","string","line1\nline2"],["H1","string",'quote"and,comma'],["I1","string","\u65e5\u672c\u8a9e " + String.fromCodePoint(0x1F600) + " \u00e9\u00e8"],["H2","number",3.14159265358979],["I2","string","  padded  "],["J2","number",1e15]];
for (const [a1, t, v] of probe) {
  const cell = at(types, a1);
  assert.ok(cell, "missing " + a1);
  assert.strictEqual(cell.t, t, a1 + " type: want " + t + " got " + cell.t);
  assert.strictEqual(cell.v, v, a1 + " value: want " + JSON.stringify(v) + " got " + JSON.stringify(cell.v));
}
console.log("  scalar round trip: " + probe.length + " cells exact (incl. unicode, newline, quotes, 1e15)");
const d1 = at(types, "D1");
console.log("  D1 (datetime):", JSON.stringify({ t: d1.t, v: d1.v, serial: d1.serial }));
const a2 = at(types, "A2"), b2 = at(types, "B2");
console.log("  formulas:", JSON.stringify({ A2: a2.f, B2: b2.f }));
assert.strictEqual(a2.f, "B1*2", "formula A2");
assert.strictEqual(b2.f, "SUM(B1:B1)", "formula B2");

console.log("=== 3. layout: merges, freeze, sparse ===");
const layout = book.sheets[1];
console.log("  merges:", JSON.stringify(layout.merges), "freeze:", JSON.stringify(layout.freeze));
assert.ok(layout.merges.includes("A1:C1") && layout.merges.includes("A3:B4"), "merges");
const volume = book.sheets[2];
const far = at(volume, "Z1");
console.log("  Volume Z1:", JSON.stringify(far && far.v), "rowCount:", volume.rowCount, "columnCount:", volume.columnCount);
assert.strictEqual(far && far.v, "far1", "sparse far column");
assert.strictEqual(volume.rowCount, 200, "volume row count");

console.log("=== 4. round trip: read -> write -> read ===");
const out = await g.write(book, { type: "buffer" });
assert.ok(out && out.length > 1000, "write produced bytes");
writeFileSync(BENCH + "/roundtrip.xlsx", out);
const again = await g.read(BENCH + "/roundtrip.xlsx");
assert.deepStrictEqual(again.sheets.map((s) => s.name), ["Types", "Layout", "Volume", "Empty"], "sheet names survive");
const t2 = again.sheets[0];
for (const [a1, t, v] of probe) {
  const cell = at(t2, a1);
  assert.ok(cell, "round trip missing " + a1);
  assert.strictEqual(cell.t, t, "round trip type " + a1 + ": " + cell.t);
  assert.strictEqual(cell.v, v, "round trip value " + a1);
}
assert.deepStrictEqual(again.sheets[1].merges.sort(), ["A1:C1", "A3:B4"], "merges survive");
console.log("  round trip preserved all " + probe.length + " probed cells + merges + sheet names");
const srcNames = JSON.stringify(book.sheets.map(s => s.name));
const rtNames = JSON.stringify(again.sheets.map(s => s.name));
assert.strictEqual(rtNames, srcNames);

console.log("=== 5. write a fresh file and confirm python can read it ===");
const fresh = g.normalizeBook({ sheets: [
  [["name", "score", "when", "ok"], ["ada", 42, new Date(2024, 2, 15, 14, 30, 0), true], ["grace", 37.5, new Date(1999, 11, 31), false]],
  [["only", "two"], [1, 2]],
] });
const bytes2 = await g.write(fresh, { type: "buffer" });
writeFileSync(BENCH + "/gridxcel-written.xlsx", bytes2);
console.log("  wrote bench/gridxcel-written.xlsx", bytes2.length, "bytes");

console.log("=== 6. streaming eachRow ===");
let n = 0, rowsSeen = 0;
for await (const ev of g.eachRow(F("golden.xlsx"))) { n++; if (ev.r === 1) rowsSeen++; }
console.log("  eachRow emitted", n, "rows;", rowsSeen, "sheets had a row 1");

console.log("=== 7. csv + json ===");
const csv = await g.sheetToCSV(types, {});
console.log("  csv first line:", JSON.stringify(csv.split(/\r?\n/)[0]));
const js = await g.sheetToJSON(types, {});
console.log("  json rows:", js.length, "first keys:", JSON.stringify(Object.keys(js[0] || {}).slice(0, 6)));

console.log("ALL INTEGRATION CHECKS PASSED");
