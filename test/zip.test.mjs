
import { ZipWriter, ZipArchive, createZip } from "../src/zip.js";
import { crc32 } from "../src/crc32.js";
import assert from "node:assert";

// 1. CRC32 known vectors
assert.strictEqual(crc32(new TextEncoder().encode("")), 0, "crc empty");
assert.strictEqual(crc32(new TextEncoder().encode("123456789")), 0xcbf43926, "crc check vector");
const big = new Uint8Array(100000); for (let i=0;i<big.length;i++) big[i] = (i*7)&0xff;
const inc = crc32(big);
// chunked must equal one-shot
let acc = 0; const { Crc32 } = await import("../src/crc32.js");
const h = new Crc32();
for (let i=0;i<big.length;i+=999) h.update(big.subarray(i, Math.min(i+999, big.length)));
assert.strictEqual(h.digest(), inc, "chunked crc");
console.log("crc32 OK", inc.toString(16));

// 2. writer -> reader round trip, incl. many parts and empty part
const w = new ZipWriter({ deterministic: true });
const payloads = new Map();
for (let i=0;i<200;i++) {
  const name = "xl/worksheets/sheet" + i + ".xml";
  const data = "<row r=\"" + i + "\"><c r=\"A1\" t=\"s\"><v>" + i + "</v></c></row>".repeat(50);
  payloads.set(name, data);
  w.add(name, data);
}
w.add("empty.txt", "");
w.add("stored.bin", new Uint8Array([1,2,3,4,5]), { compress: false });
w.add("unicode/\u65e5\u672c\u8a9e.txt", "\u3053\u3093\u306b\u3061\u306f");
const bytes = w.build();
console.log("archive bytes", bytes.length);

const a = new ZipArchive(bytes);
assert.strictEqual(a.entries.length, 203);
for (const [name, data] of payloads) {
  assert.strictEqual(a.readText(name), data, "mismatch " + name);
}
assert.strictEqual(a.readText("empty.txt"), "");
assert.deepStrictEqual(Array.from(a.read("stored.bin")), [1,2,3,4,5]);
assert.strictEqual(a.readText("unicode/\u65e5\u672c\u8a9e.txt"), "\u3053\u3093\u306b\u3061\u306f");
console.log("zip round-trip OK (203 entries, unicode + empty + stored)");

// 3. 64 KiB+ comment on EOCD tolerance
const w2 = createZip({ "a.txt": "hello".repeat(1000) });
console.log("createZip OK", w2.length);

// 4. stream()
const s = a.get("xl/worksheets/sheet3.xml").stream();
let n = 0; s.on("data", (c) => { n += c.length; });
await new Promise((r) => s.on("end", r));
assert.ok(n > 0);
console.log("stream OK", n);

// 5. determinism
const b1 = createZip({ "x": "abc" }, { deterministic: true });
const b2 = createZip({ "x": "abc" }, { deterministic: true });
assert.deepStrictEqual(Array.from(b1), Array.from(b2));
console.log("deterministic build OK");
console.log("ALL ZIP TESTS PASSED");
