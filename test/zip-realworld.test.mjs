
import { ZipArchive } from "../src/zip.js";
import { crc32 } from "../src/crc32.js";
import { readFileSync, readdirSync } from "node:fs";
import assert from "node:assert";

const dir = new URL("./fixtures/", import.meta.url);
const files = readdirSync(dir).filter((f) => f.endsWith(".xlsx"));
let total = 0;
for (const f of files) {
  const buf = readFileSync(new URL(f, dir));
  const z = new ZipArchive(buf);
  let n = 0;
  for (const e of z.entries) {
    const data = e.bytes();
    assert.strictEqual(data.length, e.size, f + " size mismatch " + e.name);
    const c = crc32(data);
    assert.strictEqual(c, e.crc, f + " CRC mismatch " + e.name + " got " + c.toString(16) + " want " + e.crc.toString(16));
    n++;
  }
  const wbXml = z.readText("xl/workbook.xml");
  assert.ok(wbXml && wbXml.includes("<workbook"), f + " workbook.xml");
  console.log(f.padEnd(26), z.entries.length, "entries, all CRCs verified,", "comment=" + JSON.stringify(z.comment.slice(0,12)));
  total += n;
}
console.log("verified", total, "parts across", files.length, "real-world fixtures");
