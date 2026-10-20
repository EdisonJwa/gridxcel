
import { XmlParser, parseTag, decodeEntities, decodeOoxmlEscapes, encodeXmlText, encodeXmlAttr } from "../src/xml.js";
import { ZipArchive } from "../src/zip.js";
import { readFileSync } from "node:fs";
import assert from "node:assert";

/* Collect a canonical event log from a parse. */
function events(chunks, opts = {}) {
  const log = [];
  const p = new XmlParser({
    ...opts,
    onOpen: (n, a, parent) => log.push(["open", n, JSON.stringify(a), p.depth]),
    onClose: (n, a, node) => log.push(["close", n]),
    onText: (t, node) => { if (t.trim()) log.push(["text", t]); },
  });
  for (const c of chunks) p.write(c);
  p.end();
  return log;
}

const SHEET = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
  '<sheetData>' +
  '<row r="1" spans="1:3"><c r="A1" t="inlineStr"><is><t xml:space="preserve">  Name  </t></is></c><c r="B1" t="s"><v>0</v></c></row>' +
  '<row r="2"><c r="A2"><v>3.5</v></c><c r="B2" t="b"><v>1</v></c><c r="C2" t="e"><v>#DIV/0!</v></c></row>' +
  '<row r="3"><c r="A3" t="str"><f>CONCAT(A1,"&amp;")</f><v>a&amp;b</v></c><c r="B3" s="2"><v>45292</v></c></row>' +
  '</sheetData>' +
  '<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>' +
  '<![CDATA[ignore me]]>' +
  '<!-- a comment with <angle> brackets -->' +
  '</worksheet>';

const whole = events([SHEET]);
assert.ok(whole.length > 20, "expected a rich event log, got " + whole.length);
assert.deepStrictEqual(events([SHEET]), whole, "single chunk");

// Property: any chunking of the same byte stream must yield identical events.
let failures = 0;
for (let size = 1; size <= 97; size++) {
  const chunks = [];
  for (let i = 0; i < SHEET.length; i += size) chunks.push(SHEET.slice(i, i + size));
  const got = events(chunks);
  if (JSON.stringify(got) !== JSON.stringify(whole)) {
    failures++;
    if (failures <= 2) {
      console.log("CHUNK FAILURE at size", size);
      for (let i = 0; i < Math.max(got.length, whole.length); i++) {
        if (JSON.stringify(got[i]) !== JSON.stringify(whole[i])) { console.log("  idx", i, "want", JSON.stringify(whole[i]), "got", JSON.stringify(got[i])); break; }
      }
    }
  }
}
console.log("chunk-boundary property (sizes 1..97):", failures === 0 ? "PASS" : failures + " FAILURES");

// Random chunking, 200 trials.
let rfail = 0;
for (let trial = 0; trial < 200; trial++) {
  const chunks = [];
  let i = 0;
  while (i < SHEET.length) { const n = 1 + ((trial * 7 + i * 13) % 23); chunks.push(SHEET.slice(i, i + n)); i += n; }
  if (JSON.stringify(events(chunks)) !== JSON.stringify(whole)) rfail++;
}
console.log("random chunking (200 trials):", rfail === 0 ? "PASS" : rfail + " FAILURES");

// Entities and OOXML escapes.
assert.strictEqual(decodeEntities("a&amp;b&lt;c&gt;d&quot;e&apos;f"), 'a&b<c>d"e\'f');
assert.strictEqual(decodeEntities("&#65;&#x42;&#x1F600;"), "AB\u{1F600}");
assert.strictEqual(decodeEntities("&unknown; &"), "&unknown; &");
assert.strictEqual(decodeOoxmlEscapes("a_x0009_b"), "a\tb");
// _x005F_ is OOXML's escape for a LITERAL escape: the leading underscore is
// escaped and the rest is text, so this is the literal "_x0009_", not a tab.
assert.strictEqual(decodeOoxmlEscapes("_x005F_x0009_"), "_x0009_");
assert.strictEqual(decodeOoxmlEscapes("_x005F_"), "_");
assert.strictEqual(encodeXmlText('a<b>&"c"'), "a&lt;b&gt;&amp;&quot;c&quot;".replace("&quot;c&quot;", '"c"'));
assert.strictEqual(encodeXmlText("tab\there"), "tab\there");
assert.strictEqual(encodeXmlText("ctl\u0001"), "ctl_x0001_");
assert.strictEqual(encodeXmlAttr('say "hi"'), "say &quot;hi&quot;");
console.log("entity/escape handling OK");

// parseTag corner cases.
const tg = parseTag('c r="A1" t="inlineStr" /');
assert.strictEqual(tg.name, "c");
assert.strictEqual(tg.attrs.r, "A1");
assert.strictEqual(tg.attrs.t, "inlineStr");
const tg2 = parseTag("x:foo a='1' b = \"two\" bare");
assert.strictEqual(tg2.local, "foo");
assert.strictEqual(tg2.attrs.a, "1");
assert.strictEqual(tg2.attrs.b, "two");
assert.strictEqual(tg2.attrs.bare, "");
console.log("parseTag OK");

// Real-world parts must parse with no chunking surprises and produce sane structure.
const buf = readFileSync(new URL("./fixtures/xlsxwriter-full.xlsx", import.meta.url));
const zip = new ZipArchive(buf);
for (const name of ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/styles.xml", "xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"]) {
  const xml = zip.readText(name);
  if (!xml) { console.log("missing", name); continue; }
  const one = events([xml]);
  let mismatch = 0;
  for (const size of [1, 3, 7, 64, 4096, 65536]) {
    const chunks = [];
    for (let i = 0; i < xml.length; i += size) chunks.push(xml.slice(i, i + size));
    if (JSON.stringify(events(chunks)) !== JSON.stringify(one)) mismatch++;
  }
  console.log(name.padEnd(34), xml.length.toString().padStart(8), "bytes,", one.length, "events,", mismatch === 0 ? "chunk-stable" : mismatch + " CHUNK MISMATCHES");
}
console.log(failures === 0 && rfail === 0 ? "XML ALL GREEN" : "XML HAS FAILURES");
