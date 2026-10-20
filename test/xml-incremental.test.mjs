/**
 * MAJOR-1: the incremental XmlParser must be linear in the input across every
 * chunk boundary, and MINOR-4: attribute values decode _xHHHH_ escapes the way
 * text nodes do.
 *
 * The equivalence tests below feed the same document at EVERY chunk size from
 * 1 byte up to the whole document. An off-by-one in a resume offset loses or
 * duplicates content instead of throwing, so a representative sample of sizes
 * is not enough - chunk sizes of 1, 2 and 3 have to be exercised explicitly.
 */
import test from "node:test";
import assert from "node:assert/strict";

import {
  XmlParser,
  parseTag,
  decodeEntities,
  decodeOoxmlEscapes,
  encodeXmlText,
  encodeXmlAttr,
} from "../src/xml.js";
import { StyleTable, parseStyles } from "../src/style-registry.js";

/** Canonical event log: every event, with text kept whole where the parser emits it whole. */
function events(chunks, opts = {}) {
  const log = [];
  const p = new XmlParser({
    ...opts,
    onOpen: (n, a) => log.push(["open", n, JSON.stringify(a), p.depth]),
    onClose: (n) => log.push(["close", n]),
    onText: (t) => log.push(["text", t]),
  });
  for (const c of chunks) p.write(c);
  p.end();
  return log;
}

function split(doc, size) {
  const out = [];
  for (let i = 0; i < doc.length; i += size) out.push(doc.slice(i, i + size));
  return out;
}

/** Every chunk size from 1 to the whole document must give the same events. */
function everyBoundary(doc, label) {
  const whole = JSON.stringify(events([doc]));
  const bad = [];
  for (let size = 1; size <= doc.length; size++) {
    if (JSON.stringify(events(split(doc, size))) !== whole) bad.push(size);
  }
  assert.deepEqual(bad, [], label + ": identical events for every chunk size 1.." + doc.length);
  return whole;
}

/* --------------------------------------------------- every-boundary equivalence */

const ALL_CONSTRUCTS =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  '<!DOCTYPE w [<!ENTITY ignored "never expanded">]>' +
  '<w:root xmlns:w="urn:w" plain="v" quoted="a&amp;b_x0041_" single=\'s\'>' +
  'lead text &amp; more' +
  '<self c="d>e"/>' +
  '<!-- comment with <angle> and -- dashes -->' +
  '<![CDATA[ raw <tag> & ]] > stuff ]]>' +
  '<?pi target with > and ? inside?>' +
  '<x d="' + "y".repeat(37) + '">tail _x005F_x0041_</x>' +
  '</w:root>';

test("MAJOR-1: every construct survives every chunk size (1..len)", () => {
  const log = everyBoundary(ALL_CONSTRUCTS, "all constructs");
  assert.ok(log.includes("_x0041_"), "the escaped literal decodes to the literal, not to A");
  assert.ok(log.includes("d>e"), "a > inside a quoted attribute is content");
});

test("MAJOR-1: a terminator that straddles a boundary at every offset", () => {
  // Each document places a multi-character terminator at a different offset by
  // padding the construct, so every possible split inside "-->" / "]]>" / "?>"
  // is covered for every chunk size as well.
  for (let pad = 0; pad < 8; pad++) {
    const p = "z".repeat(pad);
    everyBoundary("<r><!--" + p + "--></r>", "comment pad " + pad);
    everyBoundary("<r><![CDATA[" + p + "]]></r>", "cdata pad " + pad);
    everyBoundary("<r><?" + p + "?></r>", "pi pad " + pad);
    everyBoundary('<r a="' + p + '">' + p + "</r>", "attr pad " + pad);
  }
});

test("MAJOR-1: long constructs fed 1, 2 and 3 bytes at a time", () => {
  const body = "q".repeat(200);
  const docs = [
    '<t a="' + body + '">x</t>',
    "<t><!--" + body + "--></t>",
    "<t><![CDATA[" + body + "]]></t>",
    "<t><?" + body + "?></t>",
    "<t>" + body + "</t>",
    "<t>" + body + "</t >",
  ];
  for (const doc of docs) {
    const whole = JSON.stringify(events([doc]));
    for (const size of [1, 2, 3]) {
      assert.equal(JSON.stringify(events(split(doc, size))), whole, "size " + size + " for " + JSON.stringify(doc.slice(0, 12)));
    }
  }
});

test("MAJOR-1: malformed input stays chunk-invariant too", () => {
  for (const doc of ["<a><b></a>", "<a><b", "<a", "<", "a<b", "</a>", "<a/>&amp;", "<!--", "<![CDATA[", "<?"]) {
    const whole = JSON.stringify(events([doc]));
    for (let size = 1; size <= doc.length; size++) {
      assert.equal(JSON.stringify(events(split(doc, size))), whole, JSON.stringify(doc) + " at size " + size);
    }
  }
});

test("MAJOR-1: an unfinished construct is parked, not discarded", () => {
  // The trap this pins: parking with the index at the end of the buffer keeps
  // an empty buffer and silently drops the construct.
  const p = new XmlParser({});
  p.write("<a><b");
  assert.equal(p.depth, 1, "<a> is open, <b is still incomplete");
  assert.equal(p.current.name, "a");
  p.end();
  assert.equal(p.depth, 1, "the incomplete tag must not have been opened");
  const q = new XmlParser({});
  q.write("<a>text");
  q.end();
  assert.equal(q.depth, 1);
});

test("MAJOR-1: a piece is scanned once, not again by a later call", () => {
  // end() with no new data must not re-scan the newest piece: re-applying the
  // quote state to characters that were already counted flips it back open and
  // the tag never completes.
  const log = [];
  const p = new XmlParser({
    onOpen: (n, a) => log.push(["open", n, JSON.stringify(a)]),
    onClose: (n) => log.push(["close", n]),
    onText: (t) => log.push(["text", t]),
  });
  p.write('<t a="x');
  p.write('y"z');
  p.end();
  p.end();
  p.write(">tail</t>");
  p.end();
  assert.deepEqual(log, [
    ["open", "t", '{"a":"xy","z":""}'],
    ["text", "tail"],
    ["close", "t"],
  ]);
});

/* ----------------------------------------------------------------- MAJOR-1 scaling */

const MB = 1024 * 1024;
const CHUNK = 16 * 1024;

function parseThrough(doc, chunkSize) {
  const p = new XmlParser({ onText: () => {} });
  for (let i = 0; i < doc.length; i += chunkSize) p.write(doc.slice(i, i + chunkSize));
  p.end();
}

function best(doc, chunkSize, reps) {
  let low = Infinity;
  for (let r = 0; r < reps; r++) {
    const t0 = performance.now();
    parseThrough(doc, chunkSize);
    const dt = performance.now() - t0;
    if (dt < low) low = dt;
  }
  return low;
}

const LONG = {
  tag: (n) => '<t a="' + "x".repeat(n) + '">v</t>',
  comment: (n) => "<t><!--" + "x".repeat(n) + "--></t>",
  cdata: (n) => "<t><![CDATA[" + "x".repeat(n) + "]]></t>",
  pi: (n) => "<t><?" + "x".repeat(n) + "?></t>",
  text: (n) => "<t>" + "x".repeat(n) + "</t>",
};

test("MAJOR-1: 4x the input through small chunks costs about 4x, not 16x", () => {
  const N = 2 * MB;
  const FOUR_N = 4 * N;
  const lines = [];
  for (const [name, make] of Object.entries(LONG)) {
    const small = make(N);
    const large = make(FOUR_N);
    const t1 = best(small, CHUNK, 2);
    const t4 = best(large, CHUNK, 2);
    lines.push(name + ": " + N / MB + "MB " + t1.toFixed(1) + "ms, " + FOUR_N / MB + "MB " + t4.toFixed(1) + "ms, ratio " + (t4 / t1).toFixed(2));
    // Quadratic would be ~16x; 8x plus a small constant absorbs timer noise
    // without letting a quadratic path through.
    assert.ok(t4 <= t1 * 8 + 20, name + ": " + FOUR_N / MB + "MB took " + t4.toFixed(1) + "ms against " + t1.toFixed(1) + "ms at " + N / MB + "MB (quadratic?)");
  }
  console.log("    " + lines.join("\n    "));
  // The recorded pre-fix figure was 2,904 ms for 8 MB through 16 KB chunks.
  const eightMB = best(LONG.tag(8 * MB), CHUNK, 2);
  console.log("    tag, 8MB through 16KB chunks: " + eightMB.toFixed(1) + "ms (pre-fix 2904ms)");
  assert.ok(eightMB < 1500, "8MB tag should be far below the recorded 2904ms quadratic figure, got " + eightMB.toFixed(1) + "ms");
});

/* ------------------------------------------------------------------------ MINOR-4 */

test("MINOR-4: attribute values decode _xHHHH_ exactly like text", () => {
  assert.equal(parseTag('name val="Bad_x0001_Font"').attrs.val, "Bad\u0001Font");
  assert.equal(parseTag("name val='Bad_x0002_Font'").attrs.val, "Bad\u0002Font");
  assert.equal(parseTag("name val=Bad_x0003_Font").attrs.val, "Bad\u0003Font");
  assert.equal(parseTag('name val="a&amp;_x0041_"').attrs.val, "a&A");
  // The option mirrors the parser's: entities decode, escapes stay written.
  assert.equal(parseTag('name val="Bad_x0001_Font"', false).attrs.val, "Bad_x0001_Font");
  assert.equal(parseTag('name val="a&amp;b"', false).attrs.val, "a&b");
});

test("MINOR-4: the parser hands decoded attributes to onOpen", () => {
  const seen = [];
  const doc = '<t a="Bad_x0001_Font" b="0.00_x0002_%"/>';
  const p = new XmlParser({ onOpen: (n, a) => seen.push(a) });
  for (let size = 1; size <= doc.length; size++) {
    seen.length = 0;
    const q = new XmlParser({ onOpen: (n, a) => seen.push(a) });
    for (const c of split(doc, size)) q.write(c);
    q.end();
    assert.equal(seen[0].a, "Bad\u0001Font", "font name at chunk size " + size);
    assert.equal(seen[0].b, "0.00\u0002%", "format code at chunk size " + size);
  }
  const raw = [];
  const p2 = new XmlParser({ decodeEscapes: false, onOpen: (n, a) => raw.push(a) });
  p2.write(doc);
  p2.end();
  assert.equal(raw[0].a, "Bad_x0001_Font");
});

test("MINOR-4: _x005F_ is the escape for a literal escape, decoded in one pass", () => {
  assert.equal(decodeOoxmlEscapes("_x005F_x0041_"), "_x0041_", "a literal _x0041_, not A");
  assert.equal(decodeOoxmlEscapes("_x005F_"), "_");
  assert.equal(decodeOoxmlEscapes("_x005F_x005F_"), "_x005F_");
  assert.equal(decodeOoxmlEscapes("_x0041_"), "A");
  assert.equal(decodeOoxmlEscapes("plain"), "plain");
  // Exactly once: the parser must not decode its own output again.
  const log = events(['<t a="_x005F_x0041_">_x005F_x0041_</t>']);
  assert.equal(JSON.stringify(log), JSON.stringify([["open", "t", '{"a":"_x0041_"}', 0], ["text", "_x0041_"], ["close", "t"]]));
});

test("MINOR-4: control characters and literals survive an encode/parse round trip", () => {
  const values = ["Bad\u0001Font", "0.00\u0002%", "_x0041_", "_x005F_", "a_x0009_b", "tab\there", 'q"q', "<&>"];
  for (const v of values) {
    assert.equal(parseTag('t v="' + encodeXmlAttr(v) + '"').attrs.v, v, "attribute round trip: " + JSON.stringify(v));
    const log = events(["<t>" + encodeXmlText(v) + "</t>"]);
    assert.equal(log[1][1], v, "text round trip: " + JSON.stringify(v));
  }
});

test("MINOR-4: a styles.xml round trip keeps the control characters", () => {
  const table = new StyleTable();
  const idx = table.xf({ font: { name: "Bad\u0001Font" }, numFmt: "0.00\u0002%" });
  const xml = table.toXml();
  const parsed = parseStyles(xml);
  const style = parsed.styles[idx];
  assert.equal(style.font.name, "Bad\u0001Font", "font name decoded out of its attribute");
  assert.equal(style.numFmtCode, "0.00\u0002%", "format code decoded out of its attribute");
  assert.equal(parsed.numFmts.get(style.numFmt), "0.00\u0002%");
  assert.ok(xml.includes("_x0001_") && xml.includes("_x0002_"), "the part still carries the escapes, not raw controls");
  assert.equal(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(xml), false, "no raw control character reaches the part");
});
