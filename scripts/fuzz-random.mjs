/**
 * Random, time-bounded fuzzing of every parser that consumes untrusted input.
 *
 * The suites in test/ use a fixed corpus: each case is one somebody thought of.
 * That is enough to catch a regression and not enough to find a new defect. Every
 * serious bug in this project's history was found by an adversary rather than by
 * a fixture - a sparse row index that walked two billion holes, a three-character
 * number format that exhausted the heap, a 29-byte SYLK file that allocated forty
 * gigabytes. A fixed corpus cannot produce any of those.
 *
 * This runs the parsers against generated input for a fixed time, then reports.
 * A failing seed is printed, so the failure reproduces exactly:
 *
 *   node scripts/fuzz-random.mjs --seed 12345 --seconds 60
 *
 * Law under test, for every parser and every input:
 *   - it returns a value of the documented shape, or
 *   - it throws a GridxcelError (a stable code),
 * and never: throws a non-Error, throws a raw RangeError/TypeError, hangs, or
 * allocates without bound.
 */
import { formatNumber, formatDateTime, isDateFormat } from "../src/numfmt.js";
import { parseDSV, csvToSheet } from "../src/csv.js";
import { sylkToSheet, sheetToSYLK } from "../src/sylk.js";
import { difToSheet, sheetToDIF } from "../src/dif.js";
import { splitSections, pickSection } from "../src/numfmt.js";
import { decodeRange, encodeCell, decodeCell } from "../src/address.js";
import { GridxcelError } from "../src/errors.js";
import { storedZip } from "./zip-surgery.mjs";
import { ZipArchive } from "../src/zip.js";
import { readWorkbookSync } from "../src/read.js";
import { writeWorkbookSync } from "../src/write.js";

/* --------------------------------------------------------------- arguments */

const argv = process.argv.slice(2);
const argOf = (name, fallback) => {
  const i = argv.indexOf("--" + name);
  return i >= 0 && argv[i + 1] !== undefined ? argv[i + 1] : fallback;
};
const SEED = Number(argOf("seed", process.env.GRIDXCEL_FUZZ_SEED || 1));
const SECONDS = Number(argOf("seconds", process.env.GRIDXCEL_FUZZ_SECONDS || 20));
const DEADLINE = Date.now() + SECONDS * 1000;

/* ------------------------------------------------------------------ random */

/** Deterministic PRNG: the same seed reproduces the same sequence exactly. */
function makeRandom(seed) {
  let s = (seed >>> 0) || 1;
  return () => {
    s ^= s << 13; s >>>= 0;
    s ^= s >>> 17;
    s ^= s << 5; s >>>= 0;
    return s / 4294967296;
  };
}
const rnd = makeRandom(SEED);
/** Sentinel: this target's closure takes no argument, so its generator has none to give. */
const IGNORE = Symbol("no-argument");
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
const int = (n) => Math.floor(rnd() * n);

const CRLF = String.fromCharCode(13) + String.fromCharCode(10);
const LF = String.fromCharCode(10);
const CR = String.fromCharCode(13);

const ALPHABET = "0123456789abcdefXYZ;,/\\\"'[]()%#?*_@ .:-+=<>|{}~!$^&\n\r\t";

/** Random text, biased toward the punctuation each grammar actually uses. */
function randomText(maxLen, bias = ALPHABET) {
  const len = int(maxLen);
  let out = "";
  for (let i = 0; i < len; i++) out += bias[int(bias.length)];
  return out;
}

/* ------------------------------------------------------------------ harness */

let cases = 0;
const failures = [];

/**
 * Run one case under a wall-clock bound. A parser that hangs is killed by the
 * caller's timeout in CI; here we detect it by measuring, because a hang inside
 * this process would take the whole run with it.
 */
function probe(parser, label, input) {
  const started = Date.now();
  try {
    parser(input);
  } catch (error) {
    if (!(error instanceof GridxcelError)) {
      failures.push({
        parser: label,
        input: typeof input === "string" ? input.slice(0, 120) : String(input).slice(0, 120),
        why: "non-GridxcelError: " + (error && error.constructor ? error.constructor.name : typeof error) + ": " + String(error && error.message).slice(0, 120),
      });
    }
  }
  const ms = Date.now() - started;
  if (ms > 2000) {
    failures.push({
      parser: label,
      input: typeof input === "string" ? input.slice(0, 120) : String(input).slice(0, 120),
      why: "took " + ms + "ms",
    });
  }
  cases++;
}

/* ------------------------------------------------------------------- targets */

/** A SYLK-shaped skeleton with random field content. */
function randomSylk() {
  const records = ["ID;P"];
  const n = 1 + int(8);
  for (let i = 0; i < n; i++) {
    const kind = pick(["C", "C", "C", "B", "P", "F", "E"]);
    const fields = [kind];
    if (rnd() < 0.8) fields.push("Y" + (int(50) - 5));
    if (rnd() < 0.8) fields.push("X" + (int(50) - 5));
    if (rnd() < 0.8) fields.push("K" + (rnd() < 0.5 ? randomText(8) : int(1000)));
    records.push(fields.join(";"));
  }
  if (rnd() < 0.5) records.push("E");
  return records.join(pick([CRLF, LF, CR]));
}

/** A DIF-shaped skeleton with random chunk content. */
function randomDif() {
  const lines = ["TABLE", "0,1", '""'];
  if (rnd() < 0.7) lines.push("VECTORS", "0," + int(100), '""');
  if (rnd() < 0.7) lines.push("TUPLES", "0," + int(100), '""');
  lines.push("DATA", "0,0", '""');
  const n = 1 + int(10);
  for (let i = 0; i < n; i++) {
    const v = rnd();
    if (v < 0.3) lines.push("-1,0", pick(["BOT", "EOD", "EOT", randomText(6)]));
    else if (v < 0.6) lines.push("0," + (int(1000) - 500), pick(["V", "NA", "ERROR", "TRUE", "FALSE", randomText(4)]));
    else lines.push("1,0", rnd() < 0.5 ? '"' + randomText(10) + '"' : randomText(10));
  }
  return lines.join(pick([CRLF, LF])) + CRLF;
}

/** A number-format code from the characters the grammar gives meaning to. */
function randomFormat() {
  const bias = ["0", "#", "?", ".", ",", "%", "E", "+", "-", "/", "\\", String.fromCharCode(34), "[", "]", "@", "*", "_", ";", " ", String.fromCharCode(0)].join("");
  return randomText(20, bias);
}

/* ------------------------------------------------- workbook mutation base */

/**
 * A real workbook, in memory. Container fuzzing mutates one part of this and
 * re-reads it: that is where the styles, shared-strings and relationship loaders
 * live, and they are the paths a fixture-based test only ever sees intact.
 */
const BASE_PARTS = (() => {
  const bytes = writeWorkbookSync({
    sheetNames: ["S", "T"],
    sheets: [
      { name: "S", rows: [[{ t: "string", v: "a" }, { t: "number", v: 1 }], [{ t: "string", v: "b" }]], originRow: 1, originCol: 0, rowCount: 2, columnCount: 2, merges: ["A1:B1"], columns: [{ width: 12 }], rowInfo: new Map() },
      { name: "T", rows: [[{ t: "number", v: 9 }]], originRow: 1, originCol: 0, rowCount: 1, columnCount: 1, merges: [], columns: [], rowInfo: new Map() },
    ],
  });
  const zip = new ZipArchive(bytes);
  const parts = [];
  for (const entry of zip.entries) {
    parts.push([entry.name, /\.(xml|rels)$/.test(entry.name) ? entry.text() : entry.bytes()]);
  }
  return parts;
})();

/** Mutate one XML part: flip characters, truncate, inject markup, or blank it. */
function mutateWorkbook() {
  const parts = BASE_PARTS.map(([name, data]) => [name, data]);
  const xmlIndexes = parts
    .map(([name], i) => (/\.(xml|rels)$/.test(name) ? i : -1))
    .filter((i) => i >= 0);
  const which = xmlIndexes[int(xmlIndexes.length)];
  const [name, original] = parts[which];
  let text = String(original);
  switch (int(6)) {
    case 0: text = ""; break;
    case 1: text = text.slice(0, int(text.length + 1)); break;
    case 2: text = text.replace(/</g, () => (rnd() < 0.2 ? "" : "<")); break;
    case 3: {
      const at = int(text.length + 1);
      text = text.slice(0, at) + randomText(20) + text.slice(at);
      break;
    }
    case 4: text = text.replace(/="[^"]*"/g, () => (rnd() < 0.3 ? '="' + randomText(10) + '"' : '=""')); break;
    default: {
      // Duplicate a random slice: cheap way to grow nesting and repetition.
      const at = int(text.length + 1);
      text = text.slice(0, at) + text.slice(at, at + 40) + text.slice(at);
      break;
    }
  }
  parts[which] = [name, text];
  try {
    return storedZip(parts);
  } catch {
    return null;
  }
}

/**
 * Relative time share per target.
 *
 * An equal split sounds fair and is not: the workbook target rebuilds a zip per
 * case, so it gets a few thousand cases while "datetime" gets a million. The
 * expensive targets are the ones covering the least-tested code, so they get a
 * larger share.
 */
const TARGET_WEIGHT = {
  workbook: 4,
  sylk: 1, "sylk-raw": 1, dif: 1, "dif-raw": 1, dsv: 1,
  csvToSheet: 1, numfmt: 1, "numfmt-fixed": 1, datetime: 1, address: 1,
};

const TARGETS = [
  ["sylk", (text) => sylkToSheet(text), randomSylk],
  ["sylk-raw", (text) => sylkToSheet(text), () => randomText(80)],
  ["dif", (text) => difToSheet(text), randomDif],
  ["dif-raw", (text) => difToSheet(text), () => randomText(80)],
  ["dsv", (text) => parseDSV(text, { delimiter: pick([",", ";", "\t", "|"]) }), () => randomText(120, "ab,\"\n\r; \t|\\")],
  ["csvToSheet", (text) => csvToSheet(text), () => randomText(120, "ab,\"\n\r\n")],
  ["numfmt", (value) => formatNumber(value, randomFormat()), () => (rnd() - 0.5) * Math.pow(10, int(20))],
  ["numfmt-fixed", () => { const code = randomFormat(); formatNumber(1234.5678, code); isDateFormat(code); splitSections(code); pickSection(splitSections(code), 1, false); }, () => IGNORE],
  ["datetime", () => formatDateTime(int(100000), randomFormat()), () => IGNORE],
  ["address", () => { const a = randomText(8, "A-Z0-9$:"); try { decodeRange(a); } catch (e) { if (!(e instanceof GridxcelError)) throw e; } encodeCell(int(100), int(100)); }, () => IGNORE],
  ["workbook", (bytes) => readWorkbookSync(bytes), mutateWorkbook],
];

/* --------------------------------------------------------------------- run */

console.log("random fuzz: seed " + SEED + ", " + SECONDS + "s per target");
const perTarget = {};

for (const [label, parser, generator] of TARGETS) {
  let n = 0;
  const totalWeight = TARGETS.reduce((sum, [label]) => sum + (TARGET_WEIGHT[label] || 1), 0);
  const share = Math.floor((SECONDS * 1000) * ((TARGET_WEIGHT[label] || 1) / totalWeight));
  const until = Date.now() + share;
  while (Date.now() < until) {
    const input = generator();
    // A generator may decline a case (an archive it could not build); that is not
    // a parser result, so it is not counted.
    if (input === null || input === undefined) continue;
    probe(parser, label, input === IGNORE ? undefined : input);
    n++;
  }
  perTarget[label] = n;
  console.log("  " + label.padEnd(14) + n + " cases");
}

/* --------------------------------------------------------------- reporting */

console.log("\n" + cases + " cases total");
if (failures.length === 0) {
  console.log("OK: no untyped error, no hang, no unbounded allocation (seed " + SEED + ")");
  process.exit(0);
}
console.error("FAIL: " + failures.length + " problem(s) at seed " + SEED);
for (const f of failures.slice(0, 20)) {
  console.error("  [" + f.parser + "] " + f.why);
  console.error("    input: " + JSON.stringify(f.input));
}
console.error("\nReproduce with: node scripts/fuzz-random.mjs --seed " + SEED + " --seconds " + SECONDS);
process.exit(1);
