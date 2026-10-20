#!/usr/bin/env node
/**
 * Parser scaling and head-to-head benchmark.
 *
 * Two independent measurements:
 *
 *   1. Chunk-size scaling. The same worksheet is parsed from 16 KiB up to one
 *      single chunk. A linear scanner is flat across those sizes; a quadratic
 *      one grows with the buffer. This is the measurement that found the
 *      String#indexOf quadratic (see docs/PERFORMANCE.md), and it is the
 *      cheapest way to detect a regression of that class.
 *
 *   2. Head-to-head against SheetJS CE. Requires the vendor artifact in
 *      bench/sheetjs-0.20.3/ (see bench/sheetjs-0.20.3/PROVENANCE.md); the
 *      comparison is skipped with a notice when it is absent.
 *
 * Both libraries run in one process on the same bytes and the same machine, so
 * the ratio is meaningful even though absolute times move with the host.
 *
 * Usage: node --expose-gc scripts/bench-scaling.mjs [--file=path] [--runs=6]
 *        [--json]
 */

import { readFileSync, existsSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
const require = createRequire(import.meta.url);

const arg = (name, fallback) => {
  const hit = process.argv.find((a) => a.startsWith("--" + name + "="));
  return hit === undefined ? fallback : hit.slice(name.length + 3);
};
const asJson = process.argv.includes("--json");
const runs = Number(arg("runs", 6));
/** How many times to repeat the whole head-to-head comparison. */
const passes_wanted = Number(arg("passes", 5));
const file = arg("file", join(root, "bench", "fair-50000x8.xlsx"));

const gridxcel = join(root, "src", "index.js").replace(/\\/g, "/");
const { readWorkbook, readWorkbookSync, eachRow } = await import("file:///" + gridxcel.replace("/src/index.js", "/src/read.js"));
const { parseSheetXml } = await import("file:///" + gridxcel.replace("/src/index.js", "/src/sheet-reader.js"));
const { writeWorkbookSync } = await import("file:///" + gridxcel.replace("/src/index.js", "/src/write.js"));
const { ZipArchive } = await import("file:///" + gridxcel.replace("/src/index.js", "/src/zip.js"));

/** Median of \`n\` runs after one warm-up, which is what makes the ratio stable. */
async function median(fn, n = runs) {
  await fn();
  const ts = [];
  for (let i = 0; i < n; i++) {
    const s = performance.now();
    await fn();
    ts.push(performance.now() - s);
  }
  ts.sort((a, b) => a - b);
  return ts[Math.floor(ts.length / 2)];
}

async function* inChunks(text, size) {
  for (let i = 0; i < text.length; i += size) yield text.slice(i, i + size);
}

const report = { file: file.split(/[\\/]/).pop(), bytes: 0, scaling: {}, headToHead: null };

if (!existsSync(file)) {
  console.error("input not found: " + file);
  console.error("");
  console.error("This file is committed to the repository. If it is missing, regenerate it");
  console.error("with an independent implementation (openpyxl, not gridxcel or SheetJS):");
  console.error("");
  console.error("    python scripts/make-bench-input.py");
  console.error("");
  process.exit(2);
}
const bytes = readFileSync(file);
report.bytes = bytes.length;
const zip = new ZipArchive(bytes, { verifyCrc: false });
const part = zip.get("xl/worksheets/sheet1.xml");
const xml = new TextDecoder().decode(part.bytes());
const opts = { name: "S", sharedStrings: [], styles: { isDate: () => false } };

/* ---------------------------------------------------------- chunk scaling */

for (const kib of [16, 64, 256, 1024, 0]) {
  const size = kib === 0 ? xml.length : kib * 1024;
  const label = kib === 0 ? "whole" : kib + " KiB";
  report.scaling[label] = await median(() => parseSheetXml(inChunks(xml, size), opts), 3);
}
const base = report.scaling["16 KiB"];
let worst = 0;
for (const [k, v] of Object.entries(report.scaling)) {
  const ratio = v / base;
  if (ratio > worst) worst = ratio;
}

/* ----------------------------------------------------------- head to head */

/**
 * Repeat the whole comparison \`passes\` times and report the DISTRIBUTION of
 * each ratio, not a single number.
 *
 * A single pass on a shared machine can put the async reader anywhere between
 * about 1.7x and 4.7x ahead of SheetJS, because process-level effects (GC
 * timing, CPU frequency, whatever else is running) move each measurement
 * independently. Quoting one pass as "the" speedup is how a benchmark stops
 * being reproducible, so the ratio range is what gets published.
 */
const sjsDir = join(root, "bench", "sheetjs-0.20.3", "package");
if (existsSync(join(sjsDir, "xlsx.js"))) {
  const XLSX = require(join(sjsDir, "xlsx.js"));
  const book = readWorkbookSync(bytes);
  const medianOf = (xs) => {
    const s = [...xs].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  };

  const ops = {
    sheetjsRead: () => XLSX.read(bytes, { type: "buffer" }),
    gridxcelReadSync: () => readWorkbookSync(bytes),
    gridxcelRead: () => readWorkbook(bytes),
    gridxcelEachRow: async () => { for await (const r of eachRow(bytes)) void r; },
    sheetjsReadWrite: () => {
      const wb = XLSX.read(bytes, { type: "buffer" });
      return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    },
    gridxcelWriteSync: () => writeWorkbookSync(book, { type: "buffer" }),
  };

  // Warm every path so the first pass is not measuring JIT compilation.
  for (const fn of Object.values(ops)) await fn();

  const passes = [];
  for (let p = 0; p < passes_wanted; p++) {
    const sample = {};
    for (const [name, fn] of Object.entries(ops)) {
      const ts = [];
      for (let i = 0; i < 3; i++) {
        const s = performance.now();
        await fn();
        ts.push(performance.now() - s);
      }
      sample[name] = medianOf(ts);
    }
    passes.push(sample);
  }

  const times = {};
  for (const name of Object.keys(ops)) times[name] = medianOf(passes.map((p) => p[name]));

  const ratio = (name) => passes.map((p) => p.sheetjsRead / p[name]);
  const ratios = {
    gridxcelReadSync: ratio("gridxcelReadSync"),
    gridxcelRead: ratio("gridxcelRead"),
    gridxcelEachRow: ratio("gridxcelEachRow"),
  };

  report.headToHead = {
    passes: passes_wanted,
    times,
    ratios: Object.fromEntries(
      Object.entries(ratios).map(([k, v]) => [k, {
        median: medianOf(v),
        min: Math.min(...v),
        max: Math.max(...v),
      }]),
    ),
  };
}

/* ---------------------------------------------------------------- output */

if (asJson) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(report.file + "  " + (report.bytes / 1048576).toFixed(2) + " MiB  worksheet XML " + (xml.length / 1048576).toFixed(2) + " MiB");
  console.log("");
  console.log("  chunk-size scaling (same document, different feed sizes)");
  for (const [k, v] of Object.entries(report.scaling)) {
    console.log("    " + k.padEnd(10) + v.toFixed(0).padStart(7) + " ms   " + (v / base).toFixed(2) + "x vs 16 KiB");
  }
  console.log("    worst " + worst.toFixed(2) + "x - a linear scanner stays near 1.00");
  if (report.headToHead) {
    const h = report.headToHead;
    console.log("");
    console.log("  head to head (median of " + h.passes + " passes, each the median of 3)");
    const row = (k, v, extra) => console.log("    " + k.padEnd(22) + v.toFixed(0).padStart(7) + " ms" + (extra ? "   " + extra : ""));
    row("SheetJS read", h.times.sheetjsRead);
    row("gridxcel readSync", h.times.gridxcelReadSync);
    row("gridxcel read", h.times.gridxcelRead);
    row("gridxcel eachRow", h.times.gridxcelEachRow);
    row("SheetJS read+write", h.times.sheetjsReadWrite);
    row("gridxcel writeSync", h.times.gridxcelWriteSync, "write only");
    console.log("");
    console.log("  speedup versus SheetJS read, across " + h.passes + " passes");
    for (const [k, r] of Object.entries(h.ratios)) {
      console.log("    " + k.padEnd(22) + "median " + r.median.toFixed(2) + "x   range " + r.min.toFixed(2) + "x - " + r.max.toFixed(2) + "x");
    }
  } else {
    console.log("");
    console.log("  head to head skipped: bench/sheetjs-0.20.3/package/xlsx.js not present");
  }
}

if (process.env.BENCH_OUT) writeFileSync(process.env.BENCH_OUT, JSON.stringify(report, null, 2));
