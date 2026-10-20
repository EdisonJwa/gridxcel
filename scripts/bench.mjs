#!/usr/bin/env node
/**
 * gridxcel benchmark: parse / stream-parse / write / CSV-export on a large
 * workbook, with peak-heap and peak-RSS measurement.
 *
 *   (i)   readWorkbook(path)  -> wall time + peak heap for a full parse
 *         eachRow(...)        -> wall time + peak heap for the streaming path
 *   (ii)  WorkbookWriter      -> streaming write, same shape
 *         writeWorkbookSync   -> whole-book write, same shape (opt-in)
 *   (iii) sheetToCSV          -> rows/second for CSV export
 *
 * Peak memory is sampled from process.memoryUsage() on a 5 ms interval AND,
 * when node runs with --expose-gc (the script re-executes itself with that
 * flag), the retained heap is measured after a forced collection. Interval
 * sampling cannot observe inside a synchronous CPU block; the retained-heap
 * numbers cover that gap, and both are reported.
 *
 * The input workbook is generated once by python/openpyxl into bench/ and
 * reused, so the input bytes never come from gridxcel.
 *
 * Usage: node scripts/bench.mjs [--rows=200000] [--cols=10] [--quick]
 *        [--no-materialise] [--json]
 */

import { existsSync, mkdirSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { execFileSync } from "node:child_process";
import { fileURLToPath, pathToFileURL } from "node:url";

import { generate } from "./oracle-client.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "..");

/* --------------------------------------------------------------- options */

function parseArgs(argv) {
  const o = { rows: 200000, cols: 10, quick: false, materialise: true, json: false };
  for (const a of argv) {
    if (a === "--quick") o.quick = true;
    else if (a === "--json") o.json = true;
    else if (a === "--no-materialise") o.materialise = false;
    else if (a.startsWith("--rows=")) o.rows = Number(a.slice(7));
    else if (a.startsWith("--cols=")) o.cols = Number(a.slice(7));
  }
  if (o.quick) { o.rows = Math.min(o.rows, 20000); o.cols = Math.min(o.cols, 10); }
  return o;
}

/* --------------------------------------------------- re-exec with --expose-gc */

const GC_FLAG = "--expose-gc";
function ensureGc(argv) {
  if (typeof globalThis.gc === "function") return true;
  if (process.env.GRIDXCEL_BENCH_REEXEC === "1") return false;
  try {
    execFileSync(process.execPath, [GC_FLAG, fileURLToPath(import.meta.url), ...argv], {
      stdio: "inherit",
      env: { ...process.env, GRIDXCEL_BENCH_REEXEC: "1" },
    });
    process.exit(0);
  } catch (err) {
    if (typeof err?.status === "number") process.exit(err.status);
    return false;
  }
  return true;
}

/* ------------------------------------------------------------ measurement */

export function startSampler(intervalMs = 5) {
  const state = { peakHeap: 0, peakRss: 0, samples: 0, minHeap: Infinity };
  const sample = () => {
    const m = process.memoryUsage();
    if (m.heapUsed > state.peakHeap) state.peakHeap = m.heapUsed;
    if (m.heapUsed < state.minHeap) state.minHeap = m.heapUsed;
    if (m.rss > state.peakRss) state.peakRss = m.rss;
    state.samples++;
  };
  sample();
  const timer = setInterval(sample, intervalMs);
  if (typeof timer.unref === "function") timer.unref();
  return {
    state,
    stop() { clearInterval(timer); sample(); return state; },
  };
}

function gcNow() {
  if (typeof globalThis.gc === "function") { globalThis.gc(); globalThis.gc(); return true; }
  return false;
}

async function measure(fn, { gc = true } = {}) {
  if (gc) gcNow();
  const before = process.memoryUsage();
  const sampler = startSampler();
  const t0 = performance.now();
  let value;
  let error = null;
  try {
    value = await fn();
  } catch (err) {
    error = err;
  }
  const ms = performance.now() - t0;
  const s = sampler.stop();
  const after = process.memoryUsage();
  const hadGc = gcNow();
  const retained = process.memoryUsage();
  return {
    ms,
    peakHeap: s.peakHeap,
    peakRss: s.peakRss,
    samples: s.samples,
    // Peak minus the baseline this stage started from: the only fair way to
    // compare stages that run one after another in the same process.
    peakHeapExcess: Math.max(0, s.peakHeap - before.heapUsed),
    peakRssExcess: Math.max(0, s.peakRss - before.rss),
    heapBefore: before.heapUsed,
    heapAfter: after.heapUsed,
    heapRetained: retained.heapUsed,
    rssAfter: after.rss,
    forcedGc: hadGc,
    value,
    error,
  };
}

export const mb = (bytes) => (bytes / (1024 * 1024)).toFixed(1);

/* --------------------------------------------------------------- modules */

async function tryImport(rel) {
  const abs = path.join(ROOT, rel);
  if (!existsSync(abs)) return { ok: false, missing: "NOT YET AVAILABLE: " + rel };
  try {
    return { ok: true, mod: await import(pathToFileURL(abs).href) };
  } catch (err) {
    return { ok: false, missing: "IMPORT FAILED: " + rel + ": " + (err?.message ?? err) };
  }
}

/** Find eachRow wherever the public API exposes it. */
async function findEachRow() {
  for (const rel of ["src/index.js", "src/api.js", "src/read.js"]) {
    const r = await tryImport(rel);
    if (r.ok && typeof r.mod.eachRow === "function") return { fn: r.mod.eachRow, from: rel };
  }
  return { fn: null, from: null };
}

/* ------------------------------------------------------------------- run */

export async function runBench(options = parseArgs([])) {
  const { rows, cols } = options;
  const file = path.join(ROOT, "bench", "volume-" + rows + "x" + cols + ".xlsx");
  mkdirSync(path.dirname(file), { recursive: true });
  const generated = !existsSync(file);
  if (generated) generate(ROOT, file, rows, cols);
  const inputBytes = statSync(file).size;

  const result = {
    bench: "gridxcel",
    rows, cols,
    file: path.relative(ROOT, file).replace(/\\/g, "/"),
    inputBytes,
    generatedNow: generated,
    gcAvailable: typeof globalThis.gc === "function",
    node: process.version,
    results: {},
    missing: [],
  };
  const lines = [];
  const record = (key, entry) => {
    result.results[key] = entry;
    lines.push(entry);
    if (!options.json) printLine(entry);
  };
  const fail = (key, message) => {
    result.missing.push(message);
    record(key, { key, status: "skipped", note: message });
  };

  if (!options.json) {
    console.log("");
    console.log("gridxcel bench  node " + process.version + "  gc=" + (result.gcAvailable ? "on" : "off") +
      "  input=" + result.file + " (" + mb(inputBytes) + " MB, " + rows + "x" + cols + ")");
    console.log("");
    printHeader();
  }

  /* (i) streaming parse FIRST: it must be measured against a clean heap, ---
     otherwise the book kept alive by the full-parse stage inflates it. */
  const read = await tryImport("src/read.js");
  let parsedBook = null;
  const each = await findEachRow();
  if (!each.fn) {
    fail("eachRow", "NOT YET AVAILABLE: eachRow (src/index.js, src/api.js, src/read.js)");
  } else {
    let seen = 0;
    const m = await measure(async () => {
      seen = 0;
      await callEachRow(each.fn, file, (row) => { seen++; }, "volume");
    }, { gc: true });
    record("eachRow", m.error
      ? { key: "eachRow", status: "error", note: String(m.error?.message ?? m.error) + " (from " + each.from + ")" }
      : {
          key: "eachRow", status: "ok", ms: m.ms, rowsPerSec: seen / (m.ms / 1000),
          peakHeap: m.peakHeap, peakHeapExcess: m.peakHeapExcess, peakRssExcess: m.peakRssExcess,
          heapRetained: m.heapRetained, peakRss: m.peakRss, samples: m.samples,
          rowsSeen: seen, note: "streaming, from " + each.from,
        });
  }

  /* (i-b) full parse ----------------------------------------------------- */
  if (!read.ok) {
    fail("readWorkbook", read.missing);
  } else {
    const m = await measure(() => read.mod.readWorkbook(file, { cellFormula: true }), { gc: true });
    if (m.error) {
      record("readWorkbook", { key: "readWorkbook", status: "error", note: String(m.error?.message ?? m.error) });
    } else {
      parsedBook = m.value;
      m.value = null; // the stage result must not keep the book alive
      const cells = countCells(parsedBook);
      record("readWorkbook", {
        key: "readWorkbook",
        status: "ok",
        ms: m.ms,
        rowsPerSec: rows / (m.ms / 1000),
        peakHeap: m.peakHeap,
        peakHeapExcess: m.peakHeapExcess,
        peakRssExcess: m.peakRssExcess,
        heapRetained: m.heapRetained,
        peakRss: m.peakRss,
        samples: m.samples,
        cells,
        sheets: parsedBook.sheets?.length ?? 0,
        note: "full materialised parse",
      });
    }
  }

  /* (iii) CSV export, while the parsed book is still alive -------------- */
  const csv = await tryImport("src/csv.js");
  if (!csv.ok || typeof csv.mod.sheetToCSV !== "function") {
    fail("sheetToCSV", csv.ok ? "NOT YET AVAILABLE: sheetToCSV export in src/csv.js" : csv.missing);
  } else if (!parsedBook?.sheets?.length) {
    fail("sheetToCSV", "skipped: no parsed workbook to export (readWorkbook unavailable)");
  } else {
    const sheet = parsedBook.sheets[0];
    const m = await measure(async () => csv.mod.sheetToCSV(sheet, {}), { gc: true });
    record("sheetToCSV", m.error
      ? { key: "sheetToCSV", status: "error", note: String(m.error?.message ?? m.error) }
      : {
          key: "sheetToCSV", status: "ok", ms: m.ms, rowsPerSec: rows / (m.ms / 1000),
          peakHeap: m.peakHeap, peakHeapExcess: m.peakHeapExcess, heapRetained: m.heapRetained,
          peakRss: m.peakRss, samples: m.samples,
          outBytes: typeof m.value === "string" ? Buffer.byteLength(m.value) : 0,
        });
    m.value = null;
  }

  /* (ii) writes --------------------------------------------------------- */
  parsedBook = null;
  gcNow();
  const ww = await tryImport("src/workbook-writer.js");
  const write = await tryImport("src/write.js");
  const { volumeRows } = await import("./specs.mjs");

  if (!ww.ok) {
    fail("WorkbookWriter", ww.missing);
  } else {
    const shape = detectWriterShape(ww.mod);
    const m = await measure(async () => {
      const w = new ww.mod.WorkbookWriter({ compressionLevel: 6 });
      const sw = w.addSheet("volume");
      let i = 1;
      for (const line of volumeRows(rows, cols)) {
        sw.addRow(i++, line);
      }
      sw.end();
      return w.finish();
    }, { gc: true });
    record("WorkbookWriter", m.error
      ? { key: "WorkbookWriter", status: "error", note: String(m.error?.message ?? m.error) }
      : {
          key: "WorkbookWriter", status: "ok", ms: m.ms, rowsPerSec: rows / (m.ms / 1000),
          peakHeap: m.peakHeap, heapRetained: m.heapRetained, peakRss: m.peakRss, samples: m.samples,
          outBytes: m.value?.length ?? 0, shape,
        });
  }

  if (options.materialise && write.ok) {
    // Materialise the whole book first: this measures the "everything in
    // memory" write path separately from the writer itself.
    const build = await measure(async () => {
      const sheets = [{ name: "volume", rows: [], originRow: 1, originCol: 0, columnCount: cols, rowCount: rows, merges: [], columns: [], rowInfo: new Map() }];
      let i = 0;
      for (const line of volumeRows(rows, cols)) sheets[0].rows.push(line), i++;
      return { sheets, dateSystem: 1900, numFmts: new Map(), styles: [], props: {}, definedNames: [] };
    }, { gc: true });
    record("buildBookInMemory", {
      key: "buildBookInMemory", status: build.error ? "error" : "ok",
      ms: build.ms, peakHeap: build.peakHeap, heapRetained: build.heapRetained, peakRss: build.peakRss,
      note: build.error ? String(build.error?.message ?? build.error) : "rows materialised before writing",
    });
    if (!build.error) {
      const m = await measure(async () => write.mod.writeWorkbookSync(build.value, { type: "buffer" }), { gc: true });
      record("writeWorkbookSync", m.error
        ? { key: "writeWorkbookSync", status: "error", note: String(m.error?.message ?? m.error) }
        : { key: "writeWorkbookSync", status: "ok", ms: m.ms, rowsPerSec: rows / (m.ms / 1000), peakHeap: m.peakHeap, heapRetained: m.heapRetained, peakRss: m.peakRss, outBytes: m.value?.length ?? 0 });
    } else {
      fail("writeWorkbookSync", "skipped: in-memory book could not be built");
    }
  } else if (!write.ok) {
    fail("writeWorkbookSync", write.missing);
  }

  result.durationMs = undefined;
  delete result.durationMs;
  if (!options.json) {
    console.log("");
    printSummary(result);
  }
  const outFile = path.join(ROOT, "bench", "bench-" + rows + "x" + cols + ".json");
  writeFileSync(outFile, JSON.stringify(result, null, 1));
  result.report = path.relative(ROOT, outFile).replace(/\\/g, "/");
  return result;
}

/**
 * eachRow(input, options) is an async generator of { sheet, r, cells }
 * (src/index.d.ts). Iterate it when it is async-iterable; fall back to the
 * callback shapes so a different implementation still gets measured instead
 * of silently reporting zero rows.
 */
export async function callEachRow(fn, file, onRow, sheetName) {
  const opts = sheetName ? { sheets: [sheetName] } : {};
  const iterable = fn(file, opts);
  if (iterable && typeof iterable[Symbol.asyncIterator] === "function") {
    let n = 0;
    for await (const ev of iterable) { onRow(ev); n++; }
    return n;
  }
  const errors = [];
  for (const attempt of [() => fn(file, sheetName, onRow), () => fn(file, onRow), () => fn(file, opts, onRow)]) {
    try {
      return await attempt();
    } catch (err) {
      errors.push(err?.message ?? String(err));
    }
  }
  throw new Error("unsupported eachRow shape (" + fn.length + " params); attempts failed: " + errors.join(" | "));
}

function countCells(book) {
  let n = 0;
  for (const sheet of book?.sheets ?? []) {
    for (const row of sheet.rows ?? []) {
      if (!row) continue;
      for (const cell of row) if (cell) n++;
    }
  }
  return n;
}

function detectWriterShape(mod) {
  const out = {};
  try {
    const w = new mod.WorkbookWriter({ compressionLevel: 6 });
    out.workbookWriter = typeof w.addSheet === "function" ? "addSheet" : "unknown";
    out.finish = typeof w.finish === "function" ? "finish" : "unknown";
    out.parts = typeof w.parts === "function" ? "parts" : "unknown";
    const s = w.addSheet("probe");
    out.sheetWriter = typeof s.addRow === "function" ? "addRow" : "unknown";
    out.end = typeof s.end === "function" ? "end" : "unknown";
    out.rowCountProp = s.rowCount !== undefined ? typeof s.rowCount : "missing";
  } catch (err) {
    out.error = String(err?.message ?? err);
  }
  return out;
}

function printHeader() {
  console.log(
    "STAGE".padEnd(20) + "STATUS".padEnd(9) + "MS".padStart(10) + "ROWS/S".padStart(12) +
    "PEAKHEAP".padStart(11) + "RETAINED".padStart(11) + "PEAKRSS".padStart(10) + "  NOTE"
  );
  console.log("-".repeat(110));
}

function printLine(e) {
  if (e.status !== "ok") {
    console.log(String(e.key).padEnd(20) + String(e.status).padEnd(9) + "".padStart(10) + "".padStart(12) + "".padStart(11) + "".padStart(11) + "".padStart(10) + "  " + (e.note ?? ""));
    return;
  }
  console.log(
    String(e.key).padEnd(20) + "ok".padEnd(9) +
    e.ms.toFixed(1).padStart(10) +
    (e.rowsPerSec ? Math.round(e.rowsPerSec).toLocaleString("en-US") : "-").padStart(12) +
    (mb(e.peakHeap) + " MB").padStart(11) +
    (mb(e.heapRetained) + " MB").padStart(11) +
    (mb(e.peakRss) + " MB").padStart(10) +
    "  " + (e.note ?? (e.cells ? e.cells.toLocaleString("en-US") + " cells" : ""))
  );
}

function printSummary(result) {
  const r = result.results;
  console.log("-".repeat(110));
  const rd = r.readWorkbook, er = r.eachRow;
  if (rd?.status === "ok" && er?.status === "ok") {
    console.log("streaming vs full parse: eachRow peak heap EXCESS is " +
      (er.peakHeapExcess / rd.peakHeapExcess).toFixed(2) + "x of readWorkbook (" +
      mb(er.peakHeapExcess) + " MB vs " + mb(rd.peakHeapExcess) + " MB above each stage's own baseline); speed " +
      (rd.ms / er.ms).toFixed(2) + "x");
    console.log("  retained after forced GC: eachRow " + mb(er.heapRetained) + " MB, readWorkbook " + mb(rd.heapRetained) + " MB");
  }
  if (rd?.status === "ok") {
    console.log("throughput: parse " + Math.round(rd.rowsPerSec).toLocaleString("en-US") + " rows/s, " +
      (mb(result.inputBytes) + " MB input, " + (mb(rd.peakHeap) / (result.inputBytes / 1048576)).toFixed(2) + "x input size peak heap"));
  }
  if (r.WorkbookWriter?.status === "ok") {
    console.log("write: " + Math.round(r.WorkbookWriter.rowsPerSec).toLocaleString("en-US") + " rows/s, peak heap " + mb(r.WorkbookWriter.peakHeap) + " MB, output " + mb(r.WorkbookWriter.outBytes) + " MB");
  }
  if (r.sheetToCSV?.status === "ok") {
    console.log("csv: " + Math.round(r.sheetToCSV.rowsPerSec).toLocaleString("en-US") + " rows/s, output " + mb(r.sheetToCSV.outBytes) + " MB");
  }
  for (const m of result.missing) console.log(m);
  console.log("");
}

/* ------------------------------------------------------------------ main */

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const opts = parseArgs(process.argv.slice(2));
  ensureGc(process.argv.slice(2));
  const result = await runBench(opts);
  console.log("#JSON# " + JSON.stringify(result));
  process.exit(result.missing.length ? 2 : 0);
}
