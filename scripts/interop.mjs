#!/usr/bin/env node
/**
 * gridxcel interop harness - the master independent-verification entry point.
 *
 *   (a) WRITE PATH: gridxcel writes a workbook from a frozen-shape Book, then
 *       PYTHON (openpyxl + zipfile) reads it back and every cell value, type,
 *       number format, merge, sheet name and row count is compared against the
 *       spec. zipfile.ZipFile.testzip() must return None.
 *   (b) READ PATH: gridxcel reads the python-written fixtures and the result is
 *       diffed against an openpyxl dump of the very same file.
 *   (c) One machine-readable JSON summary + a human table; non-zero exit on any
 *       mismatch.
 *
 * Exit codes
 *   0  every check ran and passed
 *   1  at least one mismatch (bug found)
 *   2  a required module or the oracle is not available yet (nothing proven)
 *
 * Usage: node scripts/interop.mjs [--json] [--only=write|read|self]
 *        [--medium-rows=N] [--keep] [--quiet]
 *
 * No gridxcel module is ever used to check another gridxcel module: the only
 * gridxcel imports are the modules under test (src/read.js, src/write.js).
 */

import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  cachedMatches,
  canonFromGridxcel,
  canonFromOracle,
  colName,
  diffCanon,
  generate,
  inspect,
  normFormula,
  normRange,
  oracleAvailable,
  pythonPath,
} from "./oracle-client.mjs";
import { writeSpecs } from "./specs.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_ROOT = path.resolve(HERE, "..");

/* ------------------------------------------------------------ utilities */

const fmt = (v) => {
  if (v === undefined) return "-";
  if (typeof v === "string") return JSON.stringify(v);
  return JSON.stringify(v);
};

/** Deterministic fixture order; the generated volume file is handled separately. */
function listFixtures(dir) {
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith(".xlsx") && !f.startsWith("~$"))
    .filter((f) => statSync(path.join(dir, f)).size < 8 * 1024 * 1024)
    .sort();
}

/** Import a module under test; returns { ok, mod, error }. */
async function tryImport(absPath) {
  if (!existsSync(absPath)) return { ok: false, missing: true, error: "NOT YET AVAILABLE: " + rel(absPath) };
  try {
    return { ok: true, mod: await import(pathToFileURL(absPath).href) };
  } catch (err) {
    return { ok: false, missing: false, error: "IMPORT FAILED: " + rel(absPath) + ": " + (err?.message ?? err) };
  }
}

function rel(p) {
  return path.relative(DEFAULT_ROOT, p).replace(/\\/g, "/");
}

/* ------------------------------------------------- (a) write-path checks */

/**
 * Compare a python dump of a gridxcel-written file against the spec that
 * produced it. Returns { mismatches, checks, dateConvention }.
 */
export function checkWrittenFile(spec, doc) {
  const mismatches = [];
  const notes = [];
  let checks = 0;
  let dateConvention = null;
  const push = (m) => mismatches.push(m);

  if (doc.zip.testzip !== null) {
    push({ check: "zip", target: spec.name, field: "testzip", expected: null, actual: doc.zip.testzip, detail: "zipfile.testzip() reported a bad entry" });
  }
  checks++;
  if ((doc.zip.duplicateEntries ?? []).length) {
    push({ check: "zip", target: spec.name, field: "duplicateEntries", expected: [], actual: doc.zip.duplicateEntries });
  }

  const byName = new Map(doc.sheets.map((s) => [s.name, s]));
  const expectedNames = spec.expect.map((s) => s.name);
  checks++;
  if (expectedNames.join("|") !== doc.book.sheetNames.join("|")) {
    push({ check: "sheets", target: spec.name, field: "sheetNames", expected: expectedNames.join(","), actual: doc.book.sheetNames.join(",") });
  }

  for (const es of spec.expect) {
    const os = byName.get(es.name);
    checks++;
    if (!os) {
      push({ check: "sheets", target: spec.name + "/" + es.name, field: "sheet", expected: "present", actual: "missing" });
      continue;
    }
    checks++;
    if (os.trueMaxRow !== es.rowCount) {
      push({ check: "geometry", target: spec.name + "/" + es.name, field: "rowCount", expected: es.rowCount, actual: os.trueMaxRow });
    }
    checks++;
    if (os.trueMaxCol !== es.columnCount) {
      push({ check: "geometry", target: spec.name + "/" + es.name, field: "columnCount", expected: es.columnCount, actual: os.trueMaxCol });
    }
    checks++;
    const actualMerges = os.merges.map(normRange).sort();
    if (actualMerges.join("|") !== es.merges.join("|")) {
      push({ check: "merges", target: spec.name + "/" + es.name, field: "merges", expected: es.merges.join(","), actual: actualMerges.join(",") });
    }

    const cellsByRef = new Map(os.cells.map((c) => [c.ref, c]));
    for (const ec of es.cells) {
      checks++;
      const oc = cellsByRef.get(ec.ref);
      const where = spec.name + "/" + es.name + "!" + ec.ref;
      if (!oc) {
        push({ check: "cell", target: where, field: "cell", expected: describeExpect(ec), actual: "missing" });
        continue;
      }
      if (ec.kind === "empty") continue;

      if (ec.formula) {
        checks++;
        if (normFormula(oc.formula ?? "") !== ec.formula) {
          push({ check: "formula", target: where, field: "formula", expected: ec.formula, actual: oc.formula ?? "(none)" });
        }
        if (oc.cached === null || oc.cached === undefined) {
          notes.push({ check: "formula", target: where, field: "cachedValue", detail: "writer emitted no cached result; Excel must recalc (recalc option)" });
        } else if (ec.value !== undefined && !cachedMatches(oc.cached, { value: ec.value })) {
          push({ check: "formula", target: where, field: "cachedValue", expected: ec.value, actual: oc.cached });
        }
      }

      if (ec.kind === "date") {
        checks++;
        const serial = oc.serial;
        const dUTC = Math.abs((serial ?? NaN) - ec.serialUTC) <= 1e-9;
        const dLocal = Math.abs((serial ?? NaN) - ec.serialLocal) <= 1e-9;
        if (dUTC) dateConvention = dateConvention === "local" ? "mixed" : "utc";
        else if (dLocal) dateConvention = dateConvention === "utc" ? "mixed" : "local";
        else {
          push({
            check: "date", target: where, field: "serial",
            expected: "utc=" + ec.serialUTC.toFixed(9) + " local=" + ec.serialLocal.toFixed(9),
            actual: serial === null || serial === undefined ? "(no date cell)" : String(serial),
            detail: "date did not survive as an Excel serial under either convention; python reports " + describeOracleCell(oc),
          });
        }
      } else if (ec.kind === "error") {
        checks++;
        if (oc.kind !== "error") {
          push({ check: "type", target: where, field: "kind", expected: "error", actual: oc.kind });
        }
        checks++;
        const actualErr = oc.value === null || oc.value === undefined ? "" : String(oc.value);
        if (actualErr !== String(ec.value)) {
          push({ check: "value", target: where, field: "errorText", expected: String(ec.value), actual: actualErr });
        }
      } else if (ec.formula && oc.kind === "formula") {
        // openpyxl reports kind "formula" with the formula text as the value;
        // the result, if the writer stored one, is in `cached` (checked above).
        checks++;
      } else {
        checks++;
        if (oc.kind !== ec.kind) {
          push({ check: "type", target: where, field: "kind", expected: ec.kind, actual: oc.kind });
        }
        checks++;
        const ov = oracleScalar(oc);
        if (!scalarEqual(ec.value, ov)) {
          push({ check: "value", target: where, field: "value", expected: fmt(ec.value), actual: fmt(oc.value), actualRaw: fmt(ov) });
        }
      }

      checks++;
      const actualFormat = oc.format ?? "General";
      if (actualFormat !== ec.format) {
        push({ check: "numfmt", target: where, field: "numberFormat", expected: ec.format, actual: actualFormat });
      }
    }
  }
  return { mismatches, checks, notes, dateConvention };
}

function describeExpect(ec) {
  return JSON.stringify({ kind: ec.kind, value: ec.value ?? ec.serialUTC, format: ec.format });
}

function describeOracleCell(oc) {
  return JSON.stringify({ kind: oc.kind, value: oc.value, format: oc.format, serial: oc.serial });
}

/** oracle value -> plain JS scalar (dates and errors carry wrappers). */
function oracleScalar(oc) {
  const v = oc.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") return v;
  return v;
}

function scalarEqual(a, b) {
  if (a === b) return true;
  if (typeof a === "number" && typeof b === "number") return a === b;
  if (a === null || a === undefined) return b === null || b === undefined;
  if (b === null || b === undefined) return false;
  if (typeof a === "object" || typeof b === "object") return JSON.stringify(a) === JSON.stringify(b);
  return String(a) === String(b);
}

/* -------------------------------------------------- (b) read-path checks */

export async function checkReadFile(root, mod, file, opts = {}) {
  let doc;
  try {
    doc = inspect(root, file, { cached: true });
  } catch (err) {
    // python itself refuses the file: it cannot serve as a reference for this
    // fixture (deliberately malformed reader-eng fixtures live here). Report it
    // as a skip, never as a gridxcel mismatch.
    if (err?.oracleUnreadable) return { file: rel(file), oracleError: err.oracleMessage ?? err.message, skipped: true };
    throw err;
  }
  const expected = canonFromOracle(doc);
  const book = await mod.readWorkbook(file, { cellFormula: true, ...(opts.readOptions ?? {}) });
  const actual = canonFromGridxcel(book, {
    builtins: doc.book.builtinFormats,
    numFmts: doc.book.numFmts,
  });
  const diff = diffCanon(expected, actual, { maxDetail: opts.maxDetail ?? 25 });
  return { file: rel(file), doc, expected, actual, diff };
}

/* ------------------------------------------------------------- the run */

export async function runInterop(options = {}) {
  const root = options.root ?? DEFAULT_ROOT;
  const started = Date.now();
  const summary = {
    harness: "gridxcel-interop",
    version: 1,
    root,
    python: pythonPath(),
    status: "pass",
    dateConvention: null,
    write: { specs: [], checks: 0, mismatches: [], notes: [] },
    read: { files: [], checks: 0, mismatches: [], warnings: [], skipped: [] },
    self: { checks: 0, mismatches: [] },
    missing: [],
    durationMs: 0,
  };

  const only = options.only ?? "all";
  const outDir = options.outDir ?? path.join(os.tmpdir(), "gridxcel-interop");
  mkdirSync(outDir, { recursive: true });

  if (!oracleAvailable()) {
    summary.status = "not-available";
    summary.missing.push("python oracle (openpyxl/lxml) at " + pythonPath());
    summary.durationMs = Date.now() - started;
    return summary;
  }

  runSelfCheck(root, summary);

  const writeMod = await tryImport(path.join(root, "src", "write.js"));
  const readMod = await tryImport(path.join(root, "src", "read.js"));
  if (!writeMod.ok && !readMod.ok) {
    for (const m of [writeMod, readMod]) summary.missing.push(m.error);
    summary.status = "not-available";
    summary.durationMs = Date.now() - started;
    return summary;
  }
  if (!writeMod.ok) summary.missing.push(writeMod.error);
  if (!readMod.ok) summary.missing.push(readMod.error);

  const wantWrite = (only === "all" || only === "write") && writeMod.ok;
  const wantRead = (only === "all" || only === "read") && readMod.ok;
  // "self" checks the harness itself (comparator sanity); always worth running.
  if (wantWrite) {
    const mod = writeMod.mod;
    for (const spec of writeSpecs()) {
      const outFile = path.join(outDir, "gridxcel-" + spec.name + ".xlsx");
      rmSync(outFile, { force: true });
      let wrote = null;
      let error = null;
      try {
        const bytes = mod.writeWorkbookSync(spec.book, { type: "buffer", compressionLevel: 6 });
        if (!bytes || bytes.length === 0) throw new Error("writeWorkbookSync returned no bytes");
        writeFileSync(outFile, bytes);
        wrote = bytes.length;
      } catch (err) {
        error = (err?.stack ?? String(err)).split("\n").slice(0, 3).join(" | ");
      }
      const entry = { spec: spec.name, file: rel(outFile), bytes: wrote, checks: 0, ok: false, mismatches: [], notes: [] };
      if (error) {
        entry.mismatches.push({ check: "write", target: spec.name, field: "writeWorkbookSync", expected: "bytes", actual: error });
      } else {
        const doc = inspect(root, outFile, { cached: true });
        const res = checkWrittenFile(spec, doc);
        entry.checks = res.checks;
        entry.mismatches = res.mismatches;
        entry.notes = res.notes;
        if (res.dateConvention && !summary.dateConvention) summary.dateConvention = res.dateConvention;
        else if (res.dateConvention && summary.dateConvention && res.dateConvention !== summary.dateConvention) summary.dateConvention = "mixed";
      }
      entry.ok = entry.mismatches.length === 0;
      summary.write.checks += entry.checks;
      summary.write.mismatches.push(...entry.mismatches);
      summary.write.notes.push(...entry.notes);
      summary.write.specs.push(entry);
    }
  }

  if (wantRead) {
    const mod = readMod.mod;
    const fixtureDir = path.join(root, "test", "fixtures");
    const files = listFixtures(fixtureDir);
    const mediumRows = options.mediumRows ?? 2000;
    if (mediumRows > 0) {
      const medium = path.join(root, "bench", "oracle-medium-" + mediumRows + "x8.xlsx");
      mkdirSync(path.dirname(medium), { recursive: true });
      if (!existsSync(medium)) generate(root, medium, mediumRows, 8);
      files.push(medium);
    }
    for (const name of files) {
      const file = path.isAbsolute(name) ? name : path.join(fixtureDir, name);
      const entry = { file: rel(file), checks: 0, ok: false, mismatches: [], warnings: [] };
      try {
        const res = await checkReadFile(root, mod, file, { maxDetail: 25 });
        if (res.skipped) {
          entry.skipped = true;
          entry.ok = true;
          entry.reason = "openpyxl cannot read this fixture, so there is no oracle for it: " + res.oracleError;
          summary.read.skipped.push({ file: rel(file), reason: res.oracleError });
          summary.read.files.push(entry);
          continue;
        }
        entry.checks = res.expected.sheets.reduce((n, s) => n + s.cells.length + 3, 0);
        entry.mismatches = res.diff.mismatches.filter((m) => m.severity !== "warn");
        entry.warnings = res.diff.mismatches.filter((m) => m.severity === "warn");
        entry.sheets = res.expected.sheetNames;
        entry.cells = res.expected.sheets.reduce((n, s) => n + s.cells.length, 0);
      } catch (err) {
        entry.mismatches.push({ check: "read", target: rel(file), field: "readWorkbook", expected: "Book", actual: (err?.stack ?? String(err)).split("\n").slice(0, 3).join(" | ") });
      }
      entry.ok = entry.mismatches.length === 0;
      summary.read.checks += entry.checks;
      summary.read.mismatches.push(...entry.mismatches);
      summary.read.warnings.push(...entry.warnings);
      summary.read.files.push(entry);
    }

    // Buffer / stream input paths must agree with the path input on one file.
    const probe = path.join(fixtureDir, "openpyxl-written.xlsx");
    if (existsSync(probe)) {
      const fromPath = await checkReadFile(root, mod, probe, {});
      const buf = readFileSync(probe);
      let bufDiff = null;
      let bufError = null;
      try {
        const book = await mod.readWorkbook(new Uint8Array(buf), { cellFormula: true });
        const actual = canonFromGridxcel(book, {
          builtins: fromPath.doc.book.builtinFormats,
          numFmts: fromPath.doc.book.numFmts,
        });
        bufDiff = diffCanon(fromPath.expected, actual, { maxDetail: 10 });
      } catch (err) {
        bufError = (err?.stack ?? String(err)).split("\n").slice(0, 3).join(" | ");
      }
      const entry = { file: rel(probe) + " [Uint8Array input]", checks: 1, ok: false, mismatches: [], warnings: [] };
      if (bufError) entry.mismatches.push({ check: "read", target: rel(probe), field: "readWorkbook(Uint8Array)", expected: "same as path input", actual: bufError });
      else entry.mismatches = bufDiff.mismatches.filter((m) => m.severity !== "warn");
      entry.ok = entry.mismatches.length === 0;
      summary.read.checks += 1;
      summary.read.mismatches.push(...entry.mismatches);
      summary.read.files.push(entry);
    }
  }

  const total = summary.write.mismatches.length + summary.read.mismatches.length + summary.self.mismatches.length;
  if (total > 0) summary.status = "fail";
  else if (summary.missing.length || !wantWrite || !wantRead) summary.status = "partial";
  else summary.status = "pass";
  summary.durationMs = Date.now() - started;
  return summary;
}

/**
 * Comparator self-check: identical data must produce zero mismatches and a
 * mutated copy must produce exactly the injected mismatches. Without this a
 * green run could just mean the comparator never fires.
 */
export function runSelfCheck(root, summary) {
  const probe = path.join(root, "test", "fixtures", "openpyxl-written.xlsx");
  const doc = inspect(root, probe, { cached: true });
  const canon = canonFromOracle(doc);
  const clean = diffCanon(canon, canon);
  summary.self.checks++;
  if (clean.mismatches.length !== 0) {
    summary.self.mismatches.push({ check: "self", target: "identical-copy", field: "falsePositive", expected: 0, actual: clean.mismatches.length });
  }
  const mutated = JSON.parse(JSON.stringify(canon));
  mutated.sheets[0].cells[0].value = "MUTATED";
  mutated.sheets[0].merges = ["Z1:Z2"];
  mutated.sheets[0].rowCount = canon.sheets[0].rowCount + 41;
  const dirty = diffCanon(canon, mutated);
  summary.self.checks++;
  const fields = new Set(dirty.mismatches.map((m) => m.field));
  for (const want of ["value", "merges", "rowCount"]) {
    if (!fields.has(want)) {
      summary.self.mismatches.push({ check: "self", target: "mutated-copy", field: "truePositive:" + want, expected: "detected", actual: "not detected" });
    }
  }
  summary.self.checks++;
  return summary.self;
}

/* -------------------------------------------------------------- report */

export function printReport(summary, { quiet = false } = {}) {
  const rows = [];
  for (const s of summary.write.specs) {
    rows.push(["write", s.spec, s.ok ? "ok" : "FAIL", s.ok ? s.checks + " checks, " + (s.bytes ?? 0) + " bytes" : s.mismatches.length + " mismatches"]);
  }
  if (summary.write.specs.length === 0) rows.push(["write", "(skipped)", "-", "writer not available"]);
  for (const f of summary.read.files) {
    if (f.skipped) { rows.push(["read", f.file, "skip", f.reason]); continue; }
    rows.push(["read", f.file, f.ok ? "ok" : "FAIL", f.ok ? f.checks + " checks" + (f.warnings.length ? ", " + f.warnings.length + " warns" : "") : f.mismatches.length + " mismatches"]);
  }
  if (summary.read.files.length === 0) rows.push(["read", "(skipped)", "-", "reader not available"]);
  for (const m of summary.self.mismatches) rows.push(["self", m.target, "FAIL", m.field + " expected " + m.expected + " got " + m.actual]);
  if (summary.self.mismatches.length === 0) rows.push(["self", "comparator", "ok", summary.self.checks + " checks"]);

  const width = (i) => Math.max(...rows.map((r) => String(r[i]).length), 6);
  const w0 = width(0), w1 = width(1);
  if (!quiet) {
    console.log("");
    console.log("gridxcel interop report  (python: " + summary.python + ")");
    console.log("-".repeat(w0 + w1 + 40));
    console.log("CHECK ".padEnd(w0 + 1) + "TARGET".padEnd(w1 + 1) + "RESULT  DETAIL");
    console.log("-".repeat(w0 + w1 + 40));
    for (const r of rows) {
      console.log(String(r[0]).padEnd(w0 + 1) + String(r[1]).padEnd(w1 + 1) + String(r[2]).padEnd(8) + String(r[3]));
    }
    console.log("-".repeat(w0 + w1 + 40));
  }
  const all = [...summary.write.mismatches, ...summary.read.mismatches, ...summary.self.mismatches];
  if (all.length && !quiet) {
    console.log("");
    console.log("MISMATCHES (" + all.length + ")");
    for (const m of all.slice(0, 60)) {
      console.log("  [" + m.check + "] " + mismatchTarget(m) + " " + m.field + ": expected " + fmt(m.expected) + " actual " + fmt(m.actual));
      console.log("      sheet=" + (m.sheet ?? "-") + " ref=" + (m.ref ?? "-") + " xml=" + (m.xml ?? "(none)"));
    }
    if (all.length > 60) console.log("  ... " + (all.length - 60) + " more in the JSON summary");
  }
  if (summary.write.notes.length && !quiet) {
    console.log("");
    console.log("NOTES (" + summary.write.notes.length + ")");
    for (const n of summary.write.notes.slice(0, 20)) console.log("  [" + n.check + "] " + n.target + " " + n.field + ": " + n.detail);
  }
  if (summary.missing.length && !quiet) {
    console.log("");
    for (const m of summary.missing) console.log(m);
  }
  if (!quiet) {
    console.log("");
    console.log(
      "status=" + summary.status +
      "  writeChecks=" + summary.write.checks +
      "  readChecks=" + summary.read.checks +
      "  mismatches=" + all.length +
      "  dateConvention=" + (summary.dateConvention ?? "n/a") +
      "  " + summary.durationMs + "ms"
    );
  }
}

/** "sheet!REF" for read mismatches, target string for write mismatches. */
export function mismatchTarget(m) {
  if (m.target) return m.target;
  if (m.sheet && m.ref) return m.sheet + "!" + m.ref;
  if (m.sheet) return m.sheet;
  if (m.file) return m.file;
  return "(unknown)";
}

function parseArgs(argv) {
  const opts = { only: "all", json: false, quiet: false, keep: false };
  for (const a of argv) {
    if (a === "--json") opts.json = true;
    else if (a === "--quiet") opts.quiet = true;
    else if (a === "--keep") opts.keep = true;
    else if (a.startsWith("--only=")) opts.only = a.slice(7);
    else if (a.startsWith("--medium-rows=")) opts.mediumRows = Number(a.slice(14));
    else if (a.startsWith("--root=")) opts.root = path.resolve(a.slice(7));
  }
  return opts;
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const opts = parseArgs(process.argv.slice(2));
  const summary = await runInterop(opts);
  if (!opts.json) printReport(summary, { quiet: opts.quiet });
  console.log("#JSON# " + JSON.stringify(summary));
  process.exit(summary.status === "fail" ? 1 : summary.status === "pass" ? 0 : 2);
}
