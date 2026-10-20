/**
 * Interop suite: the same checks as scripts/interop.mjs (python oracle on the
 * write path, live openpyxl diff on the read path, pre-recorded golden
 * manifest) wired into node:test so "node --test test/" covers them.
 *
 * Every check compares gridxcel against python, never against another
 * gridxcel module. When src/read.js or src/write.js is missing the affected
 * test skips with a "NOT YET AVAILABLE" reason instead of failing the suite.
 */

import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { existsSync, mkdirSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";

import { checkReadFile, checkWrittenFile, runSelfCheck } from "../scripts/interop.mjs";
import { generate, inspect, oracleAvailable } from "../scripts/oracle-client.mjs";
import { writeSpecs } from "../scripts/specs.mjs";
import { checkGoldenManifest, checkGoldenStream, goldenFile, loadGolden } from "../scripts/golden.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const READ_PATH = path.join(ROOT, "src", "read.js");
const WRITE_PATH = path.join(ROOT, "src", "write.js");
const FIXTURE_DIR = path.join(ROOT, "test", "fixtures");
const MEDIUM_ROWS = Number(process.env.GRIDXCEL_INTEROP_ROWS ?? 2000);

const SKIP_READ = existsSync(READ_PATH) ? false : "NOT YET AVAILABLE: src/read.js";
const SKIP_WRITE = existsSync(WRITE_PATH) ? false : "NOT YET AVAILABLE: src/write.js";
const SKIP_ORACLE = oracleAvailable() ? false : "NOT YET AVAILABLE: python oracle";

const load = (p) => import(pathToFileURL(p).href);

/** Fail with the first few mismatches spelled out. */
function assertNoMismatches(mismatches, label) {
  if (mismatches.length === 0) return;
  const detail = mismatches
    .slice(0, 12)
    .map((m) => "  [" + m.check + "] " + m.target + " " + m.field + ": expected " + JSON.stringify(m.expected) + " got " + JSON.stringify(m.actual))
    .join("\n");
  assert.fail(label + ": " + mismatches.length + " mismatch(es)\n" + detail);
}

describe("interop: write path verified by python", { skip: SKIP_WRITE || SKIP_ORACLE }, () => {
  for (const spec of writeSpecs()) {
    test("gridxcel writes '" + spec.name + "' and openpyxl sees every cell/format/merge/row", () => {
      const mod = load(WRITE_PATH);
      return mod.then(async (m) => {
        const out = path.join(ROOT, "bench", "interop-" + spec.name + ".xlsx");
        mkdirSync(path.dirname(out), { recursive: true });
        const bytes = m.writeWorkbookSync(spec.book, { type: "buffer", compressionLevel: 6 });
        assert.ok(bytes && bytes.length > 0, "writeWorkbookSync returned no bytes");
        writeFileSync(out, bytes);
        const doc = inspect(ROOT, out, { cached: true });
        assert.equal(doc.zip.testzip, null, "zipfile.ZipFile.testzip() must return None");
        const res = checkWrittenFile(spec, doc);
        assertNoMismatches(res.mismatches, "write spec " + spec.name);
      });
    });
  }
});

describe("interop: read path diffed against live openpyxl", { skip: SKIP_READ || SKIP_ORACLE }, () => {
  const fixtures = existsSync(FIXTURE_DIR)
    ? readdirSync(FIXTURE_DIR).filter((f) => f.endsWith(".xlsx")).sort()
    : [];

  for (const file of fixtures) {
    test("readWorkbook(" + file + ") matches openpyxl on values, types, formats, merges, geometry", async (t) => {
      const mod = await load(READ_PATH);
      const res = await checkReadFile(ROOT, mod, path.join(FIXTURE_DIR, file), { maxDetail: 30 });
      if (res.skipped) {
        t.skip("openpyxl cannot read this fixture (no oracle): " + res.oracleError);
        return;
      }
      assertNoMismatches(res.diff.mismatches.filter((m) => m.severity !== "warn"), "read " + file);
    });
  }

  test("readWorkbook on a generated " + MEDIUM_ROWS + "x8 workbook matches openpyxl (volume)", async () => {
    const mod = await load(READ_PATH);
    const medium = path.join(ROOT, "bench", "oracle-medium-" + MEDIUM_ROWS + "x8.xlsx");
    mkdirSync(path.dirname(medium), { recursive: true });
    if (!existsSync(medium)) generate(ROOT, medium, MEDIUM_ROWS, 8);
    const res = await checkReadFile(ROOT, mod, medium, { maxDetail: 20 });
    assertNoMismatches(res.diff.mismatches.filter((m) => m.severity !== "warn"), "read volume file");
    assert.ok(res.actual.sheets[0].cells.length > MEDIUM_ROWS, "volume file should expose more cells than rows");
  });
});

describe("interop: pre-recorded golden manifest", { skip: SKIP_READ }, () => {
  test("golden.xlsx matches test/fixtures/golden.json (cellDates:true)", async () => {
    const mod = await load(READ_PATH);
    const book = await mod.readWorkbook(goldenFile(ROOT), { cellDates: true, cellFormula: true });
    const res = checkGoldenManifest(book, loadGolden(ROOT), { cellDates: true });
    assertNoMismatches(res.mismatches, "golden manifest (dates)");
    assert.equal(book.sheets.length, 4, "golden.xlsx has 4 sheets");
  });

  test("golden.xlsx matches the manifest as serials (cellDates:false)", async () => {
    const mod = await load(READ_PATH);
    const book = await mod.readWorkbook(goldenFile(ROOT), { cellDates: false, cellFormula: true });
    const res = checkGoldenManifest(book, loadGolden(ROOT), { cellDates: false });
    assertNoMismatches(res.mismatches, "golden manifest (serials)");
  });

  // The streaming path is gridxcel's headline claim: verify its cell values
  // against the same pre-recorded python manifest, not just its row count.
  test("eachRow streams the manifest cells (golden.xlsx)", async () => {
    const index = await import(pathToFileURL(path.join(ROOT, "src", "index.js")).href);
    assert.equal(typeof index.eachRow, "function", "src/index.js must export eachRow");
    const events = [];
    for await (const ev of index.eachRow(goldenFile(ROOT), { cellDates: true, cellFormula: true })) {
      events.push(ev);
      if (events.length > 5000) break; // safety valve, golden.xlsx has ~210 rows
    }
    // Resolve streamed number-format ids through the python oracle's tables.
    const doc = inspect(ROOT, goldenFile(ROOT), {});
    const res = checkGoldenStream(events, loadGolden(ROOT), {
      cellDates: true,
      numFmts: doc.book.numFmts,
      builtins: doc.book.builtinFormats,
    });
    assertNoMismatches(res.mismatches, "eachRow streaming");
    assert.ok(res.checks.cells >= 20, "expected the manifest cells to be streamed, saw " + res.checks.cells);
  });
});

describe("interop: harness sanity", { skip: SKIP_ORACLE }, () => {
  test("the comparator detects an injected mismatch and reports none on clean data", () => {
    const summary = { self: { checks: 0, mismatches: [] } };
    runSelfCheck(ROOT, summary);
    assert.equal(summary.self.mismatches.length, 0, JSON.stringify(summary.self.mismatches));
    assert.ok(summary.self.checks >= 3);
  });
});
