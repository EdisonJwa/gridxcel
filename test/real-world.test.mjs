/**
 * Third-party XLSX files, opened exactly as a user would.
 *
 * Every other fixture in this repository was produced by gridxcel's writer, by
 * openpyxl, or by hand. That is a closed loop: it can only exercise the OOXML this
 * project already knows how to emit. Real files from real tools differ in ways
 * nobody invents - empty inlineStr elements with no <is> child, style ids repeated
 * across thousands of cells, defined names pointing into the middle of a sheet,
 * sheets named in Chinese and Korean, and column counts that do not match any
 * header the writer would produce.
 *
 * The files themselves are not in the repository: they are third-party data of
 * unknown provenance and a few hundred kilobytes each. What is committed is a
 * fingerprint of how each one decodes. When the file is present at the path in
 * REAL_FILES, the fingerprint is verified exactly; otherwise the test skips and
 * says so, so a checkout without them neither fails nor pretends to have run.
 *
 * To verify locally, copy the files to test/fixtures/real-world/ or set
 * GRIDXCEL_REAL_FILES to the directory holding them.
 */
import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readWorkbookSync } from "../src/read.js";
import { writeWorkbookSync } from "../src/write.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FINGERPRINTS = JSON.parse(readFileSync(path.join(HERE, "fixtures", "real-world-fingerprints.json"), "utf8"));

/** Where a real file might be: an explicit env dir, or the fixtures directory. */
function locate(name) {
  const dirs = [];
  if (process.env.GRIDXCEL_REAL_FILES) dirs.push(process.env.GRIDXCEL_REAL_FILES);
  dirs.push(path.join(HERE, "fixtures", "real-world"));
  for (const dir of dirs) {
    const full = path.join(dir, name);
    if (existsSync(full)) return full;
  }
  return null;
}

/** A stable description of how a workbook decoded: counts, and a value hash. */
function fingerprint(book, bytes) {
  const sheets = book.sheets.map((sheet, index) => {
    let cells = 0, emptyStr = 0, nums = 0, strs = 0, bools = 0, errs = 0, dates = 0;
    let h = 2166136261;
    sheet.rows.forEach((row, r) => (row || []).forEach((cell, col) => {
      if (!cell) return;
      cells++;
      if (cell.t === "string" && cell.v === "") emptyStr++;
      else if (cell.t === "number") nums++;
      else if (cell.t === "string") strs++;
      else if (cell.t === "boolean") bools++;
      else if (cell.t === "error") errs++;
      else if (cell.t === "date") dates++;
      const token = r + "," + col + ":" + cell.t + ":" + String(cell.v);
      for (let k = 0; k < token.length; k++) { h ^= token.charCodeAt(k); h = Math.imul(h, 16777619) >>> 0; }
    }));
    return { name: book.sheetNames[index], rows: sheet.rowCount, cols: sheet.columnCount, cells, emptyStr, nums, strs, bools, errs, dates, hash: h.toString(16) };
  });
  return {
    sha256: createHash("sha256").update(bytes).digest("hex"),
    bytes: bytes.length,
    sheets,
  };
}

for (const expected of FINGERPRINTS) {
  const file = locate(expected.file);
  const label = expected.file.length > 40 ? expected.file.slice(0, 38) + "..." : expected.file;

  test("real-world: " + label + " decodes as recorded", { skip: file === null ? "file not present; set GRIDXCEL_REAL_FILES to verify" : false }, () => {
    const bytes = new Uint8Array(readFileSync(file));

    // The file must be the exact one the fingerprint describes, or the comparison
    // is meaningless.
    const sha = createHash("sha256").update(bytes).digest("hex");
    assert.equal(sha, expected.sha256, "the file on disk is not the one recorded");
    assert.equal(bytes.length, expected.bytes, "byte length differs from the record");

    const book = readWorkbookSync(bytes);
    assert.deepEqual(book.warnings || [], [], "a valid third-party file must not warn");

    const actual = fingerprint(book, bytes);
    assert.equal(actual.sheets.length, expected.sheets.length, "sheet count");
    for (let i = 0; i < expected.sheets.length; i++) {
      const want = expected.sheets[i];
      const got = actual.sheets[i];
      assert.equal(got.name, want.name, "sheet " + i + " name");
      assert.equal(got.rows, want.rows, got.name + ": row count");
      assert.equal(got.cols, want.cols, got.name + ": column count");
      assert.equal(got.cells, want.cells, got.name + ": cell count");
      assert.equal(got.nums, want.nums, got.name + ": numeric cells");
      assert.equal(got.strs, want.strs, got.name + ": string cells");
      assert.equal(got.bools, want.bools, got.name + ": boolean cells");
      assert.equal(got.errs, want.errs, got.name + ": error cells");
      assert.equal(got.dates, want.dates, got.name + ": date cells");
      assert.equal(got.emptyStr, want.emptyStr, got.name + ": empty-string cells");
      assert.equal(got.hash, want.hash, got.name + ": value hash");
    }
  });
}

test("real-world: an empty inlineStr element is an empty string, not a missing cell", { skip: locate(FINGERPRINTS[0].file) === null ? "file not present" : false }, () => {
  // Some writers emit <c r="O2" t="inlineStr"></c> - an inlineStr with no <is>
  // child. It is not well-formed per the intent of the schema, but it occurs in
  // real files and it means "this cell exists and holds nothing".
  //
  // Reading it as an absent cell loses the distinction between a gap and an empty
  // value, which changes how a row maps to columns. openpyxl drops these; this
  // library keeps them, and that is deliberate: a sparse row that silently
  // compacts is a data-corruption bug that is very hard to see downstream.
  const file = locate(FINGERPRINTS[0].file);
  const book = readWorkbookSync(new Uint8Array(readFileSync(file)));
  const total = book.sheets.reduce((n, s) => n + s.rows.reduce((m, row) => m + (row || []).filter((c) => c && c.t === "string" && c.v === "").length, 0), 0);
  assert.ok(total > 1000, "expected the recorded file to carry many empty-string cells, found " + total);

  // And they must round-trip: an empty string is a value, so writing it back and
  // rereading must not turn it into a hole.
  const again = readWorkbookSync(writeWorkbookSync(book));
  const count = (b) => b.sheets.reduce((n, s) => n + s.rows.reduce((m, row) => m + (row || []).filter((c) => c && c.t === "string" && c.v === "").length, 0), 0);
  assert.equal(count(again), total, "empty-string cells survived the round trip");
});
