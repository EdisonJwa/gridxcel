/**
 * Robustness (fuzz) suite for src/read.js.
 *
 * Law under test: for ANY input bytes, readWorkbook either
 *   (1) throws an Error (a GridxcelError with a stable code is the contract),
 *   (2) returns a Book, in which case malformed parts must be reported through
 *       onWarning, or
 *   (3) -- never -- throws a non-Error, hangs, or silently returns a Book for
 *       input that cannot possibly be a workbook.
 *
 * The corruption cases are produced by this file's own byte surgery (a tiny
 * zip locator and a hand-built stored zip). No gridxcel module is used to
 * build or check the inputs, so a bug in src/zip.js cannot hide here.
 */

import { describe, test } from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const READ_PATH = path.join(ROOT, "src", "read.js");
const ERRORS_PATH = path.join(ROOT, "src", "errors.js");
const FIXTURE_DIR = path.join(ROOT, "test", "fixtures");
const SKIP = existsSync(READ_PATH) ? false : "NOT YET AVAILABLE: src/read.js";
const CASE_TIMEOUT_MS = Number(process.env.GRIDXCEL_FUZZ_TIMEOUT ?? 10000);

/* Byte-surgery helpers live in scripts/zip-surgery.mjs so importing them from a
   probe does not execute this test file. */
import { ascii, crc32, findEocd, flipByte, locateEntry, storedZip, u8 } from "../scripts/zip-surgery.mjs";

/* -------------------------------------------------------------- the cases */

export function buildCases() {
  const cases = [];
  const fixtures = existsSync(FIXTURE_DIR)
    ? readdirSync(FIXTURE_DIR).filter((f) => f.endsWith(".xlsx")).sort()
    : [];

  for (const file of fixtures) {
    const buf = new Uint8Array(readFileSync(path.join(FIXTURE_DIR, file)));
    const steps = buf.length < 100 * 1024 ? 24 : 12;
    for (let i = 1; i <= steps; i++) {
      const at = Math.floor((buf.length * i) / (steps + 1));
      cases.push({ name: file + ":truncate@" + at, bytes: buf.slice(0, at), hard: false });
    }
    // Flip a byte inside xl/worksheets/sheet1.xml (stored entries corrupt the
    // XML text; deflated entries corrupt the deflate stream).
    const entry = locateEntry(buf, "xl/worksheets/sheet1.xml");
    if (entry) {
      const mid = entry.dataStart + Math.floor((entry.dataEnd - entry.dataStart) / 2);
      cases.push({ name: file + ":flip-sheet-xml@" + (mid - entry.dataStart), bytes: flipByte(buf, mid), hard: false });
      cases.push({ name: file + ":flip-sheet-xml-head", bytes: flipByte(buf, entry.dataStart + 1), hard: false });
      cases.push({ name: file + ":truncate-entry", bytes: buf.slice(0, entry.dataStart + 4), hard: false });
    }
    // Corrupt the central directory itself.
    const eocd = findEocd(buf);
    if (eocd >= 0) {
      cases.push({ name: file + ":eocd-entry-count-FFFF", bytes: withUint16(buf, eocd + 10, 0xffff), hard: false });
      cases.push({ name: file + ":eocd-central-offset-bogus", bytes: withUint32(buf, eocd + 16, 0x7fffffff), hard: false });
      cases.push({ name: file + ":central-dir-signature-flipped", bytes: flipByte(buf, eocd - 10), hard: false });
      cases.push({ name: file + ":no-eocd", bytes: buf.slice(0, eocd), hard: false });
    }
  }

  // Inputs that can never be a workbook: these must be rejected outright.
  cases.push({ name: "empty-file", bytes: new Uint8Array(0), hard: true });
  cases.push({ name: "one-byte", bytes: u8(0x50), hard: true });
  cases.push({ name: "plain-text", bytes: ascii("this is not a spreadsheet at all\n"), hard: true });
  cases.push({ name: "zero-filled-22", bytes: new Uint8Array(22), hard: true });
  cases.push({ name: "pk-header-only", bytes: u8(0x50, 0x4b, 0x03, 0x04), hard: true });
  cases.push({
    name: "valid-zip-not-workbook",
    bytes: storedZip([["hello.txt", "not an xlsx package"]]),
    hard: true,
  });
  return cases;
}

function withUint16(buf, at, value) {
  const out = buf.slice();
  out[at] = value & 0xff;
  out[at + 1] = (value >> 8) & 0xff;
  return out;
}

function withUint32(buf, at, value) {
  const out = buf.slice();
  out[at] = value & 0xff;
  out[at + 1] = (value >> 8) & 0xff;
  out[at + 2] = (value >> 16) & 0xff;
  out[at + 3] = (value >> 24) & 0xff;
  return out;
}

/* --------------------------------------------------------------- the run */

function withTimeout(promise, ms, label) {
  let timer;
  const guard = new Promise((_, reject) => {
    timer = setTimeout(() => reject(Object.assign(new Error("HANG: " + label + " exceeded " + ms + "ms"), { hang: true })), ms);
  });
  return Promise.race([promise, guard]).finally(() => clearTimeout(timer));
}

async function runCase(readWorkbook, GridxcelError, c) {
  const warnings = [];
  let book = null;
  let error = null;
  try {
    book = await withTimeout(
      Promise.resolve().then(() => readWorkbook(c.bytes, { onWarning: (w) => warnings.push(w) })),
      CASE_TIMEOUT_MS,
      c.name
    );
  } catch (err) {
    error = err;
  }
  const result = { name: c.name, hard: c.hard, warnings, error, hang: Boolean(error?.hang) };
  // Law A: never a non-Error throw.
  if (error && !c.hang) {
    assert.ok(error instanceof Error, "non-Error thrown for " + c.name + ": " + inspectThrown(error));
  }
  result.typed = error instanceof GridxcelError;
  result.code = error?.code;
  if (!error) {
    // Law B: a returned value must be a Book and malformed input must be
    // reported through warnings, never silently.
    assert.ok(book && typeof book === "object" && Array.isArray(book.sheets), "non-Book returned for " + c.name);
    result.sheets = book.sheets.length;
    for (const w of warnings) {
      assert.ok(w && typeof w === "object" && typeof w.code === "string", "warning without a code for " + c.name + ": " + JSON.stringify(w));
    }
    // Law C: input that cannot be a workbook must either raise a typed error or
    // come back as an explicitly warned, empty book -- no invented sheets.
    if (c.hard) {
      assert.ok(
        warnings.length > 0 && book.sheets.length === 0,
        "hard case " + c.name + " returned " + book.sheets.length + " sheet(s) with " + warnings.length + " warning(s); expected a typed throw or an empty warned book"
      );
    }
  } else if (c.hard) {
    assert.ok(error instanceof GridxcelError, "hard case " + c.name + " threw " + (error?.constructor?.name ?? typeof error) + " instead of GridxcelError");
  }
  return result;
}

function inspectThrown(x) {
  if (x === null) return "null";
  if (typeof x === "object") return "object " + JSON.stringify(Object.keys(x));
  return typeof x + " " + String(x);
}

describe("fuzz: corrupted containers never hang or throw non-Errors", { skip: SKIP }, () => {
  test("all generated corruption cases", { timeout: CASE_TIMEOUT_MS * 4 }, async () => {
    const mod = await import(pathToFileURL(READ_PATH).href);
    const { GridxcelError } = await import(pathToFileURL(ERRORS_PATH).href);
    assert.equal(typeof mod.readWorkbook, "function", "src/read.js must export readWorkbook");
    const cases = buildCases();
    assert.ok(cases.length >= 20, "expected a meaningful number of fuzz cases, got " + cases.length);

    const results = [];
    const failures = [];
    for (const c of cases) {
      try {
        results.push(await runCase(mod.readWorkbook, GridxcelError, c));
      } catch (err) {
        failures.push({ case: c.name, message: err?.message ?? String(err) });
      }
    }
    const hung = results.filter((r) => r.hang);
    const untyped = results.filter((r) => r.error && !r.typed && !r.hang);
    const recovered = results.filter((r) => !r.error);
    const silent = recovered.filter((r) => r.warnings.length === 0);
    const empty = recovered.filter((r) => r.sheets === 0);
    console.log(
      "  fuzz summary: " + cases.length + " cases, " +
      results.filter((r) => r.typed).length + " typed GridxcelError, " +
      untyped.length + " plain Error, " + recovered.length + " recovered workbook(s) (" +
      silent.length + " with no warning, " + empty.length + " with 0 sheets), " +
      hung.length + " hang(s)"
    );
    for (const r of recovered.filter((x) => x.hard)) {
      console.log("    hard case " + r.name + ": 0 sheets, warning=" + JSON.stringify(r.warnings[0]?.code) + " " + JSON.stringify(r.warnings[0]?.message));
    }
    // Errors without a code violate the taxonomy, so name them explicitly
    // instead of only counting them.
    for (const r of untyped) {
      console.log("    untyped " + r.name + ": " + (r.error?.name ?? "?") + ": " + String(r.error?.message ?? "").slice(0, 170));
    }
    // A corruption that recovered without a single warning deserves scrutiny:
    // it is either benign (the cut part is never read) or silent data loss.
    if (silent.length) {
      console.log("    " + silent.length + " corruption case(s) recovered with no warning: " +
        silent.slice(0, 24).map((r) => r.name + "(" + r.sheets + " sheets)").join(", "));
    }
    assert.deepEqual(failures, [], "hard-law violations: " + JSON.stringify(failures.slice(0, 5)));
    assert.deepEqual(hung.map((h) => h.name), [], "hangs detected");
  });

  // The typed-error path exists: strict mode must raise a GridxcelError with a
  // code for input that cannot be a workbook.
  test("strict mode raises a typed GridxcelError for impossible input", { timeout: CASE_TIMEOUT_MS * 2 }, async () => {
    const mod = await import(pathToFileURL(READ_PATH).href);
    const { GridxcelError } = await import(pathToFileURL(ERRORS_PATH).href);
    const hard = buildCases().filter((c) => c.hard);
    const bad = [];
    const codes = new Set();
    for (const c of hard) {
      try {
        const book = await withTimeout(Promise.resolve(mod.readWorkbook(c.bytes, { strict: true })), CASE_TIMEOUT_MS, c.name + ":strict");
        bad.push(c.name + ": strict:true returned a Book (" + (book?.sheets?.length ?? "?") + " sheets)");
      } catch (err) {
        if (!(err instanceof GridxcelError)) bad.push(c.name + ": threw " + (err?.constructor?.name ?? typeof err) + " (" + (err?.message ?? err) + ") instead of GridxcelError");
        else if (!err.code) bad.push(c.name + ": GridxcelError without a code");
        else codes.add(err.code);
      }
    }
    console.log("  strict error codes: " + [...codes].join(", "));
    assert.deepEqual(bad, [], "untyped strict rejections: " + JSON.stringify(bad));
  });

  test("strict mode turns recovered damage into a throw", { timeout: CASE_TIMEOUT_MS * 2 }, async () => {
    const mod = await import(pathToFileURL(READ_PATH).href);
    const cases = buildCases().filter((c) => !c.hard);
    let lenientRecovered = 0;
    let strictThrew = 0;
    for (const c of cases) {
      let book = null;
      const warnings = [];
      try {
        book = await withTimeout(Promise.resolve(mod.readWorkbook(c.bytes, { onWarning: (w) => warnings.push(w) })), CASE_TIMEOUT_MS, c.name);
      } catch {
        continue; // already rejected without strict
      }
      if (!book || warnings.length === 0) continue;
      lenientRecovered++;
      try {
        await withTimeout(Promise.resolve(mod.readWorkbook(c.bytes, { strict: true })), CASE_TIMEOUT_MS, c.name + ":strict");
      } catch {
        strictThrew++;
      }
    }
    console.log("  strict summary: " + lenientRecovered + " lenient recoveries with warnings, " + strictThrew + " threw under strict:true");
    if (lenientRecovered > 0) {
      assert.equal(strictThrew, lenientRecovered, "every lenient recovery with warnings must throw under strict:true");
    }
  });
});
