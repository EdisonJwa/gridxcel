/**
 * DUP-1/2/6/8: the de-duplication contract.
 *
 * Two things are asserted here:
 *   1. one implementation per concern - the writer's cell coercion and
 *      sheet-name rules are the canonical ones, and the two `utils` namespaces
 *      agree on every member they share;
 *   2. the differences that remain are deliberate, documented in the source,
 *      and pinned so they cannot drift back into "the same call returned a
 *      different answer depending on the import".
 *
 * Claims about a written file are checked by Python/openpyxl through
 * scripts/oracle-client.mjs - never by another gridxcel module.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { readFileSync } from "node:fs";

import * as g from "../src/index.js";
import * as compat from "../src/compat.js";
import { WorkbookWriter } from "../src/workbook-writer.js";
import { writeWorkbookSync } from "../src/write.js";
import { cellFromValue } from "../src/json.js";
import { emptyBook, normalizeBook, mergeBooks, validateSheetName } from "../src/api.js";
import * as platform from "../src/platform.js";
import * as platformBrowser from "../src/platform.browser.js";
import { DEFAULT_SHEET_NAME } from "../src/constants.js";
import { inspect, oracleAvailable } from "../scripts/oracle-client.mjs";

const ROOT = fileURLToPath(new URL("..", import.meta.url));

/** Write one sheet to a throwaway file, returning its path. */
function writeTempRows(rows, name = "Data") {
  const dir = mkdtempSync(path.join(tmpdir(), "gridxcel-dedup-"));
  const file = path.join(dir, "out.xlsx");
  const book = g.normalizeBook({ sheets: [{ name, rows }] });
  writeWorkbookSync(book, { path: file, overwrite: true, deterministic: true });
  return { file, dir };
}


/* ------------------------------------------- shared helpers (task-25) */

test("F4: the dead sync-reader cluster is gone and readSync still works", () => {
  // readBookSync/installSyncReader/syncCache had zero references; index.readSync
  // calls readWorkbookSync directly. Removing the public names is the point.
  assert.equal(g.readBookSync, undefined, "api.readBookSync is not exported");
  assert.equal(g.installSyncReader, undefined, "api.installSyncReader is not exported");
  for (const file of ["../src/api.js"]) {
    const text = readFileSync(new URL(file, import.meta.url), "utf8");
    assert.ok(!/readBookSync|installSyncReader|SYNC_UNAVAILABLE/.test(text), file + " no longer mentions them");
  }
  // and the supported synchronous path still reads a real workbook
  const { file, dir } = writeTempRows([[1, "a"]]);
  try {
    const book = g.readSync(readFileSync(file));
    assert.deepEqual(book.sheets.map((s) => s.name), ["Data"]);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("F8: one empty-Book helper, used by every caller in this change", () => {
  const book = emptyBook();
  assert.deepEqual(Object.keys(book).sort(), ["dateSystem", "definedNames", "numFmts", "props", "sheets", "styles"]);
  assert.equal(book.dateSystem, 1900);
  assert.deepEqual(book.sheets, []);
  assert.deepEqual(book.props, {});
  assert.ok(book.numFmts instanceof Map);
  assert.deepEqual(emptyBook({ title: "t" }).props, { title: "t" });

  // the three call sites that had their own copy now share this shape
  assert.deepEqual(Object.keys(normalizeBook(null)).sort(), Object.keys(book).sort());
  assert.deepEqual(Object.keys(mergeBooks()).sort(), Object.keys(book).sort());
  // compat's book is the same shape plus its two SheetJS views (pinned below),
  // so assert the canonical fields are all present rather than key-for-key equal
  const compatKeys = Object.keys(compat.utils.book_new());
  for (const key of Object.keys(book)) assert.ok(compatKeys.includes(key), "compat book carries " + key);
  assert.equal(normalizeBook(null, { date1904: true }).dateSystem, 1904);
});

test("F21: the codecs and the sync file read live in the platform seam", () => {
  // identical export surface: a bundler can swap the two files
  assert.deepEqual(Object.keys(platform).sort(), Object.keys(platformBrowser).sort());
  for (const name of ["decodeBase64", "decodeLatin1", "readFileNodeSync", "encodeBase64", "encodeLatin1"]) {
    assert.equal(typeof platform[name], "function", "platform." + name);
    assert.equal(typeof platformBrowser[name], "function", "platform.browser." + name);
  }
  // round trips, including bytes above 0x7f and a chunk-boundary-sized payload
  const bytes = new Uint8Array(70_001);
  for (let i = 0; i < bytes.length; i++) bytes[i] = (i * 37 + 200) & 0xff;
  assert.deepEqual(platform.decodeBase64(platform.encodeBase64(bytes)), bytes);
  assert.deepEqual(platform.decodeLatin1(platform.encodeLatin1(bytes)), bytes);
  assert.deepEqual(platformBrowser.decodeBase64(platformBrowser.encodeBase64(bytes)), bytes);
  // compat's read/write types still work through the seam
  assert.equal(compat.read("a,b\n1,2\n").Sheets.Sheet1.rows[0][0].v, "a");
  assert.deepEqual(
    compat.read(Buffer.from(platform.encodeBase64(new Uint8Array([1, 2, 3]))), { type: "base64" }).SheetNames,
    ["Sheet1"],
  );
  // the browser sync read refuses with the seam's typed error, not a ReferenceError
  assert.throws(() => platformBrowser.readFileNodeSync("x.xlsx"), (e) => e.code === "NOT_SUPPORTED");
});

test("F11: 'sheet not found' is NO_SHEET in the compat surface", () => {
  const book = compat.utils.book_new();
  compat.utils.book_append_sheet(book, compat.utils.aoa_to_sheet([[1]]), "Only");
  assert.throws(
    () => compat.utils.book_set_sheet_visibility(book, "Missing", 1),
    (e) => e.code === "NO_SHEET",
  );
  // the old spelling must not appear in the files this task owns
  for (const file of ["../src/api.js", "../src/compat.js"]) {
    const text = readFileSync(new URL(file, import.meta.url), "utf8");
    assert.ok(!text.includes("BAD_SHEET_SELECTOR"), file + " uses NO_SHEET only");
  }
});

test("F14: the zip sniff uses the container signature constant", () => {
  // behaviour: a zip-looking payload goes to the reader, text goes to CSV
  const csv = compat.read("a,b\n1,2\n");
  assert.equal(csv.SheetNames[0], "Sheet1");
  assert.throws(() => compat.read(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0])), (e) => e.code !== "BAD_INPUT");
  // and the signature is imported rather than re-typed
  const text = readFileSync(new URL("../src/compat.js", import.meta.url), "utf8");
  assert.ok(text.includes("SIG_LOCAL"), "compat imports the zip signature");
  assert.ok(!/bytes\[0\] === 0x50/.test(text), "no hand-written magic bytes");
});

/* --------------------------------------------------------------- DUP-1 */

test("DUP-1: the writer rejects what the canonical coercion rejects", () => {
  // These three used to reach the writer, fail every branch of its private
  // type guess and vanish from the file without an error. The canonical
  // contract (json.cellFromValue, which api.normalizeCell now delegates to)
  // rejects them, so the same value behaves the same way on every entry point.
  const rejected = [{ a: 1 }, [1, 2], Symbol("s")];
  for (const value of rejected) {
    assert.throws(() => cellFromValue(value), (e) => e.code === "BAD_CELL", String(value));
    assert.throws(() => writeWorkbookSync(g.normalizeBook({ sheets: [{ name: "S", rows: [[value]] }] })), (e) => e.code === "BAD_CELL", String(value));
  }
});

test("DUP-1: values the caller supplied reach the file (openpyxl)", (t) => {
  const when = new Date(2024, 0, 15);
  const { file, dir } = writeTempRows([[1, { v: 3.5 }, "text", true, when, { v: "rich", t: "string" }]]);
  try {
    if (!oracleAvailable()) {
      console.log("SKIP DUP-1 file check: the python oracle is not available");
      t.skip("no oracle");
      return;
    }
    const doc = inspect(ROOT, file);
    const cells = new Map(doc.sheets[0].cells.map((c) => [c.ref, c]));
    assert.equal(cells.get("A1").value, 1);
    assert.equal(cells.get("B1").value, 3.5, "a {v} cell without t is written, not dropped");
    assert.equal(cells.get("C1").value, "text");
    assert.equal(cells.get("D1").value, true);
    assert.equal(cells.get("E1").kind, "date");
    assert.equal(cells.get("F1").value, "rich");
    // A bare Date keeps the writer's own default code (date vs date-time),
    // which is richer than the "yyyy-mm-dd" the canonical cell carries.
    assert.equal(cells.get("E1").format, "yyyy-mm-dd");
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

/* --------------------------------------------------------------- DUP-6 */

test("DUP-6: one sheet-name rule for the writer and both utils surfaces", () => {
  const illegal = ["'quoted'", "History", "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "a:b", "x".repeat(32)];
  for (const name of illegal) {
    assert.throws(() => validateSheetName(name), (e) => e.code === "BAD_SHEET_NAME", "api: " + JSON.stringify(name));
    assert.throws(() => new WorkbookWriter().addSheet(name), (e) => e.code === "BAD_SHEET_NAME", "writer: " + JSON.stringify(name));
    assert.throws(
      () => g.utils.book_append_sheet(g.utils.book_new(), [["a"]], name),
      (e) => e.code === "BAD_SHEET_NAME",
      "index.utils: " + JSON.stringify(name),
    );
    assert.throws(
      () => compat.utils.book_append_sheet(compat.utils.book_new(), compat.utils.aoa_to_sheet([["a"]]), name),
      (e) => e.code === "BAD_SHEET_NAME",
      "compat.utils: " + JSON.stringify(name),
    );
  }
  // An empty name means "name it for me" on both utils surfaces, which is what
  // SheetJS's falsy-name rule does; a sheet that really is called "" is still
  // rejected by the rule itself and by the writer.
  assert.throws(() => validateSheetName(""), (e) => e.code === "BAD_SHEET_NAME");
  assert.throws(() => new WorkbookWriter().addSheet(""), (e) => e.code === "BAD_SHEET_NAME");
  // Book normalisation applies the same falsy-name rule upstream of the writer,
  // so an empty name never reaches addSheet() to be rejected there.
  assert.equal(g.normalizeBook({ sheets: [{ name: "", rows: [[1]] }] }).sheets[0].name, DEFAULT_SHEET_NAME);
  const emptyNamed = g.utils.book_new();
  g.utils.book_append_sheet(emptyNamed, [["a"]], "");
  assert.equal(emptyNamed.sheets[0].name, DEFAULT_SHEET_NAME, "index.utils names it");
  const emptyNamedShim = compat.utils.book_new();
  compat.utils.book_append_sheet(emptyNamedShim, compat.utils.aoa_to_sheet([["a"]]), "");
  assert.equal(emptyNamedShim.SheetNames[0], DEFAULT_SHEET_NAME, "compat.utils names it the same way");

  // Names that ARE legal stay legal on every path.
  for (const name of ["Data", "Sheet1", "a b", "123", "x".repeat(31)]) {
    assert.equal(validateSheetName(name), name);
    new WorkbookWriter().addSheet(name);
  }
});

test("DUP-6: uniqueness is case-insensitive, as Excel's is", () => {
  const writer = new WorkbookWriter();
  writer.addSheet("Data");
  assert.throws(() => writer.addSheet("data"), (e) => e.code === "BAD_SHEET_NAME");
  assert.throws(() => writer.addSheet("DATA"), (e) => e.code === "BAD_SHEET_NAME");
  // the same rule through both namespaces
  const viaIndex = g.utils.book_new();
  g.utils.book_append_sheet(viaIndex, [["a"]], "Data");
  assert.throws(() => g.utils.book_append_sheet(viaIndex, [["b"]], "data"), (e) => e.code === "BAD_SHEET_NAME");
  const viaCompat = compat.utils.book_new();
  compat.utils.book_append_sheet(viaCompat, compat.utils.aoa_to_sheet([["a"]]), "Data");
  assert.throws(
    () => compat.utils.book_append_sheet(viaCompat, compat.utils.aoa_to_sheet([["b"]]), "data"),
    (e) => e.code === "BAD_SHEET_NAME",
  );
});

/* --------------------------------------------------------------- DUP-2 */

test("DUP-2: the two utils namespaces agree on every shared member", async () => {
  const rows = [["name", "qty"], ["a", 1], ["b", 2]];
  const nativeSheet = g.utils.aoa_to_sheet(rows);
  const shimSheet = compat.utils.aoa_to_sheet(rows);
  assert.deepEqual(nativeSheet, shimSheet, "aoa_to_sheet produces the same cells");

  assert.deepEqual(await g.utils.sheet_to_json(nativeSheet), compat.utils.sheet_to_json(shimSheet));
  assert.deepEqual(await g.utils.sheet_to_aoa(nativeSheet), compat.utils.sheet_to_aoa(shimSheet));
  assert.equal(await g.utils.sheet_to_csv(nativeSheet), compat.utils.sheet_to_csv(shimSheet));
  assert.deepEqual(await g.utils.json_to_sheet([{ a: 1, b: 2 }]), compat.utils.json_to_sheet([{ a: 1, b: 2 }]));

  // format_cell: the measured defect. Cell.w wins on BOTH surfaces now.
  const withText = { t: "number", v: 3.14159, z: "0.00", w: "SHOWN" };
  const without = { t: "number", v: 3.14159, z: "0.00" };
  assert.equal(g.utils.format_cell(withText), "SHOWN");
  assert.equal(compat.utils.format_cell(withText), "SHOWN");
  assert.equal(g.utils.format_cell(without), "3.14");
  assert.equal(compat.utils.format_cell(without), "3.14");

  // address helpers
  assert.equal(g.utils.encode_cell({ r: 0, c: 1 }), compat.utils.encode_cell({ r: 0, c: 1 }));
  assert.equal(g.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }), compat.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }));
  assert.deepEqual(g.utils.decode_range("A1:B2"), compat.utils.decode_range("A1:B2"));

  // book_append_sheet: same naming and the same acceptance decision
  const a = g.utils.book_new();
  const b = compat.utils.book_new();
  g.utils.book_append_sheet(a, [["x"]]);
  compat.utils.book_append_sheet(b, compat.utils.aoa_to_sheet([["x"]]));
  assert.equal(a.sheets[0].name, b.SheetNames[0]);
});

test("DUP-2: the two deliberate differences are pinned", async () => {
  const sheet = g.utils.aoa_to_sheet([["a"], [1]]);

  // 1. converters: promise here, value there (docs/API.md vs SheetJS)
  assert.equal(typeof g.utils.sheet_to_json(sheet).then, "function");
  assert.equal(typeof g.utils.sheet_to_csv(sheet).then, "function");
  assert.equal(compat.utils.sheet_to_json(sheet).then, undefined);
  assert.equal(compat.utils.sheet_to_csv(sheet).then, undefined);
  assert.equal(typeof (await g.utils.sheet_to_json(sheet)).length, "number");

  // 2. book_new: the native Workbook, versus the SheetJS shape (views)
  const native = g.utils.book_new();
  const shim = compat.utils.book_new();
  assert.equal(native.SheetNames, undefined, "the native book has no SheetJS views");
  assert.deepEqual(native.sheets, []);
  assert.deepEqual(shim.SheetNames, []);
  assert.deepEqual(shim.Sheets, {});
  assert.deepEqual(shim.sheets, [], "the shim's book is a native book with views over it");

  // compat also carries the members only a migration needs
  for (const name of ["sheet_to_txt", "book_set_sheet_visibility", "decode_cell", "encode_col"]) {
    assert.equal(typeof compat.utils[name], "function", "compat.utils." + name);
    assert.equal(g.utils[name], undefined, "index.utils." + name + " is not part of the native surface");
  }
});

/* --------------------------------------------------------------- DUP-8 */

test("DUP-8: unnamed sheets use DEFAULT_SHEET_NAME and the first free number", () => {
  assert.equal(DEFAULT_SHEET_NAME, "Sheet1");
  assert.equal(g.utils.aoa_to_sheet([[1]]).name, DEFAULT_SHEET_NAME);

  const book = g.utils.book_new();
  g.utils.book_append_sheet(book, [["a"]]);
  g.utils.book_append_sheet(book, [["b"]]);
  assert.deepEqual(book.sheets.map((s) => s.name), ["Sheet1", "Sheet2"]);

  // SheetJS's rule is "first free", not "count + 1": a taken Sheet1 is skipped
  const shim = compat.utils.book_new();
  const firstName = compat.utils.book_append_sheet(shim, compat.utils.aoa_to_sheet([["a"]]), "Sheet2");
  const secondName = compat.utils.book_append_sheet(shim, compat.utils.aoa_to_sheet([["b"]]));
  assert.equal(firstName, "Sheet2");
  assert.equal(secondName, DEFAULT_SHEET_NAME);
  assert.deepEqual(shim.SheetNames, ["Sheet2", "Sheet1"]);
});
