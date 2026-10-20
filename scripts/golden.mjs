/**
 * Manifest-driven checks for test/fixtures/golden.xlsx.
 *
 * test/fixtures/golden.json was produced by python/openpyxl (see
 * test/make_golden.py): it records what openpyxl itself sees, so a reader can
 * be diffed against a python-written oracle without spawning python. That
 * makes it usable inside node --test, and it is a second, independent oracle
 * for the same file (the first being a live openpyxl dump).
 *
 * The manifest is JSON with t values from the OOXML vocabulary:
 *   s string, n number, b boolean, d date, f formula.
 */

import { readFileSync } from "node:fs";
import path from "node:path";

export const TYPE_MAP = { s: "string", n: "number", b: "boolean", d: "date", f: "formula" };

/** Standard 1900-system serial from an ISO local timestamp, computed here. */
export function serialFromIso(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?)?$/.exec(iso);
  if (!m) return NaN;
  const [, y, mo, d, hh = "0", mm = "0", ss = "0", frac = "0"] = m;
  const ms = Date.UTC(+y, +mo - 1, +d, +hh, +mm, +ss, Number(("0." + frac)) * 1000);
  return (ms - Date.UTC(1899, 11, 30)) / 86400000;
}

export function loadGolden(root) {
  return JSON.parse(readFileSync(path.join(root, "test", "fixtures", "golden.json"), "utf8"));
}

export function goldenFile(root) {
  return path.join(root, "test", "fixtures", "golden.xlsx");
}

/**
 * Check a gridxcel Book against the manifest.
 * opts.cellDates picks which representation the reader was asked for.
 * Returns { checks, mismatches: [{ check, target, expected, actual }] }.
 */
export function checkGoldenManifest(book, manifest, { cellDates = true } = {}) {
  const mismatches = [];
  const checks = { sheets: 0, cells: 0, merges: 0, dimensions: 0, volume: 0, formulas: 0 };
  const push = (check, target, expected, actual) => mismatches.push({ check, target, expected, actual });

  checks.sheets++;
  const names = (book.sheets ?? []).map((s) => s.name);
  if (names.join("|") !== manifest.sheets.join("|")) {
    push("sheetNames", "-", manifest.sheets.join(","), names.join(","));
  }

  const byName = new Map((book.sheets ?? []).map((s) => [s.name, s]));

  for (const [sheetName, cells] of Object.entries(manifest.cellDates)) {
    const sheet = byName.get(sheetName);
    if (!sheet) {
      push("sheet", sheetName, "present", "missing");
      continue;
    }
    const index = indexSheet(sheet);
    for (const [ref, want] of Object.entries(cells)) {
      checks.cells++;
      const cell = index.get(ref);
      const where = sheetName + "!" + ref;
      if (!cell) {
        push("cell", where, JSON.stringify(want), "missing");
        continue;
      }
      compareManifestCell(book, where, cell, want, cellDates, (c, w, e, a) => push(c, w, e, a), checks);
    }
  }

  for (const [sheetName, ranges] of Object.entries(manifest.merges ?? {})) {
    checks.merges++;
    const sheet = byName.get(sheetName);
    if (!sheet && ranges.length === 0) continue;
    if (!sheet) {
      push("merges", sheetName, ranges.join(","), "missing sheet");
      continue;
    }
    const actual = (sheet.merges ?? []).map((m) => String(m).toUpperCase().replace(/\$/g, "")).sort();
    const want = [...ranges].map((m) => m.toUpperCase()).sort();
    if (actual.join("|") !== want.join("|")) push("merges", sheetName, want.join(","), actual.join(","));
  }

  for (const [sheetName, dims] of Object.entries(manifest.dimensions ?? {})) {
    checks.dimensions++;
    const sheet = byName.get(sheetName);
    if (!sheet) {
      push("dimensions", sheetName, dims, "missing sheet");
      continue;
    }
    const actual = sheet.dims ? String(sheet.dims).toUpperCase().replace(/\$/g, "") : null;
    // dims is optional in the data model; fall back to the computed extent.
    const computed = "A1:" + colLetters(sheet.columnCount ?? 0) + String(sheet.rowCount ?? 0);
    if (actual !== dims && computed !== dims.toUpperCase()) {
      push("dimensions", sheetName, dims, actual ?? computed);
    }
  }

  const vol = manifest.volume;
  if (vol) {
    const sheet = byName.get("Volume");
    checks.volume++;
    if (!sheet) {
      push("volume", "Volume", "present", "missing");
    } else {
      const index = indexSheet(sheet);
      const a1 = index.get("A1");
      const b1 = index.get("B1");
      const z1 = index.get("Z1");
      const d1 = index.get("D1");
      if (!a1 || a1.v !== vol.colA) push("volume", "Volume!A1", vol.colA, a1 ? JSON.stringify(a1.v) : "missing");
      if (!b1 || b1.v !== vol.colB1) push("volume", "Volume!B1", vol.colB1, b1 ? JSON.stringify(b1.v) : "missing");
      if (!z1 || z1.v !== vol.colZ1) push("volume", "Volume!Z1 (sparse far column)", vol.colZ1, z1 ? JSON.stringify(z1.v) : "missing");
      if (!d1 || d1.v !== vol.colD1) push("volume", "Volume!D1", vol.colD1, d1 ? JSON.stringify(d1.v) : "missing");
      if ((sheet.rowCount ?? 0) < vol.rows) push("volume", "Volume rowCount", ">= " + vol.rows, String(sheet.rowCount));
    }
  }

  return { checks, mismatches };
}

/**
 * Compare one cell against one manifest entry. Shared by the whole-book check
 * and the streaming (eachRow) check so both apply identical rules.
 */
export function compareManifestCell(book, where, cell, want, cellDates, push, checks = null) {
  if (want.t === "f") {
    const formula = stripEq(cell.f);
    if (formula !== stripEq(String(want.v))) push("formula", where, stripEq(String(want.v)), formula || "(none)");
  } else if (want.t === "d") {
    if (cellDates) {
      if (!(cell.v instanceof Date)) {
        push("dateType", where, "Date", typeof cell.v + " " + String(cell.v));
      } else if (isoNoZone(cell.v) !== want.v) {
        push("dateValue", where, want.v, isoNoZone(cell.v));
      }
      // Timezone-free cross-check: the cell's serial against the manifest
      // value converted with the python (date - 1899-12-30) rule.
      const expectSerial = serialFromIso(want.v);
      const actualSerial = typeof cell.serial === "number" ? cell.serial : NaN;
      if (!(Math.abs(actualSerial - expectSerial) < 1e-9)) {
        push("dateSerial", where, expectSerial, cell.serial ?? "(absent)");
      }
    } else {
      if (cell.t !== "date") push("dateType", where, "date cell", String(cell.t));
      const serial = cell.serial ?? (typeof cell.v === "number" ? cell.v : NaN);
      const expect = serialFromIso(want.v);
      if (!(Math.abs(serial - expect) < 1e-9)) push("dateSerial", where, expect, serial);
    }
  } else {
    const wantKind = TYPE_MAP[want.t] ?? want.t;
    if (cell.t !== wantKind) push("cellType", where, wantKind, String(cell.t));
    if (!valueEqual(cell.v, want.v)) push("cellValue", where, JSON.stringify(want.v), JSON.stringify(cell.v));
  }
  if (want.nf && want.nf !== "General") {
    if (checks) checks.formulas++;
    const z = typeof cell.z === "string" ? cell.z : numberFormatCode(book, cell.z);
    if (z !== want.nf) push("numberFormat", where, want.nf, z ?? "(none)");
  }
}

/**
 * Same manifest rules applied to eachRow events ({ sheet, r, cells }).
 * Proves the streaming path delivers the same cell values as a full read,
 * with the manifest as the independent reference.
 */
export function checkGoldenStream(events, manifest, { cellDates = true, numFmts = null, builtins = null } = {}) {
  const mismatches = [];
  const checks = { cells: 0, formulas: 0, rows: 0 };
  // Streamed cells carry number-format IDS (Cell.z is an id per the data
  // model); the id -> code table comes from the python oracle, so the
  // resolution never depends on a gridxcel module.
  const styleSource = numFmts || builtins ? { numFmts: numFmts ?? {}, builtins: builtins ?? {} } : null;
  const seen = new Set();
  const push = (check, target, expected, actual) => mismatches.push({ check, target, expected, actual });
  for (const ev of events) {
    checks.rows++;
    const cells = ev.cells ?? [];
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      if (!cell) continue;
      const ref = cell.a ? String(cell.a).toUpperCase() : colLetters(c) + ev.r;
      const key = ev.sheet + "!" + ref;
      seen.add(key);
      const want = manifest.cellDates[ev.sheet]?.[ref];
      if (!want) continue;
      checks.cells++;
      compareManifestCell(styleSource, key, cell, want, cellDates, push, checks);
    }
  }
  for (const [sheetName, cells] of Object.entries(manifest.cellDates)) {
    for (const ref of Object.keys(cells)) {
      if (!seen.has(sheetName + "!" + ref)) {
        push("streamMissing", sheetName + "!" + ref, "streamed", "absent");
      }
    }
  }
  return { checks, mismatches };
}

/** ref -> cell for a gridxcel Sheet (sparse rows anchored at originRow/originCol). */
export function indexSheet(sheet) {
  const map = new Map();
  const originRow = sheet.originRow ?? 1;
  const originCol = sheet.originCol ?? 0;
  (sheet.rows ?? []).forEach((row, ri) => {
    if (!row) return;
    for (let ci = 0; ci < row.length; ci++) {
      const cell = row[ci];
      if (!cell) continue;
      const ref = cell.a ? String(cell.a).toUpperCase() : colLetters(originCol + ci) + String(originRow + ri);
      map.set(ref, cell);
    }
  });
  return map;
}

export function numberFormatCode(book, z) {
  if (z === undefined || z === null) return "General";
  if (typeof z === "string") return z;
  const numFmts = book?.numFmts;
  const code = numFmts && typeof numFmts.get === "function" ? numFmts.get(z) : numFmts?.[z];
  if (code) return code;
  const builtins = book?.builtins;
  if (builtins && builtins[z] !== undefined) return builtins[z];
  return BUILTIN[z] ?? "id:" + z;
}

/** ECMA-376 builtin formats; kept here so a read check needs no gridxcel module. */
export const BUILTIN = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
  11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy",
  16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm",
  21: "h:mm:ss", 22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
  46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

export function colLetters(col0) {
  let n = col0 + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Wall-clock ISO string (local fields). Date cells carry the Excel wall clock
 * in local time, so using UTC fields here would introduce a timezone offset
 * and break the harness in any zone other than UTC.
 */
export function isoNoZone(d) {
  const p = (n, w = 2) => String(n).padStart(w, "0");
  return (
    d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) +
    "T" + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds())
  );
}

const stripEq = (s) => String(s ?? "").trim().replace(/^=/, "").replace(/\s+/g, "");

function valueEqual(a, b) {
  if (a === b) return true;
  if (typeof a === "number" && typeof b === "number") return a === b;
  if (a === null || a === undefined) return b === null || b === undefined;
  return String(a) === String(b);
}
