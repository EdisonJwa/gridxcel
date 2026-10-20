/**
 * Thin wrapper around scripts/oracle.py plus the canonicalisation rules that
 * let a Python view and a gridxcel view of the same workbook be compared.
 *
 * Nothing in here imports a gridxcel module: the oracle must stay independent
 * of the code under test, and the canonical form is defined by this file only.
 */

import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import process from "node:process";

/** Bundled interpreter from docs/ARCHITECTURE.md; GRIDXCEL_PYTHON overrides. */
export const BUNDLED_PYTHON =
  "C:\\Users\\admin\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe";

export function pythonPath() {
  if (process.env.GRIDXCEL_PYTHON) return process.env.GRIDXCEL_PYTHON;
  if (existsSync(BUNDLED_PYTHON)) return BUNDLED_PYTHON;
  return "python";
}

export function oracleScript(root) {
  return path.join(root, "scripts", "oracle.py");
}

/** True when the interpreter can import openpyxl (i.e. the oracle is usable). */
export function oracleAvailable() {
  try {
    execFileSync(pythonPath(), ["-c", "import openpyxl, lxml"], { stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

/**
 * Run scripts/oracle.py and parse its single JSON line.
 * Throws OracleError with the raw stderr attached when python fails.
 */
export function runOracle(root, args, { timeoutMs = 300000 } = {}) {
  const script = oracleScript(root);
  if (!existsSync(script)) throw new Error("oracle script missing: " + script);
  let stdout;
  try {
    stdout = execFileSync(pythonPath(), [script, ...args], {
      cwd: root,
      encoding: "utf8",
      // Keep python's tracebacks out of the console: stderr is inherited by
      // default and deliberately-malformed fixtures make python noisy.
      stdio: ["ignore", "pipe", "pipe"],
      timeout: timeoutMs,
      maxBuffer: 1024 * 1024 * 1024,
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: "utf-8" },
    });
  } catch (err) {
    const e = new Error("oracle failed: python " + [script, ...args].join(" "));
    e.stderr = err.stderr?.toString?.() ?? String(err.stderr ?? "");
    e.stdout = err.stdout?.toString?.() ?? String(err.stdout ?? "");
    // The python side prints {"error": "..."} on stdout before exiting 1.
    try {
      const line = e.stdout.trim().split("\n").filter((l) => l.trim().startsWith("{")).pop();
      if (line) e.oracleMessage = JSON.parse(line).error;
    } catch { /* keep the raw streams */ }
    e.oracleUnreadable = true;
    throw e;
  }
  const text = stdout.trim();
  const start = text.indexOf("{");
  if (start < 0) throw new Error("oracle produced no JSON: " + text.slice(0, 200));
  const doc = JSON.parse(text.slice(start));
  if (doc.error) throw new Error("oracle error: " + doc.error);
  return doc;
}

export const inspect = (root, file, opts = {}) =>
  runOracle(root, ["inspect", file, ...(opts.cached ? ["--cached"] : []), ...(opts.readOnly ? ["--read-only"] : [])]);

export const selftest = (root, file) => runOracle(root, ["selftest", file]);
export const generate = (root, out, rows, cols, seed = 7) =>
  runOracle(root, ["gen", out, "--rows", String(rows), "--cols", String(cols), "--seed", String(seed)]);

/* ------------------------------------------------------------ canonical */

/** 1-based column number -> letters, independent of src/address.js. */
export function colName(col) {
  let n = col;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function refOf(row, col) {
  return colName(col) + String(row);
}

/** "a1:c3" / "$A$1:$C$3" -> "A1:C3". */
export function normRange(text) {
  if (!text) return "";
  return String(text)
    .toUpperCase()
    .replace(/\$/g, "")
    .split(":")
    .map((p) => p.trim())
    .join(":");
}

/** Number-format identity: builtin id or code, normalised to the code string. */
export function formatCode(z, builtins, numFmts) {
  if (z === undefined || z === null) return "General";
  if (typeof z === "string") return z;
  if (typeof z === "number") {
    const key = String(z);
    if (numFmts && numFmts[key] !== undefined) return numFmts[key];
    if (builtins && builtins[key] !== undefined) return builtins[key];
    return "id:" + z;
  }
  return String(z);
}

/** gridxcel cell type -> oracle kind, per docs/ARCHITECTURE.md data model. */
export function kindOfCell(cell) {
  switch (cell?.t) {
    case "string": return "string";
    case "number": return "number";
    case "boolean": return "boolean";
    case "date": return "date";
    case "error": return "error";
    case "empty": return "empty";
    default: return "unknown:" + String(cell?.t);
  }
}

/**
 * Canonical JSON view of one workbook, from EITHER side.
 * oracle: { cells:[{ref,row,col,kind,value,format,formula,serial,date_iso}], merges:[], trueMaxRow, ... }
 * gridxcel: Book from src/read.js.
 */
export function canonFromOracle(doc) {
  const builtins = doc.book.builtinFormats;
  const numFmts = doc.book.numFmts;
  return {
    dateSystem: doc.book.dateSystem,
    sheetNames: doc.book.sheetNames,
    numFmts,
    builtinFormats: builtins,
    sheets: doc.sheets.map((s) => ({
      name: s.name,
      rowCount: s.trueMaxRow,
      columnCount: s.trueMaxCol,
      merges: [...s.merges].map(normRange).sort(),
      cells: s.cells.map((c) => ({
        ref: c.ref,
        kind: c.kind,
        value: oracleValue(c),
        format: c.format ?? "General",
        formula: c.formula ? normFormula(c.formula) : undefined,
        cached: c.cached === undefined ? undefined : c.cached,
        serial: c.serial,
        date_iso: c.date_iso,
        // Raw XML of the <c> element, for triage in mismatch reports.
        xml: c.xml,
        source: c.source,
      })),
    })),
  };
}

function oracleValue(c) {
  const v = c.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if ("datetime" in v) return { date_iso: v.datetime };
    if ("date" in v) return { date_iso: v.date + "T00:00:00" };
    if ("timedelta_seconds" in v) return { seconds: v.timedelta_seconds };
    return v;
  }
  return v;
}

/**
 * Does the reader's value match the oracle's cached formula result?
 * The oracle emits explicitly tagged scalars: {t:"n"|"s"|"b"|"d"}. A date
 * result may legitimately appear on the reader side as a serial number (the
 * default) or as a Date (cellDates), so both representations are accepted.
 */
export function cachedMatches(cached, actualCell) {
  const v = actualCell?.value;
  switch (cached.t) {
    case "d": {
      if (typeof v === "number" && typeof cached.serial === "number") return Math.abs(v - cached.serial) < 1e-9;
      if (v && typeof v === "object" && (v.date_iso || v.datetime)) return sameInstant(String(v.date_iso ?? v.datetime), String(cached.iso));
      return false;
    }
    case "n": {
      if (cached.nan) return typeof v === "number" && Number.isNaN(v);
      return typeof v === "number" ? v === cached.v : false;
    }
    case "b": return v === cached.v;
    case "s": return typeof v === "string" ? v === cached.v : String(v) === String(cached.v);
    default: return false;
  }
}

export function describeCached(cached) {
  if (cached.t === "d") return "date serial " + cached.serial + " (" + cached.iso + ")";
  return JSON.stringify(cached.v);
}

/**
 * Canonical view of a gridxcel Book (src/read.js output) for the same file.
 * opts.builtins / opts.numFmts come from the oracle document so number-format
 * ids resolve to codes without consulting src/numfmt.js.
 */
export function canonFromGridxcel(book, opts) {
  const { builtins = {}, numFmts = {} } = opts ?? {};
  const sheets = (book.sheets ?? []).map((sheet) => {
    const originCol = sheet.originCol ?? 0;
    const originRow = sheet.originRow ?? 1;
    const cells = [];
    let maxRow = 0;
    let maxCol = 0;
    (sheet.rows ?? []).forEach((row, ri) => {
      if (!row) return;
      const rowNumber = originRow + ri;
      for (let ci = 0; ci < row.length; ci++) {
        const cell = row[ci];
        if (!cell) continue;
        const col = originCol + ci + 1;
        // Stub cells carry no data; a cell whose only content is a formula is
        // data even when the file stores no cached result.
        const hasFormula = cell.f !== undefined && cell.f !== null && cell.f !== "";
        if (cell.t === "empty" && (cell.v === undefined || cell.v === null) && !hasFormula) continue;
        const ref = cell.a ? String(cell.a).toUpperCase() : refOf(rowNumber, col);
        const kind = kindOfCell(cell);
        const entry = {
          ref,
          kind,
          value: gridxcelValue(cell),
          format: formatCode(cell.z, builtins, numFmts),
        };
        if (cell.f !== undefined && cell.f !== null && cell.f !== "") entry.formula = normFormula(cell.f);
        if (kind === "date") {
          entry.serial = typeof cell.serial === "number" ? cell.serial : (typeof cell.v === "number" ? cell.v : undefined);
          if (cell.v instanceof Date) entry.date_iso = isoNoZone(cell.v);
          if (cell.t === "date" && typeof cell.v === "string") entry.date_iso = cell.v;
        }
        if (kind === "error") entry.error = cell.e ?? cell.v;
        cells.push(entry);
        maxRow = Math.max(maxRow, rowNumber);
        maxCol = Math.max(maxCol, col);
      }
    });
    return {
      name: sheet.name,
      rowCount: sheet.rowCount ?? maxRow,
      columnCount: sheet.columnCount ?? maxCol,
      merges: [...(sheet.merges ?? [])].map(normRange).sort(),
      cells,
    };
  });
  return { dateSystem: book.dateSystem, sheetNames: sheets.map((s) => s.name), sheets };
}

function gridxcelValue(cell) {
  const v = cell.v;
  if (v === undefined || v === null) return null;
  if (v instanceof Date) return { date_iso: isoNoZone(v) };
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v;
  return String(v);
}

/**
 * Wall-clock ISO string for a Date, matching python's naive isoformat().
 *
 * The date contract is that a cell showing 2024-03-15 14:30:45 in Excel reads
 * back as a Date whose LOCAL wall clock is 2024-03-15 14:30:45 in every
 * timezone. Using toISOString()/getUTC* here would shift the value by the host
 * offset and make the harness timezone-dependent, which is a harness bug, not
 * a library one.
 */
export function isoNoZone(d) {
  const p = (n, w = 2) => String(n).padStart(w, "0");
  return (
    d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) +
    "T" + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds()) +
    (d.getMilliseconds() ? "." + p(d.getMilliseconds(), 3) : "")
  );
}

export function normFormula(f) {
  let s = String(f).trim();
  if (s.startsWith("=")) s = s.slice(1);
  return s.replace(/\s+/g, "");
}

function valEqual(a, b) {
  if (a === b) return true;
  // The oracle wraps date-like cached values as {datetime|date: iso}.
  if (a && typeof a === "object" && !(a instanceof Date)) {
    const iso = a.datetime ?? a.date;
    if (iso) return sameInstant(String(iso), String(b?.date_iso ?? b));
  }
  if (b && typeof b === "object" && b.date_iso) return sameInstant(String(a), b.date_iso);
  if (a === null || b === null || a === undefined || b === undefined) return a === b || (a == null && b == null);
  if (typeof a === "number" && typeof b === "number") return a === b; // IEEE-754 exact
  if (typeof a === "object" && typeof b === "object") {
    const ad = a.date_iso ?? a.datetime;
    const bd = b.date_iso ?? b.datetime;
    if (ad && bd) return sameInstant(String(ad), String(bd));
    if ("seconds" in a && "seconds" in b) return a.seconds === b.seconds;
    return JSON.stringify(a) === JSON.stringify(b);
  }
  if (typeof a === "number" && typeof b === "string") return String(a) === b;
  if (typeof a === "string" && typeof b === "number") return a === String(b);
  return false;
}

/** Compare two ISO local timestamps, tolerating a 1 ms rounding difference. */
export function sameInstant(a, b) {
  if (a === b) return true;
  const ta = Date.parse(a + (a.endsWith("Z") ? "" : "Z"));
  const tb = Date.parse(b + (b.endsWith("Z") ? "" : "Z"));
  if (Number.isNaN(ta) || Number.isNaN(tb)) return false;
  return Math.abs(ta - tb) <= 1;
}

/**
 * Diff two canonical views. Returns { mismatches:[{sheet,ref,field,expected,actual,severity}], counts }.
 * severity "error" fails the harness; "warn" is reported but tolerated.
 */
export function diffCanon(expected, actual, { maxDetail = 50, dateTolerance = 1 } = {}) {
  const mismatches = [];
  const push = (m) => {
    // Every mismatch must be triageable: sheet, reference and the raw <c> XML.
    if (m.xml === undefined) m.xml = undefined;
    if (mismatches.length < maxDetail) mismatches.push(m);
  };
  if (expected.sheetNames.join("|") !== actual.sheetNames.join("|")) {
    push({ sheet: "-", ref: "-", field: "sheetNames", expected: expected.sheetNames.join(","), actual: actual.sheetNames.join(","), severity: "error" });
  }
  const byName = new Map(actual.sheets.map((s) => [s.name, s]));
  for (const es of expected.sheets) {
    const as = byName.get(es.name);
    if (!as) {
      push({ sheet: es.name, ref: "-", field: "sheet", expected: "present", actual: "missing", severity: "error" });
      continue;
    }
    if (es.rowCount !== as.rowCount) {
      push({ sheet: es.name, ref: "-", field: "rowCount", expected: es.rowCount, actual: as.rowCount, severity: "error" });
    }
    if (es.columnCount !== as.columnCount) {
      push({ sheet: es.name, ref: "-", field: "columnCount", expected: es.columnCount, actual: as.columnCount, severity: "error" });
    }
    if (es.merges.join("|") !== as.merges.join("|")) {
      push({ sheet: es.name, ref: "-", field: "merges", expected: es.merges.join(","), actual: as.merges.join(","), severity: "error" });
    }
    const aByRef = new Map(as.cells.map((c) => [c.ref, c]));
    const eByRef = new Map(es.cells.map((c) => [c.ref, c]));
    for (const ec of es.cells) {
      const ac = aByRef.get(ec.ref);
      if (!ac) {
        push({ sheet: es.name, ref: ec.ref, field: "cell", expected: describe(ec), actual: "missing", severity: "error", xml: ec.xml });
        continue;
      }
      compareCell(es.name, ec, ac, push, dateTolerance);
    }
    for (const ac of as.cells) {
      if (!eByRef.has(ac.ref)) {
        push({ sheet: es.name, ref: ac.ref, field: "cell", expected: "absent", actual: describe(ac), severity: "error" });
      }
    }
  }
  return { mismatches, count: mismatches.length, truncated: mismatches.length >= maxDetail };
}

function describe(c) {
  return JSON.stringify({ kind: c.kind, value: c.value, format: c.format, formula: c.formula });
}

function compareCell(sheet, ec, ac, push, dateTolerance) {
  const at = (field, expected, actual, severity = "error") =>
    push({ sheet, ref: ec.ref, field, expected, actual, severity, xml: ec.xml });

  if (ec.formula || ac.formula) {
    if ((ec.formula ?? "") !== (ac.formula ?? "")) {
      at("formula", ec.formula ?? "(none)", ac.formula ?? "(none)");
    }
  }
  if (ec.kind === "formula") {
    // python reports the formula text as the cell value; the cached result (if
    // the file stores one) is what a reader should expose as the value.
    const cached = ec.cached === undefined ? null : ec.cached;
    if (cached !== null && !cachedMatches(cached, ac)) {
      at("cachedValue", describeCached(cached), JSON.stringify(ac.value));
    }
    if ((ec.format ?? "General") !== (ac.format ?? "General")) at("format", ec.format, ac.format, "warn");
    return;
  }
  const ecKind = ec.kind;
  const acKind = ac.kind;
  if (ecKind && acKind && ecKind !== acKind) {
    at("type", ecKind, acKind);
    return;
  }
  if (ecKind === "date" || acKind === "date") {
    const es = ec.serial;
    const as = ac.serial;
    if (typeof es === "number" && typeof as === "number") {
      if (Math.abs(es - as) > 1e-9 * Math.max(1, Math.abs(es))) at("serial", es, as);
    } else if (ec.date_iso && ac.date_iso && !sameInstant(ec.date_iso, ac.date_iso)) {
      at("date", ec.date_iso, ac.date_iso);
    } else if (ec.value !== null && ac.value !== null && !valEqual(ec.value, ac.value)) {
      at("value", JSON.stringify(ec.value), JSON.stringify(ac.value));
    }
    if ((ec.format ?? "General") !== (ac.format ?? "General")) {
      at("format", ec.format, ac.format, "warn");
    }
    return;
  }
  if (ec.value !== null && !valEqual(ec.value, ac.value)) {
    at("value", JSON.stringify(ec.value), JSON.stringify(ac.value));
  } else if (ec.value === null && ac.value !== null && ecKind !== "formula") {
    at("value", "null", JSON.stringify(ac.value));
  }
  if ((ec.format ?? "General") !== (ac.format ?? "General")) {
    at("format", ec.format, ac.format, "warn");
  }
}
