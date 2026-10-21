/**
 * Robustness suite for the number-format engine.
 *
 * A format code is not developer input: it arrives inside a workbook's
 * styles.xml, so it is attacker-controlled file content applied to every cell
 * that carries the style. The law is the same one the container fuzz suite holds
 * the reader to:
 *
 *   for ANY format code, formatNumber either returns a string or throws a
 *   GridxcelError. It never throws a raw RangeError, never takes unbounded time,
 *   and never allocates proportionally to a declared width.
 *
 * Three real defects were found this way and are pinned below:
 *   - "# ?/?????????" drove a one-billion-iteration loop via maxDen = 10^n - 1
 *   - "0" repeated 100,000 times cost 7.5 s of quadratic padding
 *   - a code with >100 decimal places threw a raw RangeError from toFixed
 */

import test from "node:test";
import assert from "node:assert/strict";
import { formatNumber, formatDateTime, isDateFormat } from "../src/numfmt.js";
import { GridxcelError } from "../src/errors.js";

/** A deterministic PRNG so a failure is reproducible. */
function rng(seed) {
  let s = seed >>> 0;
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
}

const TOKENS = ["0", "#", "?", ".", ",", "%", "E+", "E-", "/", "\\", '"x"', "[Red]", "[>=100]", "@", "*", "_", ";", "-", "+", " "];

function corpus() {
  const out = [];
  const r = rng(0x5eed1234);
  // The known-abusive shapes first.
  out.push({ code: "# ?/" + "?".repeat(9), value: 0.4142135 });
  out.push({ code: "0".repeat(100000), value: 1234.5 });
  out.push({ code: "0." + "0".repeat(100000), value: 1234.5 });
  out.push({ code: "0." + "0".repeat(5000) + "E+" + "0".repeat(5000), value: 1234.5 });
  out.push({ code: "# ?/999999999999999999", value: 0.5 });
  out.push({ code: "x".repeat(200000), value: 1 });
  out.push({ code: "0;".repeat(50000), value: 1 });
  out.push({ code: "@".repeat(100000), value: 1 });
  // Then random assemblies, bounded so the corpus itself stays small.
  for (let i = 0; i < 300; i++) {
    let code = "";
    const n = 1 + Math.floor(r() * 10);
    for (let k = 0; k < n; k++) code += TOKENS[Math.floor(r() * TOKENS.length)];
    out.push({ code, value: (r() - 0.5) * Math.pow(10, Math.floor(r() * 12)) });
  }
  return out;
}

test("numfmt: any format code returns a string or a typed error, quickly", { timeout: 60000 }, () => {
  const cases = corpus();
  const untyped = [];
  const slow = [];
  const notString = [];

  for (const { code, value } of cases) {
    const started = Date.now();
    try {
      const out = formatNumber(value, code);
      if (typeof out !== "string") notString.push(JSON.stringify(code.slice(0, 20)));
    } catch (error) {
      if (!(error instanceof GridxcelError)) {
        untyped.push(JSON.stringify(code.slice(0, 24)) + " -> " + (error && error.constructor && error.constructor.name) + ": " + String(error && error.message).slice(0, 60));
      }
    }
    const ms = Date.now() - started;
    if (ms > 250) slow.push(JSON.stringify(code.slice(0, 24)) + " took " + ms + "ms");
  }

  assert.deepEqual(untyped, [], "a non-GridxcelError escaped formatNumber");
  assert.deepEqual(slow, [], "a format code took over 250 ms");
  assert.deepEqual(notString, [], "formatNumber returned a non-string");
  console.log("  numfmt fuzz: " + cases.length + " codes, 0 untyped, 0 slow");
});

test("numfmt: a fraction denominator is bounded like Excel's", () => {
  // maxDen was 10^len - 1 straight from the code: nine '?' meant 999,999,999
  // iterations per cell. Excel renders at most three denominator digits.
  const started = Date.now();
  const out = formatNumber(0.4142135, "# ?/?????????");
  const ms = Date.now() - started;
  assert.ok(ms < 100, "must not scan a billion denominators, took " + ms + "ms");
  assert.match(out, /^\d+\s*/);
  // The three-digit answer is the correct one, not a truncation artefact.
  assert.ok(out.includes("985"), "expected the 3-digit convergent, got " + JSON.stringify(out));
});

test("numfmt: decimal places are capped, and past the cap is a typed error or a clamp", () => {
  for (const code of ["0." + "0".repeat(200), "0." + "0".repeat(5000)]) {
    let out;
    assert.doesNotThrow(() => { out = formatNumber(1.5, code); }, "must not throw a raw RangeError for " + code.length);
    assert.strictEqual(typeof out, "string");
  }
});

test("numfmt: an absurdly long code is refused rather than rendered", () => {
  // Excel caps a cell at 32,767 characters, so a longer code cannot describe
  // anything displayable - it only makes formatting cost seconds per cell.
  assert.throws(
    () => formatNumber(1, "0".repeat(40000)),
    (e) => e instanceof GridxcelError && e.code === "BAD_FORMAT",
  );
  // Just under the limit still works.
  assert.strictEqual(typeof formatNumber(1, "0".repeat(1000)), "string");
});

test("numfmt: an unpaired literal marker does not loop forever", () => {
  // THE WORST BUG IN THIS FILE. "*x", "_x" and "\x" are directives that consume
  // "this character and the next". They ran AFTER quoted runs had been replaced
  // by a NUL-delimited marker, so "*\"" ate half of one and left a lone NUL.
  // The render loop then did i = indexOf(NUL, i+1) === -1, and its own i++ put it
  // back on the same character - appending "" forever until the heap was gone.
  // Three characters of format code aborted the process: not catchable, so no
  // assertion could observe it. Verified by running each in its own process with
  // a 256 MB heap cap; before the fix all three exited 134.
  const inputs = [
    String.fromCharCode(42, 34, 120, 34),   // *"x"
    String.fromCharCode(95, 34, 120, 34),   // _"x"
    String.fromCharCode(92, 34, 120, 34),   // \"x"
  ];
  for (const code of inputs) {
    const started = Date.now();
    const out = formatNumber(1234.5678, code);
    assert.strictEqual(typeof out, "string", JSON.stringify(code));
    assert.ok(Date.now() - started < 100, JSON.stringify(code) + " must not loop");
    // A marker must never reach the caller.
    assert.ok(!out.includes("\u0000"), "the output leaked a literal marker: " + JSON.stringify(out));
  }
});

test("numfmt: fill and spacing directives leave no residue", () => {
  // "*" fills to the column width and "_" reserves one character; neither
  // contributes text of its own.
  assert.strictEqual(formatNumber(42, "*x0"), "42");
  assert.strictEqual(formatNumber(42, "_x0"), " 42");
  assert.match(formatNumber(42, "0*x"), /^42\s*$/);
});

test("numfmt: real formats still render correctly after the bounds", () => {
  assert.strictEqual(formatNumber(1234.5, "#,##0.00"), "1,234.50");
  assert.strictEqual(formatNumber(0.25, "0.0%"), "25.0%");
  assert.strictEqual(formatNumber(-1234.5, "#,##0.00;[Red](#,##0.00)"), "(1,234.50)");
  assert.strictEqual(formatNumber(1234567, "#,##0"), "1,234,567");
  assert.strictEqual(formatNumber(1234.5678, "0.00E+00"), "1.23E+03");
  // The common fraction case must be unaffected by the cap.
  assert.match(formatNumber(0.5, "# ?/?"), /1\s*\/\s*2|1\/2/);
});

test("numfmt: formatDateTime survives hostile date codes", () => {
  for (const code of ["yyyy-mm-dd", "yyyy".repeat(5000), "m".repeat(100000), "[h]:mm:ss", "dddd".repeat(1000)]) {
    const started = Date.now();
    try {
      const out = formatDateTime(45000.5, code);
      assert.strictEqual(typeof out, "string");
    } catch (error) {
      assert.ok(error instanceof GridxcelError, "typed error expected for " + code.slice(0, 16) + ", got " + error);
    }
    assert.ok(Date.now() - started < 1000, "date code " + code.slice(0, 16) + " was slow");
  }
});

test("numfmt: isDateFormat is total over hostile input", () => {
  for (const code of ["", "yyyy", "m".repeat(100000), "[".repeat(50000), "\u0000\u0001"]) {
    assert.strictEqual(typeof isDateFormat(code), "boolean");
  }
});
