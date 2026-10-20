
import { formatNumber, isDateFormat, dateToSerial, serialToDate } from "../src/numfmt.js";
import { readFileSync } from "node:fs";
import assert from "node:assert";

/* --- 1. independent classifier oracle (openpyxl is_date_format) --- */
const oracle = JSON.parse(readFileSync(new URL("./numfmt_oracle.json", import.meta.url), "utf8"));
let p = 0, f = 0;
for (const [id, code, want] of oracle) {
  const got = id !== null ? isDateFormat(id) : isDateFormat(code);
  if (got === want) p++;
  else { f++; console.log("CLASSIFY MISMATCH", JSON.stringify(code), "id=" + id, "want", want, "got", got); }
}
console.log("classifier vs openpyxl oracle: pass " + p + " fail " + f);

/* --- 2. curated golden formatting set (Excel-verified semantics) --- */
const G = [
  [0, "0", "0"], [1, "0", "1"], [-1, "0", "-1"],
  [1234.5678, "#,##0", "1,235"], [1234.5678, "#,##0.00", "1,234.57"],
  [1234567, "#,##0", "1,234,567"], [-1234.5, "#,##0.00", "-1,234.50"],
  [1, "0.00", "1.00"], [0.5, "0.00", "0.50"], [0.125, "0.0", "0.1"],
  [0.125, "0.00%", "12.50%"], [0.0125, "0.00%", "1.25%"], [1.5, "0%", "150%"],
  [-0.25, "0.00%", "-25.00%"],
  [1234.5678, "0.000", "1234.568"],
  [1234.5678, "0.00E+00", "1.23E+03"], [0.000123, "0.00E+00", "1.23E-04"],
  [0.75, "# ?/?", "3/4"], [1.5, "# ?/?", "1 1/2"], [2.25, "# ??/??", "2  1/ 4"]  /* ? = width reservation */,
  [1234.5678, '"$"#,##0.00', "$1,234.57"],
  [1234.5678, "$#,##0.00", "$1,234.57"],
  [-1234.5678, "#,##0.00;[Red](#,##0.00)", "(1,234.57)"],
  [1234.5678, "#,##0.00;(#,##0.00)", "1,234.57"],
  [-1234.5678, "#,##0.00;(#,##0.00)", "(1,234.57)"],
  [0, "0.00;-0.00;\"zero\"", "zero"],
  [1, "@", ""], [1, "0.0", "1.0"], [-0.5, "0.00;[Red](0.00)", "(0.50)"],
  [45292, "yyyy-mm-dd", "2024-01-01"],
  [45292.5, "yyyy-mm-dd hh:mm:ss", "2024-01-01 12:00:00"],
  [44927, "yyyy-mm-dd", "2023-01-01"],
  [1, "yyyy-mm-dd", "1900-01-01"],
  [59, "yyyy-mm-dd", "1900-02-28"],
  [61, "yyyy-mm-dd", "1900-03-01"],
  [45292, "ddd", "Mon"], [45292, "dddd", "Monday"],
  [45292, "mmm", "Jan"], [45292, "mmmm", "January"], [45292, "m", "1"],
  [45292.75, "h:mm AM/PM", "6:00 PM"], [45292.25, "h:mm AM/PM", "6:00 AM"],
  [45292.5, "h:mm", "12:00"], [1.5, "[h]:mm:ss", "36:00:00"],
  [45292, "m/d/yy", "1/1/24"], [45292, "yy", "24"], [45292, "yyyy", "2024"],
  [45292.500011574, "hh:mm:ss", "12:00:01"],
];
let gp = 0, gf = 0;
for (const [v, code, want] of G) {
  let got;
  try { got = formatNumber(v, code); } catch (e) { got = "THREW " + e.message; }
  if (got === want) gp++;
  else { gf++; console.log("FMT MISMATCH", JSON.stringify(code), v, "want", JSON.stringify(want), "got", JSON.stringify(got)); }
}
console.log("golden formats: pass " + gp + " fail " + gf);

/* --- 3. classification must be fast (per-cell hot path) --- */
const t0 = performance.now();
for (let i = 0; i < 200000; i++) isDateFormat(i % 2 ? "yyyy-mm-dd" : "#,##0.00");
const dt = performance.now() - t0;
console.log("200k classifications in " + dt.toFixed(1) + " ms");

/* --- 4. round trips --- */
for (const dt of [new Date(1900,0,1), new Date(2024,0,1), new Date(2024,5,15,13,30,45), new Date(1999,11,31,23,59,59)]) {
  const s = dateToSerial(dt);
  const back = serialToDate(s);
  assert.strictEqual(back.getFullYear(), dt.getFullYear(), "round trip year " + dt);
  assert.strictEqual(back.getMonth(), dt.getMonth());
  assert.strictEqual(back.getDate(), dt.getDate());
  assert.strictEqual(back.getHours(), dt.getHours());
}
console.log("serial round trips OK");
console.log(gf === 0 && f === 0 ? "NUMFMT ALL GREEN" : "NUMFMT FAILURES: fmt=" + gf + " classify=" + f);
