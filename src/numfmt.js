/**
 * ECMA-376 number formats: the built-in id table, fast "is this a date
 * format?" classification, Excel serial <-> Date conversion, and a renderer
 * for the format codes that appear in real spreadsheets.
 *
 * The classification path is deliberately cheap (one cached pass over the
 * code) because every numeric cell in a large sheet consults it.
 */

import { GridxcelError } from "./errors.js";

/** Built-in numFmtId -> format code (ECMA-376 18.8.30). */
export const BUILTIN_FORMATS = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: "$#,##0_);($#,##0)",
  6: "$#,##0_);[Red]($#,##0)",
  7: "$#,##0.00_);($#,##0.00)",
  8: "$#,##0.00_);[Red]($#,##0.00)",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0_);(#,##0)",
  38: "#,##0_);[Red](#,##0)",
  39: "#,##0.00_);(#,##0.00)",
  40: "#,##0.00_);[Red](#,##0.00)",
  41: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
  42: "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
  43: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
  44: "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

/** Format ids that Excel defines as date/time formats. */
const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

export function builtinFormat(id) {
  return BUILTIN_FORMATS[id];
}

/* ----------------------------------------------------------- classification */

const cache = new Map();

/**
 * True when a format code renders a date and/or time.
 * Strips literals, colours and conditions first so that "$0.00" and
 * "[Red]0.0" are not mistaken for date formats.
 */
export function isDateFormat(code) {
  if (code === undefined || code === null) return false;
  const hit = cache.get(code);
  if (hit !== undefined) return hit;
  let v = false;
  if (typeof code === "number") {
    v = BUILTIN_DATE_IDS.has(code);
  } else {
    // Inspect only the first section: Excel picks the section per value, and
    // date formats in practice are identical across sections.
    const first = code.split(";")[0];
    let clean = "";
    for (let i = 0; i < first.length; i++) {
      const ch = first[i];
      if (ch === '"') { const c = first.indexOf('"', i + 1); i = c < 0 ? first.length : c; continue; }
      if (ch === "\\") { i++; continue; }
      if (ch === "[") {
        const c = first.indexOf("]", i + 1);
        const body = c < 0 ? first.slice(i + 1) : first.slice(i + 1, c);
        // [h]/[m]/[s] are elapsed-time tokens, everything else is a modifier
        if (/^[hms]+$/i.test(body)) clean += body.toLowerCase();
        i = c < 0 ? first.length : c;
        continue;
      }
      if (ch === "_" || ch === "*") { i++; continue; }
      clean += ch;
    }
    // Only y/m/d/h/s and AM/PM decide; "m" alone is ambiguous but with any
    // other date token present it is a month, so require a second signal.
    const hasDateToken = /[yd]/i.test(clean) || /\[h+\]/i.test(first);
    const hasTimeToken = /(h{1,2}|s{1,2}|am\/pm|a\/p)/i.test(clean);
    const mCount = (clean.match(/m/gi) || []).length;
    v = hasDateToken || hasTimeToken || (mCount > 0 && /[:]/.test(clean)) || /^\[?m+\]?$/.test(clean.replace(/\s/g, ""));
    if (!v && clean.trim() === "m") v = false;
  }
  cache.set(code, v);
  return v;
}

/* ------------------------------------------------------- serial <-> calendar */

const MS_PER_DAY = 86400000;

/** Excel serial -> UTC ms since epoch (1900 system, includes the leap bug). */
export function serialToUTCms(serial, date1904 = false) {
  // Anchor: serial 25569 === 1970-01-01 UTC. Serial 60 is the phantom
  // 1900-02-29, so everything from 1..60 sits one day further from the epoch
  // than the naive subtraction implies.
  if (date1904) return Math.round((serial - 24107) * MS_PER_DAY);
  const days = serial < 61 ? serial + 1 : serial;
  return Math.round((days - 25569) * MS_PER_DAY);
}

/** Excel serial -> Date in local time (what users expect from cellDates). */
export function serialToDate(serial, date1904 = false) {
  const utc = serialToUTCms(serial, date1904);
  const d = new Date(utc);
  // Rebase so that the wall-clock fields match the spreadsheet, not UTC.
  const local = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds(), d.getUTCMilliseconds());
  return local;
}

/** Date -> Excel serial. */
export function dateToSerial(date, date1904 = false) {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
  let serial = utc / MS_PER_DAY + 25569;
  if (date1904) serial -= 1462;
  else if (serial < 61) serial -= 1;
  return serial;
}

/** Break a serial into date/time parts without going through Date. */
export function serialParts(serial) {
  const whole = Math.floor(serial);
  let ms = Math.round((serial - whole) * MS_PER_DAY);
  let day = whole;
  // A rounded fraction can reach a full day (e.g. 1.9999999999).
  if (ms >= MS_PER_DAY) { ms -= MS_PER_DAY; day += 1; }
  // Convert the integer day on its own; feeding the fractional serial here
  // would re-round the day and shift the clock.
  const d = new Date(serialToUTCms(day, false));
  return {
    y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    dow: d.getUTCDay(),
    h: Math.floor(ms / 3600000), mi: Math.floor((ms % 3600000) / 60000),
    s: Math.floor((ms % 60000) / 1000), ms: ms % 1000,
    serial,
  };
}

/* --------------------------------------------------------------- rendering */

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const MONTHS_ABBR = MONTHS.map((m) => m.slice(0, 3));
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const DAYS_ABBR = DAYS.map((d) => d.slice(0, 3));

function groupThousands(intStr, sep = ",") {
  let out = "";
  const n = intStr.length;
  for (let i = 0; i < n; i++) {
    if (i > 0 && (n - i) % 3 === 0) out += sep;
    out += intStr[i];
  }
  return out;
}

/** Left-pad the absolute value of n to exactly width digits. */
function pad(n, width, ch = "0") {
  const s = String(Math.abs(n));
  return s.length >= width ? s : ch.repeat(width - s.length) + s;
}

/** Pad a digit string to a minimum length (never truncates, never zero-pads a non-number). */
function padStr(s, width, ch = "0") {
  return s.length >= width ? s : ch.repeat(width - s.length) + s;
}

/** Fraction approximation for "# ?/?" style codes. */
function bestFraction(value, maxDen) {
  const sign = value < 0 ? -1 : 1;
  const x = Math.abs(value);
  let bestN = 0, bestD = 1, bestErr = Infinity;
  for (let d = 1; d <= maxDen; d++) {
    const num = Math.round(x * d);
    const err = Math.abs(x - num / d);
    if (err < bestErr - 1e-12) { bestErr = err; bestN = num; bestD = d; if (err < 1e-12) break; }
  }
  return { sign, n: bestN, d: bestD };
}

/** Choose the section of a multi-part format code for a value. */
export function pickSection(sections, value, isText) {
  if (isText) return sections.length >= 4 ? 3 : 0;
  if (sections.length === 1) return 0;
  if (value > 0) return 0;
  if (value < 0) return sections.length >= 2 ? 1 : 0;
  return sections.length >= 3 ? 2 : 0;
}

/** Split a format code on top-level semicolons. */
export function splitSections(code) {
  const out = [];
  let cur = "";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') { const c = code.indexOf('"', i + 1); cur += code.slice(i, c < 0 ? code.length : c + 1); i = c < 0 ? code.length : c; continue; }
    if (ch === "[" ) { const c = code.indexOf("]", i + 1); cur += code.slice(i, c < 0 ? code.length : c + 1); i = c < 0 ? code.length : c; continue; }
    if (ch === "\\") { cur += code.slice(i, i + 2); i++; continue; }
    if (ch === ";") { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

/**
 * Render a numeric value through a format code.
 * @param {number} value
 * @param {string|number} code
 * @param {{date1904?: boolean}} [opts]
 */
export function formatNumber(value, code, opts = {}) {
  if (code === undefined || code === null || code === "General" || code === 0) return generalNumber(value);
  const codes = typeof code === "number" ? (BUILTIN_FORMATS[code] || "General") : code;
  if (codes === "General") return generalNumber(value);
  const sections = splitSections(codes);
  const idx = pickSection(sections, value, false);
  let section = sections[idx] === undefined ? sections[0] : sections[idx];
  if (!section) return "";

  // Colour / condition / locale brackets: consumed, not rendered.
  let negativeParens = idx === 1 && sections.length > 1 && /\(.*\)/.test(section) && /[#0]/.test(section);
  section = section.replace(/\[(?!\$|h+\]|m+\]|s+\])[^\]]*\]/gi, "");
  let code2 = section;

  if (isDateFormat(codes)) return formatDateTime(value, code2, opts);

  // Percent scaling: count % signs (escaped ones do not count).
  let percent = 0;
  code2 = code2.replace(/%(?!.)/g, () => { percent++; return "%"; });
  const pctCount = (code2.match(/%/g) || []).length;
  let scaled = value * Math.pow(100, pctCount);

  // Scientific
  if (/[0#?]\.?[0#?]*[eE][+-]?0+/.test(code2)) return formatScientific(scaled, code2);

  // Fractions
  const fracMatch = code2.match(/([#0?]+)\s+([#0?]+)\/([#0?]+|[0-9]+)/);
  if (fracMatch) return formatFraction(scaled, code2, fracMatch);

  // Literal text extraction
  let literals = [];
  let body = code2.replace(/"([^"]*)"/g, (m, g) => { literals.push(g); return "\u0000" + (literals.length - 1) + "\u0000"; });
  body = body.replace(/\\(.)/g, (m, c) => { literals.push(c); return "\u0000" + (literals.length - 1) + "\u0000"; });
  // Underscore = reserve the width of the next char; asterisk = fill (treated as blank).
  body = body.replace(/_.?/g, "\u0001");
  body = body.replace(/\*./g, "");
  body = body.replace(/\u0001/g, " ");
  // Padding around a fraction is part of the fraction token, not a literal,
  // so hide it from the placeholder walker (NBSP is not a number placeholder).
  body = body.replace(/ +(?=\S*\/)/g, (m) => "\u00a0".repeat(m.length));

  // Placeholder analysis
  const intPart = body.split(".")[0] || "";
  const decMatch = body.match(/\.([#0?]+)/);
  const decPart = decMatch ? decMatch[1] : "";
  const useGrouping = intPart.includes(",") && /[#0]/.test(intPart.split(",").join(""));
  const obligatoryInt = (intPart.match(/0/g) || []).length;
  const minFrac = (decPart.match(/0/g) || []).length;
  const maxFrac = decPart.length;

  const abs = Math.abs(scaled);
  let numStr;
  if (maxFrac === 0 && decMatch === null) {
    numStr = String(Math.round(abs));
    if (abs >= 1e15 && /[#0]/.test(intPart)) numStr = String(abs);
  } else {
    numStr = abs.toFixed(maxFrac);
    if (minFrac < maxFrac) {
      numStr = numStr.replace(/0+$/, "");
      const dot = numStr.indexOf(".");
      if (dot >= 0 && numStr.length - dot - 1 < minFrac) numStr = numStr.padEnd(dot + 1 + minFrac, "0");
      if (numStr.endsWith(".")) numStr = numStr.slice(0, -1);
    }
  }
  let [ip, dp] = numStr.split(".");
  // Left-pad to the number of obligatory integer digits. (pad() takes a
  // total width, so pass the shortfall, not the count.)
  if (obligatoryInt > 0 && ip.length < obligatoryInt) ip = padStr(ip, obligatoryInt);
  if (useGrouping) ip = groupThousands(ip);
  // A zero value with any numeric placeholder still needs a visible integer.
  if (ip === "" && dp === undefined && maxFrac === 0 && /[#0?]/.test(body.replace(/\u0000\d+\u0000/g, ""))) ip = "0";

  // Structural analysis must ignore the literal markers, whose embedded
  // index digits would otherwise read as number placeholders.
  const structural = body.replace(/\u0000\d+\u0000/g, "");
  // A section with no numeric placeholder at all (e.g. bare "@" or a pure
  // literal such as "zero") renders no digits for a numeric value.
  const hasNumericPlaceholder = /[#0?]/.test(structural);
  let rendered = ip + (dp !== undefined ? "." + dp : "");
  if (!hasNumericPlaceholder) rendered = "";
  // Re-apply literals and literal text, dropping placeholders.
  let out = "";
  let bodyIdx = 0;
  let numInserted = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === "\u0000") {
      const end = body.indexOf("\u0000", i + 1);
      const li = parseInt(body.slice(i + 1, end), 10);
      out += literals[li] === undefined ? "" : literals[li];
      i = end;
      continue;
    }
    if (/[#0?]/.test(ch)) { if (!numInserted) { out += rendered; numInserted = true; } continue; }
    if (ch === "@") continue;
    if (ch === "." || ch === ",") {
      if (ch === ".") { if (!numInserted) { out += rendered; numInserted = true; } continue; }
      if (!numInserted) continue;
      continue;
    }
    out += ch;
  }
  if (!numInserted) out += rendered;
  else if (hasNumericPlaceholder && !/[#0?]/.test(structural)) out = rendered + out;

  if (!hasNumericPlaceholder) return out;

  // Excel shows negatives with a leading minus only when the section has no
  // explicit sign convention; ";"-negatives and ( ) forms do not.
  const showMinus = scaled < 0 && !negativeParens && !/\(/.test(section) && !/^-/.test(section.replace(/\s/g, ""));
  if (showMinus) out = "-" + out;
  return out;
}

/**
 * Render a date/time serial through a format code.
 *
 * The tokenizer tracks the previous token because "m" means month or minute
 * depending on what precedes it - the single most error-prone rule in the
 * whole format grammar.
 */
export function formatDateTime(value, code, opts = {}) {
  const p = serialParts(value);
  const elapsedH = Math.floor(value * 24);
  const elapsedM = Math.floor(value * 1440);
  const elapsedS = Math.floor(value * 86400);
  const twelveHour = /am\/pm|a\/p/i.test(code);
  let out = "";
  let i = 0;
  let prev = "";
  const n = code.length;
  while (i < n) {
    const ch = code[i];
    if (ch === '"') { const c = code.indexOf('"', i + 1); out += code.slice(i + 1, c < 0 ? n : c); i = c < 0 ? n : c + 1; continue; }
    if (ch === "\\") { out += code[i + 1] || ""; i += 2; continue; }
    if (ch === "_") { out += " "; i += 2; continue; }
    if (ch === "*") { i += 2; continue; }
    if (ch === "[") {
      const c = code.indexOf("]", i + 1);
      const body = c < 0 ? code.slice(i + 1) : code.slice(i + 1, c);
      const lm = body.match(/^([hms]+)$/i);
      if (lm) {
        const tok = lm[1][0].toLowerCase();
        const width = lm[1].length;
        if (tok === "h") { out += pad(elapsedH, width); prev = "h"; }
        else if (tok === "m") { out += pad(elapsedM, width); prev = "m"; }
        else { out += pad(elapsedS, width); prev = "s"; }
      }
      i = c < 0 ? n : c + 1;
      continue;
    }
    const token = ch.toLowerCase();
    if (token === "y" || token === "m" || token === "d" || token === "h" || token === "s" || token === "a" || token === "p" || token === "g") {
      let j = i;
      while (j < n && code[j].toLowerCase() === token) j++;
      const width = j - i;

      // AM/PM and A/P must match as whole tokens: the run scanner stops at "/".
      if (token === "a" || token === "p") {
        const upper = code.slice(i, i + 5).toUpperCase();
        if (upper === "AM/PM") {
          const shouty = code[i] === "A";
          out += p.h < 12 ? (shouty ? "AM" : "am") : (shouty ? "PM" : "pm");
          i += 5;
          continue;
        }
        if (upper.slice(0, 3) === "A/P") {
          out += p.h < 12 ? code[i] : code[i + 2];
          i += 3;
          continue;
        }
      }
      if (token === "g" && /^general$/i.test(code.slice(i, i + 7))) { out += generalNumber(value); i += 7; continue; }

      if (token === "m") {
        // Minutes only when adjacent to hours or seconds; otherwise a month.
        const isMinute = prev === "h" || prev === "m" || prev === "s" ||
          /^[\s:]*s/i.test(code.slice(j)) || /^[\s:]*\.0/i.test(code.slice(j));
        if (isMinute) {
          out += pad(p.mi, width >= 2 ? 2 : 1);
          prev = "m";
        } else {
          if (width === 1) out += String(p.m);
          else if (width === 2) out += pad(p.m, 2);
          else if (width === 3) out += MONTHS_ABBR[p.m - 1];
          else if (width === 4) out += MONTHS[p.m - 1];
          else out += MONTHS[p.m - 1][0];
          prev = "mo";
        }
        i = j;
        continue;
      }
      if (token === "d") {
        if (width === 1) out += String(p.d);
        else if (width === 2) out += pad(p.d, 2);
        else if (width === 3) out += DAYS_ABBR[p.dow];
        else out += DAYS[p.dow];
        prev = "d";
        i = j;
        continue;
      }
      if (token === "y") {
        out += width <= 2 ? pad(p.y % 100, 2) : String(p.y);
        prev = "y";
        i = j;
        continue;
      }
      if (token === "h") {
        let h = p.h;
        if (twelveHour) { h = h % 12; if (h === 0) h = 12; }
        out += width === 1 ? String(h) : pad(h, 2);
        prev = "h";
        i = j;
        continue;
      }
      if (token === "s") {
        out += pad(p.s, width >= 2 ? 2 : 1);
        prev = "s";
        const fm = code.slice(j).match(/^\.(0+)/);
        if (fm) {
          out += "." + String(p.ms).padStart(3, "0").slice(0, fm[1].length);
          i = j + fm[0].length;
          continue;
        }
        i = j;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return out;
}

/** "General" rendering, matching Excel's 11-significant-digit behaviour. */
export function generalNumber(v) {
  if (typeof v !== "number") return String(v);
  if (!Number.isFinite(v)) return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  const abs = Math.abs(v);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-4)) {
    const exp = v.toExponential(5);
    const [m, e] = exp.split("e");
    const mant = m.replace(/0+$/, "").replace(/\.$/, "");
    const sign = e[0] === "-" ? "-" : "+";
    const digits = e.replace(/^[+-]/, "").padStart(2, "0");
    return mant + "E" + sign + digits;
  }
  const s = String(parseFloat(v.toPrecision(11)));
  return s;
}

function formatScientific(v, code) {
  const m = code.match(/([#0?]+)(?:\.([#0?]+))?[eE]([+-]?)(0+)/);
  if (!m) return generalNumber(v);
  const intLen = m[1].length;
  const decLen = m[2] ? m[2].length : 0;
  const forceSign = m[3] === "+";
  const expDigits = m[4].length;
  if (v === 0) {
    return pad(0, intLen) + (decLen ? "." + "0".repeat(decLen) : "") + "E" + (forceSign ? "+" : "") + pad(0, expDigits);
  }
  let exp = Math.floor(Math.log10(Math.abs(v)));
  let mant = v / Math.pow(10, exp);
  let s = mant.toFixed(decLen);
  if (Math.abs(parseFloat(s)) >= Math.pow(10, intLen)) { exp += 1; mant = v / Math.pow(10, exp); s = mant.toFixed(decLen); }
  const [ip, dp] = s.split(".");
  return pad(ip.replace("-", ""), intLen) + (decLen ? "." + (dp || "0".repeat(decLen)) : "") + "E" +
    (exp < 0 ? "-" : forceSign ? "+" : "") + pad(Math.abs(exp), expDigits);
}

function formatFraction(v, code, m) {
  const wholeFmt = m[1];
  const numFmt = m[2];
  const denFmt = m[3];
  const fixedDen = /^[0-9]+$/.test(denFmt) ? parseInt(denFmt, 10) : 0;
  const maxDen = fixedDen || Math.pow(10, denFmt.length) - 1;
  const sign = v < 0 ? "-" : "";
  const abs = Math.abs(v);
  const whole = Math.floor(abs);
  const rest = abs - whole;
  const fr = fixedDen ? { n: Math.round(rest * fixedDen), d: fixedDen } : bestFraction(rest, maxDen);
  const wholeStr = wholeFmt && /0/.test(wholeFmt) ? String(whole) : whole === 0 ? "" : String(whole);
  // "?" reserves one character of width so the fraction bar lines up with the
  // widest value the code can produce.
  const numPad = /^\?+$/.test(numFmt) ? numFmt.length : 0;
  const denPad = /^\?+$/.test(denFmt) ? denFmt.length : 0;
  if (fr.n === 0) return sign + (wholeStr === "" ? "0" : wholeStr);
  const numStr = padStr(String(fr.n), numPad, " ");
  const denStr = padStr(String(fr.d), denPad, " ");
  return sign + (wholeStr ? wholeStr + " " : "") + numStr + "/" + denStr;
}

/** Format any cell value for display. */
export function formatValue(value, code, opts = {}) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  if (value instanceof Date) return formatDateTime(dateToSerial(value, opts.date1904 === true), code || "m/d/yy", opts);
  return formatNumber(value, code, opts);
}
