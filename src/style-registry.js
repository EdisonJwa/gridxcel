/**
 * The style table: number formats, fonts, fills, borders and cell formats.
 *
 * One module owns both directions. Reading, \`parseStyles\` turns a styles.xml
 * into the format table a workbook carries; writing, \`StyleTable\` interns the
 * styles a workbook uses and emits styles.xml. They share the key and element
 * functions below on purpose: a key that disagrees with the element it keys
 * makes an equivalent style intern twice, which is exactly the defect that kept
 * the table from converging in one round trip (MINOR-5).
 */

import { XmlParser, encodeXmlAttr, localName } from "./xml.js";
import { BUILTIN_FORMATS, isDateFormat } from "./numfmt.js";
import { GridxcelError } from "./errors.js";
import { XML_DECL, NS_MAIN } from "./constants.js";

/* ------------------------------------------------------------- constants */

const FIRST_CUSTOM_NUMFMT_ID = 164;
const BUILTIN_BY_CODE =  new Map();
for (const key of Object.keys(BUILTIN_FORMATS)) {
  const id = Number(key);
  const code = BUILTIN_FORMATS[key];
  if (!BUILTIN_BY_CODE.has(code)) BUILTIN_BY_CODE.set(code, id);
}
const DEFAULT_FONT_NAME = "Calibri";
const DEFAULT_FONT_SIZE = 11;
const BORDER_ORDER = ["left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"];
const BORDER_ALWAYS =  new Set(["left", "right", "top", "bottom", "diagonal"]);
/** Fast membership test for "is this a legal border edge name". */
const EDGE_NAMES = new Set(BORDER_ORDER);
const HEX6 = /^[0-9A-Fa-f]{6}$/;
const HEX8 = /^[0-9A-Fa-f]{8}$/;
const DEFAULT_FONT = { name: DEFAULT_FONT_NAME, sz: DEFAULT_FONT_SIZE };
const DEFAULT_XF = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: null, protection: null };

/* --------------------------------------------------------------- helpers */

export function normalizeColor(c) {
  if (c === void 0 || c === null) return void 0;
  if (typeof c === "string") {
let s = c.trim();
if (s.length > 0 && s.charCodeAt(0) === 35) s = s.slice(1);
if (HEX8.test(s)) return s.toUpperCase();
if (HEX6.test(s)) return "FF" + s.toUpperCase();
throw new GridxcelError("invalid colour " + JSON.stringify(c) + ": expected RRGGBB or AARRGGBB hex", "BAD_STYLE");
  }
  if (typeof c === "object") {
if (typeof c.rgb === "string") return normalizeColor(c.rgb);
if (typeof c.theme === "number") {
const out = { theme: c.theme };
if (typeof c.tint === "number") out.tint = c.tint;
return out;
}
if (typeof c.indexed === "number") return { indexed: c.indexed };
if (c.auto === true) return { auto: true };
throw new GridxcelError("invalid colour object " + JSON.stringify(c), "BAD_STYLE");
  }
  throw new GridxcelError("invalid colour " + String(c), "BAD_STYLE");
}

function colorKey(c) {
  if (typeof c === "string") return "r" + c;
  if (c.theme !== void 0) return "t" + c.theme + (c.tint === void 0 ? "" : ":" + c.tint);
  if (c.indexed !== void 0) return "i" + c.indexed;
  return "auto";
}

export function colorElement(c) {
  if (typeof c === "string") return '<color rgb="' + c + '"/>';
  if (c.theme !== void 0) {
return '<color theme="' + c.theme + '"' + (c.tint === void 0 ? "" : ' tint="' + c.tint + '"') + "/>";
  }
  if (c.indexed !== void 0) return '<color indexed="' + c.indexed + '"/>';
  return '<color auto="1"/>';
}

function normalizeFont(f) {
  if (f === void 0 || f === null) return DEFAULT_FONT;
  if (typeof f !== "object") throw new GridxcelError("font must be an object", "BAD_STYLE");
  return {
// A partial font record would otherwise inherit the viewer default; fill
// the app default in so every generated file renders identically.
name: f.name === void 0 ? DEFAULT_FONT_NAME : String(f.name),
    sz: normalizeFontSize(f.sz),
b: f.b === true,
i: f.i === true,
u: f.u === true,
strike: f.strike === true,
vertAlign: f.vertAlign && f.vertAlign !== "baseline" ? f.vertAlign : void 0,
color: normalizeColor(f.color)
  };
}

/**
 * A font size in points: 1 to 409, and finite.
 *
 * @param {unknown} value the caller's `font.sz`
 * @returns {number}
 */
function normalizeFontSize(value) {
  if (value === undefined || value === null) return DEFAULT_FONT_SIZE;
  const n = Number(value);
  if (!Number.isFinite(n)) {
    throw new GridxcelError(
      "font size " + String(value) + " is not a finite number; a size must be between 1 and 409 points",
      "BAD_STYLE", { sz: value });
  }
  if (n < 1 || n > 409) {
    throw new GridxcelError(
      "font size " + n + " is outside the supported range of 1 to 409 points",
      "BAD_STYLE", { sz: value });
  }
  return n;
}

function fontKey(f) {
  return [
f.name,
f.sz,
f.b ? 1 : 0,
f.i ? 1 : 0,
f.u ? 1 : 0,
f.strike ? 1 : 0,
f.vertAlign || "",
f.color ? colorKey(f.color) : ""
  ].join("|");
}

function fontElement(f) {
  let s = "<font>";
  if (f.b) s += "<b/>";
  if (f.i) s += "<i/>";
  if (f.strike) s += "<strike/>";
  if (f.u) s += "<u/>";
  if (f.vertAlign) s += '<vertAlign val="' + f.vertAlign + '"/>';
  s += '<sz val="' + f.sz + '"/>';
  if (f.color) s += colorElement(f.color);
  s += '<name val="' + encodeXmlAttr(f.name) + '"/>';
  return s + "</font>";
}

function normalizeFill(f) {
  if (f === void 0 || f === null) return { pattern: "none", fgColor: void 0, bgColor: void 0 };
  if (typeof f !== "object") throw new GridxcelError("fill must be an object", "BAD_STYLE");
  const fgColor = normalizeColor(f.fgColor);
  const bgColor = normalizeColor(f.bgColor);
  let pattern = f.pattern;
  if (pattern === void 0) pattern = fgColor !== void 0 || bgColor !== void 0 ? "solid" : "none";
  return { pattern, fgColor, bgColor };
}

function fillKey(f) {
  return [f.pattern, f.fgColor ? colorKey(f.fgColor) : "", f.bgColor ? colorKey(f.bgColor) : ""].join("|");
}

function fillElement(f) {
  if (f.fgColor === void 0 && f.bgColor === void 0) {
return '<fill><patternFill patternType="' + f.pattern + '"/></fill>';
  }
  let s = '<fill><patternFill patternType="' + f.pattern + '">';
  if (f.fgColor) s += "<fgColor" + colorAttrs(f.fgColor) + "/>";
  if (f.bgColor) s += "<bgColor" + colorAttrs(f.bgColor) + "/>";
  return s + "</patternFill></fill>";
}

export function colorAttrs(c) {
  if (typeof c === "string") return ' rgb="' + c + '"';
  if (c.theme !== void 0) return ' theme="' + c.theme + '"' + (c.tint === void 0 ? "" : ' tint="' + c.tint + '"');
  if (c.indexed !== void 0) return ' indexed="' + c.indexed + '"';
  return ' auto="1"';
}

function normalizeBorder(b) {
  if (b === void 0 || b === null) return { edges: {}, diagonalUp: false, diagonalDown: false };
  if (typeof b !== "object") throw new GridxcelError("border must be an object", "BAD_STYLE");
  const edges = {};
  for (const name of BORDER_ORDER) {
const edge = b[name];
if (edge === void 0 || edge === null) continue;
edges[name] = { style: edge.style, color: normalizeColor(edge.color) };
  }
  return { edges, diagonalUp: b.diagonalUp === true, diagonalDown: b.diagonalDown === true };
}

function edgeIsBlank(e) {
  return e === void 0 || e === null || e.style === void 0 && e.color === void 0;
}

function borderKey(b) {
  let k = (b.diagonalUp ? "U" : "") + (b.diagonalDown ? "D" : "");
  for (const name of BORDER_ORDER) {
const e = b.edges[name];
k += "|" + name + (edgeIsBlank(e) ? "" : ":" + (e.style || "") + "," + (e.color ? colorKey(e.color) : ""));
  }
  return k;
}

function borderElement(b) {
  let s = "<border";
  if (b.diagonalUp) s += ' diagonalUp="1"';
  if (b.diagonalDown) s += ' diagonalDown="1"';
  s += ">";
  for (const name of BORDER_ORDER) {
const e = b.edges[name];
if (edgeIsBlank(e)) {
if (BORDER_ALWAYS.has(name) || e !== void 0) s += "<" + name + "/>";
continue;
}
s += "<" + name + (e.style ? ' style="' + e.style + '"' : "");
s += e.color ? ">" + colorElement(e.color) + "</" + name + ">" : "/>";
  }
  return s + "</border>";
}

function normalizeAlignment(a) {
  if (a === void 0 || a === null) return null;
  if (typeof a !== "object") throw new GridxcelError("alignment must be an object", "BAD_STYLE");
  const out = {};
  if (a.horizontal) out.horizontal = a.horizontal;
  if (a.vertical) out.vertical = a.vertical;
  if (a.textRotation !== void 0) out.textRotation = a.textRotation;
  if (a.wrapText) out.wrapText = true;
  if (a.shrinkToFit) out.shrinkToFit = true;
  if (a.indent !== void 0) out.indent = a.indent;
  if (a.readingOrder !== void 0) out.readingOrder = a.readingOrder;
  return Object.keys(out).length ? out : null;
}

function alignmentKey(a) {
  return [
a.horizontal || "",
a.vertical || "",
a.textRotation === void 0 ? "" : a.textRotation,
a.wrapText ? 1 : 0,
a.shrinkToFit ? 1 : 0,
a.indent === void 0 ? "" : a.indent,
a.readingOrder === void 0 ? "" : a.readingOrder
  ].join(",");
}

function alignmentElement(a) {
  let s = "<alignment";
  if (a.horizontal) s += ' horizontal="' + a.horizontal + '"';
  if (a.vertical) s += ' vertical="' + a.vertical + '"';
  if (a.textRotation !== void 0) s += ' textRotation="' + a.textRotation + '"';
  if (a.wrapText) s += ' wrapText="1"';
  if (a.shrinkToFit) s += ' shrinkToFit="1"';
  if (a.indent !== void 0) s += ' indent="' + a.indent + '"';
  if (a.readingOrder !== void 0) s += ' readingOrder="' + a.readingOrder + '"';
  return s + "/>";
}

function normalizeProtection(p) {
  if (p === void 0 || p === null) return null;
  if (typeof p !== "object") throw new GridxcelError("protection must be an object", "BAD_STYLE");
  const out = {};
  if (p.locked !== void 0) out.locked = p.locked !== false;
  if (p.hidden !== void 0) out.hidden = p.hidden === true;
  return Object.keys(out).length ? out : null;
}

function protectionKey(p) {
  return (p.locked === void 0 ? "" : p.locked ? "1" : "0") + (p.hidden === void 0 ? "" : p.hidden ? "1" : "0");
}

function protectionElement(p) {
  let s = "<protection";
  if (p.locked !== void 0) s += ' locked="' + (p.locked ? "1" : "0") + '"';
  if (p.hidden !== void 0) s += ' hidden="' + (p.hidden ? "1" : "0") + '"';
  return s + "/>";
}

function xfKey(x) {
  return [
x.numFmtId,
x.fontId,
x.fillId,
x.borderId,
x.alignment ? alignmentKey(x.alignment) : "",
x.protection ? protectionKey(x.protection) : ""
  ].join("/");
}

function xfElement(x) {
  let s = '<xf numFmtId="' + x.numFmtId + '" fontId="' + x.fontId + '" fillId="' + x.fillId + '" borderId="' + x.borderId + '" xfId="0"';
  if (x.numFmtId !== 0) s += ' applyNumberFormat="1"';
  if (x.fontId !== 0) s += ' applyFont="1"';
  if (x.fillId !== 0) s += ' applyFill="1"';
  if (x.borderId !== 0) s += ' applyBorder="1"';
  if (x.alignment) s += ' applyAlignment="1"';
  if (x.protection) s += ' applyProtection="1"';
  if (!x.alignment && !x.protection) return s + "/>";
  s += ">";
  if (x.alignment) s += alignmentElement(x.alignment);
  if (x.protection) s += protectionElement(x.protection);
  return s + "</xf>";
}

function attrTruthy(v) {
  if (v === void 0 || v === "") return true;
  return v === "1" || v === "true" || v === "on";
}

function toInt(v) {
  if (v === void 0 || v === "") return NaN;
  const n = Number(v);
  return Number.isInteger(n) ? n : NaN;
}

function colorFromAttrs(attrs) {
  if (attrs.rgb !== void 0) return normalizeColor(attrs.rgb);
  if (attrs.theme !== void 0) {
const out = { theme: Number(attrs.theme) };
if (attrs.tint !== void 0) out.tint = Number(attrs.tint);
return out;
  }
  if (attrs.indexed !== void 0) return { indexed: Number(attrs.indexed) };
  if (attrs.auto !== void 0) return { auto: true };
  return void 0;
}

function alignmentFromAttrs(attrs) {
  const out = {};
  if (attrs.horizontal) out.horizontal = attrs.horizontal;
  if (attrs.vertical) out.vertical = attrs.vertical;
  if (attrs.textRotation !== void 0) out.textRotation = Number(attrs.textRotation);
  if (attrTruthy(attrs.wrapText) && attrs.wrapText !== void 0) out.wrapText = true;
  if (attrTruthy(attrs.shrinkToFit) && attrs.shrinkToFit !== void 0) out.shrinkToFit = true;
  if (attrs.indent !== void 0) out.indent = Number(attrs.indent);
  if (attrs.readingOrder !== void 0) out.readingOrder = Number(attrs.readingOrder);
  return Object.keys(out).length ? out : null;
}

function protectionFromAttrs(attrs) {
  const out = {};
  if (attrs.locked !== void 0) out.locked = attrTruthy(attrs.locked);
  if (attrs.hidden !== void 0) out.hidden = attrTruthy(attrs.hidden);
  return Object.keys(out).length ? out : null;
}

export class StyleTable {
  constructor() {
    this.numFmts = /* @__PURE__ */ new Map();
    this._codeById = /* @__PURE__ */ new Map();
    this._font = { list: [], byKey: /* @__PURE__ */ new Map() };
    this._fill = { list: [], byKey: /* @__PURE__ */ new Map() };
    this._border = { list: [], byKey: /* @__PURE__ */ new Map() };
    this._xf = { list: [], byKey: /* @__PURE__ */ new Map() };
    this._cache = /* @__PURE__ */ new WeakMap();
    this._intern(this._font, fontKey(DEFAULT_FONT), fontElement(DEFAULT_FONT));
    this._intern(this._fill, fillKey({ pattern: "none" }), fillElement({ pattern: "none" }));
    this._intern(this._fill, fillKey({ pattern: "gray125" }), fillElement({ pattern: "gray125" }));
    this._intern(this._border, borderKey(normalizeBorder(null)), borderElement(normalizeBorder(null)));
    this._xf.list.push(xfElement(DEFAULT_XF));
    this._xf.byKey.set(xfKey(DEFAULT_XF), 0);
  }
  get fontCount() {
    return this._font.list.length;
  }
  get fillCount() {
    return this._fill.list.length;
  }
  get borderCount() {
    return this._border.list.length;
  }
  get xfCount() {
    return this._xf.list.length;
  }
  get numFmtCount() {
    return this.numFmts.size;
  }
  _intern(group, key, xml) {
    let i = group.byKey.get(key);
    if (i === void 0) {
      i = group.list.length;
      group.list.push(xml);
      group.byKey.set(key, i);
    }
    return i;
  }
  /** Format code or id -> numFmtId, reusing built-ins and deduping customs. */
  numFmtId(code) {
    if (code === void 0 || code === null || code === "") return 0;
    if (typeof code === "number") {
      if (!Number.isInteger(code) || code < 0) {
        throw new GridxcelError("invalid number format id " + code, "BAD_STYLE");
      }
      return code;
    }
    if (typeof code !== "string") {
      throw new GridxcelError("number format must be a code string or an id, got " + typeof code, "BAD_STYLE");
    }
    const builtin = BUILTIN_BY_CODE.get(code);
    if (builtin !== void 0) return builtin;
    const hit = this.numFmts.get(code);
    if (hit !== void 0) return hit;
    const id = FIRST_CUSTOM_NUMFMT_ID + this.numFmts.size;
    this.numFmts.set(code, id);
    this._codeById.set(id, code);
    return id;
  }
  /** Format code for an id this table allocated (used by tests/diagnostics). */
  codeForId(id) {
    return this._codeById.get(id);
  }
  /** StyleObject -> cellXfs index. Falsy styles are the mandatory default 0. */
  xf(style) {
    if (style === void 0 || style === null) return 0;
    if (typeof style === "number") {
      if (!Number.isInteger(style) || style < 0) {
        throw new GridxcelError("invalid style " + style, "BAD_STYLE");
      }
      return this._xfFor({ numFmt: style });
    }
    if (typeof style !== "object") {
      throw new GridxcelError("style must be an object, got " + typeof style, "BAD_STYLE");
    }
    const hit = this._cache.get(style);
    if (hit !== void 0) return hit;
    const idx = this._xfFor(style);
    this._cache.set(style, idx);
    return idx;
  }
  _xfFor(style) {
    const numFmtId = this.numFmtId(style.numFmt);
    const f = normalizeFont(style.font);
    const fi = normalizeFill(style.fill);
    const b = normalizeBorder(style.border);
    const x = {
      numFmtId,
      fontId: this._intern(this._font, fontKey(f), fontElement(f)),
      fillId: this._intern(this._fill, fillKey(fi), fillElement(fi)),
      borderId: this._intern(this._border, borderKey(b), borderElement(b)),
      alignment: normalizeAlignment(style.alignment),
      protection: normalizeProtection(style.protection)
    };
    return this._intern(this._xf, xfKey(x), xfElement(x));
  }
  /** The complete styles.xml for the records collected so far. */
  toXml() {
    let s = XML_DECL + '<styleSheet xmlns="' + NS_MAIN + '">';
    if (this.numFmts.size > 0) {
      s += '<numFmts count="' + this.numFmts.size + '">';
      for (const [code, id] of this.numFmts) {
        s += '<numFmt numFmtId="' + id + '" formatCode="' + encodeXmlAttr(code) + '"/>';
      }
      s += "</numFmts>";
    }
    s += '<fonts count="' + this._font.list.length + '">' + this._font.list.join("") + "</fonts>";
    s += '<fills count="' + this._fill.list.length + '">' + this._fill.list.join("") + "</fills>";
    s += '<borders count="' + this._border.list.length + '">' + this._border.list.join("") + "</borders>";
    s += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    s += '<cellXfs count="' + this._xf.list.length + '">' + this._xf.list.join("") + "</cellXfs>";
    s += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    return s + "</styleSheet>";
  }
};

export function parseStyles(xml) {
  const numFmts =  new Map();
  const fonts = [];
  const fills = [];
  const borders = [];
  const cellXfs = [];
  let font = null;
  let fill = null;
  let border = null;
  let xf = null;
  const parser = new XmlParser({
    onOpen(name, attrs, parent) {
      const local = localName(name);
      const pl = parent ? parent.local : null;
      switch (pl) {
        case "numFmts":
          if (local === "numFmt") {
            const id = toInt(attrs.numFmtId);
            if (!Number.isNaN(id)) numFmts.set(id, attrs.formatCode === void 0 ? "" : attrs.formatCode);
          }
          return;
        case "fonts":
          if (local === "font") {
            font = {};
            fonts.push(font);
          }
          return;
        case "font": {
          if (font === null) return;
          if (local === "b") font.b = attrTruthy(attrs.val);
          else if (local === "i") font.i = attrTruthy(attrs.val);
          else if (local === "strike") font.strike = attrTruthy(attrs.val);
          else if (local === "u") font.u = true;
          else if (local === "sz") {
            const sz = Number(attrs.val);
            if (Number.isFinite(sz)) font.sz = sz;
          } else if (local === "name") font.name = attrs.val;
          else if (local === "vertAlign") {
            if (attrs.val !== "baseline") font.vertAlign = attrs.val;
          } else if (local === "color") font.color = colorFromAttrs(attrs);
          return;
        }
        case "fills":
          if (local === "fill") {
            fill = { pattern: "none" };
            fills.push(fill);
          }
          return;
        case "fill":
          if (fill === null) return;
          if (local === "patternFill") fill.pattern = attrs.patternType || "none";
          return;
        case "patternFill":
          if (fill === null) return;
          if (local === "fgColor") fill.fgColor = colorFromAttrs(attrs);
          else if (local === "bgColor") fill.bgColor = colorFromAttrs(attrs);
          return;
        case "borders":
          if (local === "border") {
            border = {
              edges: {},
              diagonalUp: attrTruthy(attrs.diagonalUp) && attrs.diagonalUp !== void 0,
              diagonalDown: attrTruthy(attrs.diagonalDown) && attrs.diagonalDown !== void 0
            };
            borders.push(border);
          }
          return;
        case "cellXfs":
          if (local === "xf") {
            xf = {
              numFmtId: toInt(attrs.numFmtId) || 0,
              fontId: toInt(attrs.fontId) || 0,
              fillId: toInt(attrs.fillId) || 0,
              borderId: toInt(attrs.borderId) || 0
            };
            cellXfs.push(xf);
          }
          return;
        case "xf":
          if (xf === null) return;
          if (local === "alignment") xf.alignment = alignmentFromAttrs(attrs);
          else if (local === "protection") xf.protection = protectionFromAttrs(attrs);
          return;
        default:
          if (border !== null && EDGE_NAMES.has(local)) {
            const edge = {};
            if (attrs.style !== void 0 && attrs.style !== "") edge.style = attrs.style;
            border.edges[local] = edge;
            return;
          }
          if (border !== null && local === "color" && pl !== null && EDGE_NAMES.has(pl)) {
            const edge = border.edges[pl];
            if (edge) {
              const c = colorFromAttrs(attrs);
              if (c !== void 0) edge.color = c;
            }
            return;
          }
          if (xf !== null && (local === "alignment" || local === "protection")) return;
      }
    }
  });
  if (typeof xml === "string" && xml.length > 0) {
    parser.write(xml);
    parser.end();
  }
  for (const key of Object.keys(BUILTIN_FORMATS)) {
    const id = Number(key);
    if (!numFmts.has(id)) numFmts.set(id, BUILTIN_FORMATS[key]);
  }
  const codeForId = (id) => {
    const hit = numFmts.get(id);
    return hit === void 0 ? null : hit;
  };
  const styles = cellXfs.map((x) => {
    const out = { numFmt: x.numFmtId, numFmtCode: codeForId(x.numFmtId) };
    const f = fonts[x.fontId];
    if (f) out.font = f;
    const fi = fills[x.fillId];
    if (fi) out.fill = fi;
    const b = borders[x.borderId];
    if (b) out.border = { ...b.edges, diagonalUp: b.diagonalUp, diagonalDown: b.diagonalDown };
    if (x.alignment) out.alignment = x.alignment;
    if (x.protection) out.protection = x.protection;
    return out;
  });
  const isDate = (styleIndex) => {
    const x = cellXfs[styleIndex];
    if (x === void 0) return false;
    const code = numFmts.get(x.numFmtId);
    return isDateFormat(code === void 0 ? x.numFmtId : code);
  };
  return { numFmts, styles, isDate, fonts, fills, borders, cellXfs };
}
