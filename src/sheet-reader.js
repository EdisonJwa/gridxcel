/**
 * Worksheet part -> Sheet (materialised) or a bounded-memory row stream.
 *
 * Why not src/xml.js' XmlParser for worksheets: that parser keeps the whole
 * unconsumed tail in a string and re-slices it once per fed chunk, which is
 * O(n^2) in the number of chunks for a 100 MB part. This module runs its own
 * tag scanner over a carry buffer that is trimmed after every chunk, so peak
 * memory stays at one chunk plus one row no matter how large the sheet is.
 *
 * Cell addressing here is absolute: cells[col] with A == 0. parseSheetXml
 * rebases rows onto sheet.originCol/originRow at the end, when the minimum
 * column of the sheet is finally known; createRowStream cannot rebase
 * retroactively, so it stays absolute (originCol is effectively 0).
 */

import { NO_ATTRS, decodeEntities, decodeOoxmlEscapes, localName, tagLocalIfNoAttrs, parseTag, scanTagEnd } from "./xml.js";
import { decodeCell, decodeCol, encodeCol, makeBounds, extendBounds, boundsToRef } from "./address.js";
import { formatValue, serialToDate, dateToSerial } from "./numfmt.js";
import { GridxcelError } from "./errors.js";
import { MAX_COLUMNS, MAX_ROWS } from "./constants.js";

/** Elements whose content is character data we must capture. */
const TEXT_ELEMENTS = new Set(["v", "t", "f"]);
/** Phonetic hints carry <t> children that are NOT part of the value. */
const PHONETIC = new Set(["rPh", "phoneticPr"]);
/** 256 KiB slices keep the carry buffer small without syscall-sized chunks. */
const SLICE_BYTES = 1 << 18;

/**
 * Largest gap, in rows, that a row element may introduce before it is treated
 * as suspicious rather than as ordinary sparse layout. A crafted worksheet can
 * put one row element at r="40000000"; real sheets do not jump by a million
 * rows, and treating such a row as ordinary would set Array#length to its index
 * and hand every consumer a forty-million-row range.
 */
const EMPTY_ROW_SLACK = 65536;

const noop = () => {};

/* ------------------------------------------------------------------ scanner */

function intOf(v) {
  if (v === undefined) return undefined;
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

function boolOf(v) {
  return v === "" || v === "1" || v === "true" || v === "on";
}

/**
 * End of the tag that starts at s[lt], quote aware.
 *
 * A '>' inside a quoted attribute value is legal, so the first '>' is not
 * necessarily the end of the tag. The previous implementation looked for the
 * next '"' and '\'' from the tag start with String#indexOf, which is unbounded:
 * a worksheet with no apostrophes anywhere made every attribute-less tag scan
 * the entire remaining buffer looking for a character that was never there.
 * That is quadratic in the buffer size - measured at 724 ms for 16 KiB chunks
 * against 18,730 ms for 1 MiB chunks on the same document, and it is the whole
 * reason the synchronous reader was six times slower than the streaming one.
 *
 */
/** Start index of the close tag for "name", or -1 when it is not in yet. */
function findCloseTag(s, name, from) {
  let at = from;
  for (;;) {
    const hit = s.indexOf("</" + name, at);
    if (hit < 0) return -1;
    const after = s.charCodeAt(hit + 2 + name.length);
    // NaN means the buffer ends inside the close tag: it is still ours.
    if (!(after >= 0) || after === 62 || after === 32 || after === 9 || after === 10 || after === 13) return hit;
    at = hit + 2;
  }
}

/**
 * Character data of a text element. Values survive byte-for-byte except for a
 * newline-indented body written without xml:space="preserve", which only ever
 * comes from hand-written or pretty-printed parts.
 */
function textValue(raw, preserve) {
  if (preserve) return raw;
  if (raw.length > 1 && raw.trim() !== raw && /[\r\n]/.test(raw)) return raw.trim();
  return raw;
}

class RowScanner {
  constructor(opts = {}) {
    this.shared = opts.sharedStrings || null;
    this.styles = opts.styles || null;
    this.date1904 = opts.date1904 === true;
    this.cellDates = opts.cellDates === true;
    this.cellText = opts.cellText === true;
    this.cellFormula = opts.cellFormula !== false;
    this.sheetStubs = opts.sheetStubs === true;
    this.bounds = opts.bounds || null;
    this.collectMeta = opts.collectMeta !== false;
    this.warn = typeof opts.warn === "function" ? opts.warn : noop;
    this.onRow = typeof opts.onRow === "function" ? opts.onRow : null;
    /** xf index -> resolved format code or null; one lookup per distinct style. */
    this.codeCache = new Map();
    // Comments are known before the worksheet is read, so a streamed cell can
    // carry them exactly like a batch cell. Keyed by (0-based row << 14) | col.
    this.commentAt = null;
    if (opts.comments && typeof opts.comments.forEach === "function" && opts.comments.size > 0) {
      const index = new Map();
      opts.comments.forEach((text, ref) => {
        try {
          const a = decodeCell(ref);
          index.set(a.r * MAX_COLUMNS + a.c, text);
        } catch { /* a malformed comment ref is reported through warnings */ }
      });
      if (index.size > 0) this.commentAt = index;
    }

    this.buf = "";
    this.textEl = null;
    this.textPreserve = false;
    this.row = null;
    this.cell = null;
    this.skipText = 0;
    this.lastRowNum = 0;
    this.inSheetData = false;
    this.sawSheetData = false;
    this.ended = false;
    this.warned = new Set();

    this.dims = undefined;
    this.columns = [];
    this.rowInfo = new Map();
    this.merges = [];
    this.hyperlinks = [];
    this.freeze = undefined;
    this.autoFilter = undefined;
    this.tabColor = undefined;
    this.legacyRelId = undefined;
    this.formulaMasters = new Map();

    /** Raw text pieces of a text element still waiting for its close tag. */
    this.textParts = [];
    /** Tail kept back because it may hold the start of a split close tag. */
    this.textOverlap = "";
    /** Body pieces of a tag whose '>' has not arrived yet, or null. */
    this.tagParts = null;
    /** Quote state the unfinished tag's scan stopped in. */
    this.tagQuote = 0;

    this.minRow = 0;
    this.maxRowNum = 0;
    this.minCol = -1;
    this.maxCol = -1;
    this.rowCount = 0;
    this.strayCells = 0;
  }

  /* ------------------------------------------------------------ chunk feed */

  write(str) {
    if (this.ended) return;
    // Text that ran past the previous chunk is held in textParts/textOverlap
    // rather than in this.buf, so the buffer never grows with the text and is
    // never rescanned from the beginning.
    //
    // It used to be kept in this.buf, which meant every chunk re-concatenated
    // the whole accumulated text and rescanned it for the close tag: quadratic
    // in the length of a single text node. One cell holding a 100 MB inline
    // string cost 116 s of CPU and 1.8 GB of heap from a 104 kB upload; 8 MB
    // already took 1.2 s and quadrupled per doubling. The pieces are joined
    // once, when the close tag finally arrives, so entity references split
    // across a chunk boundary are still decoded as one string.
    let buf;
    if (this.textEl !== null) {
      buf = this.textOverlap.length ? this.textOverlap + str : str;
      this.textOverlap = "";
    } else if (this.tagParts !== null) {
      // A tag is mid-flight; its pieces live in tagParts, not in this.buf, so
      // there is nothing to prepend and no growing string to flatten.
      buf = str;
    } else {
      buf = this.buf.length ? this.buf + str : str;
    }
    const n = buf.length;
    let i = 0;
    for (;;) {
      if (this.textEl !== null) {
        const k = findCloseTag(buf, this.textEl, i);
        if (k < 0) {
          // Keep only as much as a split close tag could occupy ("</" + name
          // plus one character), and hand the rest to the pending list.
          const keep = this.textEl.length + 3;
          const cut = Math.max(i, n - keep);
          if (cut > i) this.textParts.push(buf.slice(i, cut));
          this.textOverlap = buf.slice(cut);
          i = n;
          break;
        }
        const text = this.textParts.length === 0
          ? buf.slice(i, k)
          : this.textParts.join("") + buf.slice(i, k);
        if (this.textParts.length !== 0) this.textParts.length = 0;
        this.captureText(text, this.textEl);
        this.textEl = null;
        i = k;
      }
      if (this.tagParts !== null) {
        // Continuing a tag whose attribute value ran past the last chunk. Only
        // this chunk is scanned; everything before it was scanned already and
        // is waiting in tagParts, so the work is proportional to the input.
        const r = scanTagEnd(buf, 0, this.tagQuote);
        if (!r.done) {
          this.tagQuote = r.quote;
          if (n > 0) this.tagParts.push(buf);
          i = n;
          break;
        }
        const tag = this.tagParts.join("") + buf.slice(0, r.gt);
        this.tagParts = null;
        i = r.gt + 1;
        this.tag(tag);
        continue;
      }
      const lt = buf.indexOf("<", i);
      if (lt < 0) { i = n; break; }
      const r = scanTagEnd(buf, lt + 1, 0);
      if (!r.done) {
        // The tag is not finished: park its body and the quote state it ended
        // in, and clear the buffer so nothing is concatenated or rescanned.
        this.tagParts = [buf.slice(lt + 1)];
        this.tagQuote = r.quote;
        i = n;
        break;
      }
      const tag = buf.slice(lt + 1, r.gt);
      i = r.gt + 1;
      this.tag(tag);
    }
    this.buf = i >= n ? "" : buf.slice(i);
  }

  end() {
    if (this.ended) return;
    this.ended = true;
    if (this.buf.length > 0 || this.row !== null || this.cell !== null) {
      this.warnAlways("BAD_XML", "worksheet XML ends in the middle of an element");
    }
    this.buf = "";
    this.textParts = [];
    this.textOverlap = "";
    this.tagParts = null;
    this.tagQuote = 0;
    this.row = null;
    this.cell = null;
  }

  warnOnce(code, message, extra) {
    if (this.warned.has(code)) return;
    this.warned.add(code);
    this.warn(code, message, extra);
  }

  /** Truncation is worth reporting once even if the same code was used. */
  warnAlways(code, message, extra) {
    this.warn(code, message, extra);
  }

  /* -------------------------------------------------------------- tag layer */

  tag(raw) {
    if (raw.charCodeAt(0) === 47) {
      // Close tags are half of all tags. Slice once and take the name directly:
      // a closing tag never has attributes, so there is nothing to strip.
      this.closeTag(nameOf(raw, 1));
      return;
    }
    const last = raw.length - 1;
    const selfClose = raw.charCodeAt(last) === 47;
    const body = selfClose ? raw.slice(0, last) : raw;
    // 79% of the tags in a real worksheet carry no attributes, and for those the
    // only thing wanted is the local name. One pass answers both questions: a
    // null means "has attributes, call parseTag", which is the case that needs
    // the full walk anyway.
    //
    // The two branches must then share the tail: an early return here skipped the
    // text-capture setup, so <v> and <t> content was never collected and every
    // sheet came back with zero rows.
    const quickLocal = tagLocalIfNoAttrs(body);
    const tag = quickLocal === null ? parseTag(body) : null;
    const local = quickLocal === null ? tag.local : quickLocal;
    const attrs = quickLocal === null ? tag.attrs : NO_ATTRS;
    if (selfClose) {
      this.openTag(local, attrs, true);
      return;
    }
    this.openTag(local, attrs, false);
    // Text capture is only ever needed inside a cell; gating on that keeps
    // extension markup (x14:sparklineGroup, xm:f) from hijacking the scanner.
    if (TEXT_ELEMENTS.has(local) && this.cell !== null) {
      this.textEl = local;
      this.textPreserve = attrs["xml:space"] === "preserve";
    }
  }

  captureText(raw, name) {
    if (raw.length === 0) return;
    if (name === "t" && this.skipText > 0) return;
    // "<![CDATA[" - the marker starts at offset 2, not 1. Testing for 1 meant a
    // CDATA section was never unwrapped and its delimiters were returned as part
    // of the cell value. Both the sync and streaming readers had the same bug.
    if (raw.charCodeAt(0) === 60 && raw.startsWith("<![CDATA[")) {
      const close = raw.lastIndexOf("]]" + ">");
      this.pushText(name, close < 0 ? raw.slice(9) : raw.slice(9, close), true);
      return;
    }
    // A CDATA section that is only part of the text node. "]]>" cannot occur in
    // XML character data outside CDATA, so removing the delimiters is safe.
    if (raw.indexOf("<![CDATA[") >= 0) {
      this.pushText(name, raw.split("<![CDATA[").join("").split("]]" + ">").join(""), true);
      return;
    }
    this.pushText(name, raw, this.textPreserve);
  }

  pushText(name, body, preserve) {
    const value = body.indexOf("&") >= 0 ? decodeEntities(body) : body;
    if (name === "t") {
      if (this.cell) this.cell.str += decodeOoxmlEscapes(preserve ? value : textValue(value, false));
      return;
    }
    const cell = this.cell;
    if (!cell) return;
    if (name === "v") {
      cell.vText = (cell.vText === undefined ? "" : cell.vText) + value;
    } else {
      // Formula bodies are trimmed: no writer pads <f>, and Excel strips it.
      cell.fText = ((cell.fText === undefined ? "" : cell.fText) + value).trim();
    }
  }

  openTag(local, attrs, selfClose) {
    switch (local) {
      case "dimension":
        if (attrs.ref) this.dims = attrs.ref;
        break;
      case "col":
        this.columns.push(columnInfo(attrs));
        break;
      case "f":
        if (this.cell) {
          this.cell.fType = attrs.t;
          this.cell.si = attrs.si;
          this.cell.fRef = attrs.ref;
        }
        break;
      case "row":
        this.openRow(attrs, selfClose);
        break;
      case "c":
        this.openCell(attrs, selfClose);
        break;
      case "mergeCell":
        if (attrs.ref) {
          // A merge ref is a range. Passing a malformed one straight through put
          // "NOTAREF" on sheet.merges, where a consumer would hand it to a range
          // library and fail far from the file that caused it.
          if (isRangeRef(attrs.ref)) this.merges.push(attrs.ref);
          else this.warnOnce("BAD_MERGE", "ignoring a mergeCell whose ref is not a range: " + JSON.stringify(attrs.ref), { ref: attrs.ref });
        }
        break;
      case "pane":
        this.readPane(attrs);
        break;
      case "autoFilter":
        if (attrs.ref && this.autoFilter === undefined) this.autoFilter = attrs.ref;
        break;
      case "hyperlink":
        this.hyperlinks.push({
          ref: attrs.ref, rid: attrs["r:id"], location: attrs.location,
          tooltip: attrs.tooltip, display: attrs.display,
        });
        break;
      case "tabColor":
        if (this.tabColor === undefined) this.tabColor = tabColorOf(attrs);
        break;
      case "legacyDrawing":
        if (attrs["r:id"]) this.legacyRelId = attrs["r:id"];
        break;
      case "sheetData":
        this.inSheetData = true;
        this.sawSheetData = true;
        break;
      case "rPh":
      case "phoneticPr":
        this.skipText++;
        break;
      default:
        break;
    }
  }

  closeTag(local) {
    switch (local) {
      case "row":
        this.finishRow();
        break;
      case "c":
        this.finishCell();
        break;
      case "sheetData":
        this.inSheetData = false;
        break;
      case "rPh":
      case "phoneticPr":
        if (this.skipText > 0) this.skipText--;
        break;
      default:
        break;
    }
  }

  readPane(attrs) {
    if (this.freeze !== undefined) return;
    const state = attrs.state;
    if (state !== "frozen" && state !== "frozenSplit") return;
    let row = intOf(attrs.ySplit) || 0;
    let col = intOf(attrs.xSplit) || 0;
    if (row === 0 && col === 0 && attrs.topLeftCell) {
      try {
        const a = decodeCell(attrs.topLeftCell);
        row = a.r;
        col = a.c;
      } catch { /* malformed topLeftCell: leave the pane unfrozen */ }
    }
    if (row > 0 || col > 0) this.freeze = { row, col };
  }

  /* ------------------------------------------------------------------ rows */

  openRow(attrs, selfClose) {
    if (this.row) this.finishRow();
    const explicit = intOf(attrs.r);
    const rNum = explicit !== undefined && explicit > 0 ? explicit : this.lastRowNum + 1;
    this.lastRowNum = rNum;
    const row = { r: rNum, auto: explicit === undefined, cells: [], count: 0, c0: -1, c1: -1, nextCol: 0 };
    if (this.collectMeta) {
      const info = { index: rNum };
      let any = false;
      if (attrs.ht !== undefined) {
        const h = parseFloat(attrs.ht);
        if (Number.isFinite(h) && h > 0) { info.height = h; any = true; }
      }
      if (attrs.hidden !== undefined && boolOf(attrs.hidden)) { info.hidden = true; any = true; }
      const ol = intOf(attrs.outlineLevel);
      if (ol) { info.outlineLevel = ol; any = true; }
      if (attrs.s !== undefined) { info.style = intOf(attrs.s); any = true; }
      if (any) this.rowInfo.set(rNum, info);
    }
    this.row = row;
    if (selfClose) this.finishRow();
  }

  openCell(attrs, selfClose) {
    if (!this.row) {
      this.strayCells++;
      if (this.strayCells === 1) this.warnOnce("STRAY_CELL", "cell element outside of a row");
      return;
    }
    if (this.cell) this.finishCell();
    let col = -1;
    let rowRef = 0;
    if (attrs.r) {
      try {
        const a = decodeCell(attrs.r);
        col = a.c;
        rowRef = a.r + 1;
      } catch {
        this.warnOnce("BAD_CELL_REF", "unreadable cell reference " + JSON.stringify(attrs.r) + ", using document order");
      }
    }
    if (col < 0) col = this.row.nextCol;
    if (rowRef && this.row.auto) {
      this.row.r = rowRef;
      this.row.auto = false;
      this.lastRowNum = rowRef;
    }
    this.row.nextCol = col + 1;
    const t = attrs.t;
    if (t !== undefined && t !== "n" && t !== "s" && t !== "str" && t !== "inlineStr" &&
        t !== "b" && t !== "e" && t !== "d") {
      this.warnOnce("UNKNOWN_CELL_TYPE", "unknown cell type " + JSON.stringify(t));
    }
    this.cell = {
      col,
      t,
      s: attrs.s === undefined ? undefined : intOf(attrs.s),
      vText: undefined,
      str: "",
      fText: undefined,
      fType: undefined,
      si: undefined,
      fRef: undefined,
    };
    if (selfClose) this.finishCell();
  }

  finishRow() {
    const row = this.row;
    if (!row) return;
    if (this.cell) this.finishCell();
    this.row = null;
    const rNum = row.r;
    if (row.count === 0) return;
    if (this.bounds) {
      const r0 = rNum - 1;
      if (r0 < this.bounds.minR || r0 > this.bounds.maxR) return;
    }
    if (this.minRow === 0 || rNum < this.minRow) this.minRow = rNum;
    if (rNum > this.maxRowNum) this.maxRowNum = rNum;
    if (this.collectMeta && row.c0 >= 0) {
      if (this.minCol < 0 || row.c0 < this.minCol) this.minCol = row.c0;
      if (row.c1 > this.maxCol) this.maxCol = row.c1;
    }
    this.rowCount++;
    // sheetStubs is honoured on the stream too, so eachRow and parseSheetXml
    // see the same dense rows (gaps filled from the row's first column).
    if (this.sheetStubs) fillGaps(row.cells);
    if (this.onRow) this.onRow({ r: rNum, cells: row.cells, c0: row.c0, c1: row.c1 });
  }

  finishCell() {
    const c = this.cell;
    if (!c) return;
    this.cell = null;
    const row = this.row;
    if (!row) return;
    if (this.bounds && (c.col < this.bounds.minC || c.col > this.bounds.maxC)) return;
    const cell = this.buildCell(c, row.r);
    if (cell === null) return;
    row.cells[c.col] = cell;
    row.count++;
    if (row.c0 < 0 || c.col < row.c0) row.c0 = c.col;
    if (c.col > row.c1) row.c1 = c.col;
  }

  /* ----------------------------------------------------------- cell values */

  numFmtId(style) {
    const st = this.styles;
    if (!st || style === undefined) return undefined;
    return typeof st.numFmtId === "function" ? st.numFmtId(style) : undefined;
  }

  codeOf(style) {
    const st = this.styles;
    if (!st || style === undefined) return undefined;
    return typeof st.code === "function" ? st.code(style) : undefined;
  }

  /**
   * Resolved number-format code of a style index, memoized: a 200k-row sheet
   * with one style does one lookup, not 200k.
   */
  formatCode(style) {
    if (style === undefined) return undefined;
    const cached = this.codeCache.get(style);
    if (cached !== undefined) return cached === null ? undefined : cached;
    const st = this.styles;
    let code;
    if (st) {
      if (typeof st.formatOf === "function") code = st.formatOf(style);
      else if (typeof st.code === "function") code = st.code(style);
    }
    if (typeof code !== "string") code = null;
    this.codeCache.set(style, code);
    return code === null ? undefined : code;
  }

  isDateStyle(style) {
    const st = this.styles;
    if (!st || style === undefined) return false;
    return typeof st.isDate === "function" ? st.isDate(style) === true : false;
  }

  buildCell(c, rowNum) {
    const t = c.t === undefined ? "" : c.t;
    const style = c.s;
    const z = this.numFmtId(style);
    let out = null;

    switch (t) {
      case "s": {
        if (c.vText === undefined || c.vText === "") { out = blank(); break; }
        const idx = parseInt(c.vText, 10);
        let str = "";
        if (this.shared && Number.isFinite(idx) && idx >= 0 && idx < this.shared.length) {
          str = this.shared[idx];
        } else {
          this.warnOnce("BAD_SHARED_INDEX", "shared string index " + JSON.stringify(c.vText) + " is out of range");
        }
        out = { v: str, t: "string" };
        break;
      }
      case "inlineStr": {
        out = { v: c.str, t: "string" };
        break;
      }
      case "str": {
        if (c.vText === undefined) { out = blank(); break; }
        out = { v: decodeOoxmlEscapes(c.vText), t: "string" };
        break;
      }
      case "b": {
        if (c.vText === undefined) { out = blank(); break; }
        const raw = c.vText.trim().toLowerCase();
        out = { v: raw === "1" || raw === "true", t: "boolean" };
        break;
      }
      case "e": {
        if (c.vText === undefined) { out = blank(); break; }
        const code = c.vText.trim();
        out = { v: code, t: "error", e: code };
        break;
      }
      case "d": {
        const raw = c.vText === undefined ? "" : c.vText.trim();
        const date = raw === "" ? null : new Date(raw);
        if (date === null || Number.isNaN(date.getTime())) {
          if (raw !== "") this.warnOnce("BAD_DATE", "unreadable ISO date " + JSON.stringify(raw));
          out = raw === "" ? blank() : { v: raw, t: "string" };
          break;
        }
        const serial = dateToSerial(date, this.date1904);
        out = { v: this.cellDates ? date : serial, t: "date", serial };
        break;
      }
      case "":
      case "n": {
        const raw = c.vText === undefined ? "" : c.vText.trim();
        if (raw === "") { out = blank(); break; }
        const num = Number(raw);
        if (!Number.isFinite(num)) {
          this.warnOnce("BAD_CELL_VALUE", "unreadable numeric value " + JSON.stringify(raw));
          out = { v: raw, t: "string" };
          break;
        }
        if (this.isDateStyle(style)) {
          out = { v: this.cellDates ? serialToDate(num, this.date1904) : num, t: "date", serial: num };
        } else {
          out = { v: num, t: "number" };
        }
        break;
      }
      default: {
        this.warnOnce("UNKNOWN_CELL_TYPE", "unknown cell type " + JSON.stringify(t));
        const raw = c.vText === undefined ? "" : c.vText.trim();
        out = raw === "" ? blank() : { v: raw, t: "string" };
        break;
      }
    }

    // The formula is read even for a value-less cell: a formula with no cached
    // result (openpyxl writes <f>..</f><v></v>) is still real data. Everything
    // else that carries no value is only kept when the caller asked for stubs.
    const f = this.formulaOf(c, rowNum);
    const comment = this.commentAt === null ? undefined : this.commentAt.get((rowNum - 1) * MAX_COLUMNS + c.col);
    if (out.t === "empty" && f === undefined && comment === undefined && !this.sheetStubs) return null;
    if (style !== undefined) out.s = style;
    if (z !== undefined) out.z = z;
    // zf carries the resolved format code so a streamed cell (which has no
    // Book to look the id up in) can still be rendered by the caller.
    const fmt = this.formatCode(style);
    if (fmt !== undefined) out.zf = fmt;
    if (f !== undefined) out.f = f;
    if (comment !== undefined) out.c = comment;
    if (this.cellText && out.t !== "empty") out.w = this.renderText(out);
    return out;
  }

  renderText(cell) {
    const code = this.formatCode(cell.s) ?? this.codeOf(cell.s);
    const value = cell.t === "date" ? cell.serial : cell.v;
    return formatValue(value, code, { date1904: this.date1904 });
  }

  formulaOf(c, rowNum) {
    if (!this.cellFormula) return undefined;
    const type = c.fType;
    if (type === "dataTable") return undefined;
    if (type === "shared") {
      const si = c.si === undefined ? "0" : String(c.si);
      const text = c.fText;
      if (text !== undefined && text !== "") {
        this.formulaMasters.set(si, { text, row: rowNum, col: c.col });
        return stripEquals(text);
      }
      const master = this.formulaMasters.get(si);
      if (!master) {
        this.warnOnce("MISSING_SHARED_FORMULA", "shared formula " + si + " has no master cell");
        return undefined;
      }
      if (master.row === undefined || master.col === undefined) return master.text;
      return shiftFormulaRefs(master.text, rowNum - master.row, c.col - master.col);
    }
    if (c.fText === undefined || c.fText === "") return undefined;
    return stripEquals(c.fText);
  }
}

function blank() {
  return { v: null, t: "empty" };
}

function stripEquals(text) {
  return text.charCodeAt(0) === 61 ? text.slice(1) : text;
}

/** ColumnInfo for a <col> element, carrying only the declared attributes. */
function columnInfo(attrs) {
  const min = (intOf(attrs.min) || 1) - 1;
  const info = { index: min, min, max: (intOf(attrs.max) || intOf(attrs.min) || 1) - 1 };
  if (attrs.width !== undefined) {
    const w = parseFloat(attrs.width);
    if (Number.isFinite(w)) info.width = w;
  }
  if (attrs.hidden !== undefined) info.hidden = boolOf(attrs.hidden);
  const ol = intOf(attrs.outlineLevel);
  if (ol) info.outlineLevel = ol;
  if (attrs.style !== undefined) info.style = intOf(attrs.style);
  if (attrs.customWidth !== undefined) info.customWidth = boolOf(attrs.customWidth);
  return info;
}

function tabColorOf(attrs) {
  if (attrs.rgb) return attrs.rgb.length === 8 ? "#" + attrs.rgb.slice(2) : "#" + attrs.rgb;
  if (attrs.theme !== undefined) return "theme:" + attrs.theme;
  if (attrs.indexed !== undefined) return "indexed:" + attrs.indexed;
  return undefined;
}

/**
 * The local name of a tag body that may still carry attributes. The namespace
 * rule itself lives once, in xml.js's localName; this only trims the attributes
 * first, because a tag body is not yet a bare name.
 */
function localOf(name) {
  const space = name.indexOf(" ");
  return localName(space >= 0 ? name.slice(0, space) : name);
}

/**
 * True when a string is a usable A1 range reference ("A1", "A1:B2", "$A$1:$B$2").
 *
 * The reader's job is to reject nonsense, not to normalise a ref that is already
 * reasonable, so this is structural. Before it, a malformed <mergeCell ref> was
 * pushed onto sheet.merges verbatim - a consumer would hand "NOTAREF" to a range
 * library and fail far from the file that caused it.
 */
function isRangeRef(ref) {
  if (typeof ref !== "string" || ref.length === 0 || ref.length > 32) return false;
  const parts = ref.split(":");
  if (parts.length > 2) return false;
  return parts.every((part) => /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(part));
}

/**
 * The element name at the start of a tag body, from `from` up to the first
 * space or colon. One pass, and no substring at all when the name is already
 * bare - which is the common case for close tags.
 */
function nameOf(raw, from) {
  const n = raw.length;
  for (let i = from; i < n; i++) {
    const c = raw.charCodeAt(i);
    if (c === 32 || c === 58) return raw.slice(from, i);
  }
  return from === 0 ? raw : raw.slice(from);
}

/* ------------------------------------------------------------ formula shift */

const REF_RE = /(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})/g;

/**
 * Shift relative A1 references in a shared-formula body by a row/column
 * offset. String literals and quoted sheet names are copied verbatim; tokens
 * glued to a name or function (LOG10, MyA1) are left alone.
 */
export function shiftFormulaRefs(text, dRow, dCol) {
  if (!text || (dRow === 0 && dCol === 0)) return text;
  let out = "";
  let start = 0;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (ch === '"' || ch === "'") {
      const end = text.indexOf(ch, i + 1);
      const stop = end < 0 ? n : end + 1;
      out += shiftRun(text.slice(start, i), dRow, dCol) + text.slice(i, stop);
      i = stop;
      start = stop;
      continue;
    }
    i++;
  }
  out += shiftRun(text.slice(start), dRow, dCol);
  return out;
}

function shiftRun(seg, dRow, dCol) {
  if (!/[A-Za-z][0-9]/.test(seg) && !/\$/.test(seg)) return seg;
  return seg.replace(REF_RE, (m, absCol, letters, absRow, digits, off, whole) => {
    const before = off > 0 ? whole[off - 1] : "";
    if (before && /[A-Za-z0-9_.$]/.test(before)) return m;
    const after = whole[off + m.length];
    if (after && /[A-Za-z0-9_(]/.test(after)) return m;
    let col;
    try {
      col = decodeCol(letters);
    } catch {
      return m;
    }
    const row = parseInt(digits, 10);
    const col2 = absCol ? col : col + dCol;
    const row2 = absRow ? row : row + dRow;
    if (col2 < 0 || col2 > MAX_COLUMNS - 1) return "#REF!";
    if (row2 < 1 || row2 > MAX_ROWS) return "#REF!";
    return absCol + encodeCol(col2) + absRow + row2;
  });
}

/* ------------------------------------------------------------- chunk source */

async function* textChunks(chunks) {
  if (typeof chunks === "string") {
    yield chunks;
    return;
  }
  if (chunks instanceof Uint8Array) {
    yield* byteChunks(chunks);
    return;
  }
  const dec = new TextDecoder("utf-8");
  const iterable = chunks && (chunks[Symbol.asyncIterator] || chunks[Symbol.iterator]);
  if (!iterable) {
    throw new GridxcelError("worksheet chunks must be bytes, a string or an iterable", "BAD_INPUT");
  }
  for await (const part of chunks) {
    if (typeof part === "string") yield part;
    else if (part instanceof Uint8Array) yield dec.decode(part, { stream: true });
    else if (part instanceof ArrayBuffer) yield dec.decode(new Uint8Array(part), { stream: true });
    else throw new GridxcelError("unsupported worksheet chunk type", "BAD_INPUT");
  }
  const tail = dec.decode();
  if (tail.length) yield tail;
}

function* byteChunks(bytes) {
  const dec = new TextDecoder("utf-8");
  for (let i = 0; i < bytes.length; i += SLICE_BYTES) {
    yield dec.decode(bytes.subarray(i, Math.min(i + SLICE_BYTES, bytes.length)), { stream: true });
  }
  const tail = dec.decode();
  if (tail.length) yield tail;
}

/* ------------------------------------------------------------------ results */

/** Sheet-level metadata accrued while scanning, in absolute coordinates. */
function metaOf(sc) {
  return {
    dims: sc.dims,
    columns: sc.columns,
    rowInfo: sc.rowInfo,
    merges: sc.merges,
    hyperlinks: sc.hyperlinks,
    freeze: sc.freeze,
    autoFilter: sc.autoFilter,
    tabColor: sc.tabColor,
    legacyRelId: sc.legacyRelId,
    minRow: sc.minRow,
    maxRowNum: sc.maxRowNum,
    minCol: sc.minCol,
    maxCol: sc.maxCol,
    rowCount: sc.rowCount,
    sawSheetData: sc.sawSheetData,
  };
}

function rebaseCells(cells, originCol) {
  const out = [];
  for (let c = originCol; c < cells.length; c++) {
    const cell = cells[c];
    if (cell !== undefined) out[c - originCol] = cell;
  }
  return out;
}

function fillGaps(cells) {
  const len = cells.length;
  for (let c = 0; c < len; c++) {
    if (cells[c] === undefined) cells[c] = { v: null, t: "empty" };
  }
  return cells;
}

function attachHyperlinks(sheet, meta, opts, warn) {
  if (!meta.hyperlinks.length) return;
  const resolve = typeof opts.relTarget === "function" ? opts.relTarget : null;
  for (const link of meta.hyperlinks) {
    if (!link.ref) continue;
    let target;
    let external = false;
    if (link.rid && resolve) {
      const hit = resolve(link.rid);
      if (hit !== undefined && hit !== null) {
        target = typeof hit === "string" ? hit : hit.target;
        external = typeof hit === "object" && hit.mode === "External";
      }
    }
    if (target === undefined && link.location) target = "#" + link.location;
    else if (target !== undefined && link.location && !external) target = target + "#" + link.location;
    if (target === undefined) {
      warn("BAD_HYPERLINK", "hyperlink on " + link.ref + " has no resolvable target");
      continue;
    }
    const cell = ensureCell(sheet, link.ref, 2);
    if (!cell) continue;
    cell.l = link.tooltip === undefined ? { target } : { target, tooltip: link.tooltip };
  }
}

function attachComments(sheet, opts, warn) {
  const comments = opts.comments;
  if (!comments || typeof comments.forEach !== "function") return;
  comments.forEach((text, ref) => {
    const cell = ensureCell(sheet, ref, 1);
    if (cell) cell.c = text;
    else warn("BAD_COMMENT", "comment on " + ref + " is outside the parsed range");
  });
}

/**
 * Locate (or create) the cell at an absolute A1 reference.
 * mode 0 = existing cell only, 1 = inside an existing row, 2 = create the row.
 */
function ensureCell(sheet, ref, mode) {
  let addr;
  try {
    addr = decodeCell(ref);
  } catch {
    return null;
  }
  const ri = addr.r - (sheet.originRow - 1);
  const ci = addr.c - sheet.originCol;
  if (ri < 0 || ci < 0) return null;
  let row = sheet.rows[ri];
  if (row === undefined) {
    if (mode < 2) return null;
    row = [];
    sheet.rows[ri] = row;
  }
  let cell = row[ci];
  if (cell === undefined) {
    if (mode < 1) return null;
    cell = { v: null, t: "empty" };
    row[ci] = cell;
  }
  return cell;
}

/** True when rows already arrive in ascending row order, which is the norm. */
function isAscending(rows) {
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].r < rows[i - 1].r) return false;
  }
  return true;
}

function buildSheet(meta, rows, opts) {
  const warn = typeof opts.warn === "function" ? opts.warn : noop;
  const originRow = meta.minRow === 0 ? 1 : meta.minRow;
  const originCol = meta.minCol < 0 ? 0 : meta.minCol;
  const rowCount = meta.maxRowNum;
  const columnCount = meta.maxCol < 0 ? 0 : meta.maxCol + 1;
  const rebased = [];
  let lastIdx = -1;
  // Rows whose index sits far beyond the others are kept in a sparse side table
  // instead of extending the array. Assigning to a far index silently sets
  // Array#length to that index, which is free for a handful of rows but makes
  // rows.length meaningless and gives every consumer a range it may densify.
  // The declared extent stays on sheet.rowCount either way.
  let farRows = null;
  // Rows are placed by absolute index, so document order cannot change where a
  // cell ends up - but it DID change whether a large gap was treated as hostile.
  // The same two cells at rows 1 and 1048576 produced rows.length 1 with a
  // SPARSE_ROW_GAP warning when they appeared ascending, and rows.length
  // 1,048,576 with no warning when they appeared descending. Sorting first makes
  // the gap check a property of the data rather than of the file's byte order.
  const ordered = isAscending(rows) ? rows : [...rows].sort((a, b) => a.r - b.r);
  for (const row of ordered) {
    const idx = row.r - originRow;
    if (idx < 0) continue;
    let cells = originCol > 0 ? rebaseCells(row.cells, originCol) : row.cells;
    if (opts.sheetStubs) cells = fillGaps(cells);
    if (lastIdx >= 0 && idx - lastIdx > EMPTY_ROW_SLACK) {
      if (farRows === null) farRows = new Map();
      farRows.set(idx, cells);
      continue;
    }
    rebased[idx] = cells;
    if (idx > lastIdx) lastIdx = idx;
  }
  // Anything that landed beyond the array is re-evaluated now that the
  // contiguous extent is known: a row just past the last one is ordinary.
  if (farRows !== null) {
    for (const [idx, cells] of farRows) {
      if (idx - lastIdx <= EMPTY_ROW_SLACK) {
        rebased[idx] = cells;
        lastIdx = idx;
      }
    }
  }
  // Keep the declared extent when it is within slack of the rows present, so a
  // sheet that legitimately declares more rows than it writes keeps the
  // rows.length === rowCount invariant. Beyond slack the declared extent is
  // untrusted and only the rows that exist are sized for.
  const declaredLen = rowCount >= originRow ? rowCount - originRow + 1 : 0;
  const presentLen = lastIdx + 1;
  rebased.length = declaredLen > presentLen && declaredLen - presentLen <= EMPTY_ROW_SLACK ? declaredLen : presentLen;

  const sheet = {
    name: opts.name === undefined ? "" : opts.name,
    rows: rebased,
    originRow,
    originCol,
    columnCount,
    rowCount,
    merges: meta.merges,
    columns: meta.columns,
    rowInfo: meta.rowInfo,
  };
  if (farRows !== null && farRows.size > 0) {
    // Rows kept out of the dense array, keyed by 0-based offset from originRow.
    for (const [idx, cells] of farRows) {
      if (idx >= rebased.length) { sheet.sparseRows = farRows; break; }
    }
    if (sheet.sparseRows) {
      warn(
        "SPARSE_ROW_GAP",
        "sheet " + JSON.stringify(sheet.name) + " has " + farRows.size +
          " row(s) more than " + EMPTY_ROW_SLACK + " rows beyond the others; they are kept in sheet.sparseRows",
      );
    }
  }
  if (!meta.sawSheetData) {
    warn("NO_SHEET_DATA", "sheet " + JSON.stringify(sheet.name) + " has no sheetData element");
  }
  if (meta.freeze !== undefined) sheet.freeze = meta.freeze;
  if (meta.autoFilter !== undefined) sheet.autoFilter = meta.autoFilter;
  if (meta.tabColor !== undefined) sheet.tabColor = meta.tabColor;
  if (opts.hidden === true) sheet.hidden = true;
  if (meta.dims !== undefined && !opts.bounds) sheet.dims = meta.dims;
  else if (meta.maxCol >= 0) {
    const b = makeBounds();
    extendBounds(b, meta.minRow - 1, meta.minCol);
    extendBounds(b, meta.maxRowNum - 1, meta.maxCol);
    sheet.dims = boundsToRef(b);
  } else sheet.dims = "A1";

  attachHyperlinks(sheet, meta, opts, warn);
  attachComments(sheet, opts, warn);
  return sheet;
}

/* -------------------------------------------------------------------- API */

/**
 * Parse a whole worksheet part.
 * @param {Uint8Array|string|AsyncIterable<Uint8Array|string>} chunks
 * @param {object} opts sharedStrings, styles, date1904, bounds, cellDates,
 *   cellText, cellFormula, sheetStubs, comments, relTarget, name, hidden, warn
 * @returns {Promise<object>} Sheet
 */
export async function parseSheetXml(chunks, opts = {}) {
  if (typeof chunks === "string" || chunks instanceof Uint8Array) {
    return parseSheetXmlSync(chunks, opts);
  }
  const rows = [];
  const stream = createRowStream(chunks, opts);
  for await (const row of stream) rows.push(row);
  const meta = await stream.meta;
  return buildSheet(meta, rows, opts);
}

/** Synchronous twin of parseSheetXml, for byte or string input. */
export function parseSheetXmlSync(chunks, opts = {}) {
  const rows = [];
  const sc = new RowScanner(opts);
  sc.onRow = (row) => rows.push(row);
  if (typeof chunks === "string") sc.write(chunks);
  else if (chunks instanceof Uint8Array) {
    for (const text of byteChunks(chunks)) sc.write(text);
  } else {
    throw new GridxcelError("parseSheetXmlSync needs bytes or a string", "BAD_INPUT");
  }
  sc.end();
  return buildSheet(metaOf(sc), rows, opts);
}

/**
 * Lazily iterate the rows of one worksheet part. Yields { r, cells, c0, c1 }
 * with cells indexed by absolute column (A == 0); memory stays at one chunk
 * plus the rows that chunk produced. The returned iterable also carries a
 * "meta" promise that resolves once the part has been consumed.
 */
export function createRowStream(chunks, opts = {}) {
  const sc = new RowScanner(opts);
  let resolveMeta;
  let resolved = false;
  const meta = new Promise((resolve) => { resolveMeta = resolve; });
  const it = (async function* () {
    let pending = [];
    sc.onRow = (row) => pending.push(row);
    try {
      for await (const text of textChunks(chunks)) {
        sc.write(text);
        if (pending.length) {
          const batch = pending;
          pending = [];
          for (const row of batch) yield row;
        }
      }
      sc.end();
      if (pending.length) {
        const batch = pending;
        pending = [];
        for (const row of batch) yield row;
      }
    } finally {
      if (!resolved) { resolved = true; resolveMeta(metaOf(sc)); }
    }
  })();
  it.meta = meta;
  return it;
}

/* ------------------------------------------------------- shared string table */

/**
 * Stream xl/sharedStrings.xml into a plain string array. Rich text runs are
 * concatenated in document order; phonetic rPh hints are ignored.
 */
export async function parseSharedStrings(chunks, warn) {
  const collector = sstCollector(warn);
  for await (const text of textChunks(chunks)) collector.write(text);
  collector.finish();
  return collector.out;
}

/** Synchronous twin of parseSharedStrings, for byte or string input. */
export function parseSharedStringsSync(bytes, warn) {
  const collector = sstCollector(warn);
  if (typeof bytes === "string") collector.write(bytes);
  else if (bytes instanceof Uint8Array) {
    for (const text of byteChunks(bytes)) collector.write(text);
  } else {
    throw new GridxcelError("parseSharedStringsSync needs bytes or a string", "BAD_INPUT");
  }
  collector.finish();
  return collector.out;
}

function sstCollector(warn) {
  const out = [];
  let cur = null;
  let skip = 0;
  const report = typeof warn === "function" ? warn : noop;
  const sc = new RawScanner({
    onOpen(local) {
      if (local === "si") cur = "";
      else if (PHONETIC.has(local)) skip++;
    },
    onClose(local) {
      if (PHONETIC.has(local)) { if (skip > 0) skip--; return; }
      if (local === "si") { out.push(cur === null ? "" : cur); cur = null; }
    },
    onText(local, body, preserve) {
      if (local !== "t" || skip > 0 || cur === null) return;
      const value = body.indexOf("&") >= 0 ? decodeEntities(body) : body;
      cur += decodeOoxmlEscapes(preserve ? value : textValue(value, false));
    },
  });
  return {
    out,
    write: (t) => sc.write(t),
    finish: () => {
      const truncated = sc.textEl !== null || sc.buf.length > 0 || cur !== null;
      sc.end();
      if (truncated) {
        report("BAD_SHARED_STRINGS", "sharedStrings.xml ends in the middle of an entry");
      }
    },
  };
}

/**
 * The chunked tag scanner again, reduced to open/close/text callbacks. Kept
 * separate so the worksheet hot path stays branch-light.
 */
class RawScanner {
  constructor(handlers) {
    this.h = handlers;
    this.buf = "";
    this.textEl = null;
    this.preserve = false;
    /** Body pieces of a tag whose '>' has not arrived yet, or null. */
    this.tagParts = null;
    this.tagQuote = 0;
    /** Raw text pieces of a text element still waiting for its close tag. */
    this.textParts = [];
    this.textOverlap = "";
  }

  write(str) {
    // Same discipline as RowScanner: an unfinished construct is buffered in
    // pieces with its scan state, never re-concatenated and rescanned. This
    // scanner reads sharedStrings.xml, where one very long string would
    // otherwise hit the same quadratic the worksheet scanner had.
    let buf;
    if (this.textEl !== null) {
      buf = this.textOverlap.length ? this.textOverlap + str : str;
      this.textOverlap = "";
    } else if (this.tagParts !== null) {
      buf = str;
    } else {
      buf = this.buf.length ? this.buf + str : str;
    }
    const n = buf.length;
    let i = 0;
    for (;;) {
      if (this.textEl !== null) {
        const k = buf.indexOf("</" + this.textEl, i);
        if (k < 0) {
          const keep = this.textEl.length + 3;
          const cut = Math.max(i, n - keep);
          if (cut > i) this.textParts.push(buf.slice(i, cut));
          this.textOverlap = buf.slice(cut);
          i = n;
          break;
        }
        const text = this.textParts.length === 0
          ? buf.slice(i, k)
          : this.textParts.join("") + buf.slice(i, k);
        if (this.textParts.length !== 0) this.textParts.length = 0;
        this.h.onText(this.textEl, text, this.preserve);
        this.textEl = null;
        i = k;
      }
      if (this.tagParts !== null) {
        const r = scanTagEnd(buf, 0, this.tagQuote);
        if (!r.done) {
          this.tagQuote = r.quote;
          if (n > 0) this.tagParts.push(buf);
          i = n;
          break;
        }
        const raw = this.tagParts.join("") + buf.slice(0, r.gt);
        this.tagParts = null;
        i = r.gt + 1;
        this.emit(raw);
        continue;
      }
      const lt = buf.indexOf("<", i);
      if (lt < 0) { i = n; break; }
      const r = scanTagEnd(buf, lt + 1, 0);
      if (!r.done) {
        this.tagParts = [buf.slice(lt + 1)];
        this.tagQuote = r.quote;
        i = n;
        break;
      }
      const raw = buf.slice(lt + 1, r.gt);
      i = r.gt + 1;
      this.emit(raw);
    }
    this.buf = i >= n ? "" : buf.slice(i);
  }

  /** One complete tag body, from '<' to '>'. */
  emit(raw) {
    if (raw.charCodeAt(0) === 47) {
      this.h.onClose(localOf(raw.slice(1)));
      return;
    }
    const selfClose = raw.charCodeAt(raw.length - 1) === 47;
    const tag = parseTag(selfClose ? raw.slice(0, raw.length - 1) : raw);
    if (selfClose) {
      this.h.onOpen(tag.local);
      this.h.onClose(tag.local);
      return;
    }
    this.h.onOpen(tag.local);
    if (TEXT_ELEMENTS.has(tag.local)) {
      this.textEl = tag.local;
      this.preserve = tag.attrs["xml:space"] === "preserve";
    }
  }

  end() {
    this.buf = "";
    this.textParts = [];
    this.textOverlap = "";
    this.tagParts = null;
    this.textEl = null;
  }
}
