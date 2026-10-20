/**
 * A small, fast, allocation-conscious XML reader for OOXML parts.
 *
 * Why hand-rolled: worksheet parts reach hundreds of megabytes and often
 * arrive inflating straight off a socket, so we need (a) incremental feeding,
 * (b) an event API that never materialises a DOM, and (c) predictable
 * behaviour on the parts Excel writers actually emit, including namespace
 * prefixes, self-closing tags, CDATA, numeric character references and the
 * _xHHHH_ escapes used inside shared strings.
 */

const ENTITY_RE = /&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g;
const OOXML_ESCAPE_RE = /_x([0-9A-Fa-f]{4})_/g;
const ENTITIES = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: "\u00a0",
  cr: "\r", tab: "\t",
};

/** Decode the five XML entities plus numeric character references. */
export function decodeEntities(s) {
  if (s.indexOf("&") < 0) return s;
  return s.replace(ENTITY_RE, (m, body) => {
    if (body.charCodeAt(0) === 35) {
      const isHex = body.charCodeAt(1) === 120 || body.charCodeAt(1) === 88;
      const cp = isHex ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : m;
    }
    const hit = ENTITIES[body];
    return hit === undefined ? m : hit;
  });
}

/**
 * Decode the OOXML _xHHHH_ escape sequences used to smuggle control
 * characters through XML. Excel writes these for tabs/newlines inside text.
 *
 * A *literal* "_x0041_" in the data is written as "_x005F_x0041_": the leading
 * underscore is escaped like any other underscore that starts an escape, and
 * the remaining "x0041_" is ordinary text. One left-to-right pass handles both
 * cases correctly: it turns "_x005F_" into "_" and never rescans what it
 * produced, so the following "x0041_" survives as text and the literal comes
 * back intact. (The old special case returned the "_x005F_" match unchanged,
 * which left the sequence encoded and did not consume the escape it protects.)
 *
 * Deliberately single-pass: decoding a second time would turn the literal
 * "_x0041_" into "A", which is exactly the corruption this rule exists to
 * prevent.
 */
export function decodeOoxmlEscapes(s) {
  if (s.indexOf("_x") < 0) return s;
  return s.replace(OOXML_ESCAPE_RE, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
}

/**
 * True when a literal "_xHHHH_" starts at s[i].
 *
 * Such a sequence is data, not an escape: the caller's text happens to look
 * like one. OOXML requires its leading underscore to be escaped (RFC-style
 * "escape the escape"), otherwise every reader - Excel included - decodes the
 * user's text into a control character and the value is corrupted.
 */
function isLiteralEscapeAt(s, i) {
  if (s.charCodeAt(i) !== 95 || s.charCodeAt(i + 1) !== 120) return false; // "_x"
  for (let k = i + 2; k < i + 6; k++) {
    const c = s.charCodeAt(k);
    if (!((c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102))) return false;
  }
  return s.charCodeAt(i + 6) === 95; // trailing "_"
}

/** The OOXML escape for a literal sequence: "_x0041_" -> "_x005F_x0041_". */
const LITERAL_ESCAPE_PREFIX = "_x005F_";

export function encodeXmlText(s) {
  if (!/[&<>\x00-\x08\x0b\x0c\x0e-\x1f]|_x[0-9A-Fa-f]{4}_/.test(s)) return s;
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "&") out += "&amp;";
    else if (ch === "<") out += "&lt;";
    else if (ch === ">") out += "&gt;";
    else if (isLiteralEscapeAt(s, i)) {
      out += LITERAL_ESCAPE_PREFIX + s.slice(i + 1, i + 7);
      i += 6;
    } else {
      const code = s.charCodeAt(i);
      if (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) {
        out += "_x" + code.toString(16).toUpperCase().padStart(4, "0") + "_";
      } else out += ch;
    }
  }
  return out;
}

export function encodeXmlAttr(s) {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "&") out += "&amp;";
    else if (ch === "<") out += "&lt;";
    else if (ch === ">") out += "&gt;";
    else if (ch === '"') out += "&quot;";
    else if (ch === "\n") out += "&#10;";
    else if (ch === "\r") out += "&#13;";
    else if (ch === "\t") out += "&#9;";
    else if (isLiteralEscapeAt(s, i)) {
      out += LITERAL_ESCAPE_PREFIX + s.slice(i + 1, i + 7);
      i += 6;
    } else {
      const code = s.charCodeAt(i);
      if (code < 0x20) out += "_x" + code.toString(16).toUpperCase().padStart(4, "0") + "_";
      else out += ch;
    }
  }
  return out;
}

/* ------------------------------------------------------------------ parser */

const isWs = (c) => c === 32 || c === 9 || c === 10 || c === 13;

/**
 * True when the construct described by `state` finishes inside `chunk`.
 *
 * Only the newest piece is examined. `prev` supplies the tail of the piece
 * before it, because a multi-character terminator may straddle the boundary.
 * state.quote is updated for a tag, so the piece after this one resumes in the
 * quote state this one ended in.
 */
function completesIn(chunk, tail, state) {
  switch (state.kind) {
    case "text":
      return chunk.indexOf("<") >= 0;
    case "tag": {
      const scan = scanTagEnd(chunk, 0, state.quote);
      state.quote = scan.quote;
      return scan.gt >= 0;
    }
    case "comment":
      return hasTerminator(chunk, tail, "-->");
    case "cdata":
      return hasTerminator(chunk, tail, "]]>");
    case "pi":
      return hasTerminator(chunk, tail, "?>");
    default: // decl and close are terminated by a single ">"
      return chunk.indexOf(">") >= 0;
  }
}

/** The last BOUNDARY characters of tail+chunk, without copying a whole piece. */
function tailOf(tail, chunk) {
  if (chunk.length >= BOUNDARY) return chunk.slice(chunk.length - BOUNDARY);
  const joined = tail + chunk;
  return joined.length > BOUNDARY ? joined.slice(joined.length - BOUNDARY) : joined;
}

/** True when `term` occurs in `chunk`, including a match whose first
 * character is in `tail`, at the boundary before this piece. */
function hasTerminator(chunk, tail, term) {
  if (chunk.indexOf(term) >= 0) return true;
  if (tail.length === 0) return false;
  return (tail + chunk.slice(0, term.length - 1)).indexOf(term) >= 0;
}

/**
 * Shared attribute map for tags that have none. Frozen so an accidental write
 * fails loudly instead of corrupting every subsequent no-attribute tag.
 */
/**
 * The shared empty attribute map.
 *
 * Frozen and null-prototype so a consumer cannot accidentally add a key to it or
 * inherit one from Object.prototype. Exported because the worksheet scanner takes
 * its own fast path for attribute-free tags and must hand the same object to
 * openTag, rather than allocating a second empty map per tag.
 */
export const NO_ATTRS = Object.freeze(Object.create(null));

/**
 * While the un-consumed construct is this short the parser joins its pieces and
 * re-classifies it from the start, which keeps the "<!-" -> "<!--" and
 * "<![CDA" -> "<![CDATA[" prefixes honest. Above the window the kind is settled
 * and only the newest piece is scanned, which is what makes a construct longer
 * than a chunk linear instead of quadratic.
 */
const JOIN_WINDOW = 16;

/** Characters of the scanned prefix a terminator may reach back into. */
const BOUNDARY = 2;

export class XmlParser {
  constructor(opts = {}) {
    this.onOpen = opts.onOpen || null;
    this.onClose = opts.onClose || null;
    this.onText = opts.onText || null;
    this.decodeEscapes = opts.decodeEscapes !== false;
    // Un-consumed input in the pieces it arrived in, and the construct (or text
    // run) still in flight. The pieces are deliberately NOT joined while a
    // construct is unfinished. Appending to one growing string is quadratic even
    // when the scan itself is resumable: each append makes the buffer a new cons
    // string, and the next read has to flatten the whole of it, so the construct
    // is copied once per chunk (measured with a resumable scan: a 2 MB tag took
    // 43 ms and an 8 MB tag 656 ms, 15x for 4x the input). The pieces are joined
    // once, when the construct completes. See _parse.
    this._parts = [];
    this._length = 0;
    // The last two characters scanned before the newest piece. A multi-character
    // terminator can span more than two pieces (with 1-byte chunks "-->" spans
    // three), so the window has to outlive a single piece boundary.
    this._tail = "";
    this._pending = null;
    // Whether a piece has arrived since the last _parse. A piece is scanned
    // exactly once: re-scanning the newest piece would apply the quote state to
    // characters that were already counted.
    this._fresh = false;
    this._stack = [];
    this._closed = false;
  }

  get depth() { return this._stack.length; }
  get current() { return this._stack.length ? this._stack[this._stack.length - 1] : null; }

  /** Feed a chunk of the document. */
  write(chunk) {
    if (chunk.length === 0) return;
    this._parts.push(chunk);
    this._length += chunk.length;
    this._fresh = true;
    this._parse();
  }

  /** Signal end of document and flush. */
  end(chunk) {
    if (chunk) {
      this._parts.push(chunk);
      this._length += chunk.length;
      this._fresh = true;
    }
    this._parse();
    this._closed = true;
  }

  /**
   * Parse everything that is complete.
   *
   * A construct in flight is resumed from the newest piece alone: everything
   * before it has been scanned already, so reading it again would repeat work.
   * The exception is a construct still inside JOIN_WINDOW, where joining is
   * cheap and the construct can still be re-classified from its start.
   *
   * Every construct is resumable: the text run (the search for the next "<"),
   * a tag (through scanTagEnd, the one tag scanner), and the "-->", "]]>", "?>"
   * and ">" terminators of comments, CDATA, processing instructions,
   * declarations and closing tags. Restarting a construct from its beginning on
   * every chunk is quadratic in its length - the same defect class as HIGH-1 in
   * sheet-reader.js, and reachable here through parseChunks(), which accepts
   * any async iterable.
   *
   * An unfinished construct is parked from its OWN START, never from the end of
   * the buffer; parking from the end would discard it instead of keeping it.
   *
   * A piece is scanned once. The newest piece is only examined when it has just
   * arrived (see _fresh): on a call that brings no new input there is nothing
   * new to scan, and the construct is simply re-read from its start, once.
   */
  _parse() {
    const fresh = this._fresh;
    this._fresh = false;
    if (fresh && this._pending !== null && this._length > JOIN_WINDOW) {
      const last = this._parts[this._parts.length - 1];
      if (!completesIn(last, this._tail, this._pending)) {
        this._tail = tailOf(this._tail, last);
        return; // still in flight
      }
    }

    const buf = this._parts.length === 1 ? this._parts[0] : this._parts.join("");
    this._parts = [];
    this._length = 0;
    this._pending = null;
    const n = buf.length;
    let i = 0;

    while (i < n) {
      const lt = buf.indexOf("<", i);
      if (lt < 0) {
        // A text run whose "<" has not arrived. Text is emitted only once the
        // "<" is in hand, so a run still arrives as one onText call however it
        // was chunked, and the run is never searched twice.
        this._park(buf, i, { kind: "text" });
        return;
      }
      if (lt > i && this.onText && this._stack.length) this._emitText(buf.slice(i, lt));
      i = lt;
      const next = buf.charCodeAt(i + 1);

      if (next === 33) { // comment, CDATA or declaration
        if (buf.startsWith("<!--", i)) {
          const close = buf.indexOf("-->", i + 4);
          if (close < 0) { this._park(buf, i, { kind: "comment" }); return; }
          i = close + 3;
        } else if (buf.startsWith("<![CDATA[", i)) {
          const close = buf.indexOf("]]>", i + 9);
          if (close < 0) { this._park(buf, i, { kind: "cdata" }); return; }
          if (this.onText && this._stack.length) this.onText(buf.slice(i + 9, close), this._stack[this._stack.length - 1]);
          i = close + 3;
        } else {
          const close = buf.indexOf(">", i + 2);
          if (close < 0) { this._park(buf, i, { kind: "decl" }); return; }
          i = close + 1;
        }
      } else if (next === 63) { // processing instruction
        const close = buf.indexOf("?>", i + 2);
        if (close < 0) { this._park(buf, i, { kind: "pi" }); return; }
        i = close + 2;
      } else if (next === 47) { // closing tag
        const close = buf.indexOf(">", i + 2);
        if (close < 0) { this._park(buf, i, { kind: "close" }); return; }
        this._close(buf.slice(i + 2, close).trim());
        i = close + 1;
      } else {
        // Opening tag. Scan for the terminating ">" while tracking quote state:
        // an attribute value may legally contain ">", and a ">" seen inside an
        // unterminated value means the tag continues in the next chunk.
        const scan = findTagEnd(buf, i);
        if (scan.gt < 0) { this._park(buf, i, { kind: "tag", quote: scan.quote }); return; }
        const selfClose = buf.charCodeAt(scan.gt - 1) === 47;
        this._open(parseTag(buf.slice(i + 1, selfClose ? scan.gt - 1 : scan.gt), this.decodeEscapes), selfClose);
        i = scan.gt + 1;
      }
    }
  }

  /**
   * Keep an unfinished construct for the next chunk, from its own start.
   * The pieces stay in _parts; the state says what the next piece must be
   * checked for. Nothing is copied here.
   */
  _park(buf, i, state) {
    const rest = buf.slice(i);
    this._parts = [rest];
    this._length = rest.length;
    this._tail = tailOf("", rest);
    this._pending = state;
  }

  _emitText(raw) {
    if (raw.length === 0) return;
    let t = decodeEntities(raw);
    if (this.decodeEscapes) t = decodeOoxmlEscapes(t);
    this.onText(t, this._stack[this._stack.length - 1]);
  }

  _open(tag, selfClose) {
    if (this.onOpen) this.onOpen(tag.name, tag.attrs, this._stack[this._stack.length - 1]);
    if (!selfClose) {
      const node = { local: tag.local, name: tag.name, attrs: tag.attrs };
      this._stack.push(node);
    } else if (this.onClose) {
      this.onClose(tag.name, tag.attrs, this._stack.length ? this._stack[this._stack.length - 1] : null);
    }
  }

  _close(name) {
    const node = this._stack.pop();
    if (node && this.onClose) this.onClose(name, node.attrs, node);
  }
}

/**
 * Scan for the ">" that ends the tag whose "<" is at `start`, resumably.
 *
 * The parser's one tag call site: scanTagEnd is the state machine, this names
 * the operation and supplies the offset the scan starts from (just past the
 * "<"). Kept as a function rather than inlined so there is a single place that
 * knows how a tag scan is entered, and so the reader (which imports
 * scanTagEnd directly) and the parser cannot drift.
 *
 * @param {string} buf buffer to scan
 * @param {number} start index of the "<" that opens the tag
 * @returns {{gt: number, quote: number, scanned: number, done: boolean}}
 */
function findTagEnd(buf, start) {
  return scanTagEnd(buf, start + 1, 0);
}

/**
 * Scan for the '>' that ends a tag, resumably.
 *
 * A tag whose attribute value is longer than one chunk cannot be finished until
 * the rest arrives. A scanner that throws away what it learned and restarts the
 * tag on the next chunk is quadratic in the length of that value: measured at
 * 16 MB, 11.9 s from a 17.7 kB upload, ~3.9x per doubling.
 *
 * So the scan reports where it stopped and the quote state it stopped in, and
 * the caller resumes from there with the next chunk. This is the single tag
 * scanner in the library; the worksheet reader imports it rather than keeping a
 * second copy, which is how one of the two quadratic paths survived a fix
 * applied to the other.
 *
 * @param {string} s buffer to scan
 * @param {number} from index to start at, normally just past the '<'
 * @param {number} quote quote state to begin in: 0, 34 for '"', or 39 for "'"
 * @returns {{gt: number, quote: number, scanned: number, done: boolean}}
 *   gt is the index of the terminating '>', or -1 when the tag is unfinished.
 */
export function scanTagEnd(s, from, quote) {
  const n = s.length;
  for (let i = from; i < n; i++) {
    const c = s.charCodeAt(i);
    if (quote !== 0) {
      if (c === quote) quote = 0;
      continue;
    }
    if (c === 34) { quote = 34; continue; }
    if (c === 39) { quote = 39; continue; }
    if (c === 62) return { gt: i, quote: 0, scanned: i + 1, done: true };
  }
  return { gt: -1, quote, scanned: n, done: false };
}

/**
 * Split a tag body into name + attributes without regex backtracking.
 *
 * The no-attribute case is worth a dedicated branch: a worksheet is mostly
 * <v>, <t>, <is> and </c>, so most tags carry nothing to parse, and allocating
 * an attributes object plus a result object for each of them dominated the
 * parse. When the body holds no whitespace there is no attribute syntax to
 * find, and the attribute map is shared and frozen.
 *
 * Attribute values are decoded exactly like text nodes - entities and then the
 * OOXML _xHHHH_ escapes - because OOXML uses the same escaping in both places
 * (MINOR-4). `decodeEscapes` false returns the values with entities decoded
 * but the _xHHHH_ escapes left as written, mirroring XmlParser's option of the
 * same name.
 *
 * @param {string} inner the text between "<" and ">" (without a self-closing
 *   slash, which callers strip)
 * @param {boolean} [decodeEscapes] decode _xHHHH_ escapes. Default true.
 */
export function parseTag(inner, decodeEscapes = true) {
  let i = 0;
  const n = inner.length;
  while (i < n && !isWs(inner.charCodeAt(i))) i++;
  const name = inner.slice(0, i);
  if (i === n) {
    return { name, local: localName(name), attrs: NO_ATTRS };
  }
  const attrs = Object.create(null);
  const decodeValue = decodeEscapes ? decodeAttrValue : decodeEntities;
  while (i < n) {
    while (i < n && isWs(inner.charCodeAt(i))) i++;
    if (i >= n) break;
    const keyStart = i;
    while (i < n && inner.charCodeAt(i) !== 61 && !isWs(inner.charCodeAt(i))) i++;
    const key = inner.slice(keyStart, i);
    while (i < n && isWs(inner.charCodeAt(i))) i++;
    if (inner.charCodeAt(i) !== 61) {
      if (key) attrs[key] = "";
      continue;
    }
    i++;
    while (i < n && isWs(inner.charCodeAt(i))) i++;
    const q = inner.charCodeAt(i);
    if (q === 34 || q === 39) {
      const quote = String.fromCharCode(q);
      const close = inner.indexOf(quote, i + 1);
      attrs[key] = decodeValue(close < 0 ? inner.slice(i + 1) : inner.slice(i + 1, close));
      i = close < 0 ? n : close + 1;
    } else {
      const start = i;
      while (i < n && !isWs(inner.charCodeAt(i))) i++;
      attrs[key] = decodeValue(inner.slice(start, i));
    }
  }
  return { name, local: localName(name), attrs };
}

/**
 * Decode an attribute value the way a text node is decoded: entities first,
 * then the _xHHHH_ escapes. Both are legal in attribute values - a font name
 * or number-format code carrying a control character is written as _x0001_ and
 * has to come back as the control character, or the value does not survive its
 * own round trip.
 */
function decodeAttrValue(s) {
  return decodeOoxmlEscapes(decodeEntities(s));
}

/**
 * Local part of a possibly prefixed element or attribute name.
 * Single implementation on purpose: parseTag used to inline this twice, which
 * is how "x:c" and "c" drifted apart in the first place.
 */
export function localName(name) {
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

/**
 * Local name of an attribute-free tag, in a single pass; null when it has attributes.
 *
 * A worksheet is mostly `<v>`, `<t>`, `<is>` and `</c>` - 79% of the tags in a
 * 50,000-row sheet carry no attributes - and the scanner only needs the name for
 * those. parseTag handles that case without walking attributes, but it still
 * allocates the result object for every one of two million tags.
 *
 * The caller must be able to fall back cheaply, so this returns null the moment
 * it sees whitespace, which is the only thing that can introduce an attribute.
 * That early exit matters: an earlier version asked one predicate whether the
 * body had whitespace and a second function for the name, which scanned the body
 * three times per tag and measured slower than not doing this at all.
 *
 * Unprefixed names return the input string itself, so the common case allocates
 * nothing.
 *
 * @param {string} body the text between "<" and ">" (no self-closing slash)
 * @returns {string | null} the local name, or null when the tag has attributes
 */
export function tagLocalIfNoAttrs(body) {
  const n = body.length;
  let colon = -1;
  for (let i = 0; i < n; i++) {
    const c = body.charCodeAt(i);
    if (c === 32 || c === 9 || c === 10 || c === 13) return null;
    if (c === 58 && colon < 0) colon = i;
  }
  return colon < 0 ? body : body.slice(colon + 1);
}

/** Drive a parser over an async iterable of string|Uint8Array chunks. */
export async function parseChunks(chunks, handlers) {
  const p = new XmlParser(handlers);
  const dec = new TextDecoder("utf-8");
  for await (const c of chunks) p.write(typeof c === "string" ? c : dec.decode(c, { stream: true }));
  p.end();
  return p;
}
