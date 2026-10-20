/**
 * gridxcel/xml - the streaming SAX parser and the XML entity codec.
 *
 * This is the layer the worksheet and shared-string readers are built on. It is
 * exported because writing an OOXML part by hand is occasionally the right
 * answer, and because the parser is useful on its own for any XML that fits in
 * the subset it supports: elements, attributes, text, CDATA, comments, and the
 * five predefined entities plus numeric character references.
 *
 * It is not a general XML processor. It does not resolve external entities,
 * process DTDs, or validate namespaces.
 */

/**
 * Attribute map for one element.
 *
 * For an element with no attributes this is a single shared, frozen object
 * rather than a fresh empty map - the parser sees millions of attribute-less
 * tags and allocating one object each dominated the parse. Treat every
 * attribute map as read-only.
 */
export type XmlAttributes = Readonly<Record<string, string>>;

/** An element that has been opened and not yet closed. */
export interface XmlNode {
  /** The name including any namespace prefix, e.g. `x14:sparklineGroup`. */
  name: string;
  /** The name with the prefix removed, e.g. `sparklineGroup`. */
  local: string;
  attrs: XmlAttributes;
}

/** A parsed tag body: the name, the local name, and the attributes. */
export interface XmlTag {
  name: string;
  local: string;
  attrs: XmlAttributes;
}

/** Callbacks driving the parser. All are optional. */
export interface XmlHandlers {
  /**
   * An element opens.
   * @param parent the enclosing element, or undefined at the document root
   */
  onOpen?(name: string, attrs: XmlAttributes, parent: XmlNode | undefined): void;
  /**
   * An element closes, or a self-closing element is complete.
   *
   * For a self-closing element the third argument is the enclosing element,
   * because the element itself was never pushed; otherwise it is the element
   * being closed. It is null only for a self-closing element at the root.
   */
  onClose?(name: string, attrs: XmlAttributes, node: XmlNode | null): void;
  /**
   * Character data.
   * @param node the element the text belongs to
   */
  onText?(text: string, node: XmlNode): void;
}

export interface XmlParserOptions extends XmlHandlers {
  /**
   * Decode the `_xHHHH_` escapes OOXML uses to smuggle control characters
   * through XML. Default true. Set false to see the raw text.
   */
  decodeEscapes?: boolean;
}

/**
 * An incremental, non-recursive XML parser.
 *
 * Feed it chunks with {@link XmlParser.write} and finish with
 * {@link XmlParser.end}. Nesting is tracked on an explicit stack, so deeply
 * nested input cannot overflow the call stack.
 *
 * Custom entities are never expanded, so a billion-laughs document resolves the
 * entity name as literal text rather than expanding it.
 */
export declare class XmlParser {
  constructor(opts?: XmlParserOptions);
  onOpen: ((name: string, attrs: XmlAttributes, parent: XmlNode | undefined) => void) | null;
  onClose: ((name: string, attrs: XmlAttributes, node: XmlNode | null) => void) | null;
  onText: ((text: string, node: XmlNode) => void) | null;
  /** Number of currently open elements. */
  readonly depth: number;
  /** The innermost open element, or null at the document root. */
  readonly current: XmlNode | null;
  /** Feed a chunk. Text may be split across chunks at any point. */
  write(chunk: string): void;
  /** Signal end of input and flush any buffered text. */
  end(chunk?: string): void;
}

/**
 * Parse a tag body (the text between `<` and `>`) into name and attributes.
 *
 * Exported because callers writing their own scanner need the same attribute
 * parsing, including the entity decoding inside quoted values. Quote-aware: a
 * `>` inside a quoted attribute value does not end the tag.
 *
 * Attribute values decode both entities and the OOXML \`_xHHHH_\` escape, which is
 * how a control character in a font name or a format code survives a round trip.
 * \`_x005F_x0041_\` decodes to the literal \`_x0041_\`, not to "A".
 *
 * @param decodeEscapes pass false to leave \`_xHHHH_\` sequences as written
 */
export declare function parseTag(inner: string, decodeEscapes?: boolean): XmlTag;

/**
 * Find the `>` that ends a tag, resumably.
 *
 * A tag whose attribute value is longer than one chunk cannot be finished until
 * the rest arrives. A scanner that restarts the tag on the next chunk is
 * quadratic in the length of that value, so this reports where it stopped and
 * the quote state it stopped in and the caller resumes from there.
 *
 * @param s buffer to scan
 * @param from index to start at, normally just past the `<`
 * @param quote quote state to begin in: 0, 34 for `"`, or 39 for `'`
 * @returns `gt` is the index of the terminating `>`, or -1 when the tag is
 *   unfinished; `scanned` is the offset to resume from; `quote` is the state
 *   to resume in.
 */
export declare function scanTagEnd(s: string, from: number, quote: number): {
  gt: number;
  quote: number;
  scanned: number;
  done: boolean;
};

/** Strip a namespace prefix: `x14:foo` becomes `foo`. */
export declare function localName(name: string): string;

/**
 * Decode XML entities. Recognises the five predefined entities and numeric
 * character references (`&#65;`, `&#x41;`). An unrecognised entity is left
 * as written rather than expanding to nothing or throwing.
 */
export declare function decodeEntities(s: string): string;

/**
 * Decode OOXML `_xHHHH_` escapes, the mechanism Excel uses to store control
 * characters that XML 1.0 cannot represent literally.
 */
export declare function decodeOoxmlEscapes(s: string): string;

/** Escape text content for XML, including the `_xHHHH_` form for control characters. */
export declare function encodeXmlText(s: string): string;

/** Escape an attribute value, including quotes and control characters. */
export declare function encodeXmlAttr(s: string): string;

/**
 * Drive a parser over an async iterable of string or byte chunks.
 * @returns the parser, for inspecting {@link XmlParser.depth} afterwards
 */
export declare function parseChunks(
  chunks: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>,
  handlers: XmlHandlers,
): Promise<XmlParser>;
