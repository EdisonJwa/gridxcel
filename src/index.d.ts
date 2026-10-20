/**
 * gridxcel public type declarations.
 *
 * This file describes src/index.js exactly. It is hand written so that the
 * types stay readable and so cell narrowing actually works:
 *
 *   const cell = sheet.rows[0][0];
 *   if (cell?.t === "date") cell.v satisfies Date | number;
 */

/**
 * The streaming subset the library relies on.
 *
 * Declared structurally rather than as \`NodeJS.ReadableStream\` so the published
 * types do not require @types/node. A browser consumer, or a Node consumer who
 * has not installed the Node types, can use every API without a stub - and
 * Node's \`stream.Readable\` satisfies this interface as it is.
 */
export interface ByteStream extends AsyncIterable<Uint8Array> {
  pause?(): unknown;
  resume?(): unknown;
  destroy?(error?: Error): unknown;
  on?(event: string, listener: (...args: unknown[]) => void): unknown;
}

export type CellType = "string" | "number" | "boolean" | "date" | "error" | "empty";

export interface Cell {
  /** Value. For t === "date" this is a Date when cellDates was set, else a serial number. */
  v: string | number | boolean | Date | null;
  t: CellType;
  /** Raw A1 address when the cell was read from a file. */
  a?: string;
  /** Formula text, without the leading "=". */
  f?: string;
  /** Number-format id of the cell style. */
  z?: number;
  /** Index into the workbook style table. */
  s?: number;
  /** Error literal for t === "error", e.g. "#DIV/0!". */
  e?: string;
  /** Present when t === "date": the raw Excel serial. */
  serial?: number;
  /** Rendered display text when cellText was requested. */
  w?: string;
  /** Comment text when cellComments was requested. */
  c?: string;
  /** Hyperlink target and tooltip, when the cell has one. */
  l?: { target: string; tooltip?: string };
  /** Resolved number-format code, e.g. "0.00%". Added by the reader on both
   * batch and streaming reads, so a streamed cell can be rendered without a
   * workbook to resolve {@link Cell.z} against. */
  zf?: string;
}

export type CellValue = string | number | boolean | Date | null | undefined;

export interface RichTextRun { v: string; b?: boolean; i?: boolean; u?: boolean; sz?: number; color?: string }

export interface CellObject {
  v?: CellValue;
  t?: CellType | "n" | "s" | "b" | "d" | "z" | "e" | "str";
  z?: string | number;
  s?: number | StyleObject;
  f?: string;
  r?: RichTextRun[];
  l?: { target: string; tooltip?: string };
  c?: string | { t: string; a?: string };
}

export type CellInput = CellValue | CellObject;
export type RowInput = ReadonlyArray<CellInput> | Record<string, CellInput>;
export type AOA = ReadonlyArray<RowInput>;

export interface ColorValue { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean }
export type ColorSpec = string | ColorValue;

export interface FontStyle {
  name?: string; sz?: number; b?: boolean; i?: boolean; u?: boolean; strike?: boolean;
  color?: ColorSpec; vertAlign?: "superscript" | "subscript" | "baseline";
}

export interface FillStyle {
  pattern?: "none" | "solid" | "mediumGray" | "darkGray" | "lightGray" | "darkHorizontal" | "darkVertical"
    | "darkDown" | "darkUp" | "darkGrid" | "darkTrellis" | "lightHorizontal" | "lightVertical" | "lightDown"
    | "lightUp" | "lightGrid" | "lightTrellis" | "gray125" | "gray0625";
  fgColor?: ColorSpec;
  bgColor?: ColorSpec;
}

export type BorderLineStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair"
  | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
export interface BorderEdge { style?: BorderLineStyle; color?: ColorSpec }
export interface BorderStyle { top?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge; right?: BorderEdge; diagonal?: BorderEdge; diagonalUp?: boolean; diagonalDown?: boolean }

export interface AlignmentStyle {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean; shrinkToFit?: boolean; indent?: number; readingOrder?: number; textRotation?: number;
}

export interface StyleObject {
  numFmt?: string | number;
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  protection?: { locked?: boolean; hidden?: boolean };
}

export interface ColumnInfo { index: number; min: number; max: number; width?: number; hidden?: boolean; outlineLevel?: number; style?: number }
export interface RowInfo { index: number; height?: number; hidden?: boolean; outlineLevel?: number; style?: number }

export interface Sheet {
  name: string;
  /** Sparse rows; row i is sheet row originRow + i, cells are indexed from originCol. */
  rows: Cell[][];
  originRow: number;
  originCol: number;
  columnCount: number;
  rowCount: number;
  merges: string[];
  columns: ColumnInfo[];
  rowInfo: Map<number, RowInfo>;
  freeze?: { row: number; col: number };
  autoFilter?: string;
  dims?: string;
}

export interface WorkbookProps {
  title?: string; subject?: string; creator?: string; lastModifiedBy?: string; keywords?: string;
  description?: string; category?: string; created?: Date; modified?: Date; application?: string; company?: string;
}

export interface DefinedName { name: string; ref: string; sheet?: number; hidden?: boolean; comment?: string }

export interface Workbook {
  sheets: Sheet[];
  /** Sheet names as declared by the workbook, available even with sheets:false. */
  sheetNames?: string[];
  /** Non-fatal problems found while reading. Always an array, possibly empty. */
  warnings?: Warning[];
  dateSystem: 1900 | 1904;
  numFmts: Map<number, string>;
  styles: StyleObject[];
  props: WorkbookProps;
  definedNames: DefinedName[];
  parts?: Map<string, Uint8Array>;
}

export interface Warning { code: string; message: string; part?: string; row?: number; col?: number }

export interface ReadOptions {
  sheets?: boolean | Array<string | number>;
  sheetStubs?: boolean;
  cellDates?: boolean;
  cellText?: boolean;
  cellFormula?: boolean;
  cellComments?: boolean;
  raw?: boolean;
  range?: string | { s: { r: number; c: number }; e: { r: number; c: number } };
  strict?: boolean;
  onWarning?: (w: Warning) => void;
  keepParts?: boolean;
}

export interface WriteOptions {
  bookType?: "xlsx" | "csv" | "json";
  path?: string;
  overwrite?: boolean;
  compressionLevel?: number;
  recalc?: boolean;
  type?: "buffer" | "array" | "base64" | "binary" | "string";
  sheet?: string;
  dateNF?: string;
  indent?: number;
  deterministic?: boolean;
}

export interface SheetToJSONOptions extends ReadOptions {
  header?: string[] | number;
  defval?: unknown;
  blankrows?: boolean;
  rawNumbers?: boolean;
}

export interface CSVOptions {
  delimiter?: string;
  quote?: string;
  newline?: string;
  header?: boolean;
  bom?: boolean;
  alwaysQuote?: boolean;
  escapeFormulae?: boolean;
  dateNF?: string;
  skipHidden?: boolean;
  range?: string;
  raw?: boolean;
}

export class GridxcelError extends Error { code: string; details?: unknown }
export class BadZipError extends GridxcelError {}
/** Container-level failure. Subclass of BadZipError, so instanceof GridxcelError holds. */
export class ZipError extends BadZipError { offset?: number }
export class FormatError extends GridxcelError {}
export class SchemaError extends GridxcelError {}
export class NotImplementedError extends GridxcelError {}
export class StrictWarningError extends GridxcelError { warning: Warning }
/** Thrown when a host capability is missing: outside Node, by the synchronous API and the path helpers. Code `NOT_SUPPORTED`. */
export class PlatformError extends GridxcelError {}

/** Read a workbook from bytes, a path, a stream, or an async iterable. */
export function read(input: Uint8Array | ArrayBuffer | string | ReadableStream | AsyncIterable<Uint8Array>, options?: ReadOptions): Promise<Workbook>;
/** Read synchronously from bytes (Node only). */
export function readSync(bytes: Uint8Array | ArrayBuffer, options?: ReadOptions): Workbook;
/** Read only the sheet names, without parsing cells. */
export function sheetNames(input: Uint8Array | ArrayBuffer | string): Promise<string[]>;
export function readFile(path: string, options?: ReadOptions): Promise<Workbook>;
/** Write a workbook. With options.path it writes to disk and resolves undefined. */
export function write(book: Workbook | AOA | object, options?: WriteOptions): Promise<Uint8Array | string | undefined>;
/**
 * Write a workbook to a file path.
 *
 * Resolves to the encoded bytes for `xlsx` and `xlsm`, and to undefined for
 * `csv` and `json`, which write text. Mirrors {@link write}, which returns the
 * same value instead of writing when no `path` is given.
 */
export function writeFile(book: Workbook | AOA | object, path: string, options?: WriteOptions): Promise<Uint8Array | undefined>;
export function writeRows(rows: AOA, options?: WriteOptions): Promise<Uint8Array | string | undefined>;
export function sheetNames(input: Uint8Array | string): Promise<string[]>;
export function readJSON<T = Record<string, unknown>>(input: Uint8Array | string, options?: SheetToJSONOptions): Promise<T[]>;

export interface RowEvent { sheet: string; r: number; cells: Cell[] }
/** Stream rows with bounded memory. */
export function eachRow(input: Uint8Array | string | ReadableStream | AsyncIterable<Uint8Array>, options?: ReadOptions): AsyncIterable<RowEvent>;

export function normalizeSheet(input: AOA | Sheet, name?: string): Sheet;
export function normalizeBook(input: Workbook | AOA | object, options?: WriteOptions): Workbook;
export function normalizeCell(input: CellInput): Cell | undefined;
export function aoaToSheet(rows: AOA, options?: { sheet?: string }): Sheet;
export function sheetToAOA(sheet: Sheet, options?: ReadOptions & { blankrows?: boolean; defval?: unknown; rawNumbers?: boolean }): Promise<unknown[][]>;
export function sheetToJSON<T = Record<string, unknown>>(sheet: Sheet, options?: SheetToJSONOptions): Promise<T[]>;
export function jsonToSheet(rows: ReadonlyArray<Record<string, unknown>>, options?: { header?: string[]; sheet?: string }): Promise<Sheet>;
export function sheetToCSV(sheet: Sheet, options?: CSVOptions): Promise<string>;
export function csvToSheet(text: string, options?: CSVOptions): Promise<Sheet>;

export const VERSION: string;

export const utils: {
  aoa_to_sheet(rows: AOA, opts?: { sheet?: string }): Sheet;
  json_to_sheet(rows: ReadonlyArray<Record<string, unknown>>, opts?: object): Promise<Sheet>;
  sheet_to_json<T = Record<string, unknown>>(sheet: Sheet, opts?: SheetToJSONOptions): Promise<T[]>;
  sheet_to_csv(sheet: Sheet, opts?: CSVOptions): Promise<string>;
  sheet_to_aoa(sheet: Sheet, opts?: object): Promise<unknown[][]>;
  book_new(): Workbook;
  sheet_add_aoa(sheet: Sheet, rows: AOA, opts?: { origin?: { r: number; c: number } }): Sheet;
  book_append_sheet(book: Workbook, sheet: Sheet | AOA, name?: string): Workbook;
  decode_range(ref: string): { s: { r: number; c: number }; e: { r: number; c: number } };
  encode_cell(c: { r: number; c: number }): string;
  encode_range(s: { r: number; c: number }, e: { r: number; c: number }): string;
  format_cell(cell: Cell, v?: unknown, opts?: object): string;
};

export function encodeCol(c: number): string;
export function decodeCol(letters: string): number;
export function encodeCell(r: number, c: number): string;
export function decodeCell(a1: string): { r: number; c: number };
export function encodeRange(s: { r: number; c: number }, e: { r: number; c: number }): string;
export function decodeRange(ref: string): { s: { r: number; c: number }; e: { r: number; c: number } };

export function isDateFormat(codeOrId: string | number | undefined): boolean;
export function formatNumber(value: number, code?: string | number, opts?: { date1904?: boolean }): string;
export function formatValue(value: unknown, code?: string | number, opts?: { date1904?: boolean }): string;
export function serialToDate(serial: number, date1904?: boolean): Date;
export function dateToSerial(date: Date, date1904?: boolean): number;

/* ---------------------------------------------------- streaming write */

export interface SheetWriter {
  /**
   * Append one row.
   *
   * @param rowIndex 1-based sheet row. Rows must be appended in ascending
   *   order; passing undefined auto-increments from the previous row.
   * @param cells cells indexed from the sheet's origin column, so cells[0]
   *   is column A unless `originCol` was set.
   * @throws GridxcelError with code BAD_ROW for a non-positive or non-integer
   *   index, and ROW_ORDER when a row arrives out of order.
   */
  addRow(rowIndex: number | undefined, cells: ReadonlyArray<CellInput | undefined>): void;
  /** Finish the sheet. */
  end(): void;
  readonly rowCount: number;
}

export interface WorkbookWriterOptions {
  compressionLevel?: number;
  deterministic?: boolean;
  inlineStrings?: boolean;
  date1904?: boolean;
  recalc?: boolean;
  application?: string;
  activeTab?: number;
  defaultRowHeight?: number;
  defaultColWidth?: number;
}

/**
 * Streaming workbook writer. Rows are appended as they are produced and the
 * parts are emitted on finish(), so a workbook larger than memory can be
 * written without materialising it.
 */
export class WorkbookWriter {
  constructor(opts?: WorkbookWriterOptions);
  addSheet(name: string, opts?: Record<string, unknown>): SheetWriter;
  finish(): Uint8Array;
  parts(): Map<string, Uint8Array>;
  stringParts(): Map<string, string>;
}

/** Low-level: build every OPC part for a workbook. */
export function buildParts(book: Workbook, opts?: Record<string, unknown>): Map<string, Uint8Array | string>;
/** Low-level: intern fonts, fills, borders and xf records, and emit styles.xml. */
export class StyleTable {
  constructor();
  /** Intern a style and return its cellXfs index. */
  xf(style: StyleObject): number;
  /** Intern a number format code and return its numFmtId. */
  numFmtId(code: string | number): number;
  toXml(): string;
}
/**
 * Parse a styles.xml part into numFmts, per-xf styles and a date predicate.
 *
 * This is the canonical parser for the part: the reader and the writer both use
 * it, which is why each style carries BOTH the raw id the writer re-emits and
 * the resolved code the reader renders and classifies with. \`numFmts\` includes
 * the ECMA-376 built-in formats, not only the ones the file declares.
 */
export function parseStyles(xml: string): {
  numFmts: Map<number, string>;
  styles: Array<StyleObject & {
    /** The raw numFmtId, which the writer re-emits unchanged. */
    numFmt: number;
    /** The resolved format code, or null when the id is neither declared nor built in. */
    numFmtCode: string | null;
  }>;
  isDate(index: number): boolean;
  fonts: unknown[];
  fills: unknown[];
  borders: unknown[];
  cellXfs: unknown[];
};

/** Write a workbook synchronously and return the bytes. */
export function writeWorkbookSync(book: Workbook | object, options?: WriteOptions): Uint8Array;
/** Alias of writeWorkbookSync for callers that want an explicit name. */
export function toBuffer(book: Workbook | object, options?: WriteOptions): Uint8Array;
/** Concatenate the sheets of several workbooks. Later sheets keep their names. */
export function mergeBooks(...books: Array<Workbook | object>): Workbook;
/** Rename a sheet, rejecting illegal names and duplicates at the call site. */
export function renameSheet(book: Workbook, from: string, to: string): Workbook;
/** Count populated cells, for stats and progress reporting. */
export function countCells(book: Workbook): number;

/** Resolve a read cell back to its number format code. */
export function cellFormatCode(book: Workbook, cell: Cell): string | number | undefined;

export function crc32(input: Uint8Array | Uint8Array[]): number;
export class Crc32 { update(buf: Uint8Array): this; digest(): number; reset(): this }
export class ZipArchive {
  constructor(input: Uint8Array);
  entries: Array<{ name: string; size: number; compressedSize: number; method: number; crc: number; bytes(): Uint8Array; text(): string; stream(): ByteStream }>;
  get(name: string): unknown;
  read(name: string): Uint8Array | undefined;
  readText(name: string): string | undefined;
  names(): string[];
}
export class ZipWriter {
  constructor(opts?: { compressionLevel?: number; deterministic?: boolean });
  add(name: string, data: Uint8Array | string, opts?: { compress?: boolean; level?: number; date?: Date }): this;
  build(): Uint8Array;
}


/* ------------------------------------------------- declared root exports */
// The runtime namespace is authoritative. The verifier compared it against these
// declarations and found fourteen exports that "gridxcel" provides and the types
// did not describe: eight were declared in a subpath but never re-exported from
// the root, and six were declared nowhere at all. `import { toRange } from
// "gridxcel"` was an error while `gridxcel.toRange` worked, and `toRange` and
// `splitSheetRef` are documented root exports in docs/API.md. Compiling is not
// the same as matching, which is why the check that found this compares the two.

/** A cell address: zero-based row and column, with the A1 spelling when known. */
export interface CellAddress { r: number; c: number; a?: string }
/** A rectangular range, corners included. */
export interface Range { s: CellAddress; e: CellAddress }

/** Split "Sheet1!A1:B2" (or a quoted sheet name) into [sheetName | null, rangeText]. */
export function splitSheetRef(ref: string): [string | null, string];

/** Normalise any accepted range spelling - "A1:B2", a row number, or a Range - into a Range. */
export function toRange(range?: string | number | Range | null): Range | undefined;

/** Async write. Returns the bytes, or the shaped value when options.type is set. */
export function writeWorkbook(
  book: Workbook | Sheet | ReadonlyArray<ReadonlyArray<unknown>>,
  options?: WriteOptions & { path?: string; type?: string },
): Promise<Uint8Array | string | undefined>;

/** Configure a WorkbookWriter from a read-model Workbook and push rows into it. */
export function createBookWriter(
  book: Workbook | { sheets: ReadonlyArray<Sheet> },
  opts?: WriteOptions,
): { addSheet(name?: string, opts?: Record<string, unknown>): { name: string; addRow(row: ReadonlyArray<unknown>): void }; sheets: Array<{ name: string }> };

/** Canonicalise a colour: "#RRGGBB" / "AARRGGBB" become an upper-case AARRGGBB string. */
export function normalizeColor(c: ColorValue | string | undefined): ColorValue | string | undefined;

/** Shortest decimal spelling that parses back to the identical double. */
export function numberLiteral(n: number): string;

export { XmlParser, parseTag, decodeEntities, encodeXmlText, encodeXmlAttr } from "./xml.js";
export { createZip } from "./zip.js";
export { BUILTIN_FORMATS, generalNumber } from "./numfmt.js";
export type { CommentObject } from "./types.js";

declare const _default: {
  VERSION: typeof VERSION;
  read: typeof read;
  readSync: typeof readSync;
  write: typeof write;
  writeFile: typeof writeFile;
  readFile: typeof readFile;
  eachRow: typeof eachRow;
  sheetNames: typeof sheetNames;
  readJSON: typeof readJSON;
  writeRows: typeof writeRows;
  aoaToSheet: typeof aoaToSheet;
  sheetToAOA: typeof sheetToAOA;
  sheetToJSON: typeof sheetToJSON;
  jsonToSheet: typeof jsonToSheet;
  sheetToCSV: typeof sheetToCSV;
  csvToSheet: typeof csvToSheet;
  utils: typeof utils;
  normalizeSheet: typeof normalizeSheet;
  normalizeBook: typeof normalizeBook;
  normalizeCell: typeof normalizeCell;
  mergeBooks: typeof mergeBooks;
  renameSheet: typeof renameSheet;
  countCells: typeof countCells;
  GridxcelError: typeof GridxcelError;
  BadZipError: typeof BadZipError;
  ZipError: typeof ZipError;
  FormatError: typeof FormatError;
  SchemaError: typeof SchemaError;
  NotImplementedError: typeof NotImplementedError;
  StrictWarningError: typeof StrictWarningError;
  PlatformError: typeof PlatformError;
};
export default _default;