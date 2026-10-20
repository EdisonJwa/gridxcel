/**
 * gridxcel — public data contracts.
 *
 * THIS FILE IS THE FROZEN INTERFACE between all modules. Do not change a
 * signature without updating docs/ARCHITECTURE.md and telling the Lead.
 */

/* ------------------------------------------------------------------ input */

/** A cell value accepted on write. */
export type CellValue = string | number | boolean | Date | null | undefined;

/**
 * Rich cell: any object carrying a `v` property.
 *  { v, t?, z?, s?, f?, r?, h?, c? }
 *   v = value, t = type, z = number format, s = style id,
 *   f = formula (without leading '='), r = rich text runs, h = HTML, c = comment
 */
export type CellObject = {
  v?: CellValue;
  t?: CellType | "n" | "s" | "b" | "d" | "z" | "e";
  z?: string | number;
  s?: number | StyleObject;
  f?: string;
  r?: RichTextRun[];
  h?: string;
  c?: string | CommentObject;
};

export type RichTextRun = { v: string; b?: boolean; i?: boolean; u?: boolean; sz?: number; color?: string };
export type CommentObject = { t: string; a?: string };

export type RowInput = CellValue[] | Record<string, CellValue | CellObject> | { [k: string]: any };
export type RowObject = Record<string, CellValue | CellObject>;

/* ------------------------------------------------------------------- cell */

export type CellType = "string" | "number" | "boolean" | "date" | "error" | "empty";

export interface Cell {
  /** Value. Date cells hold a Date when {@link ReadOptions.cellDates} is true, else a serial number. */
  v: string | number | boolean | Date | null;
  t: CellType;
  /** Raw A1 address, e.g. "B7". */
  a?: string;
  /** Formula text without the leading '=' (read) or as given (write). */
  f?: string;
  /** Number-format id for the cell's style. */
  z?: number;
  /** Style table index. */
  s?: number;
  /** Present only for error cells, e.g. "#DIV/0!". */
  e?: string;
  /** Present only when t === "date" and cellDates is true. */
  serial?: number;
  /** Pre-formatted display text, when available. */
  w?: string;
  /** Cell comment text. */
  c?: string;
  /** Resolved number-format code, e.g. "0.00%". Added by the reader on both
   * batch and streaming reads, so a streamed cell can be rendered without a
   * workbook to resolve {@link Cell.z} against. */
  zf?: string;
}

export type CellInput = CellValue | CellObject;

/* ------------------------------------------------------------------ style */

export interface StyleObject {
  numFmt?: string | number;
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  protection?: { locked?: boolean; hidden?: boolean };
}

export interface FontStyle {
  name?: string; sz?: number; b?: boolean; i?: boolean; u?: boolean;
  strike?: boolean; color?: ColorValue; vertAlign?: "superscript" | "subscript" | "baseline";
}

export interface FillStyle {
  pattern?:
    | "none" | "solid" | "mediumGray" | "darkGray" | "lightGray" | "darkHorizontal"
    | "darkVertical" | "darkDown" | "darkUp" | "darkGrid" | "darkTrellis"
    | "lightHorizontal" | "lightVertical" | "lightDown" | "lightUp" | "lightGrid"
    | "lightTrellis" | "gray125" | "gray0625";
  fgColor?: ColorValue;
  bgColor?: ColorValue;
}

export interface BorderStyle {
  top?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge; right?: BorderEdge;
  diagonal?: BorderEdge; diagonalUp?: boolean; diagonalDown?: boolean;
}
export interface BorderEdge { style?: BorderLineStyle; color?: ColorValue }
export type BorderLineStyle =
  | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair"
  | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";

export interface AlignmentStyle {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean; shrinkToFit?: boolean; indent?: number; readingOrder?: number; textRotation?: number;
}

/** ARGB hex ("FFRRGGBB"), 6-digit RGB hex ("RRGGBB"), or theme/indexed reference. */
export type ColorValue = string | { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean };

/* ----------------------------------------------------------------- column */

export interface ColumnInfo {
  /** Zero-based column index. */
  index: number;
  /** 0-based inclusive/exclusive column range this entry applies to. */
  min: number;
  max: number;
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: number;
}

export interface RowInfo {
  /** 1-based row number. */
  index: number;
  height?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: number;
}

/* ------------------------------------------------------------------ sheet */

export interface SheetData {
  name: string;
  rows: ReadonlyArray<ReadonlyArray<Cell>>;
  /** 1-based row number of rows[0]. Defaults to 1. */
  originRow: number;
  /** 0-based column of cells[0] in every row. Defaults to 0. */
  originCol: number;
  /** Max column index present + 1. */
  columnCount: number;
  /** Max row number present. */
  rowCount: number;
  merges: string[];
  columns: ColumnInfo[];
  rowInfo: Map<number, RowInfo>;
  freeze?: { row: number; col: number };
  autoFilter?: string;
  dims?: string;
}

export interface SheetMeta {
  name: string;
  hidden?: boolean;
  /** Sheet-level visibility of the tab. */
  tabColor?: string;
  rowCount?: number;
  columnCount?: number;
}

/* -------------------------------------------------------------- workbook */

export interface Workbook {
  sheets: SheetData[];
  /** 1900 (default) or 1904. */
  dateSystem: 1900 | 1904;
  /** Raw style table as written in styles.xml (read) — index -> format code. */
  numFmts: Map<number, string>;
  /** Parsed per-xf style records, index -> style (read). */
  styles: StyleObject[];
  props: WorkbookProps;
  definedNames: DefinedName[];
  /** Raw zip parts, for escape hatches / round-trip preservation. */
  parts?: Map<string, Uint8Array>;
}

export interface WorkbookProps {
  title?: string; subject?: string; creator?: string; lastModifiedBy?: string;
  keywords?: string; description?: string; category?: string;
  created?: Date; modified?: Date; application?: string; company?: string;
}

export interface DefinedName { name: string; ref: string; sheet?: number; hidden?: boolean; comment?: string }

/* --------------------------------------------------------------- options */

export interface ReadOptions {
  /** Parse worksheet XML — set false for metadata-only fast listing. Default true. */
  sheets?: boolean | number[] | string[];
  /** Empty cells are omitted (default) rather than present as {t:"empty"}. */
  sheetStubs?: boolean;
  /** Return Date objects for date-formatted cells. Default false (serial numbers). */
  cellDates?: boolean;
  /** Populate Cell.w with the number-format-rendered text. Default false. */
  cellText?: boolean;
  /** Populate Cell.f. Default true. */
  cellFormula?: boolean;
  /** Read cell comments into Cell.c. Default false. */
  cellComments?: boolean;
  /** Use raw column letters for object keys (A, B, ...) instead of headers. */
  raw?: boolean;
  /** Range to read, A1:B10 style, or {s:{r,c}, e:{r,c}}. */
  /** A range ref like "A1:C10", a single row number, or explicit corners. */
  range?: string | number | { s: { r: number; c: number }; e: { r: number; c: number } };
  /** Raise on the first malformed part instead of collecting into warnings. Default false. */
  strict?: boolean;
  /** Called for each non-fatal problem. */
  onWarning?: (w: Warning) => void;
  /** Leave the zip bytes available on Workbook.parts. Default false. */
  keepParts?: boolean;
}

export interface Warning {
  code: string;
  message: string;
  part?: string;
  /** 1-based sheet row, when the warning is cell-scoped. */
  row?: number;
  col?: number;
}

export interface WriteOptions {
  /** "xlsx" (default) or "csv" or "json". */
  bookType?: "xlsx" | "csv" | "json";
  /** Overwrite an existing file with fs.writeFile. Default false. */
  overwrite?: boolean;
  /** Compression level 0..9, or 0 to store uncompressed. Default 6. */
  compressionLevel?: number;
  /** Write a <calcPr fullCalcOnLoad="1"/> so Excel recalculates. Default true. */
  recalc?: boolean;
  /** SheetJS-style aliases, accepted for drop-in compatibility. */
  type?: string;
}

export type SheetInput = string | ReadonlyArray<ReadonlyArray<CellInput>> | { name: string; rows: ReadonlyArray<ReadonlyArray<CellInput>> } | Record<string, CellInput>;

export interface AOA2Options extends WriteOptions {
  header?: string[];
  dateNF?: string;
  skipHidden?: boolean;
}

export interface Sheet2JSONOptions extends ReadOptions {
  header?: string[] | number;
  // range is inherited, not re-declared: widening a property in an extending
  // interface is a type error, and ReadOptions.range now already allows a row
  // number, which is what toRange() has always accepted at runtime.
  defval?: unknown;
  blankrows?: boolean;
  rawNumbers?: boolean;
}

/* -------------------------------------------------------------- addressing */

export interface CellAddress { r: number; c: number; a?: string }
export interface Range { s: CellAddress; e: CellAddress }
