/** gridxcel/json - array/JSON <-> Sheet conversion. */

import type { Cell, CellInput, Sheet } from "./index.d.ts";

export declare const DEFAULT_DATE_NF: string;

export interface SheetToJSONOptions {
  /**
   * Header handling, following SheetJS:
   *   absent      -> row 1 of the range is the header row, data follows
   *   1           -> array of arrays; EVERY row is data (dense, null for gaps)
   *   "A"         -> A/B/C keys; every row is data
   *   string[]    -> these names; every row is data
   *   n > 1       -> row n (1-based) is the header row
   */
  header?: number | string | string[];
  /** Value substituted for missing cells. */
  defval?: unknown;
  /** Emit rows that are entirely empty. Default false for objects, true for header:1. */
  blankrows?: boolean;
  /** Keep numbers raw instead of applying the number format. Default false. */
  rawNumbers?: boolean;
  /** Override the date format used when materialising date values. */
  dateNF?: string;
  /** 1904 date system. Default false. */
  date1904?: boolean;
  /** Return Date objects for date cells. Default false. */
  cellDates?: boolean;
  /** Placeholder cells for gaps. Default false. */
  sheetStubs?: boolean;
  /** Omit hidden rows and columns. Default false. */
  skipHidden?: boolean;
  /** Range to read, e.g. "A1:D20". */
  range?: string;
  /** Use A/B/C keys derived from the column index. */
  raw?: boolean;
  [key: string]: unknown;
}

export declare function sheetToJSON<T = Record<string, unknown>>(sheet: Sheet, opts?: SheetToJSONOptions): T[];
export declare function sheetToAOA(sheet: Sheet, opts?: SheetToJSONOptions): unknown[][];
export declare function jsonToSheet(rows: ReadonlyArray<Record<string, unknown>>, opts?: { header?: string[]; sheet?: string; name?: string; dateNF?: string }): Sheet;
export declare function aoaToSheet(rows: ReadonlyArray<ReadonlyArray<CellInput>>, opts?: { sheet?: string; name?: string; dateNF?: string }): Sheet;
export declare function normalizeSheet(input: unknown, name?: string): Sheet;
export declare function blankSheet(name?: string): Sheet;
export declare function cellAt(sheet: Sheet, r: number, c: number): Cell | undefined;
export declare function cellValue(cell: Cell | undefined): unknown;
export declare function cellText(cell: Cell | undefined, opts?: { dateNF?: string; date1904?: boolean }): string;
export declare function cellFromValue(value: unknown, opts?: { dateNF?: string; date1904?: boolean }): Cell | undefined;
export declare function dateCell(value: Date, opts?: { dateNF?: string }): Cell;
export declare function rowAt(sheet: Sheet, index: number): Cell[] | undefined;
export declare function sheetRange(sheet: Sheet): { s: { r: number; c: number }; e: { r: number; c: number } };
export declare function resolveRange(sheet: Sheet, range: unknown): { s: { r: number; c: number }; e: { r: number; c: number } };
export declare function isRowHidden(sheet: Sheet, row: number): boolean;
export declare function isColHidden(sheet: Sheet, col: number): boolean;
export declare function worksheetToSheet(input: unknown, name?: string): Sheet;
