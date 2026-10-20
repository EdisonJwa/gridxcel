/**
 * gridxcel/compat - the SheetJS-shaped surface, for migrating an existing
 * codebase by changing one import specifier.
 *
 *   import { readFile, write, utils } from "gridxcel/compat";
 *
 * The reader and writer are imported statically, so the synchronous entry points
 * work immediately with no \`await\` in front of them, and there is no top-level
 * await - the module graph stays synchronous and the entry also builds as an
 * IIFE (\`<script>\`). Importing this entry always pulls in the reader and writer.
 *
 * ESM only, like the rest of the package; require() of this entry is not
 * supported.
 */

import type { Cell, CellObject, Sheet, StyleObject, Warning, Workbook } from "./index.d.ts";

/** SheetJS-shaped workbook view: SheetNames + Sheets, sharing the Book's storage. */
export interface SheetJSBook extends Workbook {
  SheetNames: string[];
  Sheets: Record<string, Sheet>;
}

export interface CompatReadOptions {
  cellDates?: boolean;
  cellText?: boolean;
  cellFormula?: boolean;
  cellNF?: boolean;
  cellStyles?: boolean;
  sheetStubs?: boolean;
  sheets?: boolean | number | string | Array<string | number>;
  /**
   * Accepted for SheetJS compatibility and ignored on purpose.
   *
   * In SheetJS this chooses between a keyed worksheet and an array-backed one.
   * A gridxcel Sheet is always array-backed (\`Sheet.rows\` is an array of rows),
   * which is the dense representation, so the option is effectively always true
   * and there is nothing to switch.
   */
  dense?: boolean;
  raw?: boolean;
  type?: "base64" | "binary" | "buffer" | "array" | "string";
  [key: string]: unknown;
}

export interface CompatWriteOptions {
  type?: "buffer" | "array" | "base64" | "binary" | "string";
  bookType?: string;
  compression?: boolean;
  cellDates?: boolean;
  [key: string]: unknown;
}

/** Which peer modules resolved. Useful for feature detection in tests. */
/**
 * Which optional capabilities resolved at import time.
 *
 * The `*Error` fields are null on success, or the message explaining why that
 * direction is unavailable. They are diagnostics, not a capability map: the
 * shapes here are the ones the runtime actually returns.
 */
export function available(): {
  read: boolean;
  write: boolean;
  readError: string | null;
  writeError: string | null;
};
/** Await the peer modules explicitly. */
export function preload(): Promise<void>;
/** Give a Book the SheetNames/Sheets views in place. */
export function attachViews(book: Workbook): SheetJSBook;
/** Accept a Book, a SheetJS-shaped object, raw rows, or a sheet name map. */
export function normalizeBookInput(input: unknown): Workbook;

export function read(data: Uint8Array | ArrayBuffer | string, opts?: CompatReadOptions): SheetJSBook;
export function readAsync(data: Uint8Array | ArrayBuffer | string, opts?: CompatReadOptions): Promise<SheetJSBook>;
export function readFile(path: string, opts?: CompatReadOptions): SheetJSBook;
export function readFileAsync(path: string, opts?: CompatReadOptions): Promise<SheetJSBook>;

export function write(wb: unknown, opts?: CompatWriteOptions): Uint8Array | string;
export function writeAsync(wb: unknown, opts?: CompatWriteOptions): Promise<Uint8Array | string>;
export function writeFile(wb: unknown, path: string, opts?: CompatWriteOptions): void;
export function writeFileAsync(wb: unknown, path: string, opts?: CompatWriteOptions): Promise<void>;

export function toWorksheet(sheet: Sheet): unknown;
export function toSheetJSBook(book: Workbook): SheetJSBook;

export const utils: {
  aoa_to_sheet(rows: unknown[][], opts?: { cellDates?: boolean; dateNF?: string }): Sheet;
  json_to_sheet(rows: Array<Record<string, unknown>>, opts?: { header?: string[]; dateNF?: string }): Sheet;
  sheet_to_json<T = Record<string, unknown>>(sheet: Sheet, opts?: Record<string, unknown>): T[];
  sheet_to_csv(sheet: Sheet, opts?: Record<string, unknown>): string;
  sheet_to_txt?(sheet: Sheet, opts?: Record<string, unknown>): string;
  sheet_to_aoa(sheet: Sheet, opts?: Record<string, unknown>): unknown[][];
  book_new(): SheetJSBook;
  /** Appends a sheet and returns its name, matching SheetJS. Rejects a duplicate name. */
  book_append_sheet(book: SheetJSBook, sheet: Sheet, name?: string): string;
  book_set_sheet_visibility(book: SheetJSBook, sheet: number | string, visibility: 0 | 1 | 2): void;
  decode_range(ref: string): { s: { r: number; c: number }; e: { r: number; c: number } };
  encode_cell(c: { r: number; c: number }): string;
  encode_range(s: { r: number; c: number }, e: { r: number; c: number }): string;
  format_cell(cell: Cell | CellObject, v?: unknown, opts?: Record<string, unknown>): string;
  [key: string]: unknown;
};

export const XLSX_like: {
  read: typeof read;
  readFile: typeof readFile;
  write: typeof write;
  writeFile: typeof writeFile;
  utils: typeof utils;
  version: string;
};

declare const _default: {
  read: typeof read;
  readAsync: typeof readAsync;
  readFile: typeof readFile;
  readFileAsync: typeof readFileAsync;
  write: typeof write;
  writeAsync: typeof writeAsync;
  writeFile: typeof writeFile;
  writeFileAsync: typeof writeFileAsync;
  utils: typeof utils;
  XLSX_like: typeof XLSX_like;
  available: typeof available;
  preload: typeof preload;
};
export default _default;
