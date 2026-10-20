import type { Sheet } from "./index.d.ts";

/** Options shared by the text formats. */
export interface TextFormatOptions {
  /** Sheet name. Defaults to "Sheet1". */
  name?: string;
  /** Line ending. Defaults to CRLF, which is what Excel writes. */
  eol?: string;
}

/**
 * SYLK text -> Sheet.
 *
 * Reads the value records (C;Y..;X..;K..), honouring the format's rule that an
 * omitted Y or X reuses the previous coordinate. Styles are parsed far enough to
 * skip safely and are not applied.
 */
export declare function sylkToSheet(text: string, opts?: TextFormatOptions): Sheet;

/** Sheet -> SYLK text. Dates are written as their serial: SYLK has no date type. */
export declare function sheetToSYLK(sheet: Sheet, opts?: TextFormatOptions): string;

/**
 * DIF text -> Sheet.
 *
 * Follows the 1981 technical specification: header chunks are three lines, data
 * chunks are two, and a numeric value lives in the type pair rather than on the
 * keyword line.
 */
export declare function difToSheet(text: string, opts?: TextFormatOptions): Sheet;

/** Sheet -> DIF text. Dates are written as their serial: DIF has no date type. */
export declare function sheetToDIF(sheet: Sheet, opts?: TextFormatOptions): string;
