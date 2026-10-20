/** gridxcel/csv - RFC 4180 CSV plus the extensions real files need. */

import type { Cell, Sheet } from "./index.d.ts";

export interface CSVOptions {
  /** Field separator. Default ",". */
  delimiter?: string;
  /** Quote character. Default '"'. */
  quote?: string;
  /** Line separator written by sheetToCSV. Default "\n". */
  lineEnding?: string;
  /** Alias for lineEnding accepted on input. */
  newline?: string;
  /** Emit a header row built from the sheet's first row. Default false. */
  header?: boolean;
  /** Prepend a UTF-8 BOM. Default false. */
  bom?: boolean;
  /** Quote every field, not only those that need it. Default false. */
  alwaysQuote?: boolean;
  /**
   * Prefix a leading =, +, - or @ in a TEXT field with an apostrophe, the
   * usual CSV-injection guard. Numbers and dates are untouched so "-2.5" stays
   * numeric. Default true.
   */
  escapeFormulae?: boolean;
  /** Format code used to render Date values. Default "yyyy-mm-dd". */
  dateNF?: string;
  /** 1904 date system. Default false. */
  date1904?: boolean;
  /** Omit rows and columns flagged hidden. Default false. */
  skipHidden?: boolean;
  /** Range to export, e.g. "A1:D20". */
  range?: string;
  /** Use raw values rather than the sheet's layout. Default false. */
  raw?: boolean;
  /** Sheet name for csvToSheet. Default "Sheet1". */
  name?: string;
  /**
   * Parse quoted fields as literal strings instead of inferring their type.
   * Default false (content decides the type, as SheetJS and pandas do). Set
   * true when a CSV must round-trip strings that look like numbers.
   */
  quotedStrings?: boolean;
  /** Treat the text as rows of blank-separated values. Default false. */
  blankrows?: boolean;
  [key: string]: unknown;
}

/** Serialise a sheet to CSV text. */
export declare function sheetToCSV(sheet: Sheet, opts?: CSVOptions): string;
/** Serialise a sheet with an explicit separator. */
export declare function sheetToDSV(sheet: Sheet, delimiter: string, opts?: CSVOptions): string;
/** Parse CSV text into a Sheet. */
export declare function csvToSheet(text: string, opts?: CSVOptions): Sheet;
/** Parse text with an explicit separator into a Sheet. */
export declare function dsvToSheet(text: string, delimiter: string, opts?: CSVOptions): Sheet;
/** Parse DSV text into a dense array of values. */
/** One parsed field, with the information needed to re-serialise it faithfully. */
export interface DSVField {
  /** The field's text, unquoted and with escapes resolved. */
  text: string;
  /** Whether the source wrapped this field in quotes. */
  quoted: boolean;
}

/**
 * Parse delimited text into rows of fields.
 *
 * Note the return type: fields are records carrying their own `quoted` flag,
 * not bare values. Quoting is preserved because it is the only way to tell a
 * numeric-looking string from a number on a later round trip.
 */
export declare function parseDSV(text: string, opts?: CSVOptions): DSVField[][];
/** Serialise a workbook (first sheet) or sheet to CSV. */
export declare function toCSV(input: unknown, opts?: CSVOptions): string;
/** Parse CSV text into a Sheet. */
export declare function fromCSV(text: string, opts?: CSVOptions): Sheet;
