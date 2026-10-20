/** gridxcel/numfmt - ECMA-376 number formats and Excel date serials. */

/** Built-in numFmtId -> format code (ECMA-376 18.8.30). */
export declare const BUILTIN_FORMATS: Record<number, string>;

/**
 * True when a format renders a date and/or time. Accepts a format code or a
 * built-in id. Cached per code, because every numeric cell consults it.
 */
export declare function isDateFormat(codeOrId: string | number | undefined | null): boolean;

/** Format code for a built-in id, or undefined when the id is custom. */
export declare function builtinFormat(id: number): string | undefined;

/** Render a number through a format code. */
export declare function formatNumber(value: number, code?: string | number | null, opts?: { date1904?: boolean }): string;

/** Render a date serial through a date/time format code. */
export declare function formatDateTime(serial: number, code: string, opts?: { date1904?: boolean }): string;

/** Format any cell value for display, dispatching on its runtime type. */
export declare function formatValue(value: unknown, code?: string | number | null, opts?: { date1904?: boolean }): string;

/** "General" rendering, matching Excel's 11-significant-digit behaviour. */
export declare function generalNumber(value: number | unknown): string;

/** Excel serial -> Date whose LOCAL fields are the spreadsheet wall clock. */
export declare function serialToDate(serial: number, date1904?: boolean): Date;
/** Date -> Excel serial, using the Date's local fields. */
export declare function dateToSerial(date: Date, date1904?: boolean): number;
/** Excel serial -> UTC milliseconds since the Unix epoch. */
export declare function serialToUTCms(serial: number, date1904?: boolean): number;
/** Break a serial into calendar fields without allocating a Date. */
export declare function serialParts(serial: number): {
  y: number; m: number; d: number; dow: number; h: number; mi: number; s: number; ms: number; serial: number;
};
/** Split a format code on top-level semicolons. */
export declare function splitSections(code: string): string[];
/** Which of a format code's sections applies to a value. */
export declare function pickSection(sections: string[], value: number, isText: boolean): number;
