/**
 * Spreadsheet limits, in one place.
 *
 * These numbers used to be repeated as literals in api.js, json.js, read.js and
 * workbook-writer.js. A limit that exists twice is a limit that eventually
 * disagrees with itself, so every module imports them from here.
 */

/** Maximum sheet-name length Excel accepts, in UTF-16 code units. */
export const MAX_SHEET_NAME_LENGTH = 31;

/** Maximum column count (column XFD). */
export const MAX_COLUMNS = 16384;

/** Maximum row count. */
export const MAX_ROWS = 1048576;

/**
 * Largest number of sheets the reader will assemble from one workbook.
 *
 * Excel itself allows far fewer in practice, and a workbook is untrusted input:
 * a 794 kB file declaring 100,000 <sheet> entries cost 5.9 s and 321 MB before
 * this limit existed. Parts beyond the limit are skipped with a warning rather
 * than refused outright, so a truncated read is still possible.
 */
export const MAX_SHEETS = 4096;

/** Sheet name used when the caller does not supply one. */
export const DEFAULT_SHEET_NAME = "Sheet1";

/* --------------------------------------------------------- OOXML namespaces */

/**
 * The XML declaration every part starts with.
 *
 * These live here rather than beside their first user because two modules need
 * them and neither may import the other: workbook-writer imports StyleTable from
 * style-registry, so style-registry importing anything from the writer would be
 * a cycle. constants.js depends on nothing, which is what makes it the right
 * home - and it is why style-registry previously carried its own copy.
 */
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** The spreadsheet main namespace: every worksheet, styles and workbook part. */
export const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/** Relationships, used by every .rels part. */
export const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** Package-level relationships. */
export const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
/** [Content_Types].xml. */
export const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
/** Core document properties (docProps/core.xml). */
export const NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
export const NS_DC = "http://purl.org/dc/elements/1.1/";
export const NS_DCTERMS = "http://purl.org/dc/terms/";
export const NS_DCMITYPE = "http://purl.org/dc/dcmitype/";
export const NS_XSI = "http://www.w3.org/2001/XMLSchema-instance";
/** Extended application properties (docProps/app.xml). */
export const NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
export const NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
