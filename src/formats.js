/**
 * Text-based spreadsheet formats beyond XLSX and CSV.
 *
 * These are reachable through the compat surface by content sniffing - hand it a
 * SYLK or DIF file and it does the right thing - but a caller who already knows
 * the format should say so rather than pay for a guess. One entry point for
 * "other formats" keeps the published surface small as more are added.
 */

export { sylkToSheet, sheetToSYLK } from "./sylk.js";
export { difToSheet, sheetToDIF } from "./dif.js";
