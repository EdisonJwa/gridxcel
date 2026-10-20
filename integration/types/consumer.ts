// A consumer of the published types. If this compiles, the declarations a user
// sees are valid. src/index.d.ts previously carried a type literal in a value
// position - `export default { VERSION: string; ... }` - which every
// type-checking consumer hit as a syntax error, and nothing in this repository
// noticed because the smoke test only regex-scanned the file.
//
// It also imports every root export that once had no declaration, so the gap
// cannot reopen: compiling is not the same as matching, and a runtime export the
// types do not describe is invisible to a compiler until someone names it.
import gridxcel, {
  read, write, writeWorkbook, createBookWriter, splitSheetRef, toRange,
  normalizeColor, numberLiteral, mergeBooks, renameSheet, countCells,
  XmlParser, parseTag, decodeEntities, encodeXmlText, encodeXmlAttr,
  createZip, BUILTIN_FORMATS, generalNumber,
  type Workbook, type Cell, type Range, type CellAddress, type ColorValue,
} from "gridxcel";
import * as compat from "gridxcel/compat";
import { sheetToCSV } from "gridxcel/csv";
import { formatNumber } from "gridxcel/numfmt";

export async function demo(): Promise<void> {
  const v: string = gridxcel.VERSION;
  const book: Workbook = await read("x.xlsx");
  const cell: Cell | undefined = book.sheets[0]?.rows[0]?.[0];
  if (cell?.t === "date") void cell.serial;
  // zf is added by the reader on both batch and streaming paths.
  const fmt: string | undefined = cell?.zf;
  void fmt;

  await write([["a", 1]], { bookType: "xlsx" });
  await writeWorkbook([["a", 1]], { bookType: "xlsx" });
  await sheetToCSV(book.sheets[0]);
  void compat.available();
  void mergeBooks(book, book);
  void renameSheet(book, "a", "b");
  void countCells(book);

  const [sheetName, rangeText]: [string | null, string] = splitSheetRef("Sheet1!A1:B2");
  const range: Range | undefined = toRange("A1:B2");
  const addr: CellAddress | undefined = range?.s;
  void sheetName; void rangeText; void addr;
  void toRange(3);

  const colour: ColorValue | string | undefined = normalizeColor("#FF0000");
  void colour;
  void numberLiteral(1.5);
  void parseTag('c r="A1"');
  void decodeEntities("a &amp; b");
  void encodeXmlText("<");
  void encodeXmlAttr('"');
  void new XmlParser({});
  void createZip([["a", new Uint8Array()]]);
  void BUILTIN_FORMATS;
  void generalNumber(1);
  void formatNumber(1, "0.00");
  void createBookWriter({ sheets: [] });
  void v;
}
