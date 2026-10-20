// Browser-flavoured entry: no filesystem, no Node builtins, only the byte API.
import { read, write, sheetToJSON, sheetToCSV, eachRow } from "gridxcel";

export async function run(bytes) {
  const book = await read(bytes);
  const rows = await sheetToJSON(book.sheets[0], { header: 1 });
  const csv = await sheetToCSV(book.sheets[0]);
  let streamed = 0;
  for await (const r of eachRow(bytes)) streamed += r.cells.length ? 1 : 0;
  const out = await write([["name", "score"], ["ada", 42]], { bookType: "xlsx" });
  const reread = await read(out);
  return {
    sheets: book.sheetNames,
    firstCell: rows[0] && rows[0][0],
    csvStart: csv.slice(0, 12),
    streamed,
    wroteBytes: out.length,
    wroteIsUint8: out instanceof Uint8Array,
    rereadSheets: reread.sheetNames,
  };
}
