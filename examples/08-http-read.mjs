/**
 * Read a workbook straight from an HTTP response.
 *
 * Shows: read() over a fetch() body, over an ArrayBuffer, and over a Node
 * Readable; streaming a CSV response into a sheet; and how a typed read error
 * surfaces from a bad response.
 *
 * The server is created in this process and serves a committed fixture, so the
 * example needs no network. Run from anywhere:
 *   node examples/08-http-read.mjs
 */
import { csvToSheet, GridxcelError, read, sheetToCSV } from "gridxcel";
import { createServer } from "node:http";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const FIXTURE = fileURLToPath(new URL("../test/fixtures/golden.xlsx", import.meta.url));

// A deliberately simple server: one workbook route and one CSV route.
const server = createServer((req, res) => {
  if (req.url === "/report.xlsx") {
    const body = readFileSync(FIXTURE);
    res.writeHead(200, {
      "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "content-length": body.length,
    });
    res.end(body);
  } else if (req.url === "/report.csv") {
    res.writeHead(200, { "content-type": "text/csv; charset=utf-8" });
    res.end("region,units\nnorth,1200\nsouth,950\n");
  } else {
    res.writeHead(404, { "content-type": "text/plain" });
    res.end("not found");
  }
});

await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const base = "http://127.0.0.1:" + server.address().port;
console.log("serving on", base);

try {
  /* ------------------------------------------- XLSX from a fetch() body */

  // response.body is a web ReadableStream. read() consumes it without you
  // collecting it first, so the bytes never sit in a buffer you allocated.
  const xlsxResponse = await fetch(base + "/report.xlsx");
  console.log("\nstatus       :", xlsxResponse.status, xlsxResponse.headers.get("content-type").split(";")[0]);
  const book = await read(xlsxResponse.body);
  console.log("sheets       :", book.sheets.map((s) => s.name).join(", "));

  /* ------------------------------------------- XLSX from an ArrayBuffer */

  const bufferResponse = await fetch(base + "/report.xlsx");
  const fromBuffer = await read(await bufferResponse.arrayBuffer());
  console.log("from buffer  :", fromBuffer.sheets.length, "sheets,", fromBuffer.sheets[2].rowCount, "rows on Volume");

  /* ------------------------------------------- CSV from the same client */

  const csvResponse = await fetch(base + "/report.csv");
  const sheet = await csvToSheet(await csvResponse.text(), { header: true });
  console.log("\nCSV status   :", csvResponse.status);
  console.log("CSV rows     :", JSON.stringify(await sheetToCSV(sheet)));

  /* --------------------------------------- what a bad response looks like */

  // Fetching an HTML error page and handing it to read() is the classic
  // production accident. It fails loudly rather than producing an empty book.
  const htmlResponse = await fetch(base + "/nope");
  try {
    await read(await htmlResponse.arrayBuffer());
    console.log("\nunexpected: an HTML body was accepted as a workbook");
  } catch (err) {
    console.log("\nHTML body ->", err.name, err.code, "| instanceof GridxcelError:", err instanceof GridxcelError);
    console.log("Check response.ok before reading, or let this throw.");
  }
} finally {
  server.close();
  await new Promise((resolve) => server.once("close", resolve));
  console.log("\nserver closed");
}
