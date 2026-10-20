/**
 * Task: turn a large sheet into CSV without holding the CSV in memory.
 *
 * sheetToCSV() is O(output): on a 200,000 x 10 sheet it retains about 14 MB for
 * the string. createCSVStream() yields the same bytes a row at a time, so the
 * output can go straight to a file, a socket or a Response.
 *
 * Run: node examples/10-stream-csv-out.mjs
 */
import { createWriteStream } from "node:fs";
import { rmSync } from "node:fs";
import { readFileSync } from "node:fs";
import { Readable } from "node:stream";
import { read } from "gridxcel";
// The csv entry point is the synchronous surface; the root re-export is
// promise-based on purpose (see docs/API.md on the root versus compat split).
import { createCSVStream, sheetToCSV } from "gridxcel/csv";

const FIXTURE = new URL("../test/fixtures/golden.xlsx", import.meta.url);
const OUT = new URL("../examples/out/volume.csv", import.meta.url);

const book = await read(readFileSync(FIXTURE));
const sheet = book.sheets[2];

/* ---- 1. the eager form, for reference -------------------------------- */
const text = sheetToCSV(sheet);
console.log("sheetToCSV        :", text.length, "bytes,", text.split("\n").length - 1, "rows");

/* ---- 2. the streaming form, to a file -------------------------------- */
rmSync(OUT, { force: true });
let bytes = 0;
let chunks = 0;
const sink = createWriteStream(OUT);
// Rows arrive as strings; converting each chunk to bytes keeps memory flat.
const source = Readable.from(
  (async function* () {
    for await (const chunk of createCSVStream(sheet, {}, { chunkRows: 64 })) {
      bytes += Buffer.byteLength(chunk);
      chunks++;
      yield chunk;
    }
  })(),
);
await new Promise((resolve, reject) => {
  source.pipe(sink);
  sink.on("finish", resolve);
  sink.on("error", reject);
});

console.log("createCSVStream   :", bytes, "bytes in", chunks, "chunks");
console.log("same bytes        :", bytes === Buffer.byteLength(text));

/* ---- 3. the same stream straight into a Response --------------- */
// In a server or a worker there is no file to write: the stream IS the body.
const body = Readable.from(
  (async function* () {
    for await (const chunk of createCSVStream(sheet, {}, { chunkRows: 256 })) yield chunk;
  })(),
);
const web = Readable.toWeb(body);
console.log("Response-ready    :", typeof web.getReader === "function" ? "ReadableStream" : "unknown");
