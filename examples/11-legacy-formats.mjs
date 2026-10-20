/**
 * Task: read and write the legacy text formats Excel still supports.
 *
 * SYLK and DIF predate XLSX. They are plain text, they hold typed cells, and a
 * surprising number of systems still export them. gridxcel reads and writes both
 * through gridxcel/formats, and gridxcel/compat recognises them from the content
 * so a migrating caller does not have to say which format it handed over.
 *
 * Run: node examples/11-legacy-formats.mjs
 */
import { difToSheet, sheetToDIF, sheetToSYLK, sylkToSheet } from "gridxcel/formats";
import { aoaToSheet } from "gridxcel/json";
import { read } from "gridxcel/compat";

const sheet = aoaToSheet([
  ["name", "score", "passed"],
  ["ada", 42, true],
  ["grace", 37.5, false],
]);

/* ---- SYLK ------------------------------------------------------------- */
const sylk = sheetToSYLK(sheet);
console.log("SYLK:", sylk.length, "bytes,", sylk.split("\r\n").length, "records");
console.log("  " + sylk.split("\r\n").slice(0, 4).join(" | "));
const fromSylk = sylkToSheet(sylk);
console.log("  read back:", JSON.stringify(fromSylk.rows[1].map((c) => c.v)));

/* ---- DIF -------------------------------------------------------------- */
const dif = sheetToDIF(sheet);
console.log("DIF:", dif.length, "bytes,", dif.split("\r\n").length, "lines");
const fromDif = difToSheet(dif);
console.log("  read back:", JSON.stringify(fromDif.rows[1].map((c) => c.v)));

/* ---- through the compat surface, no format named ---------------------- */
// The sniffer recognises SYLK by its "ID" record and DIF by "TABLE".
const sniffedSylk = read(sylk);
const sniffedDif = read(dif);
console.log("sniffed:", sniffedSylk.sheetNames, sniffedDif.sheetNames);
console.log("  same values:", JSON.stringify(sniffedSylk.sheets[0].rows[1].map((c) => c.v)));
