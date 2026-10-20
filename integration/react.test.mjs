/**
 * Is gridxcel actually framework-agnostic?
 *
 * The claim is that the core is plain data in and plain data out, so it works
 * inside a React server render, a Preact render and a bundler without any
 * framework-specific shim. This renders real components with the real
 * libraries and checks what came out, rather than asserting the claim.
 *
 * Run: node react.test.mjs
 */

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { createElement as h } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { renderToStaticMarkup as preactRender } from "preact-render-to-string";
import { h as preactH } from "preact";
import { read, sheetToJSON, write } from "gridxcel";

const here = dirname(fileURLToPath(import.meta.url));
let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};

const bytes = new Uint8Array(readFileSync(join(here, "..", "test", "fixtures", "golden.xlsx")));

/* React: an async server component reading a workbook and rendering a table. */
async function SheetTable({ bytes }) {
  const book = await read(bytes);
  const rows = await sheetToJSON(book.sheets[0], { header: 1 });
  const head = rows[0] || [];
  const body = rows.slice(1, 4);
  return h("table", null,
    h("thead", null, h("tr", null, head.map((c, i) => h("th", { key: i }, String(c))))),
    h("tbody", null, body.map((r, i) => h("tr", { key: i }, r.map((c, j) => h("td", { key: j }, String(c)))))),
  );
}

try {
  const markup = renderToStaticMarkup(await SheetTable({ bytes }));
  check("react server-renders a sheet", markup.startsWith("<table>") && markup.includes("<td>"), markup.slice(0, 60) + "...");
  check("react markup contains real cell text", /<td>[^<]+<\/td>/.test(markup));
} catch (err) {
  check("react server-renders a sheet", false, String(err && err.message).split("\n")[0]);
}

/* Preact: same data, different renderer, no adapter. */
try {
  const markup = preactRender(preactH("ul", null, [preactH("li", { key: "a" }, "ada"), preactH("li", { key: "b" }, "42")]));
  check("preact renders", markup === "<ul><li>ada</li><li>42</li></ul>", markup);
  const book = await read(bytes);
  const rows = await sheetToJSON(book.sheets[0], { header: 1 });
  const list = preactRender(preactH("ol", null, rows.slice(0, 3).map((r, i) => preactH("li", { key: i }, String(r[0])))));
  check("preact renders real sheet data", list.includes("<li>") && list.length > 12, list.slice(0, 60));
} catch (err) {
  check("preact renders", false, String(err && err.message).split("\n")[0]);
}

/* The core must be plain data: a workbook has to survive structuredClone. */
try {
  const book = await read(bytes);
  const cloned = structuredClone(book);
  check("a workbook survives structuredClone", cloned.sheets.length === book.sheets.length, cloned.sheets.length + " sheets");
  const out = await write([["a", 1]], { bookType: "xlsx" });
  check("the writer returns transferable bytes", out instanceof Uint8Array, out.constructor.name);
} catch (err) {
  check("a workbook survives structuredClone", false, String(err && err.message).split("\n")[0]);
}

console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
process.exit(failures === 0 ? 0 : 1);
