/**
 * Does gridxcel actually work in a browser?
 *
 * Not "does it type-check" and not "does the docs page say it does": this
 * bundles the library for the browser with esbuild, refuses the bundle if any
 * node: specifier survives, and then EXECUTES it in a sandbox that has the Web
 * platform globals but no Buffer, no process, no require and no module loader.
 * A library that reaches for node:zlib at import time fails here even though it
 * would have worked fine under Node.
 *
 * Run: node browser-bundle.test.mjs
 */

import { build } from "esbuild";
import { createContext, runInContext } from "node:vm";
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { readSync, write as gxWrite } from "gridxcel";

const here = dirname(fileURLToPath(import.meta.url));

let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};

/* 1. Bundle for the browser. A Node-only import is an error, not a warning. */
const result = await build({
  entryPoints: [join(here, "browser-entry.mjs")],
  bundle: true,
  format: "iife",
  globalName: "GridxcelBrowser",
  platform: "browser",
  target: ["es2020"],
  write: false,
  logLevel: "silent",
  // Deliberately NOT marking node: as external: if the library needs it, the
  // browser build must fail rather than emit an unresolvable import.
}).catch((err) => ({ errors: [err] }));

if (result.errors && result.errors.length) {
  check("esbuild bundles for the browser", false, "\n" + result.errors.map((e) => "      " + (e.text || e.message)).join("\n"));
  console.log("\n  " + failures + " failure(s)");
  process.exit(1);
}
const code = result.outputFiles[0].text;
check("esbuild bundles for the browser", true, "(" + (code.length / 1024).toFixed(0) + " kB)");

/* 2. No node: specifier may survive into the bundle. */
const nodeSpecifiers = [...new Set([...code.matchAll(/["'](node:[a-z/]+)["']/g)].map((m) => m[1]))];
check("no node: import in the browser bundle", nodeSpecifiers.length === 0, nodeSpecifiers.length ? nodeSpecifiers.join(", ") : "");
const nodeCalls = [...new Set([...code.matchAll(/\brequire\s*\(\s*["']([a-z:]+)["']/g)].map((m) => m[1]))];
check("no require() in the browser bundle", nodeCalls.length === 0, nodeCalls.join(", "));

/* 3. Execute it with the Web platform only. */
const sandbox = {
  TextDecoder, TextEncoder, DecompressionStream, CompressionStream, ReadableStream,
  Blob, URL, atob, btoa, console, setTimeout, clearTimeout, queueMicrotask,
  structuredClone, performance,
};
sandbox.globalThis = sandbox;
sandbox.self = sandbox;
sandbox.window = sandbox;
const ctx = createContext(sandbox);
// Buffer, process, require and module are deliberately absent.

const fixture = readFileSync(join(here, "..", "test", "fixtures", "golden.xlsx"));
try {
  runInContext(code, ctx, { timeout: 60000 });
  const out = await ctx.GridxcelBrowser.run(new Uint8Array(fixture));
  check("reads a workbook with no Node APIs present", Array.isArray(out.sheets) && out.sheets.length === 4, JSON.stringify(out.sheets));
  check("sheetToJSON works", out.firstCell === "Product" || typeof out.firstCell === "string", JSON.stringify(out.firstCell));
  check("sheetToCSV works", typeof out.csvStart === "string" && out.csvStart.length > 0);
  check("eachRow streams", out.streamed === 205, String(out.streamed) + " rows");
  check("writes xlsx bytes", out.wroteIsUint8 && out.wroteBytes > 1000, out.wroteBytes + " bytes");
  check("reads back what it wrote", Array.isArray(out.rereadSheets) && out.rereadSheets.length === 1, JSON.stringify(out.rereadSheets));
} catch (err) {
  check("runs in a browser-like sandbox", false, err && err.message ? err.message.split("\n")[0] : String(err));
}

/* 4. Tree-shaking: a one-function import must not pull in the world. */
const shakable = await build({
  stdin: { contents: 'import { sheetToCSV } from "gridxcel"; export const f = sheetToCSV;', resolveDir: here, loader: "js" },
  bundle: true, format: "esm", platform: "browser", write: false, logLevel: "silent", minify: true,
}).catch((err) => ({ errors: [err] }));
if (shakable.errors && shakable.errors.length) {
  check("tree-shaking build succeeds", false, shakable.errors.map((e) => e.text || e.message).join("; "));
} else {
  const shaken = shakable.outputFiles[0].text.length;
  check("importing one function tree-shakes", shaken < code.length * 0.8, (shaken / 1024).toFixed(0) + " kB vs " + (code.length / 1024).toFixed(0) + " kB for the full API");
}

console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
process.exit(failures === 0 ? 0 : 1);
