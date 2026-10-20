/**
 * A third bundler, because "it bundles" is a claim about a toolchain, not about
 * the library. esbuild and Vite are covered elsewhere; rolldown is the Rust
 * bundler inside Vite 8 and resolves the browser field independently.
 *
 * Run: node rolldown.test.mjs
 */

import { build } from "rolldown";
import { createContext, runInContext } from "node:vm";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};

let code = "";
try {
  // rolldown's build() resolves to the bundle directly; there is no generate().
  const bundle = await build({
    input: join(here, "browser-entry.mjs"),
    platform: "browser",
    output: { format: "iife", name: "GridxcelRolldown", minify: true },
  });
  code = bundle.output.find((c) => c.type === "chunk").code;
  check("rolldown bundles for the browser", true, "(" + (code.length / 1024).toFixed(0) + " kB)");
} catch (err) {
  check("rolldown bundles for the browser", false, String(err && err.message).split("\n").slice(0, 3).join(" | "));
}

if (code) {
  const nodeSpecs = [...new Set([...code.matchAll(/["'](node:[a-z/]+)["']/g)].map((m) => m[1]))];
  check("no node: specifier survives", nodeSpecs.length === 0, nodeSpecs.join(", "));
  check("no createRequire shim", !code.includes("createRequire"));

  const sandbox = {
    TextDecoder, TextEncoder, DecompressionStream, CompressionStream, ReadableStream,
    Blob, URL, atob, btoa, console, setTimeout, clearTimeout, queueMicrotask, structuredClone, performance,
  };
  sandbox.globalThis = sandbox; sandbox.self = sandbox; sandbox.window = sandbox;
  const ctx = createContext(sandbox);
  const fixture = readFileSync(join(here, "..", "test", "fixtures", "golden.xlsx"));
  try {
    runInContext(code, ctx, { timeout: 60000 });
    const out = await ctx.GridxcelRolldown.run(new Uint8Array(fixture));
    check("the rolldown bundle runs with no Node APIs present", Array.isArray(out.sheets) && out.sheets.length === 4, JSON.stringify(out.sheets));
    check("and round-trips a workbook", out.wroteIsUint8 && out.rereadSheets.length === 1, out.wroteBytes + " bytes");
  } catch (err) {
    check("the rolldown bundle runs with no Node APIs present", false, String(err && err.message).split("\n")[0]);
  }
}

console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
process.exit(failures === 0 ? 0 : 1);
