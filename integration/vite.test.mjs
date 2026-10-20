/**
 * Does gridxcel build AND RUN under Vite?
 *
 * esbuild proves the library has no unresolvable Node import. Vite proves it
 * survives a real application build. A library can pass one and fail the other,
 * and this test exists because it did: an earlier version of this file only
 * INSPECTED Vite's output for node: specifiers, which passed while the bundle
 * was broken. Vite had quietly substituted a browser shim for node:module, and
 * the emitted bundle threw "(0, createRequire) is not a function" the moment it
 * was loaded, with no compression code in it at all.
 *
 * So the build is now executed in a sandbox with the Web platform only. A
 * bundle that merely looks clean does not pass.
 *
 * Run: node vite.test.mjs
 */

import { build } from "vite";
import { createContext, runInContext } from "node:vm";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { rmSync, existsSync, readdirSync, readFileSync } from "node:fs";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, "out", "vite");

let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};

rmSync(outDir, { recursive: true, force: true });

try {
  await build({
    root: here,
    logLevel: "silent",
    build: {
      outDir,
      emptyOutDir: true,
      // iife, because an ES module cannot be evaluated inside a bare vm context,
      // and the point of this test is to RUN the artefact.
      lib: { entry: join(here, "browser-entry.mjs"), name: "GridxcelVite", formats: ["iife"] },
      minify: "esbuild",
      target: "es2020",
    },
  });
  check("vite builds the library", true);
} catch (err) {
  check("vite builds the library", false, String(err && err.message).split("\n").slice(0, 6).join(" | "));
  console.log("\n  " + failures + " failure(s)");
  process.exit(1);
}

const produced = existsSync(outDir) ? readdirSync(outDir).filter((f) => f.endsWith(".js")) : [];
check("vite emitted a bundle", produced.length > 0, produced.join(", "));

for (const f of produced) {
  const text = readFileSync(join(outDir, f), "utf8");
  const nodeSpecs = [...new Set([...text.matchAll(/["'](node:[a-z/]+)["']/g)].map((m) => m[1]))];
  check("no node: specifier survives (" + f + ")", nodeSpecs.length === 0, nodeSpecs.join(", "));
  // Vite's shim for node:module pulls this in. It is a browser-hostile signal.
  check("no createRequire shim in the bundle (" + f + ")", !text.includes("createRequire"));
  check("compression code is actually present (" + f + ")", /inflate|DecompressionStream|def late|deflate/i.test(text));
}

/* Run it. This is the check that the earlier version was missing. */
const code = produced.length ? readFileSync(join(outDir, produced[0]), "utf8") : "";
const sandbox = {
  TextDecoder, TextEncoder, DecompressionStream, CompressionStream, ReadableStream,
  Blob, URL, atob, btoa, console, setTimeout, clearTimeout, queueMicrotask,
  structuredClone, performance,
};
sandbox.globalThis = sandbox;
sandbox.self = sandbox;
sandbox.window = sandbox;
const ctx = createContext(sandbox);
const fixture = readFileSync(join(here, "..", "test", "fixtures", "golden.xlsx"));
try {
  runInContext(code, ctx, { timeout: 60000 });
  const out = await ctx.GridxcelVite.run(new Uint8Array(fixture));
  check("the vite bundle runs with no Node APIs present", Array.isArray(out.sheets) && out.sheets.length === 4, JSON.stringify(out.sheets));
  check("the vite bundle round-trips a workbook", out.wroteIsUint8 && out.rereadSheets.length === 1, out.wroteBytes + " bytes");
} catch (err) {
  check("the vite bundle runs with no Node APIs present", false, String(err && err.message).split("\n")[0]);
}

console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
process.exit(failures === 0 ? 0 : 1);
