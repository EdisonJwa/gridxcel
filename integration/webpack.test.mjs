/**
 * Does gridxcel build AND RUN under webpack 5?
 *
 * BROWSER.md named webpack as a bundler that honours the `browser` field. Naming
 * it is not the same as running it, so this runs it.
 *
 * Unlike the esbuild and rolldown suites, this one executes in a REAL browser
 * rather than a bare vm context, and that difference is the point. webpack
 * turns the library's lazy peer imports (\`import("./write.js")\`) into a separate
 * chunk and loads it on demand by injecting a script tag. A bare sandbox has no
 * chunk loader, so it fails with "internal module ./write.js is not available
 * yet" - even though the bundle is correct and works in a browser, which was
 * confirmed by hand before this test was written. Testing webpack the way webpack
 * works is the only way to get a true answer.
 *
 * Run: node webpack.test.mjs
 */

import webpack from "webpack";
import puppeteer from "puppeteer-core";
import { readFileSync, writeFileSync, mkdirSync, rmSync, existsSync, readdirSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, "out", "webpack");

let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};
const done = () => {
  console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
  process.exit(failures === 0 ? 0 : 1);
};

rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const stats = await new Promise((resolve) => {
  webpack({
    mode: "production",
    entry: join(here, "browser-entry.mjs"),
    target: "web",
    resolve: { aliasFields: ["browser"], conditionNames: ["browser", "import", "default"] },
    output: { path: outDir, filename: "bundle.js", publicPath: "", library: { type: "var", name: "GridxcelWebpack" } },
    optimization: { minimize: true },
    performance: { hints: false },
    stats: "errors-warnings",
  }, (err, s) => resolve(err ? { err } : s));
});

if (stats.err || stats.hasErrors?.()) {
  const detail = stats.err ? String(stats.err.message) : stats.toJson().errors.map((e) => e.message).join(" | ");
  check("webpack builds for the browser", false, detail.split("\n").slice(0, 4).join(" "));
  done();
}
check("webpack builds for the browser", true, readdirSync(outDir).join(", "));

const code = readFileSync(join(outDir, "bundle.js"), "utf8");
const nodeSpecs = [...new Set([...code.matchAll(/["'](node:[a-z/]+)["']/g)].map((m) => m[1]))];
check("no node: specifier survives", nodeSpecs.length === 0, nodeSpecs.join(", "));
// Check for an actual import of node:zlib, not the bare string: the browser
// platform module MENTIONS node:zlib in its error messages ("this host provides
// neither node:zlib nor..."), which is prose, not a dependency. Matching the
// string flagged those messages as a shim.
const zlibImports = [...code.matchAll(/(?:require\(|import\()\s*["']node:zlib["']/g)].length
  + [...code.matchAll(/from\s*["']node:zlib["']/g)].length;
check("the browser field was honoured (node:zlib not loaded)", zlibImports === 0,
  zlibImports === 0 ? "mentions only, in error text" : zlibImports + " import(s)");

/* Run it in a real browser, with webpack's chunk served alongside. */
const exe = [process.env.CHROME_PATH,
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
  "/usr/bin/google-chrome"].filter(Boolean).find((p) => { try { return existsSync(p); } catch { return false; } });

if (!exe) {
  console.log("  SKIP  no Chrome found; set CHROME_PATH to run the execution check");
  done();
}

const fixture = readFileSync(join(here, "..", "test", "fixtures", "golden.xlsx")).toString("base64");
writeFileSync(join(outDir, "index.html"), `<!doctype html><meta charset="utf-8">
<script type="text/plain" id="fixture">${fixture}</script>
<script src="./bundle.js"></script>
<script>
window.__r = (async () => {
  const raw = atob(document.getElementById("fixture").textContent.trim());
  const bytes = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
  try {
    const out = await GridxcelWebpack.run(bytes);
    return { ok: true, sheets: out.sheets, streamed: out.streamed, wrote: out.wroteBytes,
             buffer: typeof Buffer, process: typeof process, require: typeof require };
  } catch (e) { return { ok: false, err: (e && e.code ? e.code + ": " : "") + String(e && e.message) }; }
})();
</script>`);

let browser;
try {
  browser = await puppeteer.launch({ executablePath: exe, headless: true, args: ["--allow-file-access-from-files", "--no-sandbox", "--disable-gpu"] });
  const p = await browser.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(String(e.message)));
  await p.goto(pathToFileURL(join(outDir, "index.html")).href, { waitUntil: "load" });
  const out = await p.evaluate(() => window.__r);
  check("no page errors", errors.length === 0, errors.slice(0, 2).join(" | "));
  check("the webpack bundle runs in a real browser", out.ok === true, out.ok ? "" : out.err);
  check("with no Node APIs in the page", out.buffer === "undefined" && out.process === "undefined" && out.require === "undefined",
    "Buffer:" + out.buffer + " process:" + out.process + " require:" + out.require);
  check("reading a deflated workbook", Array.isArray(out.sheets) && out.sheets.length === 4, JSON.stringify(out.sheets));
  check("streaming and writing round-trip", out.streamed === 205 && out.wrote > 1000, out.streamed + " rows, " + out.wrote + " bytes");
} catch (err) {
  check("the webpack bundle runs in a real browser", false, String(err && err.message).split("\n")[0]);
} finally {
  if (browser) await browser.close();
}

done();
