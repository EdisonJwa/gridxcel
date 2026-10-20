/**
 * Does gridxcel work in a REAL browser?
 *
 * Every other browser test here runs the bundle in Node's vm with globals we
 * chose. That is a model of a browser, and the model can be wrong in exactly the
 * ways that matter. This one launches the Chrome installed on this machine,
 * loads a real page over file://, bundles with esbuild, and runs the library in
 * the page's own JavaScript context - a real engine, real streams, real atob,
 * no Buffer, no process, no require, nothing shimmed.
 *
 * It is the difference between "esbuild produced something plausible" and
 * "Chrome ran it".
 *
 * Run: node real-browser.test.mjs
 */

import puppeteer from "puppeteer-core";
import { build } from "esbuild";
import { readFileSync, existsSync, mkdirSync, writeFileSync, rmSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, "out", "real");

let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};
const done = () => {
  console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
  process.exit(failures === 0 ? 0 : 1);
};

/* Find a browser. Chrome first, then Edge, which all Windows machines have. */
const CANDIDATES = [
  process.env.CHROME_PATH,
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
  "/usr/bin/google-chrome",
  "/usr/bin/chromium",
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
].filter(Boolean);

const exe = CANDIDATES.find((p) => { try { return existsSync(p); } catch { return false; } });
if (!exe) {
  // Not a failure: a machine without a browser cannot test one. Say so loudly
  // rather than passing quietly.
  console.log("  SKIP  no Chrome or Edge found; set CHROME_PATH to run this");
  console.log("\n  SKIPPED (not a pass)");
  process.exit(0);
}
console.log("  using " + exe);

/* Build the bundle and the page. */
rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });
try {
  await build({
    stdin: {
      contents: 'export * from "gridxcel";',
      resolveDir: here,
      loader: "js",
    },
    bundle: true, format: "iife", globalName: "GridxcelReal",
    platform: "browser", target: ["es2020"],
    outfile: join(outDir, "gridxcel.iife.js"),
    logLevel: "silent",
  });
  check("esbuild produced a browser bundle", true);
} catch (err) {
  check("esbuild produced a browser bundle", false, String(err && err.message).split("\n")[0]);
  done();
}

/* Hand the fixture to the page without a server: base64 inline. */
const fixture = readFileSync(join(here, "..", "test", "fixtures", "golden.xlsx")).toString("base64");
let page = readFileSync(join(here, "real-browser-page.html"), "utf8");
// The fixture has to be in the DOM before the script that reads it runs, so it
// goes at the top of <body> rather than before </body>. Injecting it at the end
// made getElementById return null.
page = page.replace("<body>", '<body>\n<script type="text/plain" id="fixture">' + fixture + "</script>");
writeFileSync(join(outDir, "index.html"), page);

/* Launch and run. */
let browser;
try {
  browser = await puppeteer.launch({
    executablePath: exe,
    headless: true,
    args: ["--allow-file-access-from-files", "--no-sandbox", "--disable-gpu"],
  });
  const p = await browser.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(String(e.message)));
  p.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });

  await p.goto(pathToFileURL(join(outDir, "index.html")).href, { waitUntil: "load" });
  const result = await p.evaluate(() => window.__result);

  check("no page errors", errors.length === 0, errors.slice(0, 2).join(" | "));
  check("a real engine, not Node", result.hasBuffer === "undefined" && result.hasProcess === "undefined" && result.hasRequire === "undefined",
    "Buffer:" + result.hasBuffer + " process:" + result.hasProcess + " require:" + result.hasRequire);
  check("the engine is Chrome/Edge", /Chrome|Edg/.test(result.ua), result.ua.slice(0, 60));
  check("DecompressionStream is present", result.hasDecompressionStream === "function");
  check("reads a deflated workbook", Array.isArray(result.sheets) && result.sheets.length === 4, JSON.stringify(result.sheets));
  check("sheetToJSON works", typeof result.firstCell === "string", JSON.stringify(result.firstCell));
  check("sheetToCSV works", typeof result.csvStart === "string" && result.csvStart.length > 0, JSON.stringify(result.csvStart));
  check("eachRow streams", result.streamed === 205, result.streamed + " rows");
  check("writes xlsx bytes", result.wroteIsUint8 && result.wroteBytes > 1000, result.wroteBytes + " bytes");
  check("reads back what it wrote in the same engine", Array.isArray(result.rereadSheets) && result.rereadSheets.length === 1, JSON.stringify(result.rereadSheets));
  check("the sync API throws a typed error, not a ReferenceError", result.syncCode === "NOT_SUPPORTED" && result.syncName === "PlatformError",
    "code=" + result.syncCode + " name=" + result.syncName);
} catch (err) {
  check("launches and runs in a real browser", false, String(err && err.message).split("\n")[0]);
} finally {
  if (browser) await browser.close();
}

done();
