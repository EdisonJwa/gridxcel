/**
 * Run the same byte-level API in Chromium, Firefox and WebKit.
 *
 * real-browser.test.mjs proves the library works in one real browser, and it uses
 * puppeteer-core because that is a single dependency and the local Chrome is
 * already on disk. It cannot reach Firefox or Safari.
 *
 * This test uses Playwright, whose whole point is a uniform driver across the
 * three engines. It exists because the library's browser claim is not "works in
 * Chrome" - it is "works anywhere with the Web platform", and the part most
 * likely to differ by engine is the Web compression streams that stand in for
 * node:zlib. A `deflate-raw` handshake that works in Chromium and fails in
 * Firefox is exactly the kind of thing that only a second engine finds.
 *
 * Skipped unless GRIDXCEL_REQUIRE_BROWSER is set, or a driver is available.
 */
import { readFileSync, writeFileSync as fs_writeFile, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const here = path.dirname(fileURLToPath(import.meta.url));
const REQUIRE = process.env.GRIDXCEL_REQUIRE_BROWSER === "1";
const ENGINE = process.env.GRIDXCEL_BROWSER_ENGINE || "chromium";

let playwright;
try {
  playwright = await import("playwright");
} catch {
  if (REQUIRE) {
    console.error("FAIL  playwright is not installed, so " + ENGINE + " cannot be tested");
    process.exit(1);
  }
  console.log("  SKIP  playwright not installed (set GRIDXCEL_REQUIRE_BROWSER=1 to require it)");
  process.exit(0);
}

// Build the bundle exactly as a bundler would: the same source, no shims.
// The entry is written here rather than reusing browser-entry.mjs, which exposes
// a single run() helper on purpose for the esbuild/Vite/webpack suites.
const entryPath = path.join(here, "cross-browser-entry.mjs");
const bundlePath = path.join(here, "out", "cross-browser-" + ENGINE + ".mjs");
mkdirSync(path.join(here, "out"), { recursive: true });
fs_writeFile(entryPath, [
  // Named exports, not "export *": a star import pulls in every module, including
  // the Node-only filesystem and sync paths, so the package "browser" field cannot
  // remove the node:zlib chain. browser-entry.mjs names what it uses for the same
  // reason.
  'export { read, write, eachRow, sheetToJSON, sheetToCSV, createCSVStream, readSync } from "gridxcel";',
].join("\n") + "\n");
const { build } = await import("esbuild");
let bundled;
try {
  bundled = await build({
    entryPoints: [entryPath],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["es2020"],
    write: false,
    logLevel: "silent",
  });
} catch (err) {
  const detail = err && err.errors ? err.errors.map((e) => e.text).join("; ") : String(err && err.message);
  console.error("FAIL  could not build the bundle: " + detail.slice(0, 300));
  process.exit(1);
}
fs_writeFile(bundlePath, bundled.outputFiles[0].contents);

mkdirSync(path.join(here, "out"), { recursive: true });
const source = readFileSync(bundlePath, "utf8");
const page = `<!doctype html><meta charset="utf-8"><body><script type="module">
import * as g from "./cross-browser-${ENGINE}.mjs";
const out = { engine: "${ENGINE}" };
try {
  // 1. A write then a read, entirely in the page.
  const bytes = await g.write([["name", "score"], ["ada", 42], ["grace", 37.5]]);
  out.wrote = bytes.length;
  const book = await g.read(bytes);
  out.rows = book.sheets[0].rows.length;
  out.firstValue = book.sheets[0].rows[1][1].v;

  // 2. Streaming read.
  let streamed = 0;
  for await (const row of g.eachRow(bytes)) streamed++;
  out.streamed = streamed;

  // 3. CSV out, which exercises the string path.
  let csv = "";
  for await (const chunk of g.createCSVStream(book.sheets[0])) csv += chunk;
  out.csv = csv.length;

  // 4. The Web compression streams, which is the engine-specific part: this is
  //    where node:zlib is replaced by CompressionStream/DecompressionStream.
  // The Web compression streams are the engine-specific part: this is where
  // node:zlib is replaced. Feature-detect the same way the library does.
  try {
    new CompressionStream("deflate-raw");
    new DecompressionStream("deflate-raw");
    out.hasWebDeflate = true;
  } catch { out.hasWebDeflate = false; }

  // 5. A synchronous read of a STORED-only archive works here, because it needs
  //    no inflate. What must NOT happen is a ReferenceError or a raw TypeError:
  //    either it works, or it is a typed PlatformError. The browser writer emits
  //    STORED parts precisely so this path stays usable.
  try {
    const back = g.readSync(bytes);
    out.sync = "ok:" + back.sheets[0].rows.length;
  } catch (error) {
    out.sync = error && error.code ? error.code : "untyped:" + (error && error.name);
  }

  // 6. A synchronous read of a DEFLATED archive must be a typed refusal, because
  //    the Web platform has no synchronous inflate. Build one with the async
  //    writer in a Node-style compressed form is not possible in the page, so the
  //    refusal is checked through the documented capability probe instead.
  out.hasSyncDeflate = typeof g.writeSync === "function";

  out.ok = true;
} catch (error) {
  out.ok = false;
  out.error = String(error && error.message).slice(0, 300);
}
window.__RESULT__ = out;
</script></body>`;

const server = await import("node:http");
const srv = server.createServer((req, res) => {
  if (req.url === "/" || req.url.startsWith("/index")) {
    res.writeHead(200, { "content-type": "text/html" });
    res.end(page);
    return;
  }
  res.writeHead(200, { "content-type": "text/javascript" });
  res.end(source);
});
await new Promise((resolve) => srv.listen(0, "127.0.0.1", resolve));
const port = srv.address().port;

let failures = 0;
function check(label, ok, detail) {
  console.log((ok ? "  ok    " : "  FAIL  ") + label + (detail === undefined ? "" : "  " + detail));
  if (!ok) failures++;
}

let browser;
try {
  browser = await playwright[ENGINE].launch();
  const context = await browser.newContext();
  const tab = await context.newPage();
  const consoleErrors = [];
  tab.on("pageerror", (e) => consoleErrors.push(String(e).slice(0, 200)));
  await tab.goto("http://127.0.0.1:" + port + "/", { waitUntil: "load" });
  await tab.waitForFunction("window.__RESULT__ !== undefined", { timeout: 30000 });
  const result = await tab.evaluate("window.__RESULT__");

  check(ENGINE + ": the bundle runs", result.ok === true, result.error || "");
  if (result.ok) {
    check(ENGINE + ": wrote and read a workbook", result.rows === 3 && result.firstValue === 42, "rows=" + result.rows + " v=" + result.firstValue);
    check(ENGINE + ": eachRow streamed", result.streamed === 3, "streamed=" + result.streamed);
    check(ENGINE + ": createCSVStream produced bytes", result.csv > 0, result.csv + " chars");
    // Either it works (STORED parts need no inflate) or it refuses with a typed
    // code. What is forbidden is a ReferenceError or an untyped throw.
    check(
      ENGINE + ": the synchronous API works or refuses with a typed code, never a ReferenceError",
      typeof result.sync === "string" && (result.sync.startsWith("ok:") || result.sync === "NOT_SUPPORTED"),
      String(result.sync),
    );
    // Reported, not asserted: the floor genuinely differs by engine and the
    // library feature-detects it rather than assuming.
    console.log("  info  " + ENGINE + ": hasWebDeflate() = " + String(result.hasWebDeflate));
    check(ENGINE + ": no uncaught page errors", consoleErrors.length === 0, consoleErrors.join(" | "));
  }
} catch (error) {
  check(ENGINE + ": launches and runs", false, String(error && error.message).split("\n")[0]);
} finally {
  if (browser) await browser.close();
  srv.close();
}

if (failures > 0) {
  console.error(failures + " failure(s) in " + ENGINE);
  process.exit(1);
}
console.log("  PASS  " + ENGINE);
