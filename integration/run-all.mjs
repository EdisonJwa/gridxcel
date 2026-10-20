#!/usr/bin/env node
/** Run every integration check and summarise. */
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const SUITES = [
  ["browser bundle (esbuild + no-Node sandbox)", "browser-bundle.test.mjs"],
  ["REAL Chrome via puppeteer-core", "real-browser.test.mjs"],
  ["vite production build", "vite.test.mjs"],
  ["rolldown (the Rust bundler in Vite 8)", "rolldown.test.mjs"],
  ["webpack 5", "webpack.test.mjs"],
  ["react / preact / structuredClone", "react.test.mjs"],
  ["published types compile for a consumer", "types.test.mjs"],
];

// The suites import the published specifier "gridxcel", so it has to resolve.
// Idempotent: a no-op when node_modules already has it.
const link = spawnSync(process.execPath, [join(here, "link-self.mjs")], { cwd: here, encoding: "utf8" });
if (link.status !== 0) {
  console.error("could not link gridxcel into the harness:\n" + (link.stderr || ""));
  process.exit(1);
}

let failed = 0;
for (const [label, file] of SUITES) {
  console.log("\n=== " + label + " ===");
  const r = spawnSync(process.execPath, [join(here, file)], { cwd: here, stdio: "inherit" });
  if (r.status !== 0) failed++;
}
console.log("\n" + (failed === 0 ? "all integration suites passed" : failed + " suite(s) failed"));
process.exit(failed === 0 ? 0 : 1);
