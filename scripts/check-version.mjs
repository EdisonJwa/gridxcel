/**
 * VERSION must match package.json.
 *
 * The library exports VERSION as a literal, because reading package.json at runtime
 * needs require or an import assertion and this module is on the browser path. A
 * literal drifts: bumping the manifest and forgetting src/api.js is silent, and the
 * only thing that noticed was one documentation example that asserted the two agree
 * - which failed with "Expected values to be strictly equal" and no mention of the
 * version being the cause.
 *
 * Callers use VERSION to report which build they are running. A number that
 * disagrees with the published package is worse than no number at all.
 *
 * Run: node scripts/check-version.mjs [--fix]
 */
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const FIX = process.argv.includes("--fix");

const pkg = JSON.parse(readFileSync(path.join(ROOT, "package.json"), "utf8"));
const API = path.join(ROOT, "src", "api.js");
const source = readFileSync(API, "utf8");

const match = source.match(/export const VERSION = "([^"]*)";/);
if (!match) {
  console.error("FAIL  src/api.js no longer declares VERSION in a recognised form");
  process.exit(1);
}

const declared = match[1];
if (declared === pkg.version) {
  console.log("OK: VERSION matches package.json (" + declared + ")");
  process.exit(0);
}

if (FIX) {
  writeFileSync(API, source.replace(match[0], 'export const VERSION = "' + pkg.version + '";'));
  console.log("fixed src/api.js: " + declared + " -> " + pkg.version);
  process.exit(0);
}

console.error("FAIL  src/api.js declares VERSION " + declared + " but package.json says " + pkg.version);
console.error("      Run with --fix, or set them both by hand. A published package whose");
console.error("      VERSION disagrees with its manifest reports the wrong build to every caller.");
process.exit(1);
