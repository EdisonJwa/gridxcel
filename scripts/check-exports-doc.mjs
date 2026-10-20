#!/usr/bin/env node
/**
 * Every public export must be documented somewhere.
 *
 * scripts/check-docs.mjs validates the forward direction: a symbol that appears
 * in a document must exist in src/. It cannot catch the opposite and more common
 * failure - a public export that appears in no document at all. An audit found
 * 47 of 115 exports in that state.
 *
 * This is the inverse check. It enumerates the real exports of each entry point
 * and asserts each one is named in the documentation set.
 *
 * A symbol is considered documented when its name appears in a document outside
 * of a fenced code block, so that merely using a function in an example does not
 * count as documenting it.
 *
 * Usage: node scripts/check-exports-doc.mjs [--json]
 *   --json   machine-readable output
 * Exit code 1 when any export is undocumented.
 */

import { readFileSync, readdirSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const srcDir = join(root, "src");
const asJson = process.argv.includes("--json");

/**
 * Entry point -> source file, derived from package.json "exports".
 *
 * This used to be a hand-written table, which meant adding an entry point to the
 * package did not add it to this check - so its exports were never verified as
 * documented. Reading the manifest removes that failure mode.
 */
const ENTRIES = (() => {
  const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
  const out = [];
  for (const [specifier, target] of Object.entries(pkg.exports || {})) {
    if (specifier === "./package.json" || typeof target === "string") continue;
    const file = String(target.default || target.import || "").replace(/^\.\/src\//, "").replace(/^\.\//, "");
    if (!file) continue;
    const name = specifier === "." ? "gridxcel" : "gridxcel/" + specifier.replace(/^\.\//, "");
    out.push([name, file]);
  }
  return out;
})();

/**
 * Names that are exported for advanced use but are deliberately not part of the
 * documented surface. Every entry needs a reason; an empty list is preferred.
 */
const ALLOWED_UNDOCUMENTED = new Set([
  // Low-level ZIP signature constants: meaningless without the ZIP spec open,
  // and documented collectively under gridxcel/zip.
  "SIG_LOCAL", "SIG_CENTRAL", "SIG_EOCD", "SIG_EOCD64", "SIG_EOCD64_LOC", "SIG_DATA_DESC",
]);

function exportsOf(file) {
  const text = readFileSync(join(srcDir, file), "utf8");
  const names = new Set();
  for (const m of text.matchAll(/^export\s+(?:async\s+)?(?:function\*?|class|const|let|var)\s+([A-Za-z_$][\w$]*)/gm)) {
    names.add(m[1]);
  }
  for (const m of text.matchAll(/^export\s*\{([^}]*)\}/gm)) {
    for (const part of m[1].split(",")) {
      const name = part.trim().split(/\s+as\s+/).pop();
      if (name) names.add(name);
    }
  }
  return [...names].sort();
}

/** Documents that count as user-facing documentation. */
function docFiles() {
  const files = ["README.md"];
  const docsDir = join(root, "docs");
  if (existsSync(docsDir)) {
    for (const f of readdirSync(docsDir)) {
      if (f.endsWith(".md")) files.push("docs/" + f);
    }
  }
  return files;
}

/** Strip fenced code blocks so usage in an example is not mistaken for docs. */
function proseOf(text) {
  const out = [];
  let fence = null;
  for (const line of text.split("\n")) {
    const m = line.match(/^ {0,3}(`{3,}|~{3,})(.*)$/);
    if (fence === null && m) { fence = m[1][0]; continue; }
    if (fence !== null && m && m[1][0] === fence) { fence = null; continue; }
    if (fence === null) out.push(line);
  }
  return out.join("\n");
}

const prose = docFiles()
  .filter((f) => existsSync(join(root, f)))
  .map((f) => proseOf(readFileSync(join(root, f), "utf8")))
  .join("\n");

const report = [];
let total = 0;
let missing = 0;

for (const [specifier, file] of ENTRIES) {
  const names = exportsOf(file);
  const undoc = names.filter((n) => {
    if (ALLOWED_UNDOCUMENTED.has(n)) return false;
    return !new RegExp("\\b" + n.replace(/\$/g, "\\$") + "\\b").test(prose);
  });
  total += names.length;
  missing += undoc.length;
  report.push({ specifier, file, total: names.length, undocumented: undoc });
}

if (asJson) {
  console.log(JSON.stringify({ total, missing, entries: report }, null, 2));
} else {
  for (const r of report) {
    const covered = r.total - r.undocumented.length;
    console.log("  " + r.specifier.padEnd(18) + covered + "/" + r.total + " documented"
      + (r.undocumented.length ? "   missing: " + r.undocumented.join(", ") : ""));
  }
  console.log("");
  console.log("  public exports documented : " + (total - missing) + "/" + total);
}

if (missing > 0) {
  if (!asJson) console.log("\nFAIL: " + missing + " public export(s) appear in no document.");
  process.exit(1);
}
if (!asJson) console.log("\nOK: every public export is documented.");
