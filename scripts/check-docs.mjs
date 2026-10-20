#!/usr/bin/env node
/**
 * Documentation verifier.
 *
 * Docs rot silently, and a wrong example is worse than no example. This tool
 * ties every claim in the markdown to the implementation:
 *
 *   1. SYMBOLS  - every inline-code identifier that looks like a public API
 *                 name must be exported by src/, or appear in the source
 *                 somewhere (properties, option keys, internal helpers).
 *   2. EXAMPLES - every fenced block whose info string contains "doc-example"
 *                 is executed as a real ES module against the package.
 *
 * Run: node scripts/check-docs.mjs
 * Exit 0 = consistent. Exit 1 = drift, reported per file.
 */

import { readFileSync, readdirSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const TMP = join(ROOT, "bench", "doc-examples");
const BT = String.fromCharCode(96);
const FENCE = BT + BT + BT;
const QUOTE = BT;

const MODULES = ["index.js", "compat.js", "csv.js", "json.js", "numfmt.js", "zip.js", "xml.js", "address.js", "errors.js", "workbook-writer.js", "style-registry.js", "read.js", "write.js"];

async function apiSurface() {
  const names = new Set();
  for (const spec of MODULES) {
    const m = await import(new URL("../src/" + spec, import.meta.url).href);
    for (const k of Object.keys(m)) if (k !== "default") names.add(k);
  }
  return names;
}

const DOC_FILES = [
  "README.md",
  "docs/QUICKSTART.md",
  "docs/API.md",
  "docs/COOKBOOK.md",
  "docs/MIGRATION.md",
  "docs/ARCHITECTURE.md",
  "docs/LIMITATIONS.md",
  "docs/PERFORMANCE.md",
  "docs/COMPARISON.md",
  "CHANGELOG.md",
  "CONTRIBUTING.md",
  "SECURITY.md",
  "CODE_OF_CONDUCT.md",
  "examples/README.md",
  "test/README.md",
];

const readDoc = (rel) => (existsSync(join(ROOT, rel)) ? readFileSync(join(ROOT, rel), "utf8") : null);

/**
 * Pull fenced blocks whose info string mentions doc-example.
 *
 * A fence only opens or closes at the START of a line (up to three spaces of
 * indentation). Matching on the trimmed line instead would treat an inline
 * fence written inside a sentence - which documentation does constantly - as a
 * real one, and swallow every example that followed it.
 */
function extractExamples(md, rel) {
  const out = [];
  const lines = md.split("\n");
  let open = null;
  for (let i = 0; i < lines.length; i++) {
    const m = /^ {0,3}(`{3,})([^`]*)$/.exec(lines[i]);
    if (!m) {
      if (open) open.buf.push(lines[i]);
      continue;
    }
    if (open === null) {
      open = { info: m[2].trim(), start: i + 2, fence: m[1].length, buf: [] };
      continue;
    }
    if (m[1].length >= open.fence && m[2].trim() === "") {
      if (/\bdoc-example\b/.test(open.info)) out.push({ rel, line: open.start, code: open.buf.join("\n"), info: open.info });
      open = null;
      continue;
    }
    open.buf.push(lines[i]);
  }
  return out;
}

/** Inline-code identifiers that look like API names. */
function symbolCandidates(md) {
  const found = new Map();
  const re = new RegExp("(?<![A-Za-z0-9_$." + BT + "/])\\" + BT + "([A-Za-z_$][A-Za-z0-9_$]*)\\" + BT, "g");
  let m;
  while ((m = re.exec(md))) if (!found.has(m[1])) found.set(m[1], true);
  return [...found.keys()];
}

const ALLOW = new Set(
  existsSync(join(HERE, "docs-allow.json"))
    ? JSON.parse(readFileSync(join(HERE, "docs-allow.json"), "utf8")).allow
    : [],
);

const api = await apiSurface();
// .d.ts files count as src. A type-only export has no runtime presence, so
// scanning only .js made every documented type name look like a symbol that
// does not exist - ByteStream, declared in index.d.ts and zip.d.ts and nothing
// else, failed the check the moment a document mentioned it.
const srcBlob = readdirSync(join(ROOT, "src"))
  .filter((f) => f.endsWith(".js") || f.endsWith(".d.ts"))
  .map((f) => readFileSync(join(ROOT, "src", f), "utf8"))
  .join("\n");

const problems = [];
const notes = [];
let linkCount = 0;
let exampleCount = 0;

for (const rel of DOC_FILES) {
  const md = readDoc(rel);
  if (md === null) { notes.push("absent, skipped: " + rel); continue; }

  const unknown = symbolCandidates(md).filter((n) => !api.has(n) && !srcBlob.includes(n) && !ALLOW.has(n));
  if (unknown.length) problems.push({ rel, kind: "symbol not found in src", detail: unknown.join(", ") });

  checkLinks(md, rel);

  const examples = extractExamples(md, rel);
  if (examples.length) {
    mkdirSync(TMP, { recursive: true });
    examples.forEach((ex, i) => {
      exampleCount++;
      const file = join(TMP, "ex-" + rel.replace(/[^a-z0-9]/gi, "_") + "-" + i + ".mjs");
      writeFileSync(file, ex.code);
      try {
        const out = execFileSync(process.execPath, [file], { cwd: TMP, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 120000 });
        const last = out.trim().split("\n").filter(Boolean).pop();
        notes.push("example ok   " + rel + ":" + ex.line + (last ? "  -> " + last.slice(0, 64) : ""));
      } catch (e) {
        problems.push({ rel, kind: "example failed", detail: rel + ":" + ex.line + "\n" + String(e.stderr || e.message).split("\n").slice(0, 10).join("\n") });
      }
    });
  }
}

/* --------------------------------------------------------------- 4. links */

/** Strip fenced blocks and code spans so link-like text inside code is ignored. */
function proseOnly(md) {
  const out = [];
  let inFence = false;
  const fenceRe = new RegExp("^ {0,3}" + FENCE);
  const spanRe = new RegExp(QUOTE + "[^" + QUOTE + "]*" + QUOTE, "g");
  for (const line of md.split("\n")) {
    if (fenceRe.test(line)) { inFence = !inFence; continue; }
    if (inFence) continue;
    out.push(line.replace(spanRe, ""));
  }
  return out.join("\n");
}

/** Local file links must resolve; in-page anchors must match a heading. */
function checkLinks(md, rel) {
  const prose = proseOnly(md);
  linkCount += (prose.match(/\]\([^)\s]+\)/g) || []).length;
  for (const m of prose.matchAll(/\]\(([^)#\s]+)(#[^)\s]*)?\)/g)) {
    const target = m[1];
    if (/^[a-z][a-z0-9+.-]*:/i.test(target)) continue;
    if (!existsSync(resolve(dirname(join(ROOT, rel)), target))) {
      problems.push({ rel, kind: "broken link", detail: target });
    }
  }
  const headings = new Set(
    md.split("\n").filter((l) => /^#{1,6} /.test(l))
      .map((l) => l.replace(/^#{1,6} /, "").trim().toLowerCase().replace(/[^a-z0-9 -]/g, "").replace(/ +/g, "-")),
  );
  for (const m of prose.matchAll(/\]\(#([^)\s]+)\)/g)) {
    if (!headings.has(m[1])) problems.push({ rel, kind: "unresolved anchor", detail: "#" + m[1] });
  }
}

console.log("gridxcel docs check");
console.log("  public symbols known : " + api.size);
console.log("  doc examples run     : " + exampleCount);
console.log("  local links checked  : " + linkCount);
for (const n of notes) console.log("  " + n);
if (problems.length === 0) {
  console.log("");
  console.log("OK: documented symbols exist and every doc-example runs.");
  process.exit(0);
}
console.log("");
console.log(problems.length + " problem(s):");
for (const p of problems) console.log("  [" + p.kind + "] " + p.rel + "\n    " + p.detail.split("\n").join("\n    "));
process.exit(1);
