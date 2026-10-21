/**
 * Keep the test count in the docs honest.
 *
 * The number of tests is quoted in four documents. It has drifted three times -
 * 137, then 225, then 262, then 303 - and each time the number was wrong for
 * weeks before a reader noticed. A number nobody re-measures is a number that is
 * eventually false.
 *
 * This runs the suite, reads its own report, and fails when a document disagrees.
 * The docs are the source of the claim; this is the source of the number.
 *
 * Run: node scripts/check-counts.mjs [--fix]
 */
import { execFileSync } from "node:child_process";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const FIX = process.argv.includes("--fix");

let out = "";
let runError = null;
try {
  out = execFileSync(process.execPath, [path.join(ROOT, "scripts", "run-tests.mjs")], {
    cwd: ROOT, encoding: "utf8", maxBuffer: 1 << 26, stdio: ["ignore", "pipe", "pipe"],
  });
} catch (err) {
  runError = err;
  out = (err && (err.stdout || "")) || "";
  console.error("the test runner exited with status " + (err && err.status) + ": " + String((err && err.stderr) || (err && err.message) || err).split("\n").slice(-4).join(" | "));
}
const counts = {};
for (const m of out.matchAll(/ℹ (tests|pass|fail|skipped) (\d+)/g)) counts[m[1]] = Number(m[2]);
if (!counts.tests) {
  console.error("FAIL  the test runner reported no counts; cannot verify the documents");
  console.error("       captured " + out.length + " chars from the runner");
  if (runError) console.error("       runner error: " + String(runError.message || runError).split("\n")[0]);
  const tail = out.trim().split("\n").slice(-6).join("\n       ");
  if (tail) console.error("       last output:\n       " + tail);
  process.exit(1);
}
console.log("suite reports: " + counts.tests + " tests, " + counts.pass + " pass, " + counts.fail + " fail, " + counts.skipped + " skipped");

/** Every document that quotes the count, and the shapes it appears in. */
const CLAIMS = [
  ["README.md", /\| Tests \| (\d+) run, (\d+) pass, (\d+) fail, (\d+) skipped \|/,
    (m) => "| Tests | " + counts.tests + " run, " + counts.pass + " pass, " + counts.fail + " fail, " + counts.skipped + " skipped |"],
  ["docs/ARCHITECTURE.md", /\| Unit and integration tests \| `npm test` \| (\d+) tests: (\d+) pass, (\d+) skipped, (\d+) fail\. \|/,
    () => "| Unit and integration tests | `npm test` | " + counts.tests + " tests: " + counts.pass + " pass, " + counts.skipped + " skipped, " + counts.fail + " fail. |"],
  ["docs/LIMITATIONS.md", /`npm test` runs (\d+) tests \((\d+) pass, (\d+) skipped\)/,
    () => "`npm test` runs " + counts.tests + " tests (" + counts.pass + " pass, " + counts.skipped + " skipped)"],
  ["CONTRIBUTING.md", /As last measured, (\d+) tests, (\d+) pass, (\d+) fail, (\d+) skipped/,
    () => "As last measured, " + counts.tests + " tests, " + counts.pass + " pass, " + counts.fail + " fail, " + counts.skipped + " skipped"],
];

let failures = 0;
for (const [file, pattern, replacement] of CLAIMS) {
  const full = path.join(ROOT, file);
  if (!existsSync(full)) { console.log("  FAIL  " + file + " is missing"); failures++; continue; }
  const text = readFileSync(full, "utf8");
  const match = text.match(pattern);
  if (!match) {
    console.log("  FAIL  " + file + " no longer states the count in a recognised form");
    failures++;
    continue;
  }
  const stated = Number(match[1]);
  if (stated === counts.tests) {
    console.log("  ok    " + file + " states " + stated);
    continue;
  }
  if (FIX) {
    writeFileSync(full, text.replace(pattern, replacement()));
    console.log("  fixed " + file + ": " + stated + " -> " + counts.tests);
    continue;
  }
  console.log("  FAIL  " + file + " states " + stated + ", the suite has " + counts.tests);
  failures++;
}

if (failures > 0) {
  console.error("\n" + failures + " document(s) disagree with the suite. Run with --fix.");
  process.exit(1);
}
console.log("OK: every document states the current test count");
