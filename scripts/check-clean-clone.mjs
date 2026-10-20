/**
 * Run the documentation checks against a checkout of TRACKED FILES ONLY.
 *
 * A doc example that reads a fixture from `bench/` passes on a developer machine
 * and fails in CI, because `bench/` is untracked and excluded from the tarball.
 * That happened: the rewritten README's first example read
 * `bench/fair-50000x8.xlsx`, went green locally, and failed all eight CI legs on
 * the Documentation step.
 *
 * A local check cannot catch that class of bug by definition - the file it needs
 * is present. This builds the one environment where it is absent, so the check
 * can fail before CI does.
 *
 * Run: node scripts/check-clean-clone.mjs
 */
import { cpSync, mkdirSync, rmSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const TMP = path.join(ROOT, "..", ".clean-clone-check");

function fail(message) {
  console.error("FAIL: " + message);
  rmSync(TMP, { recursive: true, force: true });
  process.exit(1);
}

const listed = spawnSync("git", ["ls-files"], { cwd: ROOT, encoding: "utf8", maxBuffer: 1 << 26 });
if (listed.status !== 0) fail("git ls-files failed - is this a git checkout?");

const files = listed.stdout.split("\n").filter(Boolean);
if (files.length < 50) fail("only " + files.length + " tracked files; refusing to treat that as a clone");

rmSync(TMP, { recursive: true, force: true });
for (const file of files) {
  const dest = path.join(TMP, file);
  mkdirSync(path.dirname(dest), { recursive: true });
  cpSync(path.join(ROOT, file), dest);
}

console.log("clean clone: " + files.length + " tracked files, no bench/, no node_modules/");

const result = spawnSync(process.execPath, ["scripts/check-docs.mjs"], { cwd: TMP, encoding: "utf8", maxBuffer: 1 << 26 });
const output = (result.stdout || "") + (result.stderr || "");
rmSync(TMP, { recursive: true, force: true });

if (result.status !== 0) {
  const problems = output.split("\n").filter((line) => /problem|failed|not found|Error/.test(line));
  fail("the documentation checks fail without the untracked files:\n  " + (problems.join("\n  ") || output.slice(-800)));
}

console.log("OK: every documented symbol and example works from a tracked-files-only checkout.");
