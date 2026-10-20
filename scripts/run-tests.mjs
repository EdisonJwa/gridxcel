#!/usr/bin/env node
/**
 * Portable test runner.
 *
 * Why this exists instead of putting a glob in package.json: passing the test
 * runner a recursive glob only works from Node 21, where it learned to expand
 * globs. On Node 18 and 20 that string is treated as a literal path and the run
 * fails, even though package.json declares "node": ">=18". Letting the shell
 * expand it fixes POSIX but not Windows, where npm runs scripts through cmd.exe
 * and nothing expands the pattern.
 *
 * So the file list is built here, in Node, and passed explicitly. That behaves
 * identically on every supported version and every platform, which lets CI run
 * the same "npm test" everywhere instead of special-casing older Node.
 *
 * Usage: node scripts/run-tests.mjs [extra node --test flags...]
 */
import { readdirSync, statSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

/** Every test file under test/, at any depth, in a stable order. */
function collect(dir, out = []) {
  for (const entry of readdirSync(dir).sort()) {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) collect(full, out);
    else if (entry.endsWith(".test.mjs")) out.push(full);
  }
  return out;
}

const files = collect(join(root, "test"));

if (files.length === 0) {
  console.error("no test files found under test/");
  process.exit(1);
}

const args = ["--test", ...process.argv.slice(2), ...files];
const result = spawnSync(process.execPath, args, { cwd: root, stdio: "inherit" });

if (result.error) {
  console.error("failed to start the test runner:", result.error.message);
  process.exit(1);
}
process.exit(result.status === null ? 1 : result.status);
