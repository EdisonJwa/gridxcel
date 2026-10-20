/**
 * Run every numbered example and report which ones fail.
 *
 * Shows: nothing about the library. This is the harness that proves the other
 * files in this directory still run, so the examples cannot rot silently.
 *
 * Run from anywhere:
 *   node examples/run-all.mjs
 */
import { spawnSync } from "node:child_process";
import { readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

const HERE = fileURLToPath(new URL("./", import.meta.url));
const files = readdirSync(HERE).filter((f) => /^\d\d-.*\.mjs$/.test(f)).sort();

if (files.length === 0) {
  console.error("no examples found in " + HERE);
  process.exit(1);
}

let failed = 0;
let slowest = { file: null, ms: 0 };

for (const file of files) {
  const started = Date.now();
  // stdio pipes rather than inherit, so a passing example stays quiet.
  const run = spawnSync(process.execPath, [HERE + file], { encoding: "utf8", timeout: 300000 });
  const ms = Date.now() - started;
  if (ms > slowest.ms) slowest = { file, ms };
  if (run.status === 0) {
    console.log("ok    " + file.padEnd(30) + (ms + " ms").padStart(9));
  } else {
    failed++;
    console.log("FAIL  " + file.padEnd(30) + (ms + " ms").padStart(9));
    console.log((run.stderr || run.stdout || "").split("\n").slice(0, 12).map((l) => "      " + l).join("\n"));
  }
}

console.log("");
console.log(files.length - failed + "/" + files.length + " examples passed; slowest was " + slowest.file + " at " + slowest.ms + " ms");
process.exit(failed === 0 ? 0 : 1);
