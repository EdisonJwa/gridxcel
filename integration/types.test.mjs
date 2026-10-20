/**
 * Are the PUBLISHED types actually valid?
 *
 * Nothing in this repository checked that. test/types.smoke.mjs regex-scans
 * src/index.d.ts for symbol names, and scripts/check-docs.mjs loads src/*.js -
 * neither compiles the declarations. A type literal in a value position in the
 * default export, a missing declaration file, three exports with no declaration
 * at all, and a hard dependency on @types/node all survived that gap, and every
 * TypeScript consumer saw them.
 *
 * This compiles a real consumer against the package with skipLibCheck: false
 * and NO @types/node, which is what a browser or bundler consumer has.
 *
 * Run: node types.test.mjs
 */

import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
let failures = 0;
const check = (name, ok, detail) => {
  console.log((ok ? "  ok    " : "  FAIL  ") + name + (detail ? "  " + detail : ""));
  if (!ok) failures++;
};

// Run the compiler through Node rather than through node_modules/.bin. On
// Windows that is a .cmd shim, which spawnSync refuses to execute without a
// shell on Node 20+ - it returns a null status and no output, so the check
// failed with an empty message while the compiler had never run at all.
const tsc = join(here, "node_modules", "typescript", "lib", "tsc.js");
const r = spawnSync(process.execPath, [tsc, "--project", join(here, "types", "tsconfig.json")], { cwd: here, encoding: "utf8" });
const output = ((r.stdout || "") + (r.stderr || "")).trim();
check("the compiler ran", r.status !== null && !r.error, r.error ? String(r.error.message) : "exit " + r.status);
check("a consumer compiles against the published types", r.status === 0, output.split("\n").filter(Boolean).slice(0, 5).join(" | "));
check("with skipLibCheck disabled", !/skipLibCheck/.test(output));
check("and without @types/node", !/Cannot find name 'Buffer'|Cannot find namespace 'NodeJS'/.test(output));

/* 4. Compiling is not matching. A declaration can be perfectly valid and still
      describe something that is not there, so compare the two directly: every
      name the declarations export from the root must exist at runtime, and the
      default export's members must match the runtime object exactly. */
const { readFileSync } = await import("node:fs");
const gx = await import("gridxcel");
const dts = readFileSync(join(here, "..", "src", "index.d.ts"), "utf8");
// A .d.ts declares two different kinds of thing, and only one of them can be
// checked against the runtime. Types are erased, so they must NOT be looked for;
// values must be.
const declaredValues = new Set();
const declaredTypes = new Set();
for (const m of dts.matchAll(/^export (?:declare )?(?:async )?(?:function|class|const|let|var)\s+(\w+)/gm)) declaredValues.add(m[1]);
for (const m of dts.matchAll(/^export (?:interface|type)\s+(\w+)/gm)) declaredTypes.add(m[1]);
for (const m of dts.matchAll(/^export (?:type )?\{([^}]*)\}/gm)) {
  for (const n of m[1].split(",")) {
    const t = n.trim().split(/\s+as\s+/).pop().trim();
    if (!t) continue;
    // "export type { X }" is a type; a bare re-export is a value.
    (/^export type/.test(m[0]) ? declaredTypes : declaredValues).add(t);
  }
}

const runtime = new Set(Object.keys(gx));
const absent = [...declaredValues].filter((n) => !runtime.has(n));
check("every declared root value exists at runtime", absent.length === 0, absent.length ? absent.join(", ") : declaredValues.size + " values checked");
const declared = new Set([...declaredValues, ...declaredTypes]);

const missing = [...runtime].filter((n) => n !== "default" && !declared.has(n));
check("every runtime root export is declared", missing.length === 0, missing.length ? missing.join(", ") : runtime.size + " exports checked");

const defaultKeys = Object.keys(gx.default);
const undeclaredDefault = defaultKeys.filter((k) => !dts.includes(k + ": typeof"));
check("the default export members are declared", undeclaredDefault.length === 0, undeclaredDefault.join(", "));

console.log("\n  " + (failures === 0 ? "PASS" : failures + " failure(s)"));
process.exit(failures === 0 ? 0 : 1);
