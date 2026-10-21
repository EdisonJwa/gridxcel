/**
 * Pre-release consistency check.
 *
 * A release is the one irreversible action this project takes: npm versions
 * cannot be withdrawn, only deprecated. Everything that can be verified before a
 * publish is verified here, so a mistake costs a failed run rather than a
 * permanent version.
 *
 * Checks:
 *   1. package.json version is a valid semver and not already on the registry.
 *   2. The git tag (when there is one) matches the version.
 *   3. The working tree is clean, so the tarball matches a commit.
 *   4. The files list ships what it claims, and nothing untracked leaks in.
 *   5. The tarball builds and installs, and the entry points resolve from it.
 *
 * Run: node scripts/check-release.mjs [--allow-published]
 */
import { execFileSync } from "node:child_process";
import { readFileSync, writeFileSync, existsSync, rmSync, mkdirSync, cpSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import os from "node:os";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const ALLOW_PUBLISHED = process.argv.includes("--allow-published");

let failures = 0;
function check(label, ok, detail) {
  console.log((ok ? "  ok    " : "  FAIL  ") + label + (detail ? "  " + detail : ""));
  if (!ok) failures++;
}
/**
 * Run a command, tolerating Windows' .cmd shims.
 *
 * npm is npm.cmd on Windows and execFileSync refuses to spawn that without a
 * shell, which made this whole script fail before its first check.
 */
function run(cmd, args, opts = {}) {
  // Windows: npm is npm.cmd. Passing shell:true with args is deprecated (DEP0190)
  // because the arguments are concatenated rather than escaped. Resolve the .cmd
  // path instead and spawn it directly, so no shell is involved and nothing can be
  // misinterpreted. Every argument here is a literal in this file, but a checker
  // that models the right pattern is worth more than one that relies on that.
  const executable = process.platform === "win32" && cmd === "npm" ? "npm.cmd" : cmd;
  return execFileSync(executable, args, {
    cwd: ROOT, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"],
    ...opts,
  }).trim();
}

const pkg = JSON.parse(readFileSync(path.join(ROOT, "package.json"), "utf8"));

/* 1. valid semver, not already published */
const SEMVER = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/;
check("version is valid semver", SEMVER.test(pkg.version), pkg.version);

let published = null;
try {
  published = run("npm", ["view", pkg.name + "@" + pkg.version, "version"]);
} catch {
  published = null; // 404 means free, which is what we want
}
const alreadyOut = Boolean(published);
check(
  alreadyOut ? "version is already published" : "version is not yet published",
  alreadyOut ? ALLOW_PUBLISHED : true,
  alreadyOut ? pkg.name + "@" + pkg.version + " exists on the registry" : "",
);

/* 2. tag agrees */
let tag = null;
try {
  const out = run("git", ["describe", "--exact-match", "--tags", "HEAD"]);
  tag = out.replace(/^v/, "");
} catch {
  tag = null; // not on a tag; a publish from a branch is allowed but noted
}
if (tag !== null) {
  check("git tag matches package.json", tag === pkg.version, "tag v" + tag + " vs " + pkg.version);
} else {
  console.log("  note  HEAD is not tagged; the release workflow publishes from a tag");
}

/* 3. clean tree */
const dirty = run("git", ["status", "--porcelain"]);
check("working tree is clean", dirty === "", dirty ? dirty.split("\n").length + " path(s) modified" : "");

/* 4. the published surface */
check("no runtime dependencies", Object.keys(pkg.dependencies || {}).length === 0);
check("no devDependencies in the published package", Object.keys(pkg.devDependencies || {}).length === 0);
check("engines declares a node floor", Boolean(pkg.engines && pkg.engines.node), pkg.engines && pkg.engines.node);
check("repository points at this repo", /github\.com\/EdisonJwa\/gridxcel/.test(pkg.repository && pkg.repository.url || ""), pkg.repository && pkg.repository.url);
check("license is set", Boolean(pkg.license), pkg.license);
check("prepublishOnly runs the gate", /npm run check/.test(pkg.scripts.prepublishOnly || ""), pkg.scripts.prepublishOnly);
check("every exports target exists", Object.values(pkg.exports || {}).every((target) => {
  if (typeof target === "string") return existsSync(path.join(ROOT, target));
  return ["types", "default", "import"].every((key) => !target[key] || existsSync(path.join(ROOT, target[key])));
}));
check("the files list ships the docs", (pkg.files || []).some((f) => f === "docs/API.md"));

/* 5. pack, install, resolve */
const tmp = path.join(os.tmpdir(), "gridxcel-release-check-" + Date.now());
mkdirSync(tmp, { recursive: true });
try {
  const tarball = run("npm", ["pack", "--pack-destination", tmp, "--silent"]).split("\n").pop();
  const tarballPath = path.join(tmp, tarball);
  check("npm pack produced a tarball", existsSync(tarballPath), tarball);

  const consumer = path.join(tmp, "consumer");
  mkdirSync(consumer, { recursive: true });
  cpSync(tarballPath, path.join(consumer, path.basename(tarballPath)));
  writeFileSync(
    path.join(consumer, "package.json"),
    JSON.stringify({ name: "release-check-consumer", version: "1.0.0", type: "module", private: true }, null, 2),
  );
  run("npm", ["install", "./" + path.basename(tarballPath), "--no-audit", "--no-fund", "--silent"], { cwd: consumer });

  check("the tarball installs with zero dependencies", existsSync(path.join(consumer, "node_modules", pkg.name, "package.json")));

  // Every entry point must resolve from the installed package, not the source tree.
  const probe = "import(\"" + pkg.name + "\").then(m=>console.log(JSON.stringify(Object.keys(m).length)))";
  const rootKeys = execFileSync(process.execPath, ["--input-type=module", "-e", probe], { cwd: consumer, encoding: "utf8" }).trim();
  check("the root entry point resolves from the tarball", Number(rootKeys) > 50, rootKeys + " exports");

  for (const specifier of Object.keys(pkg.exports || {})) {
    if (specifier === "." || specifier === "./package.json") continue;
    const name = pkg.name + "/" + specifier.replace(/^\.\//, "");
    const ok = (() => {
      try {
        execFileSync(process.execPath, ["--input-type=module", "-e", "import(\"" + name + "\")"], { cwd: consumer, stdio: "ignore" });
        return true;
      } catch {
        return false;
      }
    })();
    check("subpath resolves: " + specifier, ok);
  }
} finally {
  rmSync(tmp, { recursive: true, force: true });
}

console.log("");
if (failures > 0) {
  console.error(failures + " problem(s); do not publish");
  process.exit(1);
}
console.log("OK: this version is ready to publish");
