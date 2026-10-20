/**
 * Make "gridxcel" resolvable from this harness.
 *
 * The integration tests import the PUBLISHED specifier on purpose - that is what
 * a consumer writes, and it is what checks the exports map, the types and the
 * browser field. It only resolves if node_modules contains the package, which is
 * true on a developer machine with a self-link and false on a fresh clone. This
 * creates the link, so the suite runs anywhere.
 *
 * Run by `npm run link`, and by the CI workflow before the integration step.
 */
import { symlinkSync, mkdirSync, existsSync, realpathSync, rmSync, lstatSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, "..");
const target = path.join(here, "node_modules", "gridxcel");

if (!existsSync(path.join(pkgRoot, "package.json"))) {
  console.error("link-self: " + pkgRoot + " does not look like the package root");
  process.exit(1);
}

mkdirSync(path.join(here, "node_modules"), { recursive: true });

let already = false;
try {
  const st = lstatSync(target);
  if (st.isSymbolicLink()) {
    already = realpathSync(target) === realpathSync(pkgRoot);
  } else {
    already = true; // a real install; leave it alone
  }
} catch {
  already = false;
}

if (already) {
  console.log("link-self: already linked");
  process.exit(0);
}

rmSync(target, { recursive: true, force: true });
// "junction" on Windows needs no elevation; "dir" is a relative symlink elsewhere.
symlinkSync(pkgRoot, target, process.platform === "win32" ? "junction" : "dir");
console.log("link-self: linked " + target + " -> " + pkgRoot);
