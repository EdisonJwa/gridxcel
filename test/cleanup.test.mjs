/**
 * Regressions for the code-review cleanup pass: F11 (one code for "no such
 * sheet"), F12 (internals that were accidentally exported) and F15 (a guard the
 * coercion change made unreachable).
 *
 * F12 is asserted against the real module namespaces rather than the source
 * text, because the point of the change is the shape of the public surface.
 */
import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { readWorkbook, readWorkbookSync } from "../src/read.js";
import { writeWorkbookSync } from "../src/write.js";
import * as workbookWriter from "../src/workbook-writer.js";
import * as opc from "../src/opc.js";
import * as styleRegistry from "../src/style-registry.js";

const PANDAS = "test/fixtures/pandas.xlsx";

/* ------------------------------------------------------------------------ F11 */

test("F11: a sheet selector that matches nothing warns NO_SHEET, sync and async", async () => {
  const missingName = [];
  const byName = readWorkbookSync(PANDAS, { sheets: ["nope"], onWarning: (w) => missingName.push(w) });
  assert.deepEqual(missingName.map((w) => w.code), ["NO_SHEET"], "unknown name");
  assert.deepEqual(byName.sheets, []);

  const missingIndex = [];
  await readWorkbook(PANDAS, { sheets: [7], onWarning: (w) => missingIndex.push(w) });
  assert.deepEqual(missingIndex.map((w) => w.code), ["NO_SHEET"], "unknown index (async path)");

  const badSelector = [];
  readWorkbookSync(PANDAS, { sheets: "df", onWarning: (w) => badSelector.push(w) });
  assert.deepEqual(badSelector.map((w) => w.code), ["NO_SHEET"], "non-array selector");

  const badItem = [];
  readWorkbookSync(PANDAS, { sheets: [{ sheet: "df" }], onWarning: (w) => badItem.push(w) });
  assert.deepEqual(badItem.map((w) => w.code), ["NO_SHEET"], "selector that is neither name nor index");
});

test("F11: the old BAD_SHEET_SELECTOR spelling is gone from the reader", () => {
  assert.equal(readFileSync("src/read.js", "utf8").includes("BAD_SHEET_SELECTOR"), false);
  // The option still selects sheets correctly.
  const warnings = [];
  const { sheets } = readWorkbookSync(PANDAS, { sheets: ["df"], onWarning: (w) => warnings.push(w) });
  assert.deepEqual(sheets.map((s) => s.name), ["df"]);
  assert.deepEqual(warnings, []);
});

/* ------------------------------------------------------------------------ F12 */

test("F12: internals used only inside their own module are no longer exports", () => {
  const internal = {
    "src/workbook-writer.js": ["XML_DECL", "NS_MAIN", "NS_REL", "NS_PKG_REL", "NS_CT", "NS_CP", "NS_DC", "NS_DCTERMS", "NS_DCMITYPE", "NS_XSI", "NS_EP", "NS_VT"],
    // internalPathOf is the same finding, found by the fresh inventory rather
    // than by the review's list (opc.js:351,418,433 are its only users).
    "src/opc.js": ["normalizePartName", "relKind", "internalPathOf"],
    "src/style-registry.js": ["FIRST_CUSTOM_NUMFMT_ID"],
  };
  const namespaces = { "src/workbook-writer.js": workbookWriter, "src/opc.js": opc, "src/style-registry.js": styleRegistry };
  for (const [file, names] of Object.entries(internal)) {
    for (const name of names) {
      assert.equal(name in namespaces[file], false, file + " must not export " + name);
    }
  }
});

test("F12: exports a neighbour or an entry point needs are kept", () => {
  // read.js imports these from opc.js, so they stay exported.
  for (const name of ["getPart", "parseRels", "resolveTarget", "readWorkbookMeta", "relsPathOf"]) {
    assert.equal(typeof opc[name], "function", "opc." + name + " must stay exported");
  }
  // index.js re-exports these and the documentation names them.
  for (const name of ["WorkbookWriter", "SheetWriter", "buildParts", "createBookWriter", "numberLiteral"]) {
    assert.equal(name in workbookWriter, true, "workbook-writer." + name + " must stay exported");
  }
  for (const name of ["StyleTable", "parseStyles", "normalizeColor", "colorAttrs", "colorElement"]) {
    assert.equal(name in styleRegistry, true, "style-registry." + name + " must stay exported");
  }
});

/* ------------------------------------------------------------------------ F15 */

test("F15: type aliases still resolve after the unreachable guard was removed", () => {
  // json.cellFromValue guarantees every cell it returns carries a t, so the
  // writer's "t !== undefined && t !== null" test could never fail. The alias
  // mapping it guarded must still run for every alias the writer accepts.
  const book = {
    sheets: [{
      name: "S",
      rows: [[
        { v: 1, t: "n" },
        { v: "x", t: "str" },
        { v: true, t: "b" },
        { v: "#N/A", t: "e" },
        { v: "plain" },
      ]],
      merges: [],
      columns: [],
      rowInfo: new Map(),
    }],
    dateSystem: 1900,
    numFmts: new Map(),
    styles: [],
    props: {},
    definedNames: [],
  };
  const read = readWorkbookSync(writeWorkbookSync(book, {}));
  const row = read.sheets[0].rows[0];
  assert.deepEqual({ t: row[0].t, v: row[0].v }, { t: "number", v: 1 });
  assert.deepEqual({ t: row[1].t, v: row[1].v }, { t: "string", v: "x" });
  assert.deepEqual({ t: row[2].t, v: row[2].v }, { t: "boolean", v: true });
  assert.deepEqual({ t: row[3].t, v: row[3].v }, { t: "error", v: "#N/A" });
  assert.deepEqual({ t: row[4].t, v: row[4].v }, { t: "string", v: "plain" });
});
