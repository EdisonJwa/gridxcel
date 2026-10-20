/**
 * OPC (ECMA-376 Part 2) package plumbing: content types, relationship parts,
 * part-name resolution, and the workbook-level metadata sweep (sheet list,
 * date system, defined names, document properties).
 *
 * All parts handled here are small metadata documents, so they are read into
 * memory and parsed with src/xml.js. Only worksheet and sharedStrings parts
 * (which can be hundreds of megabytes) go through the chunked scanner in
 * src/sheet-reader.js.
 */

import { XmlParser, localName } from "./xml.js";

const CORE_PROPERTIES = "core-properties";
const EXTENDED_PROPERTIES = "extended-properties";
const OFFICE_DOCUMENT = "officeDocument";

/* ------------------------------------------------------------------ helpers */

/** "/xl/workbook.xml" for any spelling of the same part name. */
function normalizePartName(name) {
  if (!name) return "";
  let n = String(name).replace(/\\/g, "/");
  const hash = n.indexOf("#");
  if (hash >= 0) n = n.slice(0, hash);
  return n.charCodeAt(0) === 47 ? n : "/" + n;
}

/** Last segment of a relationship type URI (works for strict and transitional). */
function relKind(type) {
  if (!type) return "";
  const cut = type.lastIndexOf("/");
  return cut < 0 ? type : type.slice(cut + 1);
}

function boolAttr(v) {
  return v === "" || v === "1" || v === "true" || v === "on";
}

/** Zip entry lookup that tolerates the leading slash OPC part names carry. */
export function getPart(zip, partName) {
  if (!zip || !partName) return undefined;
  const name = normalizePartName(partName);
  return zip.get(name) || zip.get(name.slice(1));
}

/** Bracket a document with the frozen XmlParser; small parts only. */
function scanXml(xml, handlers) {
  const parser = new XmlParser(handlers);
  if (xml) {
    parser.write(xml);
    parser.end();
  }
  return parser;
}

/* ------------------------------------------------------------ content types */

/**
 * [Content_Types].xml -> Map<partName ("/xl/workbook.xml"), contentType>.
 * Default (extension) mappings are attached as the returned map's "defaults"
 * property, since they are keyed by extension rather than by part name.
 */
export function parseContentTypes(xml) {
  const types = new Map();
  types.defaults = new Map();
  scanXml(xml, {
    onClose(name, attrs) {
      const local = localName(name);
      if (local === "Override") {
        if (attrs.PartName) types.set(normalizePartName(attrs.PartName), attrs.ContentType || "");
      } else if (local === "Default") {
        const ext = (attrs.Extension || "").toLowerCase();
        if (ext) types.defaults.set(ext, attrs.ContentType || "");
      }
    },
  });
  return types;
}

/** Content type of a part, falling back to the extension default mapping. */
export function contentTypeOf(types, partName) {
  if (!types) return undefined;
  const name = normalizePartName(partName);
  const hit = types.get(name);
  if (hit !== undefined) return hit;
  const dot = name.lastIndexOf(".");
  if (dot < 0 || !types.defaults) return undefined;
  return types.defaults.get(name.slice(dot + 1).toLowerCase());
}

/* ----------------------------------------------------------- relationships */

/**
 * A .rels part -> Map<Id, { id, type, kind, target, mode }>.
 * mode is "Internal" unless TargetMode says otherwise (hyperlinks, external
 * links), in which case target is an absolute URI that must not be resolved
 * against the part name.
 */
export function parseRels(xml) {
  const rels = new Map();
  scanXml(xml, {
    onClose(name, attrs) {
      if (localName(name) !== "Relationship") return;
      const id = attrs.Id;
      if (!id) return;
      const type = attrs.Type || "";
      rels.set(id, {
        id,
        type,
        kind: relKind(type),
        target: attrs.Target || "",
        mode: attrs.TargetMode || "Internal",
      });
    },
  });
  return rels;
}

/**
 * Resolve a relationship target against the part that owns it.
 * "/xl/workbook.xml" + "worksheets/sheet1.xml" -> "/xl/worksheets/sheet1.xml".
 * Absolute targets and ".." segments are both handled; a fragment is dropped.
 */
export function resolveTarget(basePart, target) {
  let t = String(target === undefined || target === null ? "" : target).replace(/\\/g, "/");
  const hash = t.indexOf("#");
  if (hash >= 0) t = t.slice(0, hash);
  try {
    t = decodeURIComponent(t);
  } catch {
    // Percent escapes that are not valid UTF-8: keep the raw spelling.
  }
  if (t.charCodeAt(0) === 47) return normalizePath(t);
  const base = normalizePartName(basePart);
  const cut = base.lastIndexOf("/");
  const dir = cut <= 0 ? "/" : base.slice(0, cut + 1);
  return normalizePath(dir + t);
}

function normalizePath(p) {
  const out = [];
  for (const seg of p.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") { out.pop(); continue; }
    out.push(seg);
  }
  return "/" + out.join("/");
}

/* -------------------------------------------------------------- properties */

/** docProps/core.xml + docProps/app.xml -> WorkbookProps. */
function parseCoreProps(xml, props) {
  scanXml(xml, {
    onText(text, node) {
      if (!node) return;
      const value = text.trim();
      if (value === "") return;
      switch (node.local) {
        case "title": props.title = value; break;
        case "subject": props.subject = value; break;
        case "creator": props.creator = value; break;
        case "lastModifiedBy": props.lastModifiedBy = value; break;
        case "keywords": props.keywords = value; break;
        case "description": props.description = value; break;
        case "category": props.category = value; break;
        case "created": {
          const d = new Date(value);
          if (!Number.isNaN(d.getTime())) props.created = d;
          break;
        }
        case "modified": {
          const d = new Date(value);
          if (!Number.isNaN(d.getTime())) props.modified = d;
          break;
        }
        default: break;
      }
    },
  });
  return props;
}

function parseAppProps(xml, props) {
  scanXml(xml, {
    onText(text, node) {
      if (!node) return;
      const value = text.trim();
      if (value === "") return;
      if (node.local === "Application") props.application = value;
      else if (node.local === "Company") props.company = value;
    },
  });
  return props;
}

/* --------------------------------------------------------------- workbook */

function relsPathFor(partName) {
  const name = normalizePartName(partName);
  const cut = name.lastIndexOf("/");
  const dir = cut < 0 ? "" : name.slice(0, cut);
  return dir + "/_rels/" + name.slice(cut + 1) + ".rels";
}

/**
 * Read everything the workbook part declares, without touching worksheets:
 * the ordered sheet list with resolved part names, the date system, defined
 * names, document properties and the workbook relationship map.
 *
 * Never throws on a malformed package: problems land in the returned
 * "warnings" array so the caller can route them through onWarning/strict.
 *
 * @returns {{sheets: object[], props: object, definedNames: object[],
 *   date1904: boolean, rels: Map, pathOf: object, workbookPart: string,
 *   types: Map, warnings: object[]}}
 */
export function readWorkbookMeta(zip) {
  const warnings = [];
  const warn = (code, message, part) => {
    warnings.push(part === undefined ? { code, message } : { code, message, part });
  };

  const types = zip && getPart(zip, "/[Content_Types].xml")
    ? parseContentTypes(getPart(zip, "/[Content_Types].xml").text())
    : new Map();

  let pkgRels = new Map();
  const pkgRelsEntry = getPart(zip, "/_rels/.rels");
  if (pkgRelsEntry) {
    try {
      pkgRels = parseRels(pkgRelsEntry.text());
    } catch (err) {
      warn("BAD_XML", "package relationships could not be read: " + err.message, "/_rels/.rels");
    }
  } else {
    warn("MISSING_PART", "package is missing _rels/.rels", "/_rels/.rels");
  }

  const workbookPart = workbookPartFrom(pkgRels, types);
  if (!workbookPart) {
    warn("NO_WORKBOOK", "package declares no workbook part");
    return {
      sheets: [], props: {}, definedNames: [], date1904: false,
      rels: new Map(), pathOf: emptyPathOf(), workbookPart: "", types, warnings,
      workbookRead: { elements: 0, hasRoot: false },
    };
  }

  const sheets = [];
  const definedNames = [];
  let date1904 = false;
  // How much of a workbook document we actually found. read.js turns "nothing
  // here" into a FormatError instead of returning an empty Book.
  const workbookRead = { elements: 0, hasRoot: false };
  const workbookEntry = getPart(zip, workbookPart);
  if (!workbookEntry) {
    warn("MISSING_PART", "workbook part is not in the package", workbookPart);
  } else {
    let textBuf = "";
    let inDefinedName = false;
    try {
      scanXml(workbookEntry.text(), {
        onOpen(name) {
          const local = localName(name);
          if (local === "definedName") { inDefinedName = true; textBuf = ""; }
        },
        onClose(name, attrs) {
          const local = localName(name);
          workbookRead.elements++;
          if (local === "workbook") workbookRead.hasRoot = true;
          if (local === "workbookPr" && attrs.date1904 !== undefined) date1904 = boolAttr(attrs.date1904);
          else if (local === "sheet") {
            sheets.push({
              name: attrs.name === undefined ? "" : attrs.name,
              sheetId: attrs.sheetId,
              state: attrs.state || "visible",
              hidden: attrs.state === "hidden" || attrs.state === "veryHidden",
              rid: attrs["r:id"] || attrs.id,
              path: undefined,
              kind: undefined,
            });
          } else if (local === "definedName") {
            inDefinedName = false;
            const dn = { name: attrs.name || "", ref: textBuf.trim() };
            if (attrs.localSheetId !== undefined) {
              const idx = parseInt(attrs.localSheetId, 10);
              if (Number.isFinite(idx)) dn.sheet = idx;
            }
            if (attrs.hidden !== undefined && boolAttr(attrs.hidden)) dn.hidden = true;
            if (attrs.comment !== undefined) dn.comment = attrs.comment;
            definedNames.push(dn);
          }
        },
        onText(text, node) {
          if (inDefinedName && node && node.local === "definedName") textBuf += text;
        },
      });
    } catch (err) {
      warn("BAD_XML", "workbook part could not be read: " + err.message, workbookPart);
    }
  }

  let rels = new Map();
  const relsPath = relsPathFor(workbookPart);
  const relsEntry = getPart(zip, relsPath);
  if (relsEntry) {
    try {
      rels = parseRels(relsEntry.text());
    } catch (err) {
      warn("BAD_XML", "workbook relationships could not be read: " + err.message, relsPath);
    }
  } else {
    warn("MISSING_PART", "workbook relationships are missing", relsPath);
  }

  const byKind = new Map();
  const byId = new Map();
  for (const rel of rels.values()) {
    byId.set(rel.id, rel);
    if (rel.mode !== "Internal") continue;
    const list = byKind.get(rel.kind);
    const path = resolveTarget(workbookPart, rel.target);
    if (list) list.push(path);
    else byKind.set(rel.kind, [path]);
  }

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const rel = sheet.rid ? rels.get(sheet.rid) : undefined;
    if (rel) {
      sheet.kind = rel.kind;
      sheet.path = rel.mode === "External" ? undefined : resolveTarget(workbookPart, rel.target);
    }
    if (!sheet.path) {
      // Conventional layout: worksheets/sheetN.xml, N = declaration order.
      const guess = "/xl/worksheets/sheet" + (i + 1) + ".xml";
      sheet.path = guess;
      sheet.kind = sheet.kind || "worksheet";
      if (!getPart(zip, guess)) {
        warn("MISSING_PART", "sheet " + JSON.stringify(sheet.name) + " has no worksheet part", guess);
      }
    }
  }

  const pathOf = {
    workbook: workbookPart,
    sheets: sheets.map((s) => s.path),
    sharedStrings: firstOf(byKind, "sharedStrings"),
    styles: internalPathOf(rels, workbookPart, "styles"),
    theme: firstOf(byKind, "theme"),
    byId: rels,
    byKind,
    relsPath,
  };

  const props = {};
  const corePath = findPkgPart(zip, pkgRels, types, CORE_PROPERTIES, "/docProps/core.xml");
  if (corePath) {
    const entry = getPart(zip, corePath);
    if (entry) {
      try { parseCoreProps(entry.text(), props); }
      catch (err) { warn("BAD_XML", "document properties could not be read: " + err.message, corePath); }
    }
  }
  const appPath = findPkgPart(zip, pkgRels, types, EXTENDED_PROPERTIES, "/docProps/app.xml");
  if (appPath) {
    const entry = getPart(zip, appPath);
    if (entry) {
      try { parseAppProps(entry.text(), props); }
      catch (err) { warn("BAD_XML", "extended properties could not be read: " + err.message, appPath); }
    }
  }

  return { sheets, props, definedNames, date1904, rels, pathOf, workbookPart, types, warnings, workbookRead };
}

/**
 * Read the small metadata parts into the entry cache, so the synchronous
 * readWorkbookMeta() sweep below can run on a host that has no synchronous
 * codec (a browser: the Web decompressor is async only).
 *
 * It reads exactly the parts that sweep will touch - content types, the root
 * relationships, the workbook part, the workbook relationships, the styles part
 * and the document properties - and nothing else, so the extra cost on Node is
 * a handful of cache hits and the streaming parts (worksheets, shared strings)
 * are still read through the chunked scanner rather than buffered here.
 *
 * Failures are swallowed on purpose: a part that cannot be read now is reported
 * by the sweep with the same warning it always produced.
 */
export async function prewarmMeta(zip) {
  if (!zip || typeof zip.get !== "function") return;
  const warm = async (entry) => {
    if (!entry || entry.isDirectory) return undefined;
    try {
      return await entry.textAsync();
    } catch {
      return undefined;
    }
  };
  const typesXml = await warm(getPart(zip, "/[Content_Types].xml"));
  const pkgRelsXml = await warm(getPart(zip, "/_rels/.rels"));
  let types = new Map();
  let pkgRels = new Map();
  // Parsing here is only to decide WHICH parts to warm; the sweep re-parses the
  // cached text and owns the warnings.
  try { if (typesXml !== undefined) types = parseContentTypes(typesXml); } catch { types = new Map(); }
  try { if (pkgRelsXml !== undefined) pkgRels = parseRels(pkgRelsXml); } catch { pkgRels = new Map(); }

  const workbookPart = workbookPartFrom(pkgRels, types);
  if (!workbookPart) return;
  await warm(getPart(zip, workbookPart));
  const relsXml = await warm(getPart(zip, relsPathFor(workbookPart)));
  let rels = new Map();
  try { if (relsXml !== undefined) rels = parseRels(relsXml); } catch { rels = new Map(); }
  await warm(getPart(zip, internalPathOf(rels, workbookPart, "styles")));
  await warm(getPart(zip, findPkgPart(zip, pkgRels, types, CORE_PROPERTIES, "/docProps/core.xml")));
  await warm(getPart(zip, findPkgPart(zip, pkgRels, types, EXTENDED_PROPERTIES, "/docProps/app.xml")));
}

function firstOf(byKind, kind) {
  const list = byKind.get(kind);
  return list && list.length ? list[0] : undefined;
}

/**
 * First Internal relationship of a kind, resolved against its base part. This is
 * the same rule pathOf applies through byKind, expressed so that prewarmMeta()
 * and readWorkbookMeta() cannot drift about which part holds the styles.
 */
function internalPathOf(rels, basePart, kind) {
  for (const rel of rels.values()) {
    if (rel.mode === "Internal" && rel.kind === kind) return resolveTarget(basePart, rel.target);
  }
  return undefined;
}

/**
 * The workbook part: the officeDocument relationship, else the part whose
 * content type declares a sheet main document. Shared by readWorkbookMeta() and
 * prewarmMeta() so both agree on what to read.
 */
function workbookPartFrom(pkgRels, types) {
  for (const rel of pkgRels.values()) {
    if (rel.kind === OFFICE_DOCUMENT) return resolveTarget("/", rel.target);
  }
  for (const [name, type] of types) {
    if (/spreadsheetml\.[a-z.]*sheet\.main\+xml$/.test(type)) return name;
  }
  return undefined;
}

/** Locate a package-level part by relationship kind, then by content type. */
function findPkgPart(zip, pkgRels, types, kind, conventional) {
  for (const rel of pkgRels.values()) {
    if (rel.kind === kind && rel.mode === "Internal") return resolveTarget("/", rel.target);
  }
  if (getPart(zip, conventional)) return conventional;
  if (types) {
    for (const [name, type] of types) {
      if (kind === CORE_PROPERTIES && /core-properties/.test(type)) return name;
      if (kind === EXTENDED_PROPERTIES && /extended-properties/.test(type)) return name;
    }
  }
  return undefined;
}

function emptyPathOf() {
  return { workbook: "", sheets: [], sharedStrings: undefined, styles: undefined, theme: undefined, byId: new Map(), byKind: new Map() };
}

/** Relationship-part path of a worksheet, used to find its comments part. */
export function relsPathOf(partName) {
  return relsPathFor(partName);
}
