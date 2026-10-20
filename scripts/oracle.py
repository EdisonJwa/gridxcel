#!/usr/bin/env python
"""Independent XLSX oracle for the gridxcel interop harness.

This file is the ONLY source of truth used by scripts/interop.mjs for the
"Python said so" side of every comparison. It deliberately uses three
independent stacks:

  * openpyxl  - cell semantics (number formats, date detection, styles)
  * zipfile   - container integrity (testzip, entry list, compression method)
  * lxml      - the raw worksheet XML, which is the authoritative cell
                inventory. openpyxl's row iteration can miss cells whose <c>
                has only a <v> child and no style, so a cell that exists in the
                sheet XML must never be reported as absent: the XML is the
                spine of the per-sheet report and openpyxl enriches it.

Commands (all print a single JSON object on stdout):

  inspect <file.xlsx> [--cached] [--read-only] [--no-xml]
      Full description of one workbook. --cached adds the cached formula
      results from a second, data_only load. --no-xml falls back to pure
      openpyxl iteration (useful for cross-checking the XML spine).

  gen <out.xlsx> --rows N --cols M [--seed S]
      Deterministic workbook used by the volume tests and the benchmarks.
      Written with openpyxl so the numbers never come from gridxcel.

  selftest <file.xlsx>
      Container-only report (CRCs, entry list). Proves the oracle path works
      without trusting openpyxl.

Exit codes: 0 ok, 1 the file could not be inspected, 2 bad usage.
"""

import argparse
import datetime
import hashlib
import json
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ORACLE_VERSION = 2
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# Date-ness heuristic used ONLY for cells that openpyxl never instantiated.
# It is deliberately crude and never overrides openpyxl's own is_date().
DATE_TOKENS = re.compile(r"(?<!\\)[ymdhs]", re.IGNORECASE)
DATE_CODE_HINT = re.compile(r"(yy|dd|mm|hh|ss|d\-|m\-|y\-)", re.IGNORECASE)

BUILTIN_DATE_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32,
                    33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}


def die(msg, code=1):
    print(json.dumps({"error": str(msg), "oracleVersion": ORACLE_VERSION}))
    sys.exit(code)


# --------------------------------------------------------------- zip layer

def zip_report(path):
    """Container facts straight from zipfile: the ground truth for (a)."""
    out = {"testzip": None, "entries": [], "comment": "", "commentLen": 0,
           "methods": {}, "crc32": {}, "sizes": {}, "uncompressedSizes": {},
           "duplicateEntries": []}
    with zipfile.ZipFile(path) as z:
        # testzip() returns None when every entry passes its CRC check, else
        # the name of the first bad file. This is the (a) acceptance check.
        out["testzip"] = z.testzip()
        out["entries"] = list(z.namelist())
        raw_comment = z.comment or b""
        out["commentLen"] = len(raw_comment)
        out["comment"] = raw_comment.decode("utf-8", "replace")
        for info in z.infolist():
            out["methods"][info.filename] = info.compress_type
            out["crc32"][info.filename] = info.CRC
            out["sizes"][info.filename] = info.compress_size
            out["uncompressedSizes"][info.filename] = info.file_size
        names = out["entries"]
        out["duplicateEntries"] = sorted({n for n in names if names.count(n) > 1})
    return out


def styles_info(path):
    """id -> code for numFmts plus the cellXfs numFmtId list, via lxml."""
    result = {"numFmts": {}, "cellXfs": [], "error": None}
    try:
        from lxml import etree
        with zipfile.ZipFile(path) as z:
            if "xl/styles.xml" not in z.namelist():
                return result
            raw = z.read("xl/styles.xml")
        root = etree.fromstring(raw)
        for node in root.findall(".//{%s}numFmts/{%s}numFmt" % (MAIN_NS, MAIN_NS)):
            fmt_id = node.get("numFmtId")
            if fmt_id is not None:
                result["numFmts"][int(fmt_id)] = node.get("formatCode") or ""
        for xf in root.findall(".//{%s}cellXfs/{%s}xf" % (MAIN_NS, MAIN_NS)):
            fmt_id = xf.get("numFmtId")
            result["cellXfs"].append(int(fmt_id) if fmt_id is not None else 0)
    except Exception as exc:  # pragma: no cover - malformed styles are reported
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
    return result


def shared_strings(path):
    """Resolved xl/sharedStrings.xml entries: si[i] -> concatenated text."""
    try:
        from lxml import etree
        with zipfile.ZipFile(path) as z:
            if "xl/sharedStrings.xml" not in z.namelist():
                return []
            root = etree.fromstring(z.read("xl/sharedStrings.xml"))
        out = []
        for si in root.findall("{%s}si" % MAIN_NS):
            out.append("".join(t.text or "" for t in si.iter("{%s}t" % MAIN_NS)))
        return out
    except Exception:
        return []


# ------------------------------------------------ raw worksheet XML spine

def col_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


REF_RE = re.compile(r"^([A-Za-z]{1,3})(\d+)$")


def split_ref(ref):
    m = REF_RE.match(ref or "")
    if not m:
        return None, None
    return int(m.group(2)), col_to_num(m.group(1))


def element_xml(node):
    """Compact raw XML for one element, for triage in mismatch reports."""
    try:
        from lxml import etree
        text = etree.tostring(node, encoding="unicode")
    except Exception:
        return None
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= 300 else text[:297] + "..."


def xml_sheet_cells(path, sheet_part, styles, sst, date_hint):
    """
    ref -> cell record, read from the worksheet XML itself.

    This is the authoritative inventory: a <c> element in sheetData exists,
    whether or not openpyxl materialised a Cell object for it. Values are
    decoded here rather than taken from openpyxl so that the harness can see
    cells openpyxl skipped.
    """
    from lxml import etree
    with zipfile.ZipFile(path) as z:
        if sheet_part not in z.namelist():
            return None
        root = etree.fromstring(z.read(sheet_part))
    cells = {}
    row_index = 0
    for row in root.findall(".//{%s}sheetData/{%s}row" % (MAIN_NS, MAIN_NS)):
        row_index = int(row.get("r")) if row.get("r") else row_index + 1
        col_index = 0
        for c in row.findall("{%s}c" % MAIN_NS):
            ref = c.get("r")
            if ref:
                r, col = split_ref(ref)
                if r is None:
                    continue
                col_index = col
            else:
                # No r attribute: positional fallback.
                col_index += 1
                r = row_index
                ref = num_to_col(col_index) + str(r)
            cells[ref] = decode_xml_cell(c, ref, r, col_index, styles, sst, date_hint)
    return cells


def num_to_col(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def decode_xml_cell(c, ref, r, col, styles, sst, date_hint):
    t_attr = c.get("t")
    style = int(c.get("s")) if c.get("s") is not None else 0
    f_node = c.find("{%s}f" % MAIN_NS)
    v_node = c.find("{%s}v" % MAIN_NS)
    is_node = c.find("{%s}is" % MAIN_NS)
    rec = {
        "ref": ref, "row": r, "col": col, "style": style,
        "xml": element_xml(c), "fromXml": True,
    }
    formula = None
    if f_node is not None:
        formula = f_node.text or ""
        rec["formula"] = "=" + formula
        rec["formulaType"] = f_node.get("t")
        rec["formulaRef"] = f_node.get("ref")

    raw = v_node.text if v_node is not None and v_node.text is not None else None
    if is_node is not None:
        raw = "".join(x.text or "" for x in is_node.iter("{%s}t" % MAIN_NS))

    if t_attr == "s" and raw is not None:
        try:
            rec["kind"] = "string"
            rec["value"] = sst[int(raw)]
        except Exception:
            rec["kind"] = "string"
            rec["value"] = ""
            rec["resolveError"] = "shared string index %r" % raw
    elif t_attr == "inlineStr":
        rec["kind"] = "string"
        rec["value"] = raw or ""
    elif t_attr == "str":
        rec["kind"] = "string"
        rec["value"] = raw if raw is not None else None
    elif t_attr == "b":
        rec["kind"] = "boolean"
        rec["value"] = (raw == "1") if raw is not None else None
    elif t_attr == "e":
        rec["kind"] = "error"
        rec["value"] = raw
    elif t_attr == "d":
        rec["kind"] = "date"
        rec["value"] = {"datetime": raw}
        rec["date_iso"] = raw
        rec["serial"] = iso_to_serial(raw)
    elif raw is None:
        rec["kind"] = "formula" if formula else "empty"
        rec["value"] = None
    else:
        try:
            num = float(raw)
        except ValueError:
            rec["kind"] = "string"
            rec["value"] = raw
            return rec
        fmt_id = styles["cellXfs"][style] if style < len(styles["cellXfs"]) else 0
        code = styles["numFmts"].get(fmt_id)
        if code is None:
            if num.is_integer() and num >= 0:
                rec["kind"] = "date" if (fmt_id in BUILTIN_DATE_IDS and date_hint) else "number"
            else:
                rec["kind"] = "number"
        else:
            rec["kind"] = "date" if DATE_CODE_HINT.search(code) else "number"
        rec["value"] = num
        if rec["kind"] == "date":
            rec["serial"] = num
            rec["date_iso"] = serial_to_iso(num)
        else:
            rec["kind"] = "formula" if (formula and rec["kind"] == "number") else rec["kind"]
    if formula:
        rec["kind"] = "formula"
    return rec


def iso_to_serial(iso):
    """1900-system serial from an ISO wall-clock string (python's own rule)."""
    try:
        dt = datetime.datetime.fromisoformat(iso)
    except Exception:
        return None
    base = datetime.datetime(1899, 12, 30)
    delta = dt - base
    return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400.0


def serial_to_iso(serial):
    base = datetime.datetime(1899, 12, 30)
    try:
        dt = base + datetime.timedelta(days=float(serial))
    except Exception:
        return None
    return dt.isoformat()


# ------------------------------------------------------------ cell layer

def cell_kind(cell):
    t = cell.data_type
    if t == "f":
        return "formula"
    if t == "b":
        return "boolean"
    if t == "d":
        return "date"
    if t == "e":
        return "error"
    if t == "n":
        return "number"
    if t in ("s", "str", "inlineStr"):
        return "string"
    return "other:" + str(t)


def formula_text(value):
    """ArrayFormula / DataTableFormula expose .text; plain formulas are str."""
    if isinstance(value, str):
        return value
    text = getattr(value, "text", None)
    if text is not None:
        return text
    return str(value)


def encode_value(cell, epoch):
    """JSON-safe canonical value plus the Excel serial for date cells."""
    v = cell.value
    out = {"json": None, "serial": None, "date_iso": None}
    if isinstance(v, datetime.datetime):
        out["json"] = {"datetime": v.isoformat()}
        out["date_iso"] = v.isoformat()
        out["serial"] = openpyxl_to_excel(v, epoch)
    elif isinstance(v, datetime.date):
        out["json"] = {"date": v.isoformat()}
        out["date_iso"] = v.isoformat() + "T00:00:00"
        out["serial"] = openpyxl_to_excel(v, epoch)
    elif isinstance(v, datetime.timedelta):
        out["json"] = {"timedelta_seconds": v.total_seconds()}
    elif isinstance(v, bool):
        out["json"] = v
    elif isinstance(v, (int, float)):
        out["json"] = v
    elif v is None:
        out["json"] = None
    else:
        out["json"] = str(v)
    return out


def typed_scalar(value, epoch):
    """A cached formula result as an explicitly tagged scalar (never a repr)."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return {"t": "d", "iso": value.isoformat(), "serial": openpyxl_to_excel(value, epoch)}
    if isinstance(value, datetime.date):
        iso = value.isoformat() + "T00:00:00"
        return {"t": "d", "iso": iso, "serial": openpyxl_to_excel(value, epoch)}
    if isinstance(value, bool):
        return {"t": "b", "v": value}
    if isinstance(value, float):
        if value != value:
            return {"t": "n", "nan": True}
        return {"t": "n", "v": value}
    if isinstance(value, int):
        return {"t": "n", "v": value}
    return {"t": "s", "v": formula_text(value)}


def openpyxl_to_excel(value, epoch):
    from openpyxl.utils.datetime import to_excel
    try:
        return float(to_excel(value, epoch))
    except Exception:
        return None


def sheet_report(ws, epoch, cached_ws=None, xml_cells=None):
    """Merge the XML spine with openpyxl's richer per-cell semantics."""
    cells = []
    true_max_row = 0
    true_max_col = 0
    cached_cells = {}
    if cached_ws is not None:
        for row in cached_ws.iter_rows():
            for c in row:
                if c.value is not None:
                    cached_cells[c.coordinate] = c.value
    if cached_ws is not None:
        # Cells openpyxl skipped in the normal load can still surface here.
        for coord, c in getattr(cached_ws, "_cells", {}).items():
            if c.value is not None:
                cached_cells.setdefault(c.coordinate, c.value)

    seen = set()
    xml_cells = xml_cells or {}

    def add_from_openpyxl(cell):
        kind = cell_kind(cell)
        enc = encode_value(cell, epoch)
        entry = {
            "ref": cell.coordinate,
            "row": cell.row,
            "col": cell.column,
            "kind": kind,
            "value": enc["json"],
            "format": cell.number_format,
            "source": "openpyxl",
        }
        if enc["serial"] is not None:
            entry["serial"] = enc["serial"]
        if enc["date_iso"] is not None:
            entry["date_iso"] = enc["date_iso"]
        if kind == "formula":
            entry["formula"] = formula_text(cell.value)
            ref = getattr(cell.value, "ref", None)
            if ref:
                entry["formulaRef"] = ref
            entry["cached"] = typed_scalar(cached_cells.get(cell.coordinate), epoch)
        xml = xml_cells.get(cell.coordinate)
        if xml:
            entry["xml"] = xml.get("xml")
            if xml.get("formula") and not entry.get("formula"):
                entry["formula"] = xml["formula"]
        return entry

    # 1) the XML spine first: every <c> in the sheet exists as a cell
    for ref, xml in sorted(xml_cells.items(), key=lambda kv: (kv[1]["row"], kv[1]["col"])):
        op_cell = ws._cells.get((xml["row"], xml["col"])) if hasattr(ws, "_cells") else None
        if op_cell is not None and (op_cell.value is not None or xml["kind"] == "empty"):
            entry = add_from_openpyxl(op_cell)
            if entry["ref"] != ref:
                entry["ref"] = ref
        else:
            entry = {
                "ref": ref, "row": xml["row"], "col": xml["col"],
                "kind": xml["kind"], "value": xml.get("value"),
                "format": xml_format(xml, ws), "source": "xml", "xml": xml.get("xml"),
            }
            if xml.get("serial") is not None:
                entry["serial"] = xml["serial"]
            if xml.get("date_iso") is not None:
                entry["date_iso"] = xml["date_iso"]
            if xml.get("formula"):
                entry["formula"] = xml["formula"]
            if xml.get("kind") == "formula":
                entry["cached"] = typed_scalar(cached_cells.get(ref), epoch)
            if xml.get("kind") == "empty":
                continue
        seen.add(entry["ref"])
        cells.append(entry)
        true_max_row = max(true_max_row, entry["row"])
        true_max_col = max(true_max_col, entry["col"])

    # 2) anything openpyxl knows that the XML scan did not cover
    for row in ws.iter_rows():
        for cell in row:
            if cell.coordinate in seen:
                continue
            kind = cell_kind(cell)
            if cell.value is None and kind in ("string", "number", "date", "boolean"):
                continue  # styled-but-empty stub
            entry = add_from_openpyxl(cell)
            seen.add(entry["ref"])
            cells.append(entry)
            true_max_row = max(true_max_row, cell.row)
            true_max_col = max(true_max_col, cell.column)

    merges = []
    try:
        merges = sorted(str(r) for r in ws.merged_cells.ranges)
    except Exception:
        merges = []

    return {
        "name": ws.title,
        "state": getattr(ws, "sheet_state", "visible"),
        "maxRow": ws.max_row,
        "maxCol": ws.max_column,
        "trueMaxRow": true_max_row,
        "trueMaxCol": true_max_col,
        "dimensions": safe_dimensions(ws),
        "merges": merges,
        "freeze": safe_attr(ws, "freeze_panes"),
        "autoFilter": safe_attr(getattr(ws, "auto_filter", None), "ref"),
        "cellCount": len(cells),
        "cells": cells,
    }


def xml_format(xml, ws):
    """Number format for an XML-only cell, from openpyxl's style table when possible."""
    try:
        fmt_id = xml.get("style", 0)
        wb = ws.parent
        if wb is not None:
            from openpyxl.styles.numbers import BUILTIN_FORMATS
            table = getattr(wb, "_number_formats", None) or []
            for item in table:
                if getattr(item, "numFmtId", None) == fmt_id:
                    return item.formatCode
            if fmt_id in BUILTIN_FORMATS:
                return BUILTIN_FORMATS[fmt_id]
    except Exception:
        pass
    return "General"


def safe_attr(obj, name):
    try:
        return getattr(obj, name)
    except Exception:
        return None


def safe_dimensions(ws):
    try:
        return ws.calculate_dimension()
    except Exception:
        return None


def inspect(path, cached=False, read_only=False, use_xml=True):
    import openpyxl
    from openpyxl.styles.numbers import BUILTIN_FORMATS
    if not os.path.exists(path):
        die("no such file: %s" % path)
    wb = openpyxl.load_workbook(path, data_only=False, read_only=read_only)
    cached_wb = None
    if cached:
        try:
            cached_wb = openpyxl.load_workbook(path, data_only=True, read_only=read_only)
        except Exception:
            cached_wb = None
    styles = styles_info(path)
    sst = shared_strings(path) if use_xml else []
    date_hint = True
    sheets = []
    for ws in wb.worksheets:
        cws = None
        if cached_wb is not None and ws.title in getattr(cached_wb, "sheetnames", []):
            cws = cached_wb[ws.title]
        xml_cells = None
        if use_xml:
            part = getattr(ws, "path", None)
            if part and part.startswith("/"):
                part = part[1:]
            if part:
                try:
                    xml_cells = xml_sheet_cells(path, part, styles, sst, date_hint)
                except Exception:
                    xml_cells = None
        sheets.append(sheet_report(ws, wb.epoch, cws, xml_cells))
    defined = []
    try:
        for name, dn in (wb.defined_names or {}).items():
            defined.append({"name": name, "ref": getattr(dn, "value", None)})
    except Exception:
        defined = []
    return {
        "oracleVersion": ORACLE_VERSION,
        "file": {
            "path": os.path.abspath(path),
            "size": os.path.getsize(path),
            "sha256": sha256(path),
        },
        "zip": zip_report(path),
        "book": {
            "dateSystem": 1904 if getattr(wb, "epoch", None) == datetime.datetime(1904, 1, 1) else 1900,
            "epoch": getattr(wb, "epoch", None).isoformat() if getattr(wb, "epoch", None) else None,
            "numFmts": {str(k): v for k, v in sorted(styles["numFmts"].items())},
            "builtinFormats": {str(k): v for k, v in sorted(BUILTIN_FORMATS.items())},
            "cellXfs": styles["cellXfs"],
            "sheetNames": list(wb.sheetnames),
            "definedNames": defined,
            "stylesError": styles["error"],
        },
        "sheets": sheets,
    }


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


# ------------------------------------------------------------ generator

def generate(out, rows, cols, seed):
    """Deterministic openpyxl-written workbook: numbers, strings, dates, formulas."""
    import openpyxl
    import random
    rnd = random.Random(seed)
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("volume")
    ws.append(["col_%d" % c for c in range(cols)])
    base = datetime.datetime(2020, 1, 1)
    for r in range(rows):
        row = []
        for c in range(cols):
            m = c % 5
            if m == 0:
                row.append(r * cols + c)
            elif m == 1:
                row.append((r * cols + c) / 7.0)
            elif m == 2:
                row.append("s%d_%d" % (r, c))
            elif m == 3:
                row.append(r % 2 == 0)
            else:
                row.append(base + datetime.timedelta(minutes=r))
        ws.append(row)
    wb.save(out)
    return {"path": os.path.abspath(out), "rows": rows, "cols": cols, "size": os.path.getsize(out)}


def selftest(path):
    rep = zip_report(path)
    ok = rep["testzip"] is None and not rep["duplicateEntries"]
    return {"file": os.path.abspath(path), "zipok": ok, "entries": len(rep["entries"]),
            "testzip": rep["testzip"], "duplicateEntries": rep["duplicateEntries"]}


def main():
    ap = argparse.ArgumentParser(description="gridxcel independent oracle")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("inspect")
    p.add_argument("file")
    p.add_argument("--cached", action="store_true")
    p.add_argument("--read-only", action="store_true")
    p.add_argument("--no-xml", action="store_true")

    p = sub.add_parser("gen")
    p.add_argument("out")
    p.add_argument("--rows", type=int, default=20000)
    p.add_argument("--cols", type=int, default=8)
    p.add_argument("--seed", type=int, default=7)

    p = sub.add_parser("selftest")
    p.add_argument("file")

    args = ap.parse_args()
    try:
        if args.cmd == "inspect":
            doc = inspect(args.file, cached=args.cached, read_only=args.read_only, use_xml=not args.no_xml)
        elif args.cmd == "gen":
            doc = generate(args.out, args.rows, args.cols, args.seed)
        else:
            doc = selftest(args.file)
    except SystemExit:
        raise
    except Exception as exc:
        import traceback
        traceback.print_exc(file=sys.stderr)
        die("%s: %s" % (type(exc).__name__, exc))
    print(json.dumps(doc, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
