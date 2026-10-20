
import os, zipfile, datetime, json, random
here = os.path.dirname(os.path.abspath(__file__))
out = os.path.join(here, "fixtures")
os.makedirs(out, exist_ok=True)
random.seed(42)

import xlsxwriter
# 1. xlsxwriter: formats, formulas, dates, merged cells, column widths, hyperlink,
#    multiple sheets, booleans, errors, unicode, 2000 rows.
p = os.path.join(out, "xlsxwriter-full.xlsx")
wb = xlsxwriter.Workbook(p, {"default_date_format": "yyyy-mm-dd"})
bold = wb.add_format({"bold": True})
money = wb.add_format({"num_format": "$#,##0.00"})
pct = wb.add_format({"num_format": "0.00%"})
datef = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
ws = wb.add_worksheet("Report")
ws.write("A1", "Product", bold); ws.write("B1", "Qty", bold); ws.write("C1", "Price", bold)
ws.write("D1", "Total", bold); ws.write("E1", "Share", bold); ws.write("F1", "When", bold)
ws.write("G1", "Flag", bold); ws.write("H1", "Note", bold)
for i in range(2, 202):
    ws.write_string(i-1, 0, "Item-%d" % i)
    ws.write_number(i-1, 1, i)
    ws.write_number(i-1, 2, i * 1.25, money)
    ws.write_formula(i-1, 3, "=B%d*C%d" % (i, i), money, i * 1.25 * i)
    ws.write_number(i-1, 3, i * 1.25 * i)
    ws.write_number(i-1, 4, 0.0125, pct)
    ws.write_datetime(i-1, 5, datetime.datetime(2024, 1, 1) + datetime.timedelta(hours=i), datef)
    ws.write_boolean(i-1, 6, i % 2 == 0)
    ws.write_string(i-1, 7, "\u3053\u3093\u306b\u3061\u306f \u4e16\u754c \U0001f600")
ws.write_formula("B203", "=SUM(B2:B201)")
ws.merge_range("A205:C205", "merged block", wb.add_format({"align": "center"}))
ws.write_url("A207", "https://example.com/", None, "Example link")
ws.write_blank("B207", None)
ws.set_column("A:A", 24); ws.set_column("B:D", 12); ws.set_row(0, 30)
ws2 = wb.add_worksheet("Errors")
ws2.write("A1", "=1/0"); ws2.write("A2", 0); ws2.write_formula("A1", "=1/0")
ws2.write("A3", "=#N/A"); ws2.write("A4", "plain")
wb.close()

import openpyxl
# 2. openpyxl write
p2 = os.path.join(out, "openpyxl-written.xlsx")
w = openpyxl.Workbook()
s = w.active; s.title = "Sheet1"
s["A1"] = "hello"; s["B1"] = 42; s["C1"] = 3.14159
s["A2"] = datetime.datetime(2020, 5, 17, 13, 45, 0)
s["B2"] = True; s["C2"] = "=A1&\" world\""
s["A3"] = "\u00e9\u00e8\u00ea \u4e2d\u6587"
s2 = w.create_sheet("Second")
for i in range(1, 51):
    s2.cell(row=i, column=1, value=i)
    s2.cell(row=i, column=2, value="row %d" % i)
w.save(p2)

# 3. pandas to_excel
import pandas as pd
p3 = os.path.join(out, "pandas.xlsx")
df = pd.DataFrame({"a": range(100), "b": [x / 3.0 for x in range(100)], "c": ["s%d" % x for x in range(100)]})
df.to_excel(p3, index=False, sheet_name="df")

# 4. stored-only zip, level 9, and a comment
src = open(p2, "rb").read()
import shutil
for name, comp in [("openpyxl-stored.xlsx", zipfile.ZIP_STORED), ("openpyxl-lvl9.xlsx", zipfile.ZIP_DEFLATED)]:
    dst = os.path.join(out, name)
    zin = zipfile.ZipFile(p2)
    with zipfile.ZipFile(dst, "w", comp, compresslevel=9 if comp else None) as zout:
        for it in zin.infolist():
            zout.writestr(it.filename, zin.read(it.filename))
    zin.close()
p4 = os.path.join(out, "openpyxl-comment.xlsx")
shutil.copyfile(p2, p4)
with zipfile.ZipFile(p4, "a") as z:
    z.comment = b"gridxcel test comment " * 3

meta = {"fixtures": sorted(os.listdir(out))}
for f in meta["fixtures"]:
    fp = os.path.join(out, f)
    with zipfile.ZipFile(fp) as z:
        meta[f] = {"size": os.path.getsize(fp), "entries": z.namelist()}
print(json.dumps(meta, indent=1))
