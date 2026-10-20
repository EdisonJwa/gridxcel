"""Generate bench/stream-big.xlsx: 200000 rows x 3 columns, written by XlsxWriter.

constant_memory mode keeps the generator itself cheap; the file is a realistic
deflated worksheet for the streaming reader benchmarks.
"""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "bench", "stream-big.xlsx")
os.makedirs(os.path.dirname(OUT), exist_ok=True)

ROWS = int(sys.argv[1]) if len(sys.argv) > 1 else 200000
wb = xlsxwriter.Workbook(OUT, {"constant_memory": True, "default_date_format": "yyyy-mm-dd"})
ws = wb.add_worksheet("Big")
ws.set_column("A:A", 18)
for i in range(1, ROWS + 1):
    ws.write_number(i - 1, 0, i)
    ws.write_number(i - 1, 1, i / 8.0)
    ws.write_string(i - 1, 2, "row-%d" % i)
wb.close()
print("wrote", OUT, os.path.getsize(OUT), "bytes for", ROWS, "rows")
