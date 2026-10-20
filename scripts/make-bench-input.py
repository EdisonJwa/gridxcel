#!/usr/bin/env python3
"""
Generate bench/fair-50000x8.xlsx, the input for the head-to-head benchmark.

The point of this fixture is that NEITHER gridxcel nor SheetJS wrote it. It is
produced here by openpyxl, which is an independent implementation, so a parser
bug that the two JavaScript libraries share cannot hide in the input.

The result is committed to the repository, so running this is only necessary to
regenerate or resize the fixture. The benchmark reads the committed file and
does not invoke python.

Usage:
    python scripts/make-bench-input.py [--rows 50000] [--cols 8]
                                       [--out bench/fair-50000x8.xlsx]
"""

import argparse
import os

try:
    from openpyxl import Workbook
except ModuleNotFoundError:  # pragma: no cover - environment guidance
    raise SystemExit(
        "openpyxl is required to regenerate this fixture, because the whole point\n"
        "of the file is that neither gridxcel nor SheetJS wrote it.\n\n"
        "  pip install openpyxl\n\n"
        "The generated file is committed to the repository, so this script is only\n"
        "needed to regenerate or resize it; the benchmark itself does not use python."
    )

# write_only streams rows out without holding them, which is also what makes
# this fixture an honest stand-in for a file a real exporter would produce.
def build(rows: int, cols: int, out: str) -> None:
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")

    # A header row of inline strings, then half text and half numeric columns so
    # both the inline-string and the bare-number paths are exercised.
    ws.append([f"col{c}" for c in range(cols)])
    for r in range(1, rows + 1):
        row = []
        for c in range(cols):
            if c % 2 == 0:
                row.append(f"row{r}c{c}")
            else:
                row.append(r * cols + c)
        ws.append(row)

    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    wb.save(out)
    size = os.path.getsize(out)
    print(f"wrote {out}: {rows} rows x {cols} cols, {size} bytes ({size / 1048576:.2f} MiB)")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--rows", type=int, default=50000)
    ap.add_argument("--cols", type=int, default=8)
    ap.add_argument("--out", default=os.path.join("bench", "fair-50000x8.xlsx"))
    args = ap.parse_args()
    build(args.rows, args.cols, args.out)


if __name__ == "__main__":
    main()
