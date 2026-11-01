"""Size and time of the HTML output for a large sheet (ARCHITECTURE.md §7.3).

    uv run python bench/html_size.py [DIR] [ROWS] [COLUMNS]

Writes a workbook of ROWS x COLUMNS cells (default 5,000 x 20, 100,000 cells) into DIR (default:
a temporary directory) unless it exists: numbers with a format, texts (some long enough to spill
into an empty column), dates, integers, a bold filled header and a fill on every 50th row. Then
reads it, lays it out in excel mode and renders HTML, and reports the size and the time of each
step and the peak RSS (``ru_maxrss``: bytes on macOS, kilobytes on Linux).
"""

from __future__ import annotations

import datetime as dt
import io
import random
import resource
import sys
import tempfile
import time
from pathlib import Path

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

import sheetpeek


def build(path: Path, rows: int, columns: int) -> None:
    rng = random.Random(3)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFDDEBF7")
    header = []
    for c in range(columns):
        cell = WriteOnlyCell(ws, f"Column {c + 1}")
        cell.font, cell.fill = bold, fill
        header.append(cell)
    ws.append(header)
    for r in range(1, rows):
        cells = []
        for c in range(columns):
            kind = c % 5
            if kind == 0:
                value: object = rng.random() * 1e6
            elif kind == 1:
                value = f"text {r}" + (" that is long enough to spill over" if r % 7 == 0 else "")
            elif kind == 2:
                value = dt.datetime(2020, 1, 1) + dt.timedelta(seconds=rng.randrange(10**8))
            elif kind == 3:
                value = rng.randrange(10**6)
            else:
                value = None if r % 7 == 0 else "x"  # empty beside the long texts
            cell = WriteOnlyCell(ws, value)
            if kind == 0:
                cell.number_format = "#,##0.00"
            elif kind == 2:
                cell.number_format = "yyyy-mm-dd hh:mm"
            if r % 50 == 0:
                cell.font = bold
                cell.fill = fill
            cells.append(cell)
        ws.append(cells)
    wb.save(path)


def main() -> None:
    directory = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(tempfile.mkdtemp())
    rows = int(sys.argv[2]) if len(sys.argv) > 2 else 5_000
    columns = int(sys.argv[3]) if len(sys.argv) > 3 else 20
    path = directory / f"html_{rows}x{columns}.xlsx"
    if not path.exists():
        build(path, rows, columns)
    start = time.perf_counter()
    sheet = sheetpeek.read(path)
    read = time.perf_counter()
    rendered = sheet.layout("excel")
    laid = time.perf_counter()
    out = io.StringIO()
    sheetpeek.backends.HTMLBackend().render(rendered, out)
    done = time.perf_counter()
    text = out.getvalue()
    cells = rows * columns
    runs = sum(len(row.runs) for row in rendered.rows)
    size = len(text.encode())
    print(f"{rows} x {columns} = {cells:,} cells, {runs:,} text runs")
    print(f"html: {size:,} bytes, {size / cells:.0f} bytes per cell")
    print(f"read {read - start:.2f} s, layout {laid - read:.2f} s, render {done - laid:.2f} s")
    print(f"peak RSS {resource.getrusage(resource.RUSAGE_SELF).ru_maxrss:,}")


if __name__ == "__main__":
    main()
