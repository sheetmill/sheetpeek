"""Time and peak memory of reading a large workbook in each mode (ARCHITECTURE.md §5.2, §5.4, §5.5).

    uv run python bench/load_modes.py [DIR] [ROWS] [COLUMNS]

Writes a workbook of ROWS x COLUMNS cells (default 100,000 x 20) into DIR (default: a temporary
directory) unless it exists, then measures each mode in a fresh process, since peak RSS only
grows within one. Peak RSS is ``ru_maxrss``: bytes on macOS, kilobytes on Linux.
"""

from __future__ import annotations

import datetime as dt
import random
import resource
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

MODES = (
    "normal",
    "normal-data-only",
    "normal-data-only-styles",
    "read-only",
    "read-only-styles",
    "read-only-raw-dates",
    "structure-scan",
    "package-check",
)


def build(path: Path, rows: int, columns: int) -> None:
    rng = random.Random(3)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    ws.column_dimensions["B"].width = 20
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFFFFF00")
    for r in range(rows):
        cells = []
        for c in range(columns):
            kind = c % 4
            if kind == 0:
                value: object = rng.random() * 1e6
            elif kind == 1:
                value = f"text {r} {c} " + "x" * (r % 13)
            elif kind == 2:
                value = dt.datetime(2020, 1, 1) + dt.timedelta(seconds=rng.randrange(10**8))
            else:
                value = rng.randrange(10**6)
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


def measure(path: Path, mode: str) -> None:
    """Run one mode in this process and print one line."""
    start = time.perf_counter()
    if mode == "package-check":
        from sheetpeek._package import ReadLimits, open_package

        with path.open("rb") as handle:
            open_package(handle, ReadLimits(), str(path)).close()
        loaded = time.perf_counter()
        cells = 0
    elif mode == "structure-scan":
        from sheetpeek._package import worksheet_parts
        from sheetpeek._structure import scan

        with zipfile.ZipFile(path) as archive:
            part = next(iter(worksheet_parts(archive).values()))
            with archive.open(part) as handle:
                scan(handle)
        loaded = time.perf_counter()
        cells = 0
    else:
        options = {"read_only": True, "data_only": True} if mode.startswith("read-only") else {}
        if mode.startswith("normal-data-only"):
            options = {"data_only": True}
        wb = openpyxl.load_workbook(path, **options)
        if mode == "read-only-raw-dates":
            from sheetpeek._read import disable_date_conversion

            assert disable_date_conversion(wb)
        loaded = time.perf_counter()
        styles = mode.endswith("-styles")
        cells = 0
        for row in wb.worksheets[0].iter_rows():
            for cell in row:
                if styles and cell.value is not None:
                    _ = (cell.font.b, cell.number_format, cell.fill.fill_type)
                    _ = cell.alignment.horizontal
                cells += 1
    done = time.perf_counter()
    scale = 1 if sys.platform == "darwin" else 1024
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss * scale / 2**20
    print(
        f"{mode:26s} load {loaded - start:6.2f} s  iterate {done - loaded:6.2f} s  "
        f"total {done - start:6.2f} s  peak RSS {peak:6.0f} MB  cells {cells}"
    )


def main() -> None:
    if len(sys.argv) == 4 and sys.argv[1] == "--mode":
        measure(Path(sys.argv[3]), sys.argv[2])
        return
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(tempfile.mkdtemp())
    rows = int(sys.argv[2]) if len(sys.argv) > 2 else 100_000
    columns = int(sys.argv[3]) if len(sys.argv) > 3 else 20
    path = folder / f"load_modes_{rows}x{columns}.xlsx"
    if not path.exists():
        print(f"writing {path}")
        build(path, rows, columns)
    print(f"openpyxl {openpyxl.__version__}, {path.stat().st_size / 2**20:.1f} MB on disk")
    for mode in MODES:
        subprocess.run([sys.executable, __file__, "--mode", mode, str(path)], check=True)


if __name__ == "__main__":
    main()
