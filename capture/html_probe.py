"""Probe: what Excel for Mac's "save as web page" writes for spill, merges, hidden columns, `####`.

Builds one workbook with openpyxl inside Excel's sandbox container, has Excel open it and save it
as HTML (``html_export.applescript``), and copies the export to ``capture/build/html_probe/``.

    uv run --group capture python -m capture.html_probe build    # write the workbook
    uv run --group capture python -m capture.html_probe run      # build, export, copy back

Each probe row's purpose is written in column J of that row, so the export is self-describing.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, PatternFill

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "html_probe"
# Workbooks go into Excel's sandbox container; anywhere else makes macOS ask for access per file.
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
)
WORKBOOK = CONTAINER_DIR / "html_probe.xlsx"
EXPORT = CONTAINER_DIR / "html_probe.htm"
LONG = "spilled text that is much wider than one column"


def build() -> Path:
    """Write the probe workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Probe"
    note = {}
    yellow = PatternFill("solid", fgColor="FFFFFF00")

    def put(row: int, purpose: str) -> None:
        note[row] = purpose

    ws["A1"] = LONG
    put(1, "left text, empty neighbors: spill right")
    ws["A2"] = LONG
    ws["B2"] = 5
    put(2, "left text, number in B: blocked")
    ws["A3"] = LONG
    ws["B3"].fill = yellow
    put(3, "left text, B styled only (fill): spill?")
    ws["A4"] = LONG
    ws["B4"] = '=""'
    put(4, 'left text, B formula ="": blocked?')
    ws["A5"] = LONG
    ws["B5"] = " "
    put(5, "left text, B a space: blocked?")
    ws["C6"] = LONG
    ws["C6"].alignment = Alignment(horizontal="right")
    put(6, "right-aligned text in C: spill left")
    ws["C7"] = LONG
    ws["C7"].alignment = Alignment(horizontal="center")
    put(7, "centered text in C: spill both ways")
    ws["A8"] = LONG
    ws["A8"].alignment = Alignment(wrap_text=True)
    put(8, "wrapped text: no spill")
    ws["A9"] = 1234567890123.0
    put(9, "General number wider than the column, no width set")
    ws["A10"] = 123456789
    ws["A10"].number_format = "#,##0.00"
    put(10, "formatted number wider than the column: ####")
    ws["A11"] = 123456789
    ws["A11"].number_format = "#,##0.00"
    ws.merge_cells("A11:C11")
    put(11, "same number merged over A:C: #### or text?")
    ws["D12"] = LONG
    put(12, "text in D spilling across hidden column E")
    ws["D13"] = "merged across a hidden column"
    ws.merge_cells("D13:F13")
    put(13, "merge D:F containing hidden E")
    ws["A14"] = 42
    ws["B14"] = "text"
    ws["C14"] = True
    ws["D14"] = "=1/0"
    ws["F14"] = 45000
    ws["F14"].number_format = "yyyy-mm-dd"
    put(14, "default alignment: number, text, bool, error, (E hidden), date")
    ws["A15"] = "centered across selection"
    for col in "ABC":
        ws[f"{col}15"].alignment = Alignment(horizontal="centerContinuous")
    put(15, "centerContinuous over A:C")
    ws["A16"] = 12345
    ws["A16"].number_format = "@"
    put(16, "number with @ format")
    ws["A17"] = "hidden row"
    ws.row_dimensions[17].hidden = True
    put(17, "hidden row")
    ws["H18"] = "=1/0"
    ws["H19"] = "=NA()"
    ws["H20"] = True
    ws["H21"] = 1234567
    put(18, "error #DIV/0! in narrow column H (width 3)")
    put(19, "error #N/A in narrow column H")
    put(20, "TRUE in narrow column H")
    put(21, "General 1234567 in narrow column H")
    ws["G22"] = "text in a width-0 column"
    put(22, "G has width 0 without the hidden flag")
    ws["A23"] = LONG
    ws["B23"] = "=A99"
    put(23, "left text, B a formula returning 0 (empty reference)")
    ws["A24"] = "  indented"
    ws["A24"].alignment = Alignment(indent=2)
    put(24, "indent 2")

    ws.column_dimensions["E"].hidden = True
    ws.column_dimensions["G"].width = 0
    ws.column_dimensions["H"].width = 3
    for row, purpose in note.items():
        ws[f"J{row}"] = purpose
    ws.column_dimensions["J"].width = 60

    CONTAINER_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(WORKBOOK)
    return WORKBOOK


def run() -> None:
    """Build, export through Excel, copy the export into the build directory."""
    build()
    for stale in (EXPORT, CONTAINER_DIR / "html_probe_files"):
        if stale.is_dir():
            shutil.rmtree(stale)
        elif stale.exists():
            stale.unlink()
    script = Path(__file__).with_name("html_export.applescript")
    result = subprocess.run(
        ["osascript", str(script), str(WORKBOOK), str(EXPORT), "A1:H24"],
        capture_output=True,
        text=True,
        check=False,
    )
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    (BUILD_DIR / "string_values.txt").write_text(result.stdout)
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    shutil.copy2(EXPORT, BUILD_DIR / EXPORT.name)
    files = CONTAINER_DIR / "html_probe_files"
    if files.is_dir():
        shutil.copytree(files, BUILD_DIR / files.name, dirs_exist_ok=True)
    print(f"export copied to {BUILD_DIR}")


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
