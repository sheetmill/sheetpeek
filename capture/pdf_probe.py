"""Probe: can Excel for Mac's PDF export show where text is painted (route B)?

Builds a small workbook inside Excel's sandbox container, has Excel open it and save it as PDF
(``pdf_export.applescript``), then reads the PDF's text runs and lines with their coordinates.

    uv run --group capture python -m capture.pdf_probe build    # write the workbook only
    uv run --group capture python -m capture.pdf_probe run      # build, export, read
    uv run --group capture python -m capture.pdf_probe read     # read an earlier export

The page setup is part of the workbook: landscape, 100 % scale, gridlines printed, small margins,
so that one point in the PDF is one point on the sheet. Rows 1 and 2 calibrate: a left-aligned
``L`` and a right-aligned ``R`` in every column give each column's text edges.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "pdf_probe"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
)
WORKBOOK = CONTAINER_DIR / "pdf_probe.xlsx"
EXPORT = CONTAINER_DIR / "pdf_probe.pdf"
COLUMNS = "ABCDEFGH"
LONG = "spilled text that is much wider than one column"


def build() -> Path:
    """Write the probe workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Probe"
    for col in COLUMNS:
        ws[f"{col}1"] = "L"
        ws[f"{col}1"].alignment = Alignment(horizontal="left")
        ws[f"{col}2"] = "R"
        ws[f"{col}2"].alignment = Alignment(horizontal="right")
    ws["A3"] = LONG
    ws["E4"] = LONG
    ws["E4"].alignment = Alignment(horizontal="right")
    ws["E5"] = LONG
    ws["E5"].alignment = Alignment(horizontal="center")
    ws.page_setup.orientation = "landscape"
    ws.page_setup.scale = 100
    ws.print_options.gridLines = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.print_area = f"A1:{COLUMNS[-1]}5"
    CONTAINER_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(WORKBOOK)
    return WORKBOOK


def export() -> None:
    """Have Excel save the workbook as PDF and copy the PDF into the build directory."""
    if EXPORT.exists():
        EXPORT.unlink()
    script = Path(__file__).with_name("pdf_export.applescript")
    result = subprocess.run(
        ["osascript", str(script), str(WORKBOOK), str(EXPORT)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(EXPORT, BUILD_DIR / EXPORT.name)


def read(path: Path) -> None:
    """Print each text line's characters with their x-coordinates, and the vertical lines."""
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LAParams, LTChar, LTCurve, LTTextContainer

    for number, page in enumerate(extract_pages(path, laparams=LAParams(char_margin=50)), 1):
        print(f"page {number}: {page.width:.1f} x {page.height:.1f} pt")
        verticals = sorted(
            {round(item.x0, 2) for item in page if isinstance(item, LTCurve) and item.width < 1}
        )
        print(f"  vertical lines at x = {verticals}")
        for item in page:
            if not isinstance(item, LTTextContainer):
                continue
            for line in item:
                chars = [c for c in line if isinstance(c, LTChar)]
                if not chars:
                    continue
                text = "".join(c.get_text() for c in chars)
                print(
                    f"  y={chars[0].y0:7.2f} x0={chars[0].x0:7.2f} x1={chars[-1].x1:7.2f} "
                    f"{chars[0].fontname} {chars[0].size:.1f}: {text!r}"
                )


if __name__ == "__main__":
    command = sys.argv[1]
    if command == "build":
        print(build())
    elif command == "run":
        build()
        export()
        read(BUILD_DIR / EXPORT.name)
    elif command == "read":
        read(BUILD_DIR / EXPORT.name)
