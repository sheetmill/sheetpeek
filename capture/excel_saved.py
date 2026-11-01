"""Have Excel save a small workbook, to see the XML Excel itself writes for formula results.

    uv run --group capture python -m capture.excel_saved run

Writes formulas with openpyxl (which stores no cached values), has Excel open the workbook and
save a copy (``save_copy.applescript``), prints each cell element of the copy's first sheet, and
keeps the copy as ``tests/files/excel_formulas.xlsx``, a real Excel-saved file for the reader's
tests, without the two personal items Excel saves: the folder it was saved in
(``x15ac:absPath``) and the user who saved it (``cp:lastModifiedBy``).
"""

from __future__ import annotations

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
)
SOURCE = CONTAINER_DIR / "excel_formulas_source.xlsx"
COPY = CONTAINER_DIR / "excel_formulas.xlsx"
KEPT = ROOT / "tests" / "files" / "excel_formulas.xlsx"
FORMULAS = ['=""', "=1+1", '="a"&"b"', "=1/0", "=A99", "=TRUE", '=" "', "=NA()", "=DATE(1900,2,29)"]


def build() -> None:
    """Write the formulas in column A, with a note on each in column B."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formulas"
    for row, formula in enumerate(FORMULAS, 1):
        ws.cell(row, 1, formula)
        ws.cell(row, 2, f"formula {formula}")
    ws["A9"].number_format = "yyyy-mm-dd"
    CONTAINER_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(SOURCE)


def run() -> None:
    """Build, have Excel save the copy, print its cell elements, keep it for the tests."""
    build()
    if COPY.exists():
        COPY.unlink()
    script = Path(__file__).with_name("save_copy.applescript")
    result = subprocess.run(
        ["osascript", str(script), str(SOURCE), str(COPY)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    with zipfile.ZipFile(COPY) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    for element in re.findall(r'<c r="A\d+"[^>]*?(?:/>|>.*?</c>)', xml):
        print(element)
    KEPT.parent.mkdir(parents=True, exist_ok=True)
    scrub(COPY, KEPT)
    print(f"kept {KEPT} ({KEPT.stat().st_size} bytes)")


def scrub(source: Path, target: Path) -> None:
    """Copy a workbook Excel saved without its saved path and its last-modified-by."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/workbook.xml":
                data = re.sub(
                    rb'<mc:AlternateContent[^>]*><mc:Choice Requires="x15">'
                    rb"<x15ac:absPath[^>]*/></mc:Choice></mc:AlternateContent>",
                    b"",
                    data,
                )
            elif info.filename == "docProps/core.xml":
                data = re.sub(rb"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", b"", data)
            zout.writestr(info, data)


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
