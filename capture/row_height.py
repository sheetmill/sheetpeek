"""Capture the row height Excel for Mac gives a row without a stored height, per Normal font.

    uv run --group capture python -m capture.row_height build   # write the workbooks only
    uv run --group capture python -m capture.row_height run     # build, read in Excel, write TSV

Each workbook has one font xlformat measures as its Normal style, written as xlformat's captures
write it (the theme's minor scheme only for Calibri, so Excel does not substitute the theme font
for Aptos Narrow), a number in A1, an empty cell element in row 3 and an empty row element for
row 5 (both written into the XML; rows 2 and 100 have no element), and a ``<sheetFormatPr>``
written exactly as the case gives it, or none. The cases:

- every measured font (Calibri and Aptos Narrow at 9, 10, 11, 12, 14 and 16 points) without
  ``<sheetFormatPr>``, then with openpyxl's own ``baseColWidth="8" defaultRowHeight="15"``;
- at Calibri 11 and 14 (whose own row height, 19 pt, differs from every stored value):
  ``customHeight="1"`` with stored heights of 15, 20, 30 and 12.75, other stored heights without
  it, and ``zeroHeight``, ``thickTop`` and ``thickBottom``;
- at Calibri 11 with openpyxl's ``<sheetFormatPr>``, A1 in another font size, bold, empty but
  styled, wrapped, merged, or in a row with a stored height (``CELLS``): whether Excel fits a row to
  its cells when it opens the file.

``row_height.applescript`` reads ``row height`` of rows 1, 2, 3, 5 and 100, the worksheet's
``standard height``, and the font Excel applied to A1, which shows a
substitution. The output is ``tests/captures/row_height.tsv``. New cases go at the end.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils.indexed_list import IndexedList

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "row_height.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "row_height"
)
FAMILIES = ("Calibri", "Aptos Narrow")
SIZES = (9, 10, 11, 12, 14, 16)
SHEET = "xl/worksheets/sheet1.xml"
OPENPYXL = 'baseColWidth="8" defaultRowHeight="15"'
PROBES = (
    'baseColWidth="8" defaultRowHeight="15" customHeight="1"',
    'baseColWidth="8" defaultRowHeight="20" customHeight="1"',
    'baseColWidth="8" defaultRowHeight="30" customHeight="1"',
    'baseColWidth="8" defaultRowHeight="12.75" customHeight="1"',
    'baseColWidth="8" defaultRowHeight="20"',
    'baseColWidth="8" defaultRowHeight="30"',
    'baseColWidth="8" defaultRowHeight="15" zeroHeight="1"',
    'baseColWidth="8" defaultRowHeight="20" customHeight="1" zeroHeight="1"',
    'baseColWidth="8" defaultRowHeight="15" thickTop="1"',
    'baseColWidth="8" defaultRowHeight="15" thickBottom="1"',
    'baseColWidth="8" defaultRowHeight="15" thickTop="1" thickBottom="1"',
)
LONG = "spilled text that is much wider than one column"
# What A1 holds, by label: (value, font size or None for the Normal font, bold, wrap, row ht,
# merged range or None)
CELLS: dict[str, tuple[object, int | None, bool, bool, float | None, str | None]] = {
    "number": (1, None, False, False, None, None),
    "font 14": (1, 14, False, False, None, None),
    "font 14 bold": (1, 14, True, False, None, None),
    "font 9": (1, 9, False, False, None, None),
    "font 14, empty": (None, 14, False, False, None, None),
    "wrap, two lines": ("first line\nsecond line", None, False, True, None, None),
    "wrap, long text": (LONG, None, False, True, None, None),
    "font 14, row ht 15": (1, 14, False, False, 15.0, None),
    "font 14, merged A1:C1": (1, 14, False, False, None, "A1:C1"),
    "wrap, two lines, merged A1:C1": ("first line\nsecond line", None, False, True, None, "A1:C1"),
}
# (family, size, the attributes of <sheetFormatPr> or None for no element, what A1 holds)
CASES: list[tuple[str, int, str | None, str]] = [
    *((f, s, None, "number") for f in FAMILIES for s in SIZES),
    *((f, s, OPENPYXL, "number") for f in FAMILIES for s in SIZES),
    *(("Calibri", s, attributes, "number") for s in (11, 14) for attributes in PROBES),
    *(("Calibri", 11, OPENPYXL, cell) for cell in list(CELLS)[1:]),
]


def build_one(number: int, family: str, size: int, sheet_format: str | None, cell: str) -> Path:
    """Workbook ``number``: the Normal style ``family`` at ``size``, ``sheet_format``, A1."""
    wb = openpyxl.Workbook()
    normal = Font(name=family, sz=size, family=2, scheme="minor" if family == "Calibri" else None)
    wb._fonts = IndexedList([normal])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = normal
    ws = wb.active
    ws.title = "rows"
    value, font_size, bold, wrap, height, merge = CELLS[cell]
    ws["A1"] = value
    if font_size is not None:
        ws["A1"].font = Font(name=family, sz=font_size, b=bold, family=2)
    if wrap:
        ws["A1"].alignment = Alignment(wrap_text=True)
    if height is not None:
        ws.row_dimensions[1].height = height
    if merge is not None:
        ws.merge_cells(merge)
    path = CONTAINER_DIR / f"row_height_{number:02}.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    element = "" if sheet_format is None else f"<sheetFormatPr {sheet_format}/>"
    _rewrite_sheet(path, element)
    return path


def _rewrite_sheet(path: Path, element: str) -> None:
    """Replace ``<sheetFormatPr>`` with ``element``; add row 3 (an empty cell) and row 5."""
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == SHEET:
                text, count = re.subn(r"<sheetFormatPr\b[^>]*/>", element, data.decode())
                if count != 1:
                    sys.exit(f"{path.name}: expected one <sheetFormatPr/>, found {count}")
                rows = '<row r="3"><c r="A3"/></row><row r="5"/></sheetData>'
                text, count = re.subn(r"</sheetData>", rows, text)
                if count != 1:
                    sys.exit(f"{path.name}: expected one </sheetData>, found {count}")
                data = text.encode()
            archive.writestr(info, data)


def build() -> list[Path]:
    """Write every case's workbook; return their paths in case order."""
    return [build_one(k, *case) for k, case in enumerate(CASES, 1)]


def run() -> None:
    """Build, have Excel read the row heights, write the TSV."""
    paths = build()
    script = Path(__file__).with_name("row_height.applescript")
    result = subprocess.run(
        ["osascript", str(script), *(str(path) for path in paths)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    lines = result.stdout.strip().split("\n")
    if len(lines) != len(CASES):
        sys.exit(f"expected {len(CASES)} lines from Excel, got {len(lines)}: {result.stdout!r}")
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: row_height (capture.row_height); Excel {version} for Mac, en_US\n")
        out.write(
            "# one workbook per record: the Normal style set to the font, a number in A1, and "
            "<sheetFormatPr> with exactly the attributes given (none: no element)\n"
        )
        out.write(f"# captured: {dt.date.today().isoformat()}; method: row height (AppleScript)\n")
        out.write(
            "# cell: what A1 holds (CELLS in capture/row_height.py); applied: the font Excel "
            "applied to A1; row1: A1's row; row2: no row "
            "element; row3: a row element with an empty cell element; row5: an empty row "
            "element; row100: no row element, far below; standard: the worksheet's standard "
            "height\n"
        )
        out.write(
            "family\tsize\tsheet_format\tcell\tapplied\trow1\trow2\trow3\trow5\trow100\tstandard\n"
        )
        for (family, size, sheet_format, cell), path, line in zip(CASES, paths, lines, strict=True):
            name, applied_name, applied_size, *heights = line.split("\t")
            if name != path.name:
                sys.exit(f"Excel answered for {name}, expected {path.name}")
            applied = f"{applied_name} {_number(applied_size)}"
            row = [family, str(size), sheet_format or "none", cell, applied]
            row += [_number(h) for h in heights]
            out.write("\t".join(row) + "\n")
            print("\t".join(row))
    print(f"wrote {OUTPUT}")


def _number(text: str) -> str:
    """``15.0`` as ``15``; other text unchanged."""
    try:
        value = float(text)
    except ValueError:
        return text
    return repr(int(value)) if value.is_integer() else repr(value)


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
