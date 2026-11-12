"""Size and time of the HTML output for a large sheet (ARCHITECTURE.md §7.3).

    uv run python bench/html_size.py [DIR] [ROWS] [COLUMNS] [plain|features]

Writes a workbook of ROWS x COLUMNS cells (default 5,000 x 20, 100,000 cells) into DIR (default:
a temporary directory) unless it exists, then reads it, lays it out in excel mode and renders
HTML, and reports the size and the time of each step and the peak RSS (``ru_maxrss``: bytes on
macOS, kilobytes on Linux).

- ``plain`` (the default): numbers with a format, texts (some long enough to spill into an empty
  column), dates, integers, a bold filled header and a fill on every 50th row: what a sheet
  without the newer features costs.
- ``features``: every cell uses a feature 0.2.0 did not draw: borders of several
  styles and diagonals, pattern and gradient fills, underlines, strikethrough, superscript and
  subscript, theme fonts, rich-text runs (as shared strings, as Excel writes them), wrapped,
  shrunk, distributed and indented text, reading order, angled and stacked text, hyperlinks,
  grouped rows and columns, merges, a grid color and a tab color.

``bench/README.md`` records the figures.
"""

from __future__ import annotations

import datetime as dt
import io
import random
import re
import resource
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, GradientFill, PatternFill, Side

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


def build_features(path: Path, rows: int, columns: int) -> None:
    rng = random.Random(3)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Features"
    thin, double = Side("thin", color="FF404040"), Side("double", color="FF0000C0")
    boxed = Border(left=thin, right=thin, top=thin, bottom=thin)
    crossed = Border(bottom=double, diagonal=Side("dashed"), diagonalUp=True)
    pattern = PatternFill("gray125", fgColor="FF4472C4", bgColor="FFFFFFFF")
    gradient = GradientFill(stop=("FFFFFFFF", "FF70AD47"))
    runs = CellRichText(
        [TextBlock(InlineFont(b=True, color="FFC00000"), "bold red"), " plain ",
         TextBlock(InlineFont(vertAlign="superscript"), "2")]
    )  # fmt: skip
    for c in range(1, columns + 1):
        ws.cell(1, c, f"Column {c}").font = Font(bold=True, underline="double")
    for r in range(2, rows + 1):
        for c in range(1, columns + 1):
            cell = ws.cell(r, c)
            kind = (c - 1) % 10
            if kind == 0:
                cell.value, cell.number_format = rng.random() * 1e6, "#,##0.00"
                cell.border = boxed
                cell.font = Font(underline="singleAccounting", color="FF7030A0")
            elif kind == 1:
                cell.value = f"styled {r}"
                cell.font = Font(italic=True, underline="single", strike=True, scheme="minor")
                cell.fill = pattern
            elif kind == 2:
                cell.value = f"wrapped text in row {r} that needs two lines"
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                cell.fill = gradient
            elif kind == 3:
                cell.value = dt.datetime(2020, 1, 1) + dt.timedelta(seconds=rng.randrange(10**8))
                cell.number_format = "yyyy-mm-dd hh:mm"
                cell.border = crossed
            elif kind == 4:
                cell.value = runs
            elif kind == 5:
                cell.value = f"angled {r}" + (" and long enough to spill" if r % 3 == 0 else "")
                cell.alignment = Alignment(textRotation=45 if r % 2 else 135)
            elif kind == 6:
                cell.value = f"shrunk to fit in its cell {r}"
                cell.alignment = Alignment(shrink_to_fit=True)
            elif kind == 7:
                cell.value = f"link {r}"
                cell.hyperlink = f"https://example.com/{r}"
            elif kind == 8:
                cell.value = f"spread words {r}"
                align = ("distributed", "justify", "right")[r % 3]
                cell.alignment = Alignment(horizontal=align, indent=1, readingOrder=1)
            elif r % 100 == 0:
                cell.value = "Stack"
                cell.alignment = Alignment(textRotation=255)
            else:
                cell.value = f"sub {r}"
                cell.font = Font(vertAlign="subscript", color="FF00B050")
        if r % 200 == 0:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    for start in range(10, rows, 20):
        ws.row_dimensions.group(start, start + 4, outline_level=1)
    ws.column_dimensions.group("C", "D", outline_level=1)
    ws.sheet_view.defaultGridColor = False
    ws.sheet_view.colorId = 10
    ws.sheet_properties.tabColor = "FF4472C4"
    wb.save(path)
    _share_strings(path)


def _share_strings(path: Path) -> None:
    """Turn the first sheet's inline strings into shared strings, as Excel writes text."""
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    table: dict[str, int] = {}

    def share(match: re.Match[str]) -> str:
        index = table.setdefault(match[3], len(table))
        return f'<c{match[1]} t="s"{match[2]}><v>{index}</v></c>'

    sheet_pattern = re.compile(r'<c([^>]*?) t="inlineStr"([^>]*)><is>(.*?)</is></c>', re.S)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                data = sheet_pattern.sub(share, data.decode()).encode()
            elif info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                    b'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
                )
            elif info.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/'
                    b'officeDocument/2006/relationships/sharedStrings" '
                    b'Target="sharedStrings.xml"/></Relationships>',
                )
            archive.writestr(info, data)
        main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        items_xml = "".join(f"<si>{inner}</si>" for inner in table)
        archive.writestr(
            "xl/sharedStrings.xml",
            f'<sst xmlns="{main}" count="{len(table)}" uniqueCount="{len(table)}">'
            f"{items_xml}</sst>",
        )


def main() -> None:
    directory = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(tempfile.mkdtemp())
    rows = int(sys.argv[2]) if len(sys.argv) > 2 else 5_000
    columns = int(sys.argv[3]) if len(sys.argv) > 3 else 20
    case = sys.argv[4] if len(sys.argv) > 4 else "plain"
    if case not in ("plain", "features"):
        raise SystemExit("the case is plain or features")
    name = "html" if case == "plain" else "features"
    path = directory / f"{name}_{rows}x{columns}.xlsx"
    if not path.exists():
        (build if case == "plain" else build_features)(path, rows, columns)
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
