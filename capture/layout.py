"""Capture how Excel for Mac paints a sheet, from its PDF export (route B).

    uv run --group capture python -m capture.layout run     # build, export, read, write the TSV
    uv run --group capture python -m capture.layout read VERSION DATE  # an earlier export again

The workbook is built with openpyxl inside Excel's sandbox container, with its page setup in the
file (landscape, 100 % scale, gridlines printed, small margins), so one PDF point is one sheet
point. Excel draws each gridline as a 1 pt rectangle and leaves a vertical one out wherever
text crosses it, so the missing lines of a row show how far its text spills; the text runs give
where the glyphs sit and their color.

Columns: A to H at the default width (53 pt in Calibri 11), I to M for the hidden-column cases
(K hidden), N narrow (3 digits) for overflow. Every row has the same height, so rows are found
from the page's top line and the pitch of its horizontal lines. The output,
``tests/captures/layout.tsv``, has one record per row: the case, the column boundaries Excel left
out, each text run with its first and last x and color, the fills, and each run's baseline.
"""

from __future__ import annotations

import datetime as dt
import itertools
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "layout"
OUTPUT = ROOT / "tests" / "captures" / "layout.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
)
WORKBOOK = CONTAINER_DIR / "layout.xlsx"
EXPORT = CONTAINER_DIR / "layout.pdf"
LAST_COLUMN = 16  # P; O has width 0 in the file (written directly: openpyxl drops a 0 width)
YELLOW = "FFFFFF00"
BLUE = "FF9BC2E6"
ROW_HEIGHT = 15.0  # points, set on every row; Excel for Mac prints these rows 14 pt apart
LONG = "spilled text that is much wider than one column"


@dataclass
class Case:
    """One row of the capture: its purpose, and the cells to write as (column, value, options)."""

    name: str
    cells: list[tuple[str, object, dict[str, object]]] = field(default_factory=list)
    merge: str | None = None  # a merged range starting in this row


def cases() -> list[Case]:
    """Every row of the workbook, in order."""
    result: list[Case] = []

    def add(
        name: str, *cells: tuple[str, object, dict[str, object]], merge: str | None = None
    ) -> None:
        result.append(Case(name, list(cells), merge))

    # 1. How far spill reaches, by alignment: zeros measure 6 pt each in Calibri 11.
    for n in range(6, 41):
        add(f"spill left zeros={n}", ("A", "0" * n, {}))
    for n in range(6, 41):
        add(f"spill right zeros={n}", ("H", "0" * n, {"horizontal": "right"}))
    for n in range(6, 41):
        add(f"spill center zeros={n}", ("E", "0" * n, {"horizontal": "center"}))
    for n in range(1, 9):
        add(f"center fits zeros={n}", ("E", "0" * n, {"horizontal": "center"}))
    for text in ("i" * 30, "W" * 12, LONG):
        add(f"spill left text={text!r}", ("A", text, {}))
        add(f"spill center text={text!r}", ("E", text, {"horizontal": "center"}))
    # Clipping: does the PDF hold the clipped characters?
    add("blocked left", ("A", LONG, {}), ("B", 5, {}))
    add("blocked right", ("D", LONG, {"horizontal": "right"}), ("C", 5, {}))
    add("blocked center", ("D", LONG, {"horizontal": "center"}), ("C", 5, {}), ("E", 5, {}))
    # 2. Competing spills and hidden values.
    add("compete A left C right", ("A", LONG, {}), ("C", LONG, {"horizontal": "right"}))
    add("compete A left D right", ("A", LONG, {}), ("D", LONG, {"horizontal": "right"}))
    add(
        "compete A left 15 C right 15",
        ("A", "0" * 15, {}),
        ("C", "0" * 15, {"horizontal": "right"}),
    )
    add("compete C center E center", ("C", LONG, {"horizontal": "center"}),
        ("E", LONG, {"horizontal": "center"}))  # fmt: skip
    add("hidden K empty", ("I", LONG, {}))
    add("hidden K holds a value", ("I", LONG, {}), ("K", "x", {}))
    add("hidden K holds a value, right", ("M", LONG, {"horizontal": "right"}), ("K", "x", {}))
    # 3. Overflow in the narrow column N.
    add("narrow -12345 0;[Red]-0", ("N", -12345, {"format": "0;[Red]-0"}))
    add("narrow 12345 0;[Red]-0", ("N", 12345, {"format": "0;[Red]-0"}))
    add("narrow -1234.5 [Blue]General", ("N", -1234.5, {"format": "[Blue]General"}))
    add("narrow 12345 @", ("N", 12345, {"format": "@"}))
    add("1234567890 @ in default-width M, N empty", ("M", 1234567890, {"format": "@"}))
    add("narrow 12345 General", ("N", 12345, {}))
    add("narrow #DIV/0! red font", ("N", "=1/0", {"color": "FFFF0000"}))
    add("wide 12345 @", ("A", 12345, {"format": "@"}))
    add("wide 12345 General", ("A", 12345, {}))
    # 4. Indent per level, left and right.
    for level in range(6):
        add(f"indent left {level}", ("A", "L", {"indent": level}))
    for level in range(6):
        add(f"indent right {level}", ("B", "R", {"horizontal": "right", "indent": level}))
    add(f"indent left 2 spill text={LONG!r}", ("A", LONG, {"indent": 2}))
    # Merged anchors over three rows, side by side: top, center, bottom, default.
    add(
        "merge vertical: A top, C center, E bottom, G default",
        ("A", "top", {"vertical": "top"}),
        ("C", "center", {"vertical": "center"}),
        ("E", "bottom", {"vertical": "bottom"}),
        ("G", "default", {}),
        merge="A:B C:D E:F G:H",
    )
    add("merge vertical (second row)")
    add("merge vertical (third row)")
    add("merge horizontal centered, A:D", ("A", "centered over four", {"horizontal": "center"}),
        merge="A:D")  # fmt: skip
    # 5. Rows 156 on: blocked centered text, cells that do not spill, line breaks, width-0
    #    columns, fills under merges and spills.
    add("center blocked left: C value, D centered", ("C", 5, {}),
        ("D", LONG, {"horizontal": "center"}))  # fmt: skip
    add("center blocked right: D centered, E value", ("D", LONG, {"horizontal": "center"}),
        ("E", 5, {}))  # fmt: skip
    add("center zeros=12 blocked left", ("C", 5, {}), ("D", "0" * 12, {"horizontal": "center"}))
    add("fill short", ("A", "ab", {"horizontal": "fill"}))
    add("fill long, B empty", ("A", LONG, {"horizontal": "fill"}))
    add("justify long, B empty", ("A", LONG, {"horizontal": "justify"}))
    add("distributed long, B empty", ("A", LONG, {"horizontal": "distributed"}))
    add("wrapped long, B empty", ("A", LONG, {"wrap": True}))
    add("spacer after wrapped")
    add("line break without wrap", ("A", "first line\nsecond line", {}))
    add("spacer after line break without wrap")
    add("line break with wrap", ("A", "first line\nsecond line", {"wrap": True}))
    add("spacer after line break with wrap")
    add("width-0 O empty, M spills", ("M", LONG, {}))
    add("width-0 O holds text", ("O", "zero width text", {}))
    add("width-0 O holds a value, M spills", ("M", LONG, {}), ("O", "x", {}))
    add("merge fill A:C: anchor yellow, covered B blue", ("A", "merged", {"fill": YELLOW}),
        ("B", None, {"fill": BLUE}), merge="A:C")  # fmt: skip
    add("merge fill A:C: anchor none, covered B blue", ("A", "merged", {}),
        ("B", None, {"fill": BLUE}), merge="A:C")  # fmt: skip
    add("spill over fill: A text, B yellow", ("A", LONG, {}), ("B", None, {"fill": YELLOW}))
    add("spill from a filled cell: A yellow text, B empty", ("A", LONG, {"fill": YELLOW}))
    # 6. Merged text wider than its merge (rows 176 on): does it spill into empty neighbors?
    add("merge A:B left, text wider, C empty", ("A", LONG, {}), merge="A:B")
    add("merge C:D right, text wider, B empty", ("C", LONG, {"horizontal": "right"}),
        merge="C:D")  # fmt: skip
    add("merge C:D centered, text wider, B and E empty", ("C", LONG, {"horizontal": "center"}),
        merge="C:D")  # fmt: skip
    add("center across A:B, text wider, C empty", ("A", LONG, {"horizontal": "centerContinuous"}),
        ("B", None, {"horizontal": "centerContinuous"}))  # fmt: skip
    return result


def build(path: Path = WORKBOOK) -> Path:
    """Write the capture workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Layout"
    for row, case in enumerate(cases(), 1):
        ws.row_dimensions[row].height = ROW_HEIGHT
        # Merge first, so that a covered cell can still be given a fill of its own.
        if case.merge:
            for span in case.merge.split():
                first, last = span.split(":")
                height = 3 if case.name.startswith("merge vertical") else 1
                ws.merge_cells(f"{first}{row}:{last}{row + height - 1}")
        for column, value, options in case.cells:
            cell = ws[f"{column}{row}"]
            if value is not None:
                cell.value = value
            if "format" in options:
                cell.number_format = str(options["format"])
            if "color" in options:
                cell.font = openpyxl.styles.Font(color=str(options["color"]))
            if "fill" in options:
                cell.fill = PatternFill("solid", fgColor=str(options["fill"]))
            alignment = {
                k: v for k, v in options.items() if k in ("horizontal", "vertical", "indent")
            }
            if "wrap" in options:
                alignment["wrap_text"] = options["wrap"]
            if alignment:
                cell.alignment = Alignment(**alignment)  # type: ignore[arg-type]
    ws.column_dimensions["K"].hidden = True
    ws.column_dimensions["N"].width = 3.7109375  # three digits in Calibri 11
    ws.page_setup.orientation = "landscape"
    ws.page_setup.scale = 100
    ws.print_options.gridLines = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.25
    ws.page_margins.header = ws.page_margins.footer = 0
    ws.print_area = f"A1:{get_column_letter(LAST_COLUMN)}{len(cases())}"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _set_zero_width(path, "O")
    return path


def _set_zero_width(path: Path, letter: str) -> None:
    """Give a column a width of 0 in the file, which openpyxl will not write."""
    index = column_index_from_string(letter)
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                col = f'<col min="{index}" max="{index}" width="0" customWidth="1"/>'
                assert text.count("</cols>") == 1
                data = text.replace("</cols>", col + "</cols>").encode()
            archive.writestr(info, data)


def export() -> str:
    """Have Excel save the workbook as PDF; return Excel's version."""
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
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True,
        text=True,
        check=True,
    )
    return version.stdout.strip()


@dataclass
class Run:
    text: str
    x0: float
    x1: float
    color: str
    baseline: float = 0.0
    """Points below the top edge of the row's top gridline."""


@dataclass
class Fill:
    x0: float
    x1: float
    color: str


def read(path: Path) -> list[tuple[list[str], list[Run], list[Fill]]]:
    """Per sheet row: the column boundaries left out, the text runs and the cell fills."""
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LAParams, LTChar, LTRect, LTTextContainer

    rows: list[tuple[list[str], list[Run], list[Fill]]] = []
    for page in extract_pages(path, laparams=LAParams(char_margin=2000, line_margin=0.01)):
        rects = [r for r in page if isinstance(r, LTRect)]
        horizontal = [r for r in rects if r.height <= 1.01 and r.width > 1.01]
        vertical = [r for r in rects if r.width <= 1.01 and r.height > 1.01]
        filled = [r for r in rects if r.width > 1.01 and r.height > 1.01]
        lines = sorted({round(r.y0, 2) for r in horizontal}, reverse=True)
        pitch = min(a - b for a, b in itertools.pairwise(lines))
        top, bottom = lines[0] + 1, lines[-1] + 1
        count = round((top - bottom) / pitch)
        boundaries = sorted({round(r.x0, 1) for r in vertical})
        chars = [
            c
            for box in page
            if isinstance(box, LTTextContainer)
            for line in box
            for c in line
            if isinstance(c, LTChar)
        ]
        for k in range(count):
            high = top - k * pitch
            low = high - pitch
            missing = []
            for index, x in enumerate(boundaries[1:-1], 1):
                drawn = any(
                    round(r.x0, 1) == x and r.y0 <= low + 2 and r.y1 >= high - 2 for r in vertical
                )
                if not drawn:
                    missing.append(str(index))
            mine = sorted((c for c in chars if low < (c.y0 + c.y1) / 2 < high), key=lambda c: c.x0)
            runs: list[Run] = []
            for char in mine:
                color = _color(char)
                if runs and abs(char.x0 - runs[-1].x1) < 1.5 and runs[-1].color == color:
                    runs[-1].text += char.get_text()
                    runs[-1].x1 = char.x1
                else:
                    runs.append(
                        Run(char.get_text(), char.x0, char.x1, color, high - char.matrix[5])
                    )
            fills = [
                Fill(r.x0, r.x1, _rgb(r.non_stroking_color))
                for r in sorted(filled, key=lambda r: r.x0)
                if low < (r.y0 + r.y1) / 2 < high and _rgb(r.non_stroking_color) != "FFFFFF"
            ]
            rows.append((missing, runs, fills))
    return rows


def _color(char: object) -> str:
    return _rgb(getattr(getattr(char, "graphicstate", None), "ncolor", None))


def _rgb(color: object) -> str:
    if isinstance(color, tuple | list) and len(color) == 3:
        return "".join(f"{round(v * 255):02X}" for v in color)
    if isinstance(color, int | float):
        return f"{round(color * 255):02X}" * 3
    return "000000"


def write(version: str, captured: str | None = None) -> None:
    """Read the export and write the TSV with its header; ``captured`` is the export's date."""
    rows = read(BUILD_DIR / EXPORT.name)
    named = cases()
    if len(rows) != len(named):
        sys.exit(f"found {len(rows)} rows in the PDF, expected {len(named)}")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", encoding="utf-8") as out:
        out.write(f"# suite: layout (capture.layout); Excel {version} for Mac, en_US\n")
        out.write("# font: Calibri 11 (openpyxl default); columns A-H default width (53 pt), ")
        out.write("K hidden, N width 3.7109375 in the file (22 pt), O width 0 in the file; rows ")
        out.write("15 pt in the file, 14 pt apart in the PDF\n")
        captured = captured or dt.date.today().isoformat()
        out.write(f"# captured: {captured}; method: PDF export, pdfminer.six\n")
        out.write("# missing: boundaries without a gridline in the row, 1 = between the first ")
        out.write("and second visible columns (A|B); runs: text@x0-x1#RRGGBB in PDF points; ")
        out.write("fills: x0-x1#RRGGBB of filled areas other than white; baselines: per run, ")
        out.write("points from the top edge of the row's top gridline down to the baseline\n")
        out.write("row\tcase\tmissing\truns\tfills\tbaselines\n")
        for number, (case, (missing, runs, fills)) in enumerate(zip(named, rows, strict=True), 1):
            text = " ".join(f"{r.text!r}@{r.x0:.2f}-{r.x1:.2f}#{r.color}" for r in runs)
            areas = " ".join(f"{f.x0:.2f}-{f.x1:.2f}#{f.color}" for f in fills)
            baselines = " ".join(f"{r.baseline:.2f}" for r in runs)
            out.write(f"{number}\t{case.name}\t{','.join(missing)}\t{text}\t{areas}\t{baselines}\n")
    print(f"{len(rows)} rows written to {OUTPUT}")


if __name__ == "__main__":
    command = sys.argv[1]
    if command == "build":
        print(build())
    elif command == "run":
        build()
        write(export())
    elif command == "read":
        write(*sys.argv[2:4]) if len(sys.argv) > 2 else write("unknown")
