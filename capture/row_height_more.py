"""Capture row heights under Normal fonts the row-height table has not measured, and wrapped rows.

    uv run --group capture python -m capture.row_height_more build PATH  # the workbooks only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.row_height_more run
    uv run --group capture python -m capture.row_height_more read  # the saved answer again

Like the suite ``row_height`` (``capture/row_height.py``), whose cases and output this suite
leaves alone, it reads what Excel itself reports, not the screen: ``row_height_more.applescript``
opens each workbook, reads ``row height`` and the font Excel applied (name and size) in column
A of the rows each workbook names, and closes it without saving. One workbook per Normal font
(the Normal style is a workbook's own), each with openpyxl's ``<sheetFormatPr
baseColWidth="8" defaultRowHeight="15">`` (no ``customHeight``, so Excel gives every row the
Normal font's height, ``ARCHITECTURE.md`` §4.7):

- ``NORMALS``: Arial, Times New Roman, Courier New and Georgia at 11 and 14 points (faces the
  table has not measured), and Calibri and Aptos Narrow at 13, 15, 17 and 22 (sizes between or
  beyond the measured ones). Row 1 holds a number in the Normal font, row 2 has no ``<row>``
  element, row 3 holds three wrapped lines in the Normal font.
- one more workbook under Calibri 11 (the measured default), ``CELL_ROWS``: a number in each
  of those fonts, a row each (whether a row is fitted to a cell font the table has not
  measured), then three wrapped lines in Calibri 11 (the captured control, 48 pt), Calibri 14
  and 20, Aptos Narrow 11, Arial 11, Courier New 11, Times New Roman 11 and Georgia 11.

It settles two rules of §4.7: a font's row is xlformat's row model at its face and size
(``_rows.font_row_height``), and the
wrapped-row height k(F + 1) plus the inset difference for fonts other than Calibri 11
(``_rows.wrapped_height`` and ``_excel._row_model``). ``EXPECTED`` is what sheetpeek lays out
today for each record, taken by running sheetpeek on the built workbooks; the capture shows
where Excel differs.

``run`` saves Excel's answer, with the Excel it came from, to ``capture/build/row_height_more/
answer.tsv``; ``read`` writes ``tests/captures/row_height_more.tsv`` from it again, one record per
row read. No screenshot, no clipboard, no marker colors; 17 workbooks, 72 rows, about two
minutes of Excel.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils.indexed_list import IndexedList

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "row_height_more.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "row_height_more"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "row_height_more"
)
SCRIPT = Path(__file__).with_name("row_height_more.applescript")
# The font family classes a font element carries (``<family val>``): 1 roman, 2 swiss, 3 modern.
FAMILY_CLASS = {"Times New Roman": 1, "Georgia": 1, "Courier New": 3}
OTHER_FACES = ("Arial", "Times New Roman", "Courier New", "Georgia")
BETWEEN = (13, 15, 17, 22)
NORMALS: tuple[tuple[str, int], ...] = (
    *((face, size) for face in OTHER_FACES for size in (11, 14)),
    *((face, size) for face in ("Calibri", "Aptos Narrow") for size in BETWEEN),
)
CELLS_NORMAL = ("Calibri", 11)
WRAPPED = "a\nb\nc"  # three lines, each far narrower than a default column at 22 pt
# The rows of the Calibri 11 workbook: (label, font face, size, wrapped)
CELL_ROWS: tuple[tuple[str, str, int, bool], ...] = (
    *((f"number in {face} {size}", face, size, False) for face, size in NORMALS),
    *(
        (f"3 wrapped lines in {face} {size}", face, size, True)
        for face, size in (
            ("Calibri", 11),
            ("Calibri", 14),
            ("Calibri", 20),
            ("Aptos Narrow", 11),
            ("Arial", 11),
            ("Courier New", 11),
            ("Times New Roman", 11),
            ("Georgia", 11),
        )
    ),
)
NORMAL_ROWS = (
    (1, "number"),
    (2, "no row element"),
    (3, "3 wrapped lines"),
)
COLUMNS = (
    "workbook",
    "normal",
    "row",
    "label",
    "font",
    "wrapped",
    "applied",
    "height",
)


@dataclass(frozen=True)
class Case:
    """One row read: its workbook, the workbook's Normal font, the row and what it holds."""

    workbook: int  # 1-based, the file ``row_height_more_<NN>.xlsx``
    normal: tuple[str, int]
    row: int
    label: str
    font: tuple[str, int]  # the cell's font (the Normal font where the cell has none)
    wrapped: bool


def _cases() -> list[Case]:
    found = []
    for k, normal in enumerate(NORMALS, 1):
        for row, label in NORMAL_ROWS:
            found.append(Case(k, normal, row, label, normal, row == 3))
    k = len(NORMALS) + 1
    for row, (label, face, size, wrapped) in enumerate(CELL_ROWS, 1):
        found.append(Case(k, CELLS_NORMAL, row, label, (face, size), wrapped))
    return found


CASES = _cases()

# What sheetpeek lays out today for each record (``excel`` mode, ``RenderedRow.height_pt``), by
# (workbook, row); ``tests/test_suite_row_height_more.py`` checks it against sheetpeek.
EXPECTED: dict[tuple[int, int], float] = {
    # Arial at 11 and 14: xlformat's 14 and 18; wrapped three lines of 15 and 19 pt
    **{(1, r): h for r, h in ((1, 14), (2, 14), (3, 45))},
    **{(2, r): h for r, h in ((1, 18), (2, 18), (3, 57))},
    # Times New Roman and Georgia at 11 and 14: xlformat's 14 and 18, wrapped three lines of 15
    # and 19 pt; Courier New's 15 and 19, wrapped 16 and 20 pt
    **{(k, r): h for k in (3, 7) for r, h in ((1, 14), (2, 14), (3, 45))},
    **{(k, r): h for k in (4, 8) for r, h in ((1, 18), (2, 18), (3, 57))},
    **{(5, r): h for r, h in ((1, 15), (2, 15), (3, 48))},
    **{(6, r): h for r, h in ((1, 19), (2, 19), (3, 60))},
    # Calibri 13, 15, 17, 22: xlformat's heights
    **{
        (k, r): h
        for k, base in ((9, 17), (10, 20), (11, 23), (12, 29))
        for r, h in ((1, base), (2, base), (3, 3 * (base + 1)))
    },
    # Aptos Narrow 13, 15, 17, 22
    **{
        (k, r): h
        for k, base in ((13, 18), (14, 20), (15, 23), (16, 29))
        for r, h in ((1, base), (2, base), (3, 3 * (base + 1)))
    },
    # Calibri 11 Normal: a number in each font, then three wrapped lines
    **{
        (17, r): h
        for r, h in enumerate((15, 18, 15, 18, 15, 19, 15, 18, 17, 20, 23, 29, 18, 20, 23, 29), 1)
    },
    **{(17, r): h for r, h in enumerate((48, 60, 81, 48, 46, 48, 46, 46), len(NORMALS) + 1)},
}


def workbook_name(number: int) -> str:
    return f"row_height_more_{number:02}.xlsx"


def _font(face: str, size: int, *, normal: bool = False) -> Font:
    # Only Calibri names the theme's minor font, so Excel does not substitute the theme's font
    # for another face (as capture/row_height.py writes them).
    scheme = "minor" if normal and face == "Calibri" else None
    return Font(name=face, sz=size, family=FAMILY_CLASS.get(face, 2), scheme=scheme)


def build_one(number: int, folder: Path) -> Path:
    """Workbook ``number`` (1-based) in ``folder``: its Normal font and its rows' cells."""
    cases = [c for c in CASES if c.workbook == number]
    wb = openpyxl.Workbook()
    normal = _font(*cases[0].normal, normal=True)
    wb._fonts = IndexedList([normal])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = normal
    ws = wb.active
    ws.title = "rows"
    for case in cases:
        if case.label == "no row element":
            continue
        cell = ws.cell(case.row, 1, WRAPPED if case.wrapped else 1)
        if case.font != case.normal:
            cell.font = _font(*case.font)
        if case.wrapped:
            cell.alignment = Alignment(wrap_text=True)
    path = folder / workbook_name(number)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    """Write every workbook into ``folder``; return their paths in workbook order."""
    return [build_one(k, folder) for k in sorted({c.workbook for c in CASES})]


def _rows_of(number: int) -> str:
    return ",".join(str(c.row) for c in CASES if c.workbook == number)


def run() -> None:
    """Build, have Excel read the rows (needs Excel), save its answer, write the TSV."""
    from capture.edition import excel_description

    paths = build()
    arguments = []
    for k, path in enumerate(paths, 1):
        arguments += [str(path), _rows_of(k)]
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), *arguments],
            capture_output=True,
            text=True,
            check=False,
            timeout=600,  # kills osascript, which ignores SIGTERM while it waits on Excel
        )
    except subprocess.TimeoutExpired:
        sys.exit("osascript did not answer in 10 minutes; check Excel before another run")
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    with ANSWER.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# excel: {excel_description()}, en_US\n")
        out.write(f"# captured: {dt.date.today().isoformat()}\n")
        out.write(result.stdout.strip() + "\n")
    write()


def records(answer: str) -> tuple[list[str], list[dict[str, str]]]:
    """The header lines and the records of a saved answer, checked against ``CASES``.

    Each answer line is ``file name, row, applied font name, applied size, row height``, tab
    separated, in case order.
    """
    lines = answer.strip().split("\n")
    head = [line for line in lines if line.startswith("# ")]
    body = [line for line in lines if line and not line.startswith("# ")]
    if len(body) != len(CASES):
        raise ValueError(f"expected {len(CASES)} lines from Excel, got {len(body)}")
    found = []
    for case, line in zip(CASES, body, strict=True):
        name, row, applied_name, applied_size, height = line.split("\t")
        if (name, int(row)) != (workbook_name(case.workbook), case.row):
            raise ValueError(f"Excel answered for {name} row {row}, expected {case}")
        found.append(
            {
                "workbook": str(case.workbook),
                "normal": f"{case.normal[0]} {case.normal[1]}",
                "row": str(case.row),
                "label": case.label,
                "font": f"{case.font[0]} {case.font[1]}",
                "wrapped": "yes" if case.wrapped else "no",
                "applied": f"{applied_name} {_number(applied_size)}",
                "height": _number(height),
            }
        )
    return head, found


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``row_height_more.tsv`` from Excel's saved answer (no Excel)."""
    if not answer.exists():
        sys.exit(f"not captured yet: no {answer}")
    head, found = records(answer.read_text(encoding="utf-8"))
    excel = next((h[len("# excel: ") :] for h in head if h.startswith("# excel: ")), "unknown")
    date = next((h[len("# captured: ") :] for h in head if h.startswith("# captured: ")), "")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: row_height_more (capture.row_height_more); {excel}\n")
        out.write(
            '# one workbook per Normal font, with openpyxl\'s <sheetFormatPr baseColWidth="8" '
            'defaultRowHeight="15"> (no customHeight); workbook 17 is Calibri 11 with a cell '
            "font a row\n"
        )
        out.write(f"# captured: {date}; method: row height (AppleScript)\n")
        out.write(
            "# normal: the workbook's Normal font; label: what column A of the row holds "
            "(CASES in capture/row_height_more.py); font: the cell's font; wrapped: three "
            "wrapped lines; applied: the font Excel applied to the cell; height: the row's "
            "height in points\n"
        )
        out.write("\t".join(COLUMNS) + "\n")
        for record in found:
            out.write("\t".join(record[c] for c in COLUMNS) + "\n")
    print(f"{len(found)} rows written to {output}")


def _number(text: str) -> str:
    """``15.0`` as ``15``; other text unchanged."""
    try:
        value = float(text)
    except ValueError:
        return text
    return repr(int(value)) if value.is_integer() else repr(value)


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        paths = build(Path(argv[1]))
        print(f"{len(paths)} workbooks, {len(CASES)} rows in {argv[1]}")
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
