"""Capture the line width at which Excel fits a wrapped row a line taller, column by column.

    uv run --group capture python -m capture.wrap_band build PATH  # the workbooks only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.wrap_band run
    uv run --group capture python -m capture.wrap_band read  # the saved answer again

The suites ``text_style`` (groups ``align`` and ``line_budget``) and ``wrap_rows`` show Excel
drawing a wrapped line while it is at most the column less its insets, and fitting most rows to
the lines it draws; three ``align`` rows on Windows are fitted a line taller than drawn. This
suite asks where, line width by line width, a row fitted on opening takes a second line, on each
platform (``ARCHITECTURE.md`` §4.7). Like ``wrap_rows`` it reads what Excel itself reports, not
the screen: the height of every case's row, the font Excel applied in column A, and column A's
width (``column width`` in characters and ``width`` in points), so that each row is compared
with the width Excel laid out, not the one the file asks for.

One workbook per column width (``COLUMNS``), under openpyxl's Calibri 11 Normal font and no
``customHeight``, one case a row in column A from row 1. A ``w`` column is a whole number of
pixels on Excel for Windows (file width ``px / 7``: the file's width includes the column's
padding), an ``m`` column a whole number of points on Excel for Mac (file width ``pt / 6``); its cases measure their text in that unit, by
xlformat's whole advances on that platform (``unit_widths``), and both platforms read every
workbook. A case (``Case``) is a ``variant`` at a ``target`` width:

- ``wrap``, ``justify``, ``distributed``: two words, together ``target`` wide with the space,
  wrapped, justified or distributed, ``target`` from the column less 16 to the column less 3;
- ``indent wrap``, ``indent distributed``: the same two words with indent 1, ``target`` from the
  column less 40 to the column less 3;
- ``word wrap``, ``word distributed``: one word ``target`` wide, from the column less 14 to the
  column less 1 (is a word too wide to fit broken?);
- ``arial wrap``: two words in Arial 10 cells, measured in Arial 10, as ``wrap``.

Every column has ``wrap``; ``w064`` and ``m066`` have every variant. No screenshot, no
clipboard, no marker colors; a few minutes of Excel.

``run`` saves Excel's answer to ``capture/build/wrap_band/answer.tsv``; ``read`` writes
``tests/captures/wrap_band.tsv`` from it again, one record per row read.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from dataclasses import dataclass
from functools import cache
from pathlib import Path

import openpyxl
import xlformat
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "wrap_band.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "wrap_band"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "wrap_band"
)
SCRIPT = Path(__file__).with_name("wrap_band.applescript")
COLUMNS = (
    "name", "workbook", "row", "column", "variant", "target", "unit", "font", "applied",
    "column_width", "width", "height", "text",
)  # fmt: skip
WINDOWS_PX = (64, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 100, 128, 160)
"""Columns in whole pixels on Excel for Windows: 87 to 96 put the line widths 80 to 89 at the
column less 7, where autofit's long-text bonus takes its first step."""
MAC_PT = (54, 66, 96, 97, 98, 99, 100, 102, 126, 160)
"""Columns in whole points on Excel for Mac: 96 to 100 put the line widths 89 to 93 at the
column less 7."""
FULL = ("w064", "m066")
"""The columns with every variant; the others have ``wrap`` only."""
VARIANTS = {
    "wrap": (Alignment(wrap_text=True), 16, 3),
    "justify": (Alignment(horizontal="justify", wrap_text=True), 16, 3),
    "distributed": (Alignment(horizontal="distributed", wrap_text=True), 16, 3),
    "indent wrap": (Alignment(horizontal="left", wrap_text=True, indent=1), 40, 3),
    "indent distributed": (Alignment(horizontal="distributed", wrap_text=True, indent=1), 40, 3),
    "word wrap": (Alignment(wrap_text=True), 14, 1),
    "word distributed": (Alignment(horizontal="distributed", wrap_text=True), 14, 1),
    "arial wrap": (Alignment(wrap_text=True), 16, 3),
}
"""Each variant's alignment and its targets, from the column less the first number to the
column less the second."""
LETTERS = "aceonsuxvzbdhkpqgwil"
"""The letters words are made of, tried in this order; not "m", whose advance on Excel for Mac
carries a fraction its whole point leaves out."""


@dataclass(frozen=True)
class Case:
    """One row: ``variant`` with a text ``target`` units wide, in the column ``column``."""

    column: str
    variant: str
    target: int

    @property
    def platform(self) -> str:
        """The platform whose unit the column and the target are in."""
        return "windows" if self.column.startswith("w") else "mac"

    @property
    def unit(self) -> str:
        return "px" if self.platform == "windows" else "pt"

    @property
    def face(self) -> tuple[str, int]:
        return ("Arial", 10) if self.variant.startswith("arial") else ("Calibri", 11)

    @property
    def name(self) -> str:
        return f"{self.column} {self.variant} {self.target}"

    @property
    def text(self) -> str:
        widths = unit_widths(self.platform, *self.face)
        if self.variant.startswith("word"):
            return _word(self.target, widths)
        space = widths[" "]
        left = (self.target - space) // 2
        return _word(left, widths) + " " + _word(self.target - space - left, widths)


@cache
def unit_widths(platform: str, face: str, size: int) -> dict[str, int]:
    """Each letter's and the space's whole advance in the platform's unit (pixels, points).

    Raises:
        ValueError: if an advance is not a whole number of units.
    """
    font = xlformat.Font(face, size)
    scale = 4 / 3 if platform == "windows" else 1
    found = {}
    for char in " " + LETTERS:
        value = xlformat.glyph_widths(char, font, platform=platform)[0] * scale
        if value != round(value):
            raise ValueError(f"{char!r} is {value} units in {face} {size} on {platform}")
        found[char] = round(value)
    return found


def _word(target: int, widths: dict[str, int]) -> str:
    """A word of ``LETTERS`` exactly ``target`` units wide: one letter repeated, then another."""
    for n in range(1, 40):
        for base in LETTERS:
            for extra in LETTERS:
                for k in range(n):
                    word = base * (n - k) + extra * k
                    if sum(widths[c] for c in word) == target:
                        return word
    raise ValueError(f"no word {target} units wide")


def column_size(column: str) -> int:
    """The column's width in its unit: ``w064`` is 64 pixels, ``m066`` 66 points."""
    return int(column[1:])


def file_width(column: str) -> float:
    """The width the file gives the column, in characters."""
    size = column_size(column)
    return size / 7 if column.startswith("w") else size / 6


def workbook_name(column: str) -> str:
    return f"wrap_band_{column}.xlsx"


def _columns() -> tuple[str, ...]:
    return tuple(f"w{px:03d}" for px in WINDOWS_PX) + tuple(f"m{pt:03d}" for pt in MAC_PT)


def cases(column: str) -> list[Case]:
    """The column's cases in row order."""
    size = column_size(column)
    found = []
    for variant, (_, low, high) in VARIANTS.items():
        if variant != "wrap" and column not in FULL:
            continue
        found += [Case(column, variant, w) for w in range(size - low, size - high + 1)]
    return found


COLUMN_KEYS = _columns()


def build_one(column: str, folder: Path) -> Path:
    """Write the column's workbook into ``folder``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "band"
    ws.column_dimensions["A"].width = file_width(column)
    for row, case in enumerate(cases(column), 1):
        cell = ws.cell(row, 1, case.text)
        if case.face != ("Calibri", 11):
            cell.font = Font(name=case.face[0], sz=case.face[1], family=2)
        cell.alignment = VARIANTS[case.variant][0]
    path = folder / workbook_name(column)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    """Write every column's workbook into ``folder``; return their paths."""
    return [build_one(column, folder) for column in COLUMN_KEYS]


def run() -> None:
    """Build, have Excel read the rows (needs Excel), save its answer, write the TSV."""
    from capture.edition import excel_description

    args = []
    for column, path in zip(COLUMN_KEYS, build(), strict=True):
        args += [str(path), ",".join(str(r) for r in range(1, len(cases(column)) + 1))]
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), *args],
            capture_output=True,
            text=True,
            check=False,
            timeout=900,  # kills osascript, which ignores SIGTERM while it waits on Excel
        )
    except subprocess.TimeoutExpired:
        sys.exit("osascript did not answer in 15 minutes; check Excel before another run")
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    with ANSWER.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# excel: {excel_description()}, en_US\n")
        out.write(f"# captured: {dt.date.today().isoformat()}\n")
        out.write(result.stdout.strip() + "\n")
    write()


def records(answer: str) -> tuple[list[str], list[dict[str, str]]]:
    """The header lines and the records of a saved answer, checked against the cases.

    Each answer line is ``file name, row, applied font name, applied size, row height, column
    width, width``, tab separated, in the order of ``COLUMN_KEYS`` and then of the rows.
    """
    lines = answer.strip().split("\n")
    head = [line for line in lines if line.startswith("# ")]
    body = [line for line in lines if line and not line.startswith("# ")]
    expected = [(column, row, case) for column in COLUMN_KEYS
                for row, case in enumerate(cases(column), 1)]  # fmt: skip
    if len(body) != len(expected):
        raise ValueError(f"expected {len(expected)} lines from Excel, got {len(body)}")
    found = []
    for (column, row, case), line in zip(expected, body, strict=True):
        name, got_row, applied_name, applied_size, height, chars, points = line.split("\t")
        if (name, int(got_row)) != (workbook_name(column), row):
            raise ValueError(f"Excel answered for {name} row {got_row}, expected {case}")
        found.append(
            {
                "name": case.name,
                "workbook": name,
                "row": str(row),
                "column": column,
                "variant": case.variant,
                "target": str(case.target),
                "unit": case.unit,
                "font": f"{case.face[0]} {case.face[1]}",
                "applied": f"{applied_name} {_number(applied_size)}",
                "column_width": _number(chars),
                "width": _number(points),
                "height": _number(height),
                "text": case.text,
            }
        )
    return head, found


def write_records(head_lines: list[str], found: list[dict[str, str]], output: Path) -> None:
    """Write ``found`` under ``head_lines`` and the column names."""
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write("".join(f"{line}\n" for line in head_lines))
        out.write("\t".join(COLUMNS) + "\n")
        for record in found:
            out.write("\t".join(record[c] for c in COLUMNS) + "\n")
    print(f"{len(found)} rows written to {output}")


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``wrap_band.tsv`` from Excel's saved answer (no Excel)."""
    if not answer.exists():
        sys.exit(f"not captured yet: no {answer}")
    head, found = records(answer.read_text(encoding="utf-8"))
    excel = next((h[len("# excel: ") :] for h in head if h.startswith("# excel: ")), "unknown")
    date = next((h[len("# captured: ") :] for h in head if h.startswith("# captured: ")), "")
    write_records(
        [
            f"# suite: wrap_band (capture.wrap_band); {excel}",
            "# a workbook per column width, openpyxl's default styles (Normal Calibri 11, no "
            "customHeight); one case a row in column A, fitted by Excel on opening",
            f"# captured: {date}; method: row height, column width and width (AppleScript)",
            "# name: the case (cases() in capture/wrap_band.py); target: the text's width in "
            "unit by xlformat's whole advances on the column's platform; applied: the font Excel "
            "applied to column A; column_width: column A's width in characters and width in "
            "points, as Excel laid it out; height: the row's height in points",
        ],
        found,
        output,
    )


def _number(text: str) -> str:
    """``15.0`` as ``15``; other text unchanged."""
    try:
        value = float(text)
    except ValueError:
        return text
    return repr(int(value)) if value.is_integer() else repr(value)


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        for path in build(Path(argv[1])):
            print(path)
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[0], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
