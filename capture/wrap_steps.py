"""Capture where a fitted wrapped row takes a second line at autofit's later bonus steps and under
another Normal font.

    uv run --group capture python -m capture.wrap_steps build PATH  # the workbooks only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.wrap_steps run
    uv run --group capture python -m capture.wrap_steps read  # the saved answer again

``wrap_band`` places the first step of autofit's long-text bonus in Excel's row fitting on each
platform: the rows take the bonus only past the step, where autofit takes it at the step
(``ARCHITECTURE.md`` §4.7). This suite asks the same at the second and third steps under a
Calibri 11 Normal font (Windows 112 and 140 px, Mac 120 and 150 pt, a column of the line plus 7
plus the bonus either way), and asks the whole rule under a Courier 10 Normal font with Arial 10
cells, with lines on both sides of the gap between their drawn width and their autofit width. It
reads what Excel reports, as ``wrap_band`` does, with the same AppleScript and the same columns,
and a ``normal`` column naming the workbook's Normal font.

One workbook per column (``COLUMNS``), one case a row in column A from row 1, wrapped, no
``customHeight``. A ``w`` column is whole pixels on Excel for Windows, an ``m`` column whole
points on Excel for Mac; its cases measure their text in that unit by xlformat's whole advances
in the cell's font (``wrap_band.unit_widths``), from the column less ``low`` to the column less
``high``. Under Calibri 11 the file width is the column over 7 (Windows) or 6 (Mac); under
Courier 10 over 8 and 6, Courier's digit on each (Windows' unmeasured by xlformat: Excel's width
is read back). Text is generated, two words of the target's width.

``run`` saves Excel's answer to ``capture/build/wrap_steps/answer.tsv``; ``read`` writes
``tests/captures/wrap_steps.tsv`` from it again.
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

from capture import wrap_band

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "wrap_steps.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "wrap_steps"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = wrap_band.CONTAINER_DIR.parent / "wrap_steps"
SCRIPT = wrap_band.SCRIPT
COLUMNS = (*wrap_band.COLUMNS[:7], "normal", *wrap_band.COLUMNS[7:])


@dataclass(frozen=True)
class Column:
    """One workbook: column A ``size`` units wide on ``platform``, under ``normal``."""

    key: str
    size: int
    normal: tuple[str, int] = ("Calibri", 11)
    cell: tuple[str, int] = ("Calibri", 11)
    low: int = 16
    high: int = 3

    @property
    def platform(self) -> str:
        return "windows" if self.key.lstrip("c").startswith("w") else "mac"

    @property
    def unit(self) -> str:
        return "px" if self.platform == "windows" else "pt"

    @property
    def file_width(self) -> float:
        """The width the file gives column A, in characters of the Normal font's digit."""
        if self.platform == "mac":
            return self.size / 6
        return self.size / (8 if self.normal[0] == "Courier" else 7)

    @property
    def workbook(self) -> str:
        return f"wrap_steps_{self.key}.xlsx"


COLUMN_LIST = (
    *(Column(f"w{px:03d}", px) for px in (119, 120, 121, 148, 149, 150)),
    *(Column(f"m{pt:03d}", pt) for pt in (127, 128, 129, 158, 159, 160)),
    *(
        Column(key, size, ("Courier", 10), ("Arial", 10), low, 1)
        for key, size, low in (
            ("cm154", 154, 40),
            ("cm239", 239, 40),
            ("cw205", 205, 70),
            ("cw319", 319, 70),
        )
    ),
)
"""Calibri 11 Normal: columns where a line at the second or third step (Windows 112 and 140 px,
Mac 120 and 150 pt) fits only if its bonus is the step's less one. Courier 10 Normal, Arial 10
cells: two widths a platform, 40 units of lines up to the column on the Mac and 70 on Windows,
whose Courier xlformat does not measure (the column Excel lays out is read back)."""


@dataclass(frozen=True)
class Case:
    """One row: two words ``target`` units wide in the column's cell font."""

    column: Column
    target: int

    @property
    def name(self) -> str:
        return f"{self.column.key} wrap {self.target}"

    @property
    def text(self) -> str:
        widths = wrap_band.unit_widths(self.column.platform, *self.column.cell)
        space = widths[" "]
        left = (self.target - space) // 2
        return (
            wrap_band._word(left, widths)
            + " "
            + wrap_band._word(self.target - space - left, widths)
        )


def cases(column: Column) -> list[Case]:
    """The column's cases in row order."""
    return [Case(column, w) for w in range(column.size - column.low, column.size - column.high + 1)]


def build_one(column: Column, folder: Path) -> Path:
    """Write the column's workbook into ``folder``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "steps"
    name, size = column.normal
    if column.normal != ("Calibri", 11):
        normal = Font(name=name, sz=size, family=3 if name == "Courier" else 2)
        wb._fonts = IndexedList([normal])
        for style in wb._named_styles:
            if style.name == "Normal":
                style.font = normal
    ws.column_dimensions["A"].width = column.file_width
    for row, case in enumerate(cases(column), 1):
        cell = ws.cell(row, 1, case.text)
        cell.font = Font(name=column.cell[0], sz=column.cell[1], family=2)
        cell.alignment = Alignment(wrap_text=True)
    path = folder / column.workbook
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    """Write every column's workbook into ``folder``; return their paths."""
    return [build_one(column, folder) for column in COLUMN_LIST]


def run() -> None:
    """Build, have Excel read the rows (needs Excel), save its answer, write the TSV."""
    from capture.edition import excel_description

    args = []
    for column, path in zip(COLUMN_LIST, build(), strict=True):
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
    width, width``, tab separated, in the order of ``COLUMN_LIST`` and then of the rows.
    """
    lines = answer.strip().split("\n")
    head = [line for line in lines if line.startswith("# ")]
    body = [line for line in lines if line and not line.startswith("# ")]
    expected = [(c, row, case) for c in COLUMN_LIST for row, case in enumerate(cases(c), 1)]
    if len(body) != len(expected):
        raise ValueError(f"expected {len(expected)} lines from Excel, got {len(body)}")
    found = []
    for (column, row, case), line in zip(expected, body, strict=True):
        name, got_row, applied_name, applied_size, height, chars, points = line.split("\t")
        if (name, int(got_row)) != (column.workbook, row):
            raise ValueError(f"Excel answered for {name} row {got_row}, expected {case.name}")
        found.append(
            {
                "name": case.name,
                "workbook": name,
                "row": str(row),
                "column": column.key,
                "variant": "wrap",
                "target": str(case.target),
                "unit": column.unit,
                "normal": "{} {}".format(*column.normal),
                "font": "{} {}".format(*column.cell),
                "applied": f"{applied_name} {wrap_band._number(applied_size)}",
                "column_width": wrap_band._number(chars),
                "width": wrap_band._number(points),
                "height": wrap_band._number(height),
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
    """Write ``wrap_steps.tsv`` from Excel's saved answer (no Excel)."""
    if not answer.exists():
        sys.exit(f"not captured yet: no {answer}")
    head, found = records(answer.read_text(encoding="utf-8"))
    excel = next((h[len("# excel: ") :] for h in head if h.startswith("# excel: ")), "unknown")
    date = next((h[len("# captured: ") :] for h in head if h.startswith("# captured: ")), "")
    write_records(
        [
            f"# suite: wrap_steps (capture.wrap_steps); {excel}",
            "# a workbook per column width, under a Calibri 11 or Courier 10 Normal font, no "
            "customHeight; one case a row in column A, fitted by Excel on opening",
            f"# captured: {date}; method: row height, column width and width (AppleScript)",
            "# name: the case (cases() in capture/wrap_steps.py); target: the text's width in "
            "unit by xlformat's whole advances in the cell's font on the column's platform; "
            "normal: the workbook's Normal font; applied: the font Excel applied to column A; "
            "column_width: column A's width in characters and width in points, as Excel laid "
            "it out; height: the row's height in points",
        ],
        found,
        output,
    )


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
