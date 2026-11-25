"""Capture the rows Excel fits to wrapped, justified and distributed text of one to six lines.

    uv run --group capture python -m capture.wrap_rows build PATH  # the workbook only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.wrap_rows run
    uv run --group capture python -m capture.wrap_rows read  # the saved answer again

The suite ``text_style`` (group ``align``) showed horizontally distributed and justified text
wrapped into three lines in rows that Excel for Windows fits a line taller than the lines it
draws; this suite asks how the fitted row grows with the lines, in each alignment, on each
platform (``ARCHITECTURE.md`` §4.7). Like ``row_height_more``, whose AppleScript it uses, it
reads what Excel itself reports, not the screen: ``row height`` of every case's row, and the
font Excel applied in column A.

One workbook under openpyxl's Calibri 11 Normal font and ``<sheetFormatPr baseColWidth="8"
defaultRowHeight="15">`` (no ``customHeight``, so Excel fits every row with a ``<row>``
element), default column widths, one case a row in column A, from row 1. A case is ``lines``
lines of text in one alignment (``ALIGNMENTS``): wrapped (``wrap``), justified, distributed, and
distributed with wrap off (``distributed unwrapped``, which wraps all the same). Its text is
either ``lines`` words of "wrapped" ("wrap" in Calibri 16), one a line of a default column
(``words``), or ``lines`` lines of "Hxg" broken by line feeds (``breaks``). Calibri 11 has both
forms; Calibri 16 and Arial 11 the words. No screenshot, no clipboard, no marker colors; one
workbook, under a minute of Excel.

``run`` saves Excel's answer to ``capture/build/wrap_rows/answer.tsv``; ``read`` writes
``tests/captures/wrap_rows.tsv`` from it again, one record per row read.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "wrap_rows.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "wrap_rows"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "wrap_rows"
)
SCRIPT = Path(__file__).with_name("row_height_more.applescript")
WORKBOOK = "wrap_rows.xlsx"
COLUMNS = ("name", "row", "alignment", "form", "lines", "font", "applied", "height", "text")
ALIGNMENTS = {
    "wrap": Alignment(wrap_text=True),
    "justify": Alignment(horizontal="justify", wrap_text=True),
    "distributed": Alignment(horizontal="distributed", wrap_text=True),
    "distributed unwrapped": Alignment(horizontal="distributed"),
}
WORDS = {("Calibri", 16): "wrap"}
"""The word of a font's ``words`` cases, "wrapped" where not listed: one fits a line of a default
column, two do not (by xlformat's widths and insets on both platforms)."""
FONTS = (("Calibri", 11), ("Calibri", 16), ("Arial", 11))


@dataclass(frozen=True)
class Case:
    """One row: ``lines`` lines of text in ``alignment``, as ``form``, in ``face`` ``size``."""

    alignment: str
    form: str
    lines: int
    face: str = "Calibri"
    size: int = 11

    @property
    def name(self) -> str:
        return f"{self.alignment} {self.form} {self.lines} {self.face} {self.size}"

    @property
    def text(self) -> str:
        if self.form == "words":
            return " ".join([WORDS.get((self.face, self.size), "wrapped")] * self.lines)
        return "\n".join(["Hxg"] * self.lines)


def _cases() -> list[Case]:
    found = []
    for face, size in FONTS:
        forms = ("words", "breaks") if (face, size) == ("Calibri", 11) else ("words",)
        for form in forms:
            for alignment in ALIGNMENTS:
                found += [Case(alignment, form, k, face, size) for k in range(1, 7)]
    return found


CASES = _cases()


def build(folder: Path = CONTAINER_DIR) -> Path:
    """Write the workbook into ``folder``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "rows"
    for row, case in enumerate(CASES, 1):
        cell = ws.cell(row, 1, case.text)
        if (case.face, case.size) != ("Calibri", 11):
            cell.font = Font(name=case.face, sz=case.size, family=2)
        cell.alignment = ALIGNMENTS[case.alignment]
    path = folder / WORKBOOK
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def run() -> None:
    """Build, have Excel read the rows (needs Excel), save its answer, write the TSV."""
    from capture.edition import excel_description

    path = build()
    rows = ",".join(str(row) for row in range(1, len(CASES) + 1))
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), str(path), rows],
            capture_output=True,
            text=True,
            check=False,
            timeout=300,  # kills osascript, which ignores SIGTERM while it waits on Excel
        )
    except subprocess.TimeoutExpired:
        sys.exit("osascript did not answer in 5 minutes; check Excel before another run")
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
    for k, (case, line) in enumerate(zip(CASES, body, strict=True), 1):
        name, row, applied_name, applied_size, height = line.split("\t")
        if (name, int(row)) != (WORKBOOK, k):
            raise ValueError(f"Excel answered for {name} row {row}, expected {case}")
        found.append(
            {
                "name": case.name,
                "row": str(k),
                "alignment": case.alignment,
                "form": case.form,
                "lines": str(case.lines),
                "font": f"{case.face} {case.size}",
                "applied": f"{applied_name} {_number(applied_size)}",
                "height": _number(height),
                "text": case.text.replace("\n", "\\n"),
            }
        )
    return head, found


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``wrap_rows.tsv`` from Excel's saved answer (no Excel)."""
    if not answer.exists():
        sys.exit(f"not captured yet: no {answer}")
    head, found = records(answer.read_text(encoding="utf-8"))
    excel = next((h[len("# excel: ") :] for h in head if h.startswith("# excel: ")), "unknown")
    date = next((h[len("# captured: ") :] for h in head if h.startswith("# captured: ")), "")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: wrap_rows (capture.wrap_rows); {excel}\n")
        out.write(
            "# one workbook, openpyxl's default styles (Normal Calibri 11, no customHeight); "
            "one case a row in column A, fitted by Excel on opening\n"
        )
        out.write(f"# captured: {date}; method: row height (AppleScript)\n")
        out.write(
            "# name: the case (CASES in capture/wrap_rows.py); lines: the lines of its text; "
            "applied: the font Excel applied to column A; height: the row's height in points; "
            "text: line feeds as \\n\n"
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
        print(build(Path(argv[1])))
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
