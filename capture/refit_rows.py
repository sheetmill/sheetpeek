"""Capture what Excel does with a row's stored height when the row has no ``customHeight``.

    uv run --group capture python -m capture.refit_rows build PATH  # the workbooks only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.refit_rows run
    uv run --group capture python -m capture.refit_rows read  # the saved answer again

Excel for Windows lays a row whose ``ht`` has no ``customHeight`` out at 15 pt beside a Calibri
11 cell, whatever the ``ht`` (``windows/stored_rows``): it refits the row, or gives it the
default height. This suite tells refitted, default and stored apart: every row stores an ``ht``
without ``customHeight``, below, equal to and above the height its cell fits, with a Calibri 11,
20 or 28 cell in column A or no cell. Excel for Mac reads it by AppleScript (``run``), Excel for
Windows by the job group ``refit_rows`` (``capture/job.py``).

- ``refit``: a row a case (``CASES``), from row 1.
- ``macsaved``: rows as Excel for Mac writes a sheet it fitted: Calibri 11, 20 and 28 cells with
  the heights the Mac fits them (15, 26, 37), no ``customHeight``, an unlisted row between, and
  an empty row element at 15.

``run`` saves Excel's answer to ``capture/build/refit_rows/answer.tsv`` and writes
``tests/captures/refit_rows.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from capture import stored_rows, wrap_band

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "refit_rows.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "refit_rows"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = wrap_band.CONTAINER_DIR.parent / "refit_rows"
COLUMNS = ("workbook", "name", "row", "cell", "stored", "row_height", "height", "top")


@dataclass(frozen=True)
class Case:
    """One row: its cell (``None``, empty; else the Calibri size) and stored ``ht`` (``None``,
    no row element)."""

    workbook: str
    row: int
    size: int | None
    stored: float | None

    @property
    def name(self) -> str:
        cell = "empty" if self.size is None else f"calibri {self.size}"
        stored = "unlisted" if self.stored is None else f"ht {self.stored:g}"
        return f"{self.workbook} {cell} {stored}"


_REFIT = (
    (11, 10.0), (11, 15.0), (11, 30.0),
    (20, 12.0), (20, 26.0), (20, 26.25), (20, 40.0),
    (28, 20.0), (28, 60.0),
    (None, 10.0), (None, 15.0), (None, 30.0),
)  # fmt: skip
_MACSAVED = ((11, 15.0), (20, 26.0), (None, None), (28, 37.0), (None, 15.0))
CASES = {
    "refit": [Case("refit", k, size, h) for k, (size, h) in enumerate(_REFIT, 1)],
    "macsaved": [Case("macsaved", k, size, h) for k, (size, h) in enumerate(_MACSAVED, 1)],
}


def workbook_name(name: str) -> str:
    return f"refit_rows_{name}.xlsx"


def build_one(name: str, folder: Path) -> Path:
    """Write one workbook, every ``ht`` without ``customHeight``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for case in CASES[name]:
        if case.size is not None:
            ws.cell(case.row, 1, "x").font = Font(name="Calibri", sz=case.size, family=2)
        if case.stored is not None:
            ws.row_dimensions[case.row].height = case.stored
    path = folder / workbook_name(name)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    path.write_bytes(stored_rows._without_custom_height(path.read_bytes()))
    return path


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    return [build_one(name, folder) for name in CASES]


def run() -> None:
    """Build, have Excel for Mac read the rows, save its answer, write the TSV."""
    from capture.edition import excel_description

    args = []
    for name, path in zip(CASES, build(), strict=True):
        args += [str(path), ",".join(str(c.row) for c in CASES[name])]
    try:
        result = subprocess.run(
            ["osascript", str(wrap_band.SCRIPT), *args],
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


def records(answer: str) -> list[dict[str, str]]:
    """The records of a Mac answer (``wrap_band.applescript``'s lines), in case order."""
    body = [line for line in answer.strip().split("\n") if line and not line.startswith("# ")]
    every = [case for name in CASES for case in CASES[name]]
    if len(body) != len(every):
        raise ValueError(f"expected {len(every)} lines from Excel, got {len(body)}")
    found = []
    for case, line in zip(every, body, strict=True):
        name, row, _, _, height, _, _ = line.split("\t")
        if (name, int(row)) != (workbook_name(case.workbook), case.row):
            raise ValueError(f"Excel answered for {name} row {row}, expected {case.name}")
        found.append(record(case, height, "", ""))
    return found


def record(case: Case, row_height: str, height: str, top: str) -> dict[str, str]:
    return {
        "workbook": case.workbook,
        "name": case.name,
        "row": str(case.row),
        "cell": "" if case.size is None else f"Calibri {case.size}",
        "stored": "" if case.stored is None else f"{case.stored:g}",
        "row_height": wrap_band._number(row_height),
        "height": height,
        "top": top,
    }


def write_records(head: list[str], found: list[dict[str, str]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write("".join(f"{line}\n" for line in head))
        out.write("\t".join(COLUMNS) + "\n")
        for found_record in found:
            out.write("\t".join(found_record[c] for c in COLUMNS) + "\n")
    print(f"{len(found)} rows written to {output}")


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``refit_rows.tsv`` from Excel for Mac's saved answer (no Excel)."""
    text = answer.read_text(encoding="utf-8")
    head = [line for line in text.split("\n") if line.startswith("# ")]
    write_records(
        [*head, "# suite: refit_rows (capture.refit_rows); method: row height (AppleScript)"],
        records(text),
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
