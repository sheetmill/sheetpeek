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
- ``zero``: a height of 0 without ``customHeight``, as ``ht="0"`` and as ``ht=""`` (which Excel
  reads as 0), on an empty row and beside a Calibri 20 cell; then the same two with
  ``customHeight="1"``, which Excel shows hidden (xlio's capture ``live_reads``).
- ``customsheet``: ``<sheetFormatPr defaultRowHeight="20" customHeight="1">``, where Excel fits no
  row: an ``ht`` without ``customHeight`` beside Calibri 11 and 28 cells and on an empty row, and an
  unlisted row (the default, 20).
- ``zerosheet``: ``<sheetFormatPr defaultRowHeight="15" zeroHeight="1">``, where unlisted rows are
  hidden: an ``ht`` without ``customHeight`` beside Calibri 11 and 20 cells and on an empty row, a
  Calibri 11 row without ``ht``, and an unlisted row (hidden).

``run`` saves Excel's answer to ``capture/build/refit_rows/answer.tsv`` and writes
``tests/captures/refit_rows.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from capture import stored_rows, wrap_band, xmlform

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "refit_rows.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "refit_rows"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = wrap_band.CONTAINER_DIR.parent / "refit_rows"
COLUMNS = ("workbook", "name", "row", "cell", "stored", "row_height", "height", "top")


@dataclass(frozen=True)
class Case:
    """One row: its cell (``None``, empty; else the Calibri size), stored ``ht`` (``None``, no
    ``ht``; with no cell either, no row element), the ``ht`` as written when it is not the
    number's own text, and whether the row keeps ``customHeight="1"``."""

    workbook: str
    row: int
    size: int | None
    stored: float | None
    text: str | None = None
    custom: bool = False

    @property
    def written(self) -> str | None:
        """The ``ht`` attribute's text, or ``None`` for no ``ht``."""
        if self.text is not None:
            return self.text
        return None if self.stored is None else f"{self.stored:g}"

    @property
    def name(self) -> str:
        cell = "empty" if self.size is None else f"calibri {self.size}"
        written = self.written
        if written is None:
            stored = (
                "no ht" if self.size is not None and self.workbook in SHEET_FORMATS else "unlisted"
            )
        elif written == "":
            stored = 'ht ""'
        else:
            stored = f"ht {written}"
        custom = " custom" if self.custom else ""
        return f"{self.workbook} {cell} {stored}{custom}"


_REFIT = (
    (11, 10.0), (11, 15.0), (11, 30.0),
    (20, 12.0), (20, 26.0), (20, 26.25), (20, 40.0),
    (28, 20.0), (28, 60.0),
    (None, 10.0), (None, 15.0), (None, 30.0),
)  # fmt: skip
_MACSAVED = ((11, 15.0), (20, 26.0), (None, None), (28, 37.0), (None, 15.0))
_ZERO = (
    (None, "0", False), (20, "0", False), (None, "", False), (20, "", False),
    (None, "0", True), (20, "", True),
)  # fmt: skip
_CUSTOMSHEET = ((11, 30.0), (28, 30.0), (None, 30.0), (11, 10.0), (None, None))
_ZEROSHEET = ((11, 30.0), (20, 10.0), (None, 30.0), (11, None), (None, None))
SHEET_FORMATS = {
    "customsheet": 'baseColWidth="8" defaultRowHeight="20" customHeight="1"',
    "zerosheet": 'baseColWidth="8" defaultRowHeight="15" zeroHeight="1"',
}
"""The ``<sheetFormatPr>`` attributes of a workbook that does not keep openpyxl's."""
CASES = {
    "refit": [Case("refit", k, size, h) for k, (size, h) in enumerate(_REFIT, 1)],
    "macsaved": [Case("macsaved", k, size, h) for k, (size, h) in enumerate(_MACSAVED, 1)],
    "zero": [
        Case("zero", k, size, 0.0, text, custom) for k, (size, text, custom) in enumerate(_ZERO, 1)
    ],
    "customsheet": [Case("customsheet", k, size, h) for k, (size, h) in enumerate(_CUSTOMSHEET, 1)],
    "zerosheet": [Case("zerosheet", k, size, h) for k, (size, h) in enumerate(_ZEROSHEET, 1)],
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
            # openpyxl leaves out a height of 0; _rewrite writes every ``ht`` as the case does
            ws.row_dimensions[case.row].height = case.stored or 1.0
    path = folder / workbook_name(name)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    path.write_bytes(stored_rows._without_custom_height(path.read_bytes()))
    if name not in ("refit", "macsaved"):
        _rewrite(path, CASES[name], SHEET_FORMATS.get(name))
    return path


def _rewrite(path: Path, cases: list[Case], sheet_format: str | None) -> None:
    """Write each case's ``ht`` as its text, with ``customHeight="1"`` where it keeps it, and
    the sheet's ``<sheetFormatPr>``; the XML in the standard library's form."""
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                text = xmlform.standard_form(data.decode("utf-8"))
                for case in cases:
                    if case.written is None:
                        continue
                    custom = ' customHeight="1"' if case.custom else ""
                    text, count = re.subn(
                        rf'<row r="{case.row}" ht="[^"]*"',
                        f'<row r="{case.row}" ht="{case.written}"{custom}',
                        text,
                    )
                    if count != 1:
                        sys.exit(f"{path.name}: expected one row {case.row} with ht, found {count}")
                if sheet_format is not None:
                    text, count = re.subn(
                        r"<sheetFormatPr\b[^>]*/>", f"<sheetFormatPr {sheet_format} />", text
                    )
                    if count != 1:
                        sys.exit(f"{path.name}: expected one <sheetFormatPr />, found {count}")
                data = text.encode("utf-8")
            archive.writestr(info, data)


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
        "stored": "" if case.written is None else case.written or '""',
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
