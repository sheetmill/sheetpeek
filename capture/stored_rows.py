"""Capture how Excel for Windows lays out row heights a file stores in points, as a capture job.

    uv run --group capture python -m capture.stored_rows build PATH  # the workbooks only

A row height is stored in points, and Excel for Windows draws rows in whole pixels at 96 dpi,
where a point is 4/3 of a pixel: a row stored at 32 pt (42.67 px) was laid out at 42 px
(``windows/row_labels``). Files saved by Excel for Mac and by openpyxl store heights that are
not whole pixels. This suite stores heights across the fractions of a pixel (``HEIGHTS``) in
three forms (``FORMS``) and reads, for every row, the height Excel reports (``row_height``) and
the row's ``height`` and ``top`` in points, which give its pixel rows.

- ``custom``: one workbook, a row a height, its ``ht`` with ``customHeight="1"`` (openpyxl's
  form), "x" in column A.
- ``plain``: the same without ``customHeight``.
- ``default``: a workbook a height, the sheet's ``defaultRowHeight`` with ``customHeight="1"``
  and no row element; rows 1 to 3 read.

The job group ``stored_rows`` (``capture/job.py``) reads them into
``tests/captures/windows/stored_rows.tsv``.
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
HEIGHTS = (
    *(round(15 + k / 10, 1) for k in range(11)),
    15.375, 16.33, 21.33, 31.875, 32.0, 32.25, 32.5, 100.1, 409.0,
)  # fmt: skip
"""Stored heights in points: 15 to 16 by tenths (20 to 21.33 px), 20.5 and 42.5 px exactly,
42.67 px (32 pt), 43 px exactly, and others."""
FORMS = ("custom", "plain", "default")
DEFAULT_ROWS = 3
COLUMNS = ("form", "name", "stored", "row", "row_height", "height", "top")


@dataclass(frozen=True)
class Case:
    """One row read: ``form``, the stored height, and the row read."""

    form: str
    stored: float
    row: int

    @property
    def name(self) -> str:
        return f"{self.form} {self.stored:g} row {self.row}"


def workbook_name(form: str, stored: float | None = None) -> str:
    if form == "default":
        return f"stored_rows_default_{str(stored).replace('.', '_')}.xlsx"
    return f"stored_rows_{form}.xlsx"


def cases(form: str, stored: float | None = None) -> list[Case]:
    """The rows read in one workbook, in order."""
    if form == "default":
        assert stored is not None
        return [Case(form, stored, row) for row in range(1, DEFAULT_ROWS + 1)]
    return [Case(form, h, row) for row, h in enumerate(HEIGHTS, 1)]


def workbooks() -> list[tuple[str, float | None]]:
    """Every workbook as ``(form, stored height or None)``."""
    return [("custom", None), ("plain", None), *(("default", h) for h in HEIGHTS)]


def _without_custom_height(data: bytes) -> bytes:
    """The workbook with ``customHeight`` taken off every ``<row>`` (standard library's form)."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            body = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = body.decode("utf-8")
                text = re.sub(r'(<row [^>]*?) customHeight="1"', r"\1", text)
                body = text.encode("utf-8")
            zout.writestr(item, body)
    return out.getvalue()


def build_one(form: str, stored: float | None, folder: Path) -> Path:
    """Write one workbook into ``folder``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    if form == "default":
        assert stored is not None
        ws.sheet_format.defaultRowHeight = stored
        ws.sheet_format.customHeight = True
    else:
        for row, height in enumerate(HEIGHTS, 1):
            ws.cell(row, 1, "x")
            ws.row_dimensions[row].height = height
    path = folder / workbook_name(form, stored)
    path.parent.mkdir(parents=True, exist_ok=True)
    buffer = io.BytesIO()
    wb.save(buffer)
    data = buffer.getvalue()
    if form == "plain":
        data = _without_custom_height(data)
    path.write_bytes(data)
    return path


def build(folder: Path) -> list[Path]:
    return [build_one(form, stored, folder) for form, stored in workbooks()]


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        for path in build(Path(argv[1])):
            print(path)
        return 0
    print(__doc__.split("\n\n")[0], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
