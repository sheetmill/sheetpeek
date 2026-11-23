"""Captures from Excel for Windows (``tests/captures/windows/``), against sheetpeek's layout with
``platform="windows"``.

Each file is the Mac suite's schema, read from a capture job's result by ``capture/job.py``; its
header names the job and holds the result's header. A comparison is skipped until its file exists.
"""

from pathlib import Path

import pytest

import sheetpeek
from tests.capture_report import REPORT

WINDOWS = Path(__file__).with_name("captures") / "windows"


def _rows(name):
    path = WINDOWS / name
    if not path.exists():
        pytest.skip(f"not captured yet: {path.name}")
    lines = path.read_text(encoding="utf-8").splitlines()
    head = [line for line in lines if line.startswith("# ")]
    assert "# platform: windows" in head
    return [line.split("\t") for line in lines if line and not line.startswith("# ")][1:]


@pytest.fixture(scope="module")
def cells(tmp_path_factory):
    pytest.importorskip("openpyxl")
    from capture.cell_text import build

    path = build(tmp_path_factory.mktemp("cell_text") / "cell_text.xlsx")
    return sheetpeek.open(path, platform="windows")


@pytest.mark.excel
def test_merge_width_matches_excel_for_windows(cells):
    rows: dict[tuple[str, bool], list[str]] = {}
    for width, merge, _value, _fmt, shown in _rows("merge_width.tsv"):
        rows.setdefault((width, merge == "A:C, B hidden"), []).append(shown)
    compared, deviations = 0, {}
    for (width, hidden), texts in rows.items():
        rendered = cells.sheet(("h" if hidden else "w") + width).layout("excel")
        for row, shown in enumerate(texts, 1):
            compared += 1
            if rendered.cell(row, 1).text != shown:
                deviations[(width, hidden, row)] = (rendered.cell(row, 1).text, shown)
    REPORT["merge width (Windows)"] = (
        f"merge width, Excel for Windows: {compared} cells compared, {len(deviations)} deviations"
    )
    assert compared == 150
    assert deviations == {}


@pytest.mark.excel
def test_error_width_matches_excel_for_windows(cells):
    from capture.cell_text import ERROR_WIDTHS, ERRORS

    rendered = cells.sheet("errors").layout("excel")
    records = _rows("error_width.tsv")
    deviations = {}
    for error, width, shown in records:
        cell = rendered.cell(ERRORS.index(error) + 1, ERROR_WIDTHS.index(float(width)) + 2)
        if cell.text != shown:
            deviations[(error, width)] = (cell.text, shown)
    REPORT["error width (Windows)"] = (
        f"error width, Excel for Windows: {len(records)} cells compared, {len(deviations)} "
        "deviations"
    )
    assert len(records) == 56
    assert deviations == {}
