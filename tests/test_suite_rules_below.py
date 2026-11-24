"""The screen suite ``rules_below`` (underlines in rows too short for them), without Excel.

Its workbook is built and read back with sheetpeek; its ``below`` measure runs on a synthetic
capture with known marks.
"""

import openpyxl
import pytest
from capture import rules_below as suite
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import Canvas

BLACK = (0, 0, 0)
GRID = (212, 212, 212)


def _at(cell):
    letters, row = coordinate_from_string(cell)
    return row, column_index_from_string(letters)


def test_the_cases_are_what_their_names_say(tmp_path):
    path = tmp_path / "rules_below.xlsx"
    found = suite.build(path)
    assert len(found) == len(suite.COLUMNS) * len(suite.HEIGHTS) * len(suite.VERTICALS) == 63
    sheet = sheetpeek.read(path)
    written = openpyxl.load_workbook(path).active
    for case in found:
        underline, size, _, height, vertical = case.name.split()
        cell = sheet.cell(*_at(case.cell))
        assert cell.font.size == float(size), case.name
        assert (cell.font.underline or "plain") == underline, case.name
        assert cell.style.v_align == vertical, case.name
        row = _at(case.cell)[0]
        assert written.row_dimensions[row].height == float(height), case.name
        assert written.row_dimensions[row + 1].height == suite.GAP, case.name
        assert written.cell(row + 1, _at(case.cell)[1]).value is None, case.name


def test_every_case_is_left_of_and_below_the_markers():
    second = suite.MARKERS[1]
    marker_column = column_index_from_string(coordinate_from_string(second)[0])
    for case in suite.cases():
        row, column = _at(case.cell)
        assert column < marker_column, case.name
        assert row > 3, case.name


def test_below_finds_a_line_under_the_cells_bottom_and_not_the_gridline():
    ws = openpyxl.Workbook().active
    suite.fill(ws)
    canvas = Canvas(ws, ["C4", "E4", "C5"])
    x0, _, x1, bottom = canvas.box("C4")
    canvas.paint(x0, bottom, x1, bottom + 1, GRID)  # the next row's gridline
    canvas.paint(x0 + 2, bottom + 4, x1 - 2, bottom + 6, BLACK)  # a line 2 pt under the edge
    found = suite.below(canvas.capture(), "C4")
    assert found["below"] == "2+1"
    assert int(found["below_ink"]) == 2 * (x1 - x0 - 12)
    assert suite.below(canvas.capture(), "E4") == {"below": "", "below_ink": "0"}


def test_a_line_shorter_than_a_quarter_of_the_cell_is_not_a_band():
    ws = openpyxl.Workbook().active
    suite.fill(ws)
    canvas = Canvas(ws, ["C4", "C5"])
    x0, _, x1, bottom = canvas.box("C4")
    canvas.paint(x0 + 6, bottom + 2, x0 + 6 + (x1 - x0 - 12) // 5, bottom + 3, BLACK)
    found = suite.below(canvas.capture(), "C4")
    assert found["below"] == ""
    assert int(found["below_ink"]) > 0


@pytest.mark.parametrize("module", ["capture.screen.__main__"])
def test_the_suite_is_a_screen_suite(module):
    import importlib

    assert "rules_below" in importlib.import_module(module).SUITES


# --- the comparison with Excel's capture

KNOWN_BELOW = {
    # Unexplained: Excel for Mac draws this case's second line 2 pt into the row below, its first
    # under that row's gridline, where it draws nothing past the row in "doubleAccounting 11 in 12
    # center", whose lines fall as far past its row; sheetpeek clips every line at the row.
    "doubleAccounting 11 in 15 top": "2+1",
}


@pytest.mark.excel
def test_excel_for_mac_draws_no_line_past_its_row_but_one():
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.screen run rules_below")
    lines = suite.OUTPUT.read_text(encoding="utf-8").split("\n")
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    records = [dict(zip(rows[0], row, strict=True)) for row in rows[1:]]
    assert len(records) == 63
    found = {r["name"]: r["below"] for r in records if r["below"]}
    assert found == KNOWN_BELOW
