"""The screen suite ``text_steps`` (indents under other Normal fonts, shrink to fit by single
characters; ARCHITECTURE.md §4.6 and §4.13), without Excel.

Its workbooks are built and read back with sheetpeek; its measuring runs on a synthetic capture
with known marks. The comparisons with ``tests/captures/text_steps.tsv`` are skipped until that
capture exists.
"""

import openpyxl
import pytest
from capture import text_steps as suite
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import SCALE, Canvas

BLACK = (0, 0, 0)


def _at(cell):
    letters, row = coordinate_from_string(cell)
    return row, column_index_from_string(letters)


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("text_steps"))


def test_each_workbook_reads_back_with_its_normal_font(built):
    for group, path in built.items():
        sheet = sheetpeek.read(path)
        name, size = suite.FACES[group]
        assert (sheet.normal_font.name, sheet.normal_font.size) == (name, size), group
        for case in suite.cases(group):
            assert sheet.cell(*_at(case.cell)).value is not None, (group, case.name)


def test_the_cases_are_what_their_names_say(built):
    sheet = sheetpeek.read(built["indent_calibri11"])
    for case in suite.cases("indent_calibri11"):
        face, size, align, _, level = case.name.rsplit(" ", 4)
        style = sheet.cell(*_at(case.cell)).style
        assert (style.h_align, style.indent) == (align, int(level)), case.name
        font = sheet.cell(*_at(case.cell)).font
        assert (font.name, font.size) == (face, float(size)), case.name
    sheet = sheetpeek.read(built["shrink"])
    for case in suite.cases("shrink"):
        cell = sheet.cell(*_at(case.cell))
        if case.name.startswith("shrunk"):
            assert len(cell.value) == int(case.name.split()[1])
            assert cell.style.shrink
        else:
            assert cell.font.size == float(case.name.split()[1])


def test_every_case_has_a_cell_of_its_own_left_of_and_below_the_markers():
    for group in suite.GROUPS:
        found = suite.cases(group)
        cells = [c.cell for c in found]
        assert len(set(cells)) == len(cells), group
        first, second = suite.MARKERS[group]
        marker_column = column_index_from_string(coordinate_from_string(second)[0])
        for cell in suite._boxes(found):
            row, column = _at(cell)
            assert column < marker_column, (group, cell)
            assert row > 3, (group, cell)
            assert cell not in (first, second)


def test_no_case_draws_magenta():
    # Every case is text in the default color on white: no font color or fill is set anywhere.
    for group in suite.GROUPS:
        ws = openpyxl.Workbook().active
        suite.fill(ws, group)
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    assert cell.font.color is None, cell.coordinate
                    assert cell.fill.fill_type is None, cell.coordinate


def test_a_case_is_measured_from_its_own_cells_edges():
    ws = openpyxl.Workbook().active
    case = suite.fill(ws, "shrink")[0]
    canvas = Canvas(ws, [*suite.references("shrink"), case.cell, case.reference])
    for cell, width in ((case.cell, 20), (case.reference, 40)):
        x0, y0, _, _ = canvas.box(cell)
        canvas.paint(x0 + 6, y0 + 8, x0 + 6 + width, y0 + 20, BLACK)
    measured = suite.measure(canvas.capture(), case)
    assert measured is not None
    assert float(measured["ink_left"]) == 6 / SCALE
    assert float(measured["width_ratio"]) == 0.5


# --- the comparison with Excel's capture


def _captured():
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.text_steps run")
    lines = suite.OUTPUT.read_text(encoding="utf-8").splitlines()
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    header = rows[0]
    return {
        (r["group"], r["cell"]): r
        for r in (dict(zip(header, row, strict=True)) for row in rows[1:])
    }


def _drawn(path, group):
    """sheetpeek's runs for each case: the first run's left and the last run's right, points."""
    rendered = sheetpeek.read(path).layout("excel")
    columns = {c.index: c for c in rendered.columns}
    found = {}
    for case in suite.cases(group):
        r, c = _at(case.cell)
        row = next(x for x in rendered.rows if x.index == r)
        runs = [t for t in row.runs if t.source == (r, c)]
        left = columns[c].x_pt
        right = left + columns[c].width_pt
        found[case.cell] = (
            min(t.x for t in runs) - left,
            max(t.x + t.width for t in runs) - right,
            runs[0].scale * 7,
        )
    return found


@pytest.mark.excel
def test_each_indent_level_moves_the_text_as_sheetpeek_moves_it(built):
    # The ink's edge at each level less its edge at level 0, against sheetpeek's runs: the side
    # bearings cancel, so half a point (one device pixel on the Mac) is the tolerance.
    captured = _captured()
    wrong = {}
    for group in suite.GROUPS:
        if not group.startswith("indent"):
            continue
        drawn = _drawn(built[group], group)
        found = suite.cases(group)
        for case in found:
            base = found[found.index(case) - int(case.name.split()[-1]) * 2]
            if (group, case.cell) not in captured:
                continue
            record, origin = captured[(group, case.cell)], captured[(group, base.cell)]
            side = 0 if " left " in case.name else 1
            key = "ink_left" if side == 0 else "ink_right"
            moved = float(record[key]) - float(origin[key])
            expected = drawn[case.cell][side] - drawn[base.cell][side]
            if abs(moved - expected) > 0.5:
                wrong[(group, case.name)] = (moved, expected)
    assert wrong == {}


KNOWN_STEPS = {"shrunk 10": (5, 4.0), "shrunk 25": (2, 1.0)}
"""Known deviations (§4.13): at the step Excel draws, xlformat measures the text at 50 pt in the
48 pt budget, so sheetpeek takes the step below."""


@pytest.mark.excel
def test_each_shrunk_length_takes_the_step_sheetpeek_gives_it(built):
    # The ink width over the unshrunk reference's, to the nearest seventh (text_style_more's
    # steps: within 0.04 of k/7 at Calibri 11, 0.14 between steps).
    captured = _captured()
    drawn = _drawn(built["shrink"], "shrink")
    wrong = {}
    for case in suite.cases("shrink"):
        if not case.name.startswith("shrunk") or ("shrink", case.cell) not in captured:
            continue
        step = round(float(captured[("shrink", case.cell)]["width_ratio"]) * 7)
        if step != round(drawn[case.cell][2]):
            wrong[case.name] = (step, drawn[case.cell][2])
    assert wrong == KNOWN_STEPS
