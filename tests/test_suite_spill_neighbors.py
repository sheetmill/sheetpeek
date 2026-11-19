"""The screen suite ``spill_neighbors`` (ARCHITECTURE.md §4.1, §4.2 and §4.18), without Excel.

Its workbooks are built and read back with sheetpeek as built, and sheetpeek's layout of each
case is the suite's ``EXPECTED``. Its measuring runs on a synthetic capture with known marks.
The comparison with ``tests/captures/spill_neighbors.tsv`` is skipped until that capture exists.
"""

import zipfile

import openpyxl
import pytest
from capture import spill_neighbors
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import Canvas

GRAY = (212, 212, 212)
BLACK = (0, 0, 0)


def drawn(path):
    """What sheetpeek draws for each case of the workbook at ``path``, as ``EXPECTED`` holds it."""
    group = path.stem.removeprefix("spill_neighbors_")
    rendered = sheetpeek.read(path).layout("excel")
    columns = {c.index: c for c in rendered.columns}
    found = {}
    for case in spill_neighbors.cases(group):
        letters, r = coordinate_from_string(case.cell)
        c = column_index_from_string(letters)
        row = next(x for x in rendered.rows if x.index == r)
        cells = {x.column: x for x in row.cells}
        (run,) = [t for t in row.runs if t.source == (r, c)]
        if case.spill:
            ink = []
            for letter in spill_neighbors.SPILL_COLUMNS:
                column = columns.get(column_index_from_string(letter))
                if column is None:  # past the columns laid out: nothing reaches it
                    ink.append("0")
                    continue
                start, end = max(run.x, run.clip[0]), min(run.x + run.width, run.clip[2])
                inside = start < column.x_pt + column.width_pt and end > column.x_pt
                shown = column.index != c and bool(cells[column.index].text.strip())
                ink.append("1" if inside or shown else "0")
            first = row.cells.index(cells[c])
            crossed = [*row.crossed[first:], *[False] * len(spill_neighbors.SPILL_COLUMNS)]
            crossed = crossed[: len(spill_neighbors.SPILL_COLUMNS) - 1]
            found[case.cell] = {
                "text": cells[c].text,
                "neighbor_text": cells[c + 1].text,
                "ink_columns": "".join(ink),
                "hidden_gridlines": "".join("1" if x else "0" for x in crossed),
            }
        else:
            column = columns[c]
            found[case.cell] = {
                "text": cells[c].text,
                "row_height": row.height_pt,
                "left": run.x - column.x_pt,
                "right": run.x + run.width - (column.x_pt + column.width_pt),
            }
    return found


# --- the workbooks


def test_the_workbooks_read_back_as_built(tmp_path):
    found = spill_neighbors.build(tmp_path)
    assert found == spill_neighbors.CASES
    for group in spill_neighbors.GROUPS:
        path = tmp_path / f"spill_neighbors_{group}.xlsx"
        sheet = sheetpeek.read(path)
        assert sheet.notices == (), group
        assert sheet.view.show_zeros is (group != "zeros")
        for case in spill_neighbors.cases(group):
            letters, row = coordinate_from_string(case.cell)
            cell = sheet.cell(row, column_index_from_string(letters))
            if case.spill:
                assert cell.value == spill_neighbors.LONG, case.name
                neighbor = sheet.cell(row, column_index_from_string(letters) + 1)
                assert (
                    neighbor.value
                    == {
                        "empty": None,
                        "number 1": 1,
                        'saved =""': "",
                        "one space": " ",
                        "=0": 0,
                        "number 0": 0,
                        'uncached =""': "",
                    }[case.neighbor]
                ), case.name
            else:
                assert cell.value == case.value, case.name
                font = spill_neighbors.FONTS[case.font]
                assert (cell.style.font.name, cell.style.font.size) == (font["name"], font["size"])
                assert cell.style.font.bold is font.get("bold", False)
        assert drawn(path) == spill_neighbors.EXPECTED[group], group


def test_each_neighbor_and_the_workbook_are_written_as_excel_saves_them(tmp_path):
    spill_neighbors.build(tmp_path)
    for group in spill_neighbors.GROUPS:
        with zipfile.ZipFile(tmp_path / f"spill_neighbors_{group}.xlsx") as package:
            book = package.read("xl/workbook.xml").decode()
            sheet = package.read("xl/worksheets/sheet1.xml").decode()
        assert spill_neighbors.FILE_VERSION in book
        assert '<calcPr calcId="191029"/>' in book
        for case in spill_neighbors.cases(group):
            xml = spill_neighbors.NEIGHBORS[case.neighbor] if case.spill else None
            if xml is not None:
                assert xml.format(ref=case.neighbor_cell) in sheet, case.name
            elif case.spill:
                assert f'r="{case.neighbor_cell}"' not in sheet, case.name
    with zipfile.ZipFile(tmp_path / "spill_neighbors_zeros.xlsx") as package:
        assert 'showZeros="0"' in package.read("xl/worksheets/sheet1.xml").decode()
    assert spill_neighbors.excel_saved(book) == book  # once is enough


def test_every_case_has_a_cell_of_its_own_inside_the_markers():
    for group in spill_neighbors.GROUPS:
        cells = spill_neighbors._cells(group)
        assert len(set(cells)) == len(cells)
        first, second = spill_neighbors.MARKERS[group]
        letters, row = coordinate_from_string(second)
        assert first not in cells
        assert second not in cells
        for cell in cells:
            column, r = coordinate_from_string(cell)
            assert r < row
            assert column_index_from_string(column) < column_index_from_string(letters)
        rows = [case.row for case in spill_neighbors.cases(group) if case.spill]
        assert len(set(rows)) == len(rows), group


def test_no_case_draws_a_color():
    # the cases draw black text on white only, far from the markers' magenta
    for font in spill_neighbors.FONTS.values():
        assert "color" not in font


# --- measuring, on a synthetic capture


def _canvas(group, reads=()):
    ws = openpyxl.Workbook().active
    spill_neighbors.fill(ws, group)
    return Canvas(ws, spill_neighbors._cells(group), list(reads))


def _case(group, cell):
    return next(c for c in spill_neighbors.cases(group) if c.cell == cell)


def test_spilled_ink_and_hidden_gridlines_are_read_per_column_and_boundary():
    canvas = _canvas("blockers", [("C4", "geometry.text", "long"), ("D4", "geometry.text", "")])
    x0, y0, _, y1 = canvas.box("C4")
    top, bottom = y0, y1
    for letter in "FGH":  # gridlines at the right of E, F and G: C|D and D|E are crossed
        left = canvas.box(f"{letter}4")[0]
        canvas.paint(left, top, left + 1, bottom, GRAY)
    canvas.paint(x0 + 6, top + 8, canvas.box("E4")[0] + 30, top + 22, BLACK)
    case = _case("blockers", "C4")
    measured = spill_neighbors.measure(canvas.capture(), case)
    assert measured["ink_columns"] == "111000"
    assert measured["hidden_gridlines"] == "11000"
    assert (measured["ink_left"], measured["ink_right"]) == ("3", "-0.5")  # cut by the inset
    found = spill_neighbors.record(canvas.capture(), case)
    assert (found["text"], found["neighbor_text"], found["font"]) == ("long", "", "Calibri 11")
    assert set(spill_neighbors.COLUMNS) >= set(found)


def test_a_rows_columns_come_from_the_groups_first_row():
    # E to H are read in row 4 only; row 6's boxes take their left and right from there
    canvas = _canvas("blockers")
    y0 = canvas.box("C6")[1]
    canvas.paint(canvas.box("G4")[0] + 10, y0 + 8, canvas.box("G4")[0] + 20, y0 + 22, BLACK)
    canvas.paint(canvas.box("D4")[2] - 12, y0 + 8, canvas.box("D4")[2] - 4, y0 + 22, BLACK)
    measured = spill_neighbors.measure(canvas.capture(), _case("blockers", "C6"))
    assert measured["ink_columns"] == "010010"
    assert "ink_left" not in measured  # nothing in C


def test_a_numbers_ink_is_measured_from_its_cells_edges():
    canvas = _canvas("numbers", [("J8", "geometry.text", "##"), ("J8", "geometry.height", "19")])
    x0, y0, x1, _ = canvas.box("J8")  # J's right edge is at 455.25 points, 910.5 pixels
    canvas.paint(x0 + 8, y0 + 10, x1 - 6, y0 + 30, BLACK)
    case = _case("numbers", "J8")
    found = spill_neighbors.record(canvas.capture(), case)
    assert (found["ink_left"], found["ink_right"]) == ("4", "-3.25")
    assert (found["text"], found["row_height"]) == ("##", "19")
    assert (found["font"], found["width"]) == ("Calibri 14", "22 pt")
    assert "ink_columns" not in found


def test_every_case_is_measured_on_a_blank_capture():
    for group in spill_neighbors.GROUPS:
        cap = _canvas(group).capture()
        for case in spill_neighbors.cases(group):
            measured = spill_neighbors.measure(cap, case)
            assert "ink_left" not in measured
            if case.spill:
                assert measured["ink_columns"] == "000000"
                assert measured["hidden_gridlines"] == "11111"  # a blank canvas has none


# --- the capture (suite spill_neighbors)


def _captured():
    """The captured rows by group and cell."""
    path = spill_neighbors.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.spill_neighbors run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found: dict[str, dict[str, dict[str, str]]] = {}
    for row in body:
        record = dict(zip(header, row, strict=False))
        if record["zoom"] == str(spill_neighbors.ZOOM):
            found.setdefault(record["group"], {})[record["cell"]] = record
    return found


def _spill_cases():
    return [c for c in spill_neighbors.CASES if c.spill]


@pytest.mark.excel
@pytest.mark.parametrize("case", _spill_cases(), ids=lambda c: f"{c.group}-{c.cell}")
def test_a_neighbor_stops_the_spill_as_captured(case):
    # §4.1 and §4.18: the long text ends 26 points inside F when nothing stops it, so which
    # columns hold ink and which gridlines are hidden are yes-or-no answers, with no edge case
    # near a boundary. A blocked text shows ink in C only (or in D too, where D draws a value)
    # and every gridline from C|D on. The controls (an empty D) must spill, and the number 1
    # must block, or the capture itself is wrong.
    found = _captured()
    if case.group not in found:
        pytest.skip(f"group {case.group} not captured yet")
    record = found[case.group][case.cell]
    expected = spill_neighbors.EXPECTED[case.group][case.cell]
    assert record["text"] == expected["text"], case.name
    assert record["hidden_gridlines"] == expected["hidden_gridlines"], case.name
    assert record["ink_columns"] == expected["ink_columns"], case.name


@pytest.mark.excel
@pytest.mark.parametrize("case", _spill_cases(), ids=lambda c: f"{c.group}-{c.cell}")
def test_a_neighbor_shows_the_text_sheetpeek_gives_it(case):
    # Excel's text of D (its own read): empty for ="" and for a hidden zero, a space for " ".
    found = _captured()
    if case.group not in found:
        pytest.skip(f"group {case.group} not captured yet")
    record = found[case.group][case.cell]
    assert (
        record["neighbor_text"] == spill_neighbors.EXPECTED[case.group][case.cell]["neighbor_text"]
    ), case.name


def _number_cases():
    return [c for c in spill_neighbors.CASES if not c.spill]


UNMEASURED = "Arial 11"
"""A font xlformat has no glyph table for: sheetpeek measures it as Calibri at the same size
(decided, ARCHITECTURE §6.2). Excel's Arial "#" and digits are wider: in the 22 pt column it
shows "##" where sheetpeek gives "###", and its scientific text starts up to 2 pt further left.
A known deviation, compared apart."""


def _measured(case):
    return case.font != UNMEASURED


@pytest.mark.excel
def test_a_font_without_a_glyph_table_is_the_known_deviation():
    found = _captured()
    if "numbers" not in found:
        pytest.skip("group numbers not captured yet")
    texts = [
        case.name
        for case in _number_cases()
        if not _measured(case)
        and found["numbers"][case.cell]["text"]
        != spill_neighbors.EXPECTED["numbers"][case.cell]["text"]
    ]
    assert texts == [f"{d} digits, Arial 11, 22 pt" for d in range(7, 13)]


@pytest.mark.excel
def test_numbers_in_other_fonts_show_the_text_sheetpeek_gives_them():
    # §4.2: a number whose font is not the Normal font is formatted in its own font against the
    # Normal font's column (xlformat). Excel's text is its own read, so it is compared exactly,
    # every cell at once so that one run shows every disagreement. The row heights are §4.7's.
    found = _captured()
    if "numbers" not in found:
        pytest.skip("group numbers not captured yet")
    captured = found["numbers"]
    texts, heights = [], []
    for case in filter(_measured, _number_cases()):
        record, expected = captured[case.cell], spill_neighbors.EXPECTED["numbers"][case.cell]
        if record["text"] != expected["text"]:
            texts.append((case.name, record["text"], expected["text"]))
        if float(record["row_height"]) != expected["row_height"]:
            heights.append((case.name, record["row_height"], expected["row_height"]))
    assert texts == []
    assert heights == []


@pytest.mark.excel
def test_numbers_in_other_fonts_are_placed_where_sheetpeek_places_them():
    # The ink lies inside the glyphs' advances by their side bearings: under a tenth of an em on
    # the right for a digit, a "#" or an exponent's digit, up to about 0.15 em on the left, where a
    # leading "1" is narrow (Calibri and Arial both). The ink ends on a device pixel: half a
    # point at zoom 100 on a 2x display. So the ink's left is at most half a point before the
    # advance's left and at most 0.15 em plus half a point after it, and likewise its right is
    # within 0.1 em plus half a point before the advance's right. Only cells whose text agrees
    # are compared, the others being the previous test's.
    found = _captured()
    if "numbers" not in found:
        pytest.skip("group numbers not captured yet")
    captured = found["numbers"]
    misplaced = []
    for case in filter(_measured, _number_cases()):
        record, expected = captured[case.cell], spill_neighbors.EXPECTED["numbers"][case.cell]
        if record["text"] != expected["text"] or not record["ink_left"]:
            continue
        em = spill_neighbors.FONTS[case.font]["size"]
        left = float(record["ink_left"]) - expected["left"]
        right = float(record["ink_right"]) - expected["right"]
        if not (-0.5 <= left <= 0.15 * em + 0.5 and -(0.1 * em + 0.5) <= right <= 0.5):
            misplaced.append((case.name, round(left, 2), round(right, 2)))
    assert misplaced == []
