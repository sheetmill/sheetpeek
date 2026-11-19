"""The screen suite ``filters_more`` (filter and sort buttons, ARCHITECTURE.md §4.19), no Excel.

Its workbooks are built and read back with sheetpeek as built, and what sheetpeek draws for each
case is the suite's ``EXPECTED`` (glyphs) and ``SIZES`` (hidden rows and column). Its measuring
runs on a synthetic capture with known marks. The comparison with
``tests/captures/filters_more.tsv`` is skipped until that capture exists; it compares glyph ids
with ``tests/captures/filters.tsv``'s, which the same measuring wrote.
"""

import zipfile

import openpyxl
import pytest
from capture import filters, filters_more
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import SCALE, Canvas


def _build(tmp_path):
    placed = filters_more.build(tmp_path)
    return placed, {g: tmp_path / f"filters_more_{g}.xlsx" for g in filters_more.GROUPS}


def test_every_case_is_drawn_as_expected_when_read_back(tmp_path):
    placed, paths = _build(tmp_path)
    assert placed == list(filters_more.CASES)
    for group, path in paths.items():
        rendered = sheetpeek.read(path).layout("excel")
        buttons = {
            f"{get_column_letter(c.column)}{c.row}": c.button
            for row in rendered.rows
            for c in row.cells
        }
        for case in filters_more.cases(group):
            assert buttons.get(case.cell) == filters_more.EXPECTED[(group, case.cell)], case
        # Every button sheetpeek draws is a case: none goes unmeasured.
        drawn = {cell for cell, button in buttons.items() if button is not None}
        assert drawn <= {case.cell for case in filters_more.cases(group)}, group


def test_the_sizes_are_what_sheetpeek_lays_out(tmp_path):
    _, paths = _build(tmp_path)
    for group, path in paths.items():
        rendered = sheetpeek.read(path).layout("excel")
        heights = {row.index: row.height_pt for row in rendered.rows}
        widths = {column.index: column.width_pt for column in rendered.columns}
        for cell, size in filters_more.SIZES[group].items():
            letters, row = coordinate_from_string(cell)
            if letters == "A":
                assert heights.get(row, 0) == size, (group, cell)
            else:
                assert widths.get(column_index_from_string(letters), 0) == size, (group, cell)


def test_the_criteria_disagree_with_the_rows_saved_hidden(tmp_path):
    _, paths = _build(tmp_path)
    ws = openpyxl.load_workbook(paths["main"]).active
    kept = set(ws.auto_filter.filterColumn[0].filters.filter)
    values = {row: ws.cell(row, 8).value for row in range(5, 10)}
    assert {row for row, v in values.items() if str(v) not in kept} == {6, 8}
    assert [row for row in range(5, 10) if ws.row_dimensions[row].hidden] == [7]


def test_the_beside_workbook_holds_its_sort_beside_the_filter(tmp_path):
    _, paths = _build(tmp_path)
    with zipfile.ZipFile(paths["beside"]) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<autoFilter ref="C4:E6" />' + filters_more.BESIDE in xml
    assert xml.count("<sortState") == 1
    with pytest.raises(SystemExit):
        filters_more.sort_beside("<worksheet/>")
    # The left-to-right sorts, which Excel refuses, each in a workbook of its own: inside the
    # sheet's filter and in a table. The main workbook holds no sort.
    with zipfile.ZipFile(paths["lr_sheet"]) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<sortState columnSort="1" ref="C5:E6">' in sheet
    assert sheet.index("<sortState") < sheet.index("</autoFilter>")
    with zipfile.ZipFile(paths["lr_table"]) as z:
        table = z.read("xl/tables/table1.xml").decode()
    assert '<sortState columnSort="1" ref="C5:E6">' in table
    with zipfile.ZipFile(paths["main"]) as z:
        assert not any(b"<sortState" in z.read(n) for n in z.namelist())


def test_no_case_draws_a_fill_or_a_font_color(tmp_path):
    # Markers are magenta: the cases draw only Excel's defaults and its filter buttons.
    _, paths = _build(tmp_path)
    for path in paths.values():
        ws = openpyxl.load_workbook(path).active
        for row in ws.iter_rows():
            for cell in row:
                assert cell.fill.fill_type is None, cell.coordinate
                assert cell.font.color is None or cell.font.color.theme == 1, cell.coordinate


def test_every_case_and_marker_has_its_place():
    for group in filters_more.GROUPS:
        found = filters_more.cases(group)
        assert len({c.cell for c in found}) == len(found)
        first, second = (coordinate_from_string(m) for m in filters_more.MARKERS[group])
        assert first[0] != second[0]
        assert first[1] != second[1]
        for case in found:
            letters, row = coordinate_from_string(case.cell)
            assert row < second[1], case
            assert column_index_from_string(letters) < column_index_from_string(second[0]), case
        references = filters_more._references(group)
        assert all(c.startswith("A") or c.endswith("1") for c in references)


# --- measuring, on a synthetic capture


def _canvas(group="main"):
    ws = openpyxl.Workbook().active
    filters_more.fill(ws, group)
    return Canvas(ws, filters_more._references(group))


def _box(canvas, cell):
    letters, row = coordinate_from_string(cell)
    left, _, right, _ = canvas.box(f"{letters}1")
    _, top, _, bottom = canvas.box(f"A{row}")
    return left, top, right, bottom


def test_a_button_and_its_glyph_are_measured_in_a_case_cell():
    canvas = _canvas()
    _, top, right, bottom = _box(canvas, "H4")
    canvas.paint(right - 30, top, right, bottom, (240, 240, 240))  # 15 pt square
    canvas.paint(right - 20, top + 10, right - 10, bottom - 14, (60, 60, 60))  # 5 by 3 pt
    records = {r["cell"]: r for r in filters_more.measure(canvas.capture(), "main")}
    found = records["H4"]
    assert found["group"] == "main"
    assert found["case"] == "values filter, criteria not applied"
    assert (found["box_w"], found["box_h"], found["right_gap"], found["bottom_gap"]) == (
        "15", "15", "0", "0"
    )  # fmt: skip
    assert (found["glyph_w"], found["glyph_h"], found["glyph_color"]) == ("5", "3", "3C3C3C")
    assert len(found["glyph"]) == 8
    # The other cases show nothing on a blank capture.
    assert all("box_w" not in r for cell, r in records.items() if cell != "H4")
    assert set(filters_more.COLUMNS) >= set(found)


def test_a_button_below_the_row_it_belongs_to_is_found_in_its_own_row():
    # E5 lies under the header merged down: a button Excel draws at the merge's bottom shows in
    # E5's record, not E4's.
    canvas = _canvas()
    _, top, right, bottom = _box(canvas, "E5")
    canvas.paint(right - 30, top, right, bottom, (240, 240, 240))
    records = {r["cell"]: r for r in filters_more.measure(canvas.capture(), "main")}
    assert records["E5"]["box_w"] == "15"
    assert "box_w" not in records["E4"]
    assert "glyph" not in records["E5"]


def test_sizes_are_read_from_the_reference_cells():
    canvas = _canvas()
    canvas.geometry["A7"]["height"] = 0.0
    canvas.geometry["A12"]["height"] = 0.0
    canvas.geometry["G1"]["width"] = 0.0
    found = {r["cell"]: r for r in filters_more.sizes(canvas.capture(), "main")}
    assert {cell: float(r["size"]) for cell, r in found.items()} == filters_more.SIZES["main"]
    assert found["G1"]["case"] == "column width"
    assert found["A7"]["case"] == "row height"


def test_a_case_cell_takes_its_box_from_its_references():
    canvas = _canvas()
    cap = canvas.capture()
    filters_more.place(cap, "D4")
    column, row = canvas.geometry["D1"], canvas.geometry["A4"]
    assert cap.edge("D4", "left") == column["left"] * SCALE
    assert cap.edge("D4", "right") == (column["left"] + column["width"]) * SCALE
    assert cap.edge("D4", "bottom") == (row["top"] + row["height"]) * SCALE


# --- the capture (suite filters_more)


def _rows(path):
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return [dict(zip(header, row, strict=False)) for row in body]


def _captured():
    """The captured records by group, and the groups Excel refused with its message."""
    if not filters_more.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.filters_more run")
    found: dict[str, list[dict[str, str]]] = {}
    refused: dict[str, str] = {}
    for record in _rows(filters_more.OUTPUT):
        if record["case"].startswith("refused:"):
            refused[record["group"]] = record["case"]
        else:
            found.setdefault(record["group"], []).append(record)
    return found, refused


def _glyphs():
    """filters.tsv's glyph id for each (state, zoom, box size), as the suite filters captured."""
    found = {}
    for record in _rows(filters.OUTPUT):
        state = filters.EXPECTED[record["cell"]]
        if state is not None:
            found[(state, record["zoom"], record["box_w"])] = record["glyph"]
    return found


NARROWEST_BUTTON = 10
"""Screen points: a button is min(row height x zoom, 17) square, 15 or 17 here. A narrower mark
is the cell's own text (L13's right-aligned "1", 2.5 points wide in black at zoom 200, where
the measurer looks for a box)."""


def _button_width(record):
    """The button's width, or ``None`` where the record holds no button."""
    width = float(record["box_w"]) if record.get("box_w") else 0.0
    return width if width >= NARROWEST_BUTTON else None


def _buttons(group):
    found, refused = _captured()
    if group in refused:
        pytest.skip(
            f"Excel refused the {group} workbook ({refused[group]}): the rule does not arise"
        )
    if group not in found:
        pytest.skip(f"{group} not captured yet")
    return [r for r in found[group] if r["zoom"]]


@pytest.mark.excel
@pytest.mark.parametrize("group", filters_more.GROUPS)
def test_each_case_draws_the_glyph_sheetpeek_gives_it_as_captured(group):
    # Glyph ids are bitmaps compared exactly: the same measuring on the same Excel draws the
    # same pixels for one state, zoom and button size (filters.tsv: every state one id at each
    # zoom). A case sheetpeek draws without a button shows no box at all.
    glyphs = _glyphs()
    records = _buttons(group)
    assert len(records) == len(filters_more.cases(group)) * len(filters_more.ZOOMS)
    for record in records:
        state = filters_more.EXPECTED[(group, record["cell"])]
        if state is None:
            assert _button_width(record) is None, record
            continue
        assert record.get("glyph") == glyphs[(state, record["zoom"], record["box_w"])], record


@pytest.mark.excel
@pytest.mark.parametrize("group", filters_more.GROUPS)
def test_each_button_is_as_large_as_the_row_up_to_17_screen_points_as_captured(group):
    # The rule test_filters checks on the suite filters, exactly (those captures measured it with
    # no deviation): every case row is 15 pt, so the square is 15 at zoom 100 and 17 at 200,
    # its right edge on its own column's boundary (each column of a header merged across too).
    for record in _buttons(group):
        if _button_width(record) is None:
            continue
        height = 15 * int(record["zoom"]) / 100
        size = min(height, 17)
        gap = 0 if size == height else 1
        assert (float(record["box_w"]), float(record["box_h"])) == (size, size), record
        assert (float(record["right_gap"]), float(record["bottom_gap"])) == (0, gap), record
        assert float(record["top"]) == height - gap - size, record


@pytest.mark.excel
@pytest.mark.parametrize("group", filters_more.GROUPS)
def test_the_rows_saved_hidden_are_what_shows_as_captured(group):
    # Excel's own row heights and column widths (exact: Excel reports a hidden row or column as
    # 0, a 15 pt row as 15): the criteria hiding rows 6 and 8 left them shown and row 7 hidden.
    found, refused = _captured()
    if group in refused or group not in found:
        pytest.skip(f"{group} not captured")
    sizes = {r["cell"]: float(r["size"]) for r in found[group] if not r["zoom"]}
    assert sizes == filters_more.SIZES[group]
