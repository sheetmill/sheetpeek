"""The screen-capture suites' workbooks and measuring, without Excel.

Each suite's workbook is built and read back as built. Its measuring runs on a synthetic capture:
a white image with cells laid out from the worksheet's own column widths and row heights at 2
pixels a point (zoom 100 on a 2x display), with a few known marks drawn in it.
"""

import dataclasses

import openpyxl
import pytest
from capture import borders, fills, filters, text_style
from capture.screen.window import GRIDLINE, NO_GRIDLINE, locate, nominal_scale
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import SCALE, Canvas

# --- borders


def test_the_border_workbook_reads_back_as_built(tmp_path):
    placed = borders.build(tmp_path / "borders.xlsx")
    sheet = sheetpeek.read(tmp_path / "borders.xlsx")
    assert len(placed) == 14
    for coordinate, style in placed:
        letters, row = coordinate_from_string(coordinate)
        sides = sheet.cell(row, column_index_from_string(letters)).style.borders
        if style == "none":
            assert sides is None
            continue
        for side in (sides.left, sides.right, sides.top, sides.bottom):
            assert side is not None
            assert side.style == style, coordinate


def test_a_dashed_bottom_reads_as_its_runs_in_points():
    ws = openpyxl.Workbook().active
    borders.fill(ws)
    canvas = Canvas(ws, list(borders.CELLS.values()))
    x0, _, x1, bottom = canvas.box(borders.CELLS["dashed"])
    for x in range(x0, x1, 8):  # 3 pt on, 1 pt off, two device rows before the boundary
        canvas.paint(x, bottom - 2, min(x + 6, x1), bottom, (0, 0, 0))
    records = [r for r in borders.measure(canvas.capture()) if r["style"] == "dashed"]
    assert [(r["edge"], r["offset"], r["pattern"]) for r in records] == [
        ("bottom", "-1", "on 3 off 1"),
        ("bottom", "-0.5", "on 3 off 1"),
    ]
    assert records[0]["ink"] == "000000"
    assert records[0]["gaps"] == "white"
    assert set(borders.COLUMNS) >= set(records[0])


# --- fills


def test_the_fill_workbook_reads_back_as_built(tmp_path):
    placed = fills.build(tmp_path / "fills.xlsx")
    sheet = sheetpeek.read(tmp_path / "fills.xlsx")
    assert len(placed) == 2 * 19 + 3
    for pattern, colors, coordinate in placed:
        letters, row = coordinate_from_string(coordinate)
        spec = sheet.cell(row, column_index_from_string(letters)).style.fill_spec
        if pattern == "none":
            assert spec is None
            continue
        fg, bg = ("000000", "FFFFFF") if colors == "anchor" else fills.SETS[colors][:2]
        assert (spec.kind, spec.fg) == (pattern, fg), coordinate
        assert spec.bg == (None if pattern == "solid" else bg), coordinate


def test_a_tile_drawn_in_whole_points_reads_one_character_a_point():
    ws = openpyxl.Workbook().active
    fills.fill(ws)
    canvas = Canvas(ws, [cell for _, _, cell in fills.cases()])
    cell = next(c for p, s, c in fills.cases() if (p, s) == ("darkVertical", "color"))
    x0, y0, x1, y1 = canvas.box(cell)
    for y in range(y0, y1):  # 2 pt of blue, 2 pt of yellow, from the sheet's origin
        for x in range(x0, x1):
            canvas.paint(x, y, x + 1, y + 1, (0, 0, 255) if x // 4 % 2 == 0 else (255, 255, 0))
    record = next(r for r in fills.measure(canvas.capture()) if r["cell"] == cell)
    assert (record["period_x"], record["period_y"], record["whole_points"]) == ("4", "0.5", "yes")
    assert record["tile_sheet"] == "##.."
    assert abs(float(record["fg_share"]) - 0.5) < 0.02
    assert (record["fill_left"], record["fill_right"]) == ("0", "0")
    assert set(fills.COLUMNS) >= set(record)


# --- text styles


def test_the_text_style_workbook_reads_back_as_built(tmp_path):
    found = text_style.build(tmp_path / "text_style.xlsx")
    assert len(found) == len(text_style.cases())
    shrink = sheetpeek.read(tmp_path / "text_style.xlsx", "Shrink")
    assert shrink.cell(4, 3).style.shrink
    assert not shrink.cell(4, 7).style.shrink
    assert [str(m) for m in shrink.merges] == ["E14:F14"]
    fonts = sheetpeek.read(tmp_path / "text_style.xlsx", "Fonts")
    for row, underline in text_style.UNDERLINES.items():
        assert fonts.cell(row, 3).style.font.underline == underline
    assert fonts.cell(14, 5).style.font.vert_align == "superscript"
    assert fonts.cell(18, 7).style.font.strike
    align = sheetpeek.read(tmp_path / "text_style.xlsx", "Align")
    assert align.cell(16, 7).style.justify_last_line
    assert align.cell(26, 5).style.reading_order == 2
    assert align.cell(10, 3).style.v_align == "distributed"


def test_the_filter_workbook_reads_back_as_built(tmp_path):
    # Every header cell's glyph as sheetpeek lays it out is the suite's expectation, which the
    # capture then checks against Excel.
    path = tmp_path / "filters.xlsx"
    assert len(filters.build(path)) == len(filters.CASES)
    rendered = sheetpeek.read(path).layout("excel")
    found = {
        f"{get_column_letter(c.column)}{c.row}": c.button
        for row in rendered.rows
        for c in row.cells
        if f"{get_column_letter(c.column)}{c.row}" in filters.CASES
    }
    assert found == filters.EXPECTED


def test_every_case_names_a_distinct_cell_in_its_group():
    for group in text_style.GROUPS:
        found = text_style.cases(group)
        assert len({case.cell for case in found}) == len(found), group


def test_each_wrap_text_measures_the_width_it_is_listed_under():
    import xlformat

    for table in (text_style.NARROW, text_style.WIDE, text_style.WORDS):
        for width, text in table.items():
            assert xlformat.text_width(text, font=None) == width, text


def test_the_second_rotation_set_leaves_the_captured_windows_as_they_were():
    from capture import rotation

    windows, read = rotation._on_screen()
    first = [case for _, chosen in rotation._windows(more=False) for case in chosen]
    assert first == [case for chosen in windows for case in chosen]
    lines = (rotation.OUTPUT).read_text(encoding="utf-8").splitlines()
    captured = [line.split("\t")[1] for line in lines if line and not line.startswith("#")][1:]
    assert captured[: len(first) + len(read)] == [case.name for case in first + read]
    names = [case.name for case in rotation.cases() + rotation.more_cases()]
    assert len(set(names)) == len(names)


def test_an_underline_reads_as_a_rule_below_the_baseline():
    ws = openpyxl.Workbook().active
    found = text_style.fill(ws, "fonts")
    canvas = Canvas(ws, text_style._cells(found))
    case = next(c for c in found if c.name == "underline single text")
    x0, y0, _, y1 = canvas.box(case.cell)
    for x in range(x0 + 6, x0 + 40, 8):  # the glyphs' stems: 7 pt high
        canvas.paint(x, y1 - 20, x + 2, y1 - 6, (0, 0, 0))
    canvas.paint(x0 + 6, y1 - 4, x0 + 40, y1 - 2, (0, 0, 0))  # a 1 pt rule 1 pt below
    measured = text_style.measure(canvas.capture(), case)
    assert measured["baseline"] == f"{(y1 - 6 - y0) / SCALE:g}"
    assert (measured["rule_offset"], measured["rule_thickness"]) == ("1", "1")
    assert measured["rule_start"] == "3"
    assert (measured["ink_left"], measured["ink_top"]) == ("3", f"{(y1 - 20 - y0) / SCALE:g}")


def test_every_text_case_is_measured_on_a_blank_capture():
    for group in text_style.GROUPS:
        ws = openpyxl.Workbook().active
        found = text_style.fill(ws, group)
        cap = Canvas(ws, text_style._cells(found)).capture()
        for case in found:
            measured = text_style.measure(cap, case)
            assert measured is not None, case.name
            assert set(text_style.COLUMNS) >= set(measured)


def _text_style_cases(group):
    """The captured rows of a text_style group at zoom 100, by cell."""
    path = text_style.OUTPUT
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found = [dict(zip(header, row, strict=False)) for row in body]
    return {r["cell"]: r for r in found if r["group"] == group and r["zoom"] == "100"}


def _drawn_lines(tmp_path, group):
    """sheetpeek's lines of each cell of a text_style group: (start, end) from the cell's left."""
    text_style.build(tmp_path / "text_style.xlsx")
    rendered = sheetpeek.read(tmp_path / "text_style.xlsx", group.capitalize()).layout("excel")
    x0 = {c.index: c.x_pt for c in rendered.columns}
    found: dict[str, dict[float, list[float]]] = {}
    for row in rendered.rows:
        for run in row.runs:
            r, c = run.source
            line = found.setdefault(f"{get_column_letter(c)}{r}", {}).setdefault(run.y, [])
            line.extend((run.x - x0[c], run.x + run.width - x0[c]))
    return {cell: [(min(v), max(v)) for _, v in sorted(ys.items())] for cell, ys in found.items()}


def test_wrapped_lines_break_where_excel_breaks_them(tmp_path):
    # Captured (text_style group wrap, zoom 100): a line fits up to 48 pt in a default column
    # and 115 in one of 20 characters, shrink to fit or not. A band under 3 pt is an i's dot.
    captured = _text_style_cases("wrap")
    ours = _drawn_lines(tmp_path, "wrap")
    assert len(captured) == 43
    for cell, row in captured.items():
        bands = [b for b in row["line_tops"].split() if float(b.split("+")[1]) >= 3]
        assert len(ours[cell]) == len(bands), (row["name"], ours[cell], bands)


def test_distributed_lines_sit_where_excel_draws_them(tmp_path):
    # Captured (text_style group lone, zoom 100): each line's ink against sheetpeek's advance
    # box, within the glyphs' side bearings; a 15 pt row shows its first line only.
    captured = _text_style_cases("lone")
    ours = _drawn_lines(tmp_path, "lone")
    assert len(captured) == 12
    for cell, row in captured.items():
        theirs = [tuple(map(float, s.split("+"))) for s in row["line_inks"].split()]
        lines = ours[cell] if row["row_height"] != "15.0" else ours[cell][:1]
        assert len(lines) == len(theirs), (row["name"], lines, theirs)
        for (start, end), (left, width) in zip(lines, theirs, strict=True):
            assert abs(left - start) <= 1.5, (row["name"], start, left)
            assert abs(left + width - end) <= 1.5, (row["name"], end, left + width)


def test_lines_sit_at_the_baselines_excel_draws(tmp_path):
    # Captured (text_style group align, zoom 100, Calibri 11 in rows of 60 pt): top, center,
    # justify, distributed and bottom, one line, three lines and a long unwrapped text. A line's
    # baseline is the bottom of its ink, less the descender (2 pt) on the line with a "g". The
    # screen puts a top-aligned baseline 12 pt below the row's top and a bottom-aligned one 3 pt
    # above its bottom, lines 15 pt apart; print has 11 and 3 (layout rows 152 to 154).
    captured = {
        cell: row
        for cell, row in _text_style_cases("align").items()
        if row["name"].startswith("vertical ")
    }
    assert len(captured) == 15
    _drawn_lines(tmp_path, "align")
    rendered = sheetpeek.read(tmp_path / "text_style.xlsx", "Align").layout("excel")
    ours: dict[str, set[float]] = {}
    for row in rendered.rows:
        for run in row.runs:
            r, c = run.source
            ours.setdefault(f"{get_column_letter(c)}{r}", set()).add(run.y - row.y_pt)
    for cell, row in captured.items():
        bands = [tuple(map(float, b.split("+"))) for b in row["line_tops"].split()]
        descent = 2.0 if "g" in row["text"] else 0.0
        theirs = [top + height - descent for top, height in bands]
        assert sorted(ours[cell]) == theirs, (row["name"], sorted(ours[cell]), theirs)


# --- grids and platforms


def _marked(cells, markers):
    ws = openpyxl.Workbook().active
    canvas = Canvas(ws, [*cells, *markers])
    for marker in markers:
        x0, y0, x1, y1 = canvas.box(marker)
        canvas.paint(x0, y0, x1, y1, (255, 0, 255))
    return canvas


def test_one_marker_and_the_nominal_scale_give_the_two_marker_grid():
    canvas = _marked(["E6"], ("B2", "K12"))
    image = canvas.capture().image
    first, second = canvas.geometry["B2"], canvas.geometry["K12"]
    nominal = nominal_scale(100, 2.0, "mac")
    two, check = locate(image, first, second, nominal)
    one, _ = locate(image, first, second, nominal, one_marker=True)
    assert check["nominal_off"] == 0
    for points in (0.0, 100.0, 333.25):
        assert one.x(points) == two.x(points)
        assert one.y(points) == two.y(points)


def test_a_screen_showing_only_the_first_marker_needs_the_nominal_scale():
    canvas = _marked(["K12"], ("B2",))
    image = canvas.capture().image
    first, second = canvas.geometry["B2"], canvas.geometry["K12"]
    grid, check = locate(image, first, second, nominal_scale(100, 2.0, "mac"))
    assert check["scale"] == "nominal"
    assert (grid.x(first["left"]), grid.y(first["top"])) == canvas.box("B2")[:2]
    with pytest.raises(RuntimeError, match="expected 2 marker regions"):
        locate(image, first, second)


def test_windows_units_are_96_pixels_an_inch():
    assert nominal_scale(200, 1.0, "windows") == pytest.approx(8 / 3)
    assert nominal_scale(200, 2.0, "mac") == 4.0


def test_the_borders_control_takes_its_platforms_gridline_or_is_skipped(capsys):
    ws = openpyxl.Workbook().active
    borders.fill(ws)
    canvas = Canvas(ws, list(borders.CELLS.values()))
    x0, _, x1, bottom = canvas.box(borders.CELLS["none"])
    canvas.paint(x0 - 2, bottom - 1, x1 + 2, bottom, GRIDLINE["mac"])

    def control(platform):
        capture = dataclasses.replace(canvas.capture(), platform=platform)
        records = borders.measure(capture)
        return [r for r in records if r["style"] == "none" and r["edge"] == "bottom"]

    assert control("mac")
    assert GRIDLINE["windows"] is None  # stated once a capture here records it
    assert control("windows") == []
    assert NO_GRIDLINE in capsys.readouterr().out
