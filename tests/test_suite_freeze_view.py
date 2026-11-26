"""The screen suite ``freeze_view``: its workbooks, its measuring, and the comparison.

ARCHITECTURE.md §4.18 (what Excel's Normal view draws, the freeze line, page break preview, the
saved zoom). The workbooks are built and read back with sheetpeek; the measuring runs on
synthetic captures; the comparison with ``tests/captures/freeze_view.tsv`` is skipped until that
capture exists.
"""

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest
from capture import freeze_view
from capture.screen.window import GRIDLINE
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from sheetpeek.backends import HTMLBackend
from sheetpeek.backends.html._panes import BREAK_LINE, FREEZE_LINE
from tests.screen_canvas import SCALE, Canvas

CAPTURED = Path(__file__).parent / "captures" / "freeze_view.tsv"


@pytest.fixture(scope="module")
def folder(tmp_path_factory):
    found = tmp_path_factory.mktemp("freeze_view")
    freeze_view.build(found)
    return found


def _sheet(folder, name):
    return sheetpeek.read(folder / f"freeze_view_{name}.xlsx", platform="mac")


def _at(cell):
    letters, row = coordinate_from_string(cell)
    return row, column_index_from_string(letters)


# --- the workbooks


def test_the_frozen_sheet_reads_back_as_built(folder):
    for name in freeze_view.SCREEN:
        sheet = _sheet(folder, name)
        view = sheet.view
        assert (view.frozen_rows, view.frozen_columns, view.pane_state) == (2, 2, "frozen")
        assert [str(a) for a in sheet.print_areas] == [freeze_view.PRINT_AREA]
        assert (sheet.row_breaks, sheet.column_breaks) == ((16,), (8,))
        for cell in ("G2", "G3", "B7", "C7"):
            spec = sheet.cell(*_at(cell)).style.fill_spec
            assert (spec.kind, spec.fg) == ("solid", freeze_view.YELLOW), (name, cell)
        assert sheet.cell(*_at("I3")).style.fill_spec.fg == freeze_view.TEAL
        assert sheet.cell(*_at("K2")).style.borders.bottom.style == "medium"
        top = sheet.cell(*_at("M3")).style.borders.top
        assert (top.style, top.color) == ("thick", freeze_view.BLUE)
        assert sheet.cell(*_at("B9")).style.borders.right.style == "medium"
        assert sheet.cell(*_at("B11")).value == freeze_view.SPILL_RIGHT
        assert sheet.cell(*_at("C13")).value == freeze_view.SPILL_LEFT


def test_the_crafted_views_are_in_the_files_as_written(folder):
    def view(name):
        with zipfile.ZipFile(folder / f"freeze_view_{name}.xlsx") as archive:
            text = archive.read("xl/worksheets/sheet1.xml").decode()
        return re.search(r"<sheetView [^>]*>", text)[0]

    assert 'view="pageBreakPreview" zoomScaleSheetLayoutView="100"' in view("page_break")
    assert "zoomScale" not in view("normal")
    for name, value in freeze_view.ZOOM_CASES.items():
        assert f'zoomScale="{value}"' in view(f"zoom_{name}")
    assert set(freeze_view.REFUSALS) == {f"zoom_{n}" for n in freeze_view.ZOOM_CASES}
    assert not set(freeze_view.REFUSALS) & set(freeze_view.SCREEN)


def test_the_expected_zooms_are_sheetpeeks(folder):
    for name, expected in freeze_view.EXPECTED_ZOOM.items():
        assert _sheet(folder, f"zoom_{name}").view.zoom == expected, name
    assert set(freeze_view.EXPECTED_ZOOM) == set(freeze_view.ZOOM_CASES)


def test_the_expected_lines_are_sheetpeeks(folder):
    assert freeze_view.EXPECTED["freeze_color"] == FREEZE_LINE.upper()
    assert freeze_view.EXPECTED["break_color"] == BREAK_LINE.upper()
    page = sheetpeek.render(
        folder / "freeze_view_normal.xlsx",
        HTMLBackend(page_breaks=True, freeze="static"),
        on_notice="ignore",
    )

    def thickness(name):
        found = re.findall(rf'class="sp-{name}" style="[^"]*width:(\d+)px;height:(\d+)px"', page)
        return [min(float(w), float(h)) for w, h in found]

    freeze, breaks = thickness("freeze"), thickness("pagebreak")
    assert freeze == [freeze_view.EXPECTED["freeze_width"]] * 2
    # The row break and the column break, both inside the drawn range.
    assert breaks == [freeze_view.EXPECTED["break_width"]] * 2
    # In the Normal view (no page_breaks option) nothing is drawn for the breaks.
    normal = sheetpeek.render(folder / "freeze_view_normal.xlsx", HTMLBackend(), on_notice="ignore")
    assert "sp-pagebreak" not in normal


def test_the_expected_spills_are_sheetpeeks(folder):
    rendered = _sheet(folder, "normal").layout("excel")
    line = next(c.x_pt for c in rendered.columns if c.index == 3)  # column C's left
    found = {}
    for item, text_cell, _ in freeze_view.SPILLS:
        source = _at(text_cell)
        (run,) = [r for row in rendered.rows for r in row.runs if r.source == source]
        found[item] = (round(run.x - line), round(run.x + run.width - line))
    assert found == freeze_view.EXPECTED_SPILL
    for left, right in found.values():
        assert left < 0 < right  # both texts cross the freeze line in sheetpeek's layout


def test_no_case_draws_the_marker_color():
    for color in (freeze_view.YELLOW, freeze_view.TEAL, freeze_view.BLUE, freeze_view.BLACK):
        r, g, b = (int(color[i : i + 2], 16) for i in (0, 2, 4))
        assert not (r > 150 and b > 150 and g < 120), color
    for color in (freeze_view.EXPECTED["freeze_color"], freeze_view.EXPECTED["break_color"]):
        r, g, b = (int(color[i : i + 2], 16) for i in (0, 2, 4))
        assert not (r > 150 and b > 150 and g < 120), color


def test_the_markers_keep_inside_the_print_area_and_off_the_lines():
    area = openpyxl.worksheet.cell_range.CellRange(freeze_view.PRINT_AREA)
    used = {cell for _, cell, _ in freeze_view.EDGES} | {"G2", "B7", "K2", "B9", "B11", "C13"}
    for marker in freeze_view.MARKERS:
        row, column = _at(marker)
        assert area.min_row < row < area.max_row
        assert area.min_col < column < area.max_col
        assert row > 3  # off the headings and off the freeze lines
        assert column > 3
        assert row not in (freeze_view.ROW_BREAK, freeze_view.ROW_BREAK + 1)
        assert column not in (freeze_view.COLUMN_BREAK, freeze_view.COLUMN_BREAK + 1)
        assert marker not in used
    first, second = (_at(m) for m in freeze_view.MARKERS)
    assert first[0] != second[0]
    assert first[1] != second[1]


def test_the_state_of_a_window_gives_its_zoom():
    assert freeze_view.zoom_of("zoom=85;view=normal view;freeze_panes=false") == "85"
    assert freeze_view.zoom_of("zoom=ERR;view=ERR") == "ERR"
    assert freeze_view.zoom_of("") == ""


# --- the measuring, on a synthetic capture

GRAY = (155, 155, 155)


def _canvas():
    ws = openpyxl.Workbook().active
    cells = freeze_view.fill(ws)
    state = "zoom=100;view=normal view;freeze_panes=true;split_row=2;split_column=2"
    cells += ["B7", "B13", "G2", "K2", "B9"]
    return Canvas(ws, [*freeze_view.MARKERS, *cells], rows=[("window", "state.100", state)])


def _records(canvas):
    records = freeze_view.measure(canvas.capture(), "normal")
    assert all(set(r) <= set(freeze_view.COLUMNS) for r in records)
    return records


def test_a_line_in_a_strip_is_found_nearest_the_boundary():
    white, gray = "FFFFFF", "9B9B9B"
    assert freeze_view.line([white] * 6 + [gray] * 2 + [white] * 4) == (0, 2, gray)
    assert freeze_view.line([white] * 5 + [gray] * 2 + [white] * 5) == (-1, 2, gray)
    assert freeze_view.line(["FFFF00"] * 5 + [gray] * 2 + ["FFFF00"] * 5) == (-1, 2, gray)
    # A blue line between the gray outside of a print area and its white inside.
    assert freeze_view.line(["C8C8C8"] * 4 + ["0000D0"] * 6 + [white] * 2) == (-2, 6, "0000D0")
    assert freeze_view.line([white] * 12) is None
    # Two marks: the one at the boundary, not the one farther out.
    assert freeze_view.line([white, "000000", white, white, white, gray, gray, white]) == (
        1,
        2,
        gray,
    )
    found = freeze_view.runs([white] * 5 + [gray] * 2 + [white] * 5)
    assert found == "FFFFFFx5 9B9B9Bx2 FFFFFFx5"
    assert freeze_view.strip(found) == [white] * 5 + [gray] * 2 + [white] * 5


def test_the_freeze_lines_read_over_open_cells_and_fills():
    canvas = _canvas()
    x = canvas.box("C5")[0]
    y = canvas.box("E3")[1]
    canvas.paint(*canvas.box("B7")[:2], canvas.box("C7")[2], canvas.box("C7")[3], (255, 255, 0))
    canvas.paint(x - 1, 0, x + 1, canvas.height, GRAY)  # one screen point, over the fill
    canvas.paint(0, y - 1, canvas.width, y + 1, GRAY)
    records = _records(canvas)
    (window,) = [r for r in records if r["kind"] == "window"]
    assert "split_row=2" in window["state"]
    edges = {r["item"]: r for r in records if r["kind"] == "edge"}
    assert len(edges) == len(freeze_view.EDGES)
    for item in ("freeze v open", "freeze h open", "freeze v fill"):
        found = edges[item]
        assert (found["line_offset"], found["line_px"], found["line_spt"]) == ("-1", "2", "1")
        assert found["line_color"] == "9B9B9B"
    assert edges["freeze v fill"]["strip"] == "FFFF00x5 9B9B9Bx2 FFFF00x5"
    assert edges["row break"]["strip"] == "FFFFFFx12"
    assert "line_px" not in edges["row break"]


def test_a_page_break_mark_reads_as_its_width():
    canvas = _canvas()
    y = canvas.box("E17")[1]
    canvas.paint(0, y - 3, canvas.width, y + 3, (0, 0, 208))  # 3 screen points
    x = canvas.box("A19")[2]  # the print area's left: gray outside, white inside
    canvas.paint(0, 0, x, canvas.height, (200, 200, 200))
    canvas.paint(x - 3, 0, x + 3, canvas.height, (0, 0, 208))
    edges = {r["item"]: r for r in _records(canvas) if r["kind"] == "edge"}
    found = edges["row break"]
    assert (found["line_offset"], found["line_px"], found["line_spt"]) == ("-3", "6", "3")
    assert found["line_color"] == "0000D0"
    left = edges["print area left"]
    assert (left["line_px"], left["line_color"]) == ("6", "0000D0")
    paper = {r["item"]: r["color"] for r in _records(canvas) if r["kind"] == "paper"}
    assert paper == {"outside print area": "C8C8C8", "inside print area": "FFFFFF"}


def test_spilled_text_reads_on_both_sides_of_the_line():
    canvas = _canvas()
    x = canvas.box("C11")[0]
    _, top, _, bottom = canvas.box("B11")
    canvas.paint(x - 1, 0, x + 1, canvas.height, GRAY)
    # Text from 46 points before the line to 88 after it, cut by the line (the line on top).
    canvas.paint(x - 46 * SCALE, top + 8, x - 2, top + 20, (0, 0, 0))
    canvas.paint(x + 2, top + 8, x + 88 * SCALE, top + 20, (0, 0, 0))
    _, top13, _, _ = canvas.box("B13")
    canvas.paint(x - 30 * SCALE, top13 + 8, x + 40 * SCALE, top13 + 20, (0, 0, 0))
    spills = {r["item"]: r for r in _records(canvas) if r["kind"] == "spill"}
    right = spills["spill right"]
    assert (right["ink_left"], right["ink_right"], right["dark_on_line"]) == ("-46", "88", "0")
    left = spills["spill left"]
    assert (left["ink_left"], left["ink_right"]) == ("-30", "40")
    assert left["dark_on_line"] == str(2 * 12)  # text over the line: 2 columns, 12 rows
    assert bottom > top


# --- the comparison with Excel's window


def _captured():
    if not CAPTURED.exists():
        pytest.skip("not captured yet: tests/captures/freeze_view.tsv")
    lines = [line.split("\t") for line in CAPTURED.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return [dict(zip(header, row, strict=False)) for row in body]


def _rgb(text):
    return tuple(int(text[i : i + 2], 16) for i in (0, 2, 4))


def _near(a, b, tolerance):
    return max(abs(x - y) for x, y in zip(_rgb(a), _rgb(b), strict=True)) <= tolerance


COLOR_TOLERANCE = 10
"""Per channel, for a line's color: the captures are converted to sRGB, and flat colors come
back within a few units (the gridline's 212 and the headings' 22 and 86 exact in the earlier
suites' windows); the line's mean color is averaged over 5 pixels along it."""
WIDTH_TOLERANCE = 0.5
"""Screen points, for a line's width: one device pixel on a 2x display, as the line's ends fall
on device pixels."""
SPILL_TOLERANCE = 3.0
"""Sheet points, for where spilled ink starts and ends against sheetpeek's glyph advances: the
first and last glyphs' side bearings (about a point each in Calibri 11) and Excel setting glyphs
on whole points."""


def _edges(rows, group, items):
    return [r for r in rows if r["group"] == group and r["kind"] == "edge" and r["item"] in items]


def _open(rows):
    """The open freeze lines by (zoom, appearance, direction): the line every other compares to."""
    found = {}
    for row in _edges(rows, "normal", ("freeze h open", "freeze v open")):
        found[row["zoom"], row["appearance"], row["side"]] = row
    return found


@pytest.mark.excel
def test_the_freeze_line_is_one_screen_point_of_gray_155():
    # At zoom 100 and 200, in the dark and light appearances.
    found = _open(_captured())
    assert {zoom for zoom, _, _ in found} >= {"100", "200"}
    for key, row in found.items():
        assert row.get("line_px"), (key, row)
        assert float(row["line_spt"]) == pytest.approx(
            freeze_view.EXPECTED["freeze_width"], abs=WIDTH_TOLERANCE
        ), (key, row)
        assert _near(row["line_color"], freeze_view.EXPECTED["freeze_color"], COLOR_TOLERANCE), (
            key,
            row,
        )


@pytest.mark.excel
def test_the_freeze_line_is_drawn_over_fills_and_borders():
    rows = _captured()
    opened = _open(rows)
    items = (
        "freeze h fill", "freeze h fill below", "freeze h medium border", "freeze h thick blue",
        "freeze v fill", "freeze v medium border",
    )  # fmt: skip
    compared = 0
    for row in _edges(rows, "normal", items):
        reference = opened.get((row["zoom"], row["appearance"], row["side"]))
        if reference is None or not reference.get("line_px"):
            continue
        colors = freeze_view.strip(row["strip"])
        middle = len(colors) // 2
        start = middle + int(reference["line_offset"])
        drawn = colors[start : start + int(reference["line_px"])]
        assert all(_near(c, reference["line_color"], COLOR_TOLERANCE) for c in drawn), row
        compared += 1
    assert compared


@pytest.mark.excel
def test_text_spills_across_the_freeze_line_under_it():
    # sheetpeek draws both texts across the line and the line over them.
    rows = _captured()
    compared = 0
    for row in rows:
        if row["kind"] != "spill" or row["group"] != "normal":
            continue
        expected_left, expected_right = freeze_view.EXPECTED_SPILL[row["item"]]
        assert row["ink_left"], row
        assert row["ink_right"], row
        assert float(row["ink_left"]) < 0 < float(row["ink_right"]), row
        assert row["dark_on_line"] == "0", row
        if row["zoom"] == "100":
            # At zoom 200 Excel lays text out again in whole screen pixels (§4.18, the zoom).
            assert float(row["ink_left"]) == pytest.approx(expected_left, abs=SPILL_TOLERANCE)
            assert float(row["ink_right"]) == pytest.approx(expected_right, abs=SPILL_TOLERANCE)
        compared += 1
    assert compared


NO_MARKS = ("row break", "column break", "print area left", "print area top", "print area right",
            "print area bottom")  # fmt: skip


@pytest.mark.excel
def test_the_normal_view_draws_neither_breaks_nor_the_print_area():
    # At every break and print-area edge only the gridline: gray 212, one screen point.
    rows = _captured()
    gray = "".join(f"{c:02X}" for c in GRIDLINE["mac"])
    compared = 0
    for row in _edges(rows, "normal", NO_MARKS):
        assert row.get("line_px"), row
        assert float(row["line_spt"]) == pytest.approx(1.0, abs=WIDTH_TOLERANCE), row
        assert _near(row["line_color"], gray, COLOR_TOLERANCE), row
        compared += 1
    views = {(r["zoom"], r["appearance"]) for r in _edges(rows, "normal", NO_MARKS)}
    assert compared == len(NO_MARKS) * len(views)


@pytest.mark.excel
def test_the_freeze_line_lies_on_the_gridline():
    # sheetpeek draws the freeze line on the boundary's gridline; the gridline's place at a
    # plain boundary of the same capture is the reference, within a device pixel.
    rows = _captured()
    plain = {
        (r["zoom"], r["appearance"], r["side"]): int(r["line_offset"])
        for r in _edges(rows, "normal", ("row break", "column break"))
        if r.get("line_offset")
    }
    compared = 0
    for (zoom, appearance, side), row in _open(rows).items():
        reference = plain.get((zoom, appearance, side))
        if reference is None or not row.get("line_offset"):
            continue
        assert abs(int(row["line_offset"]) - reference) <= 1, row
        compared += 1
    assert compared


@pytest.mark.excel
def test_page_break_preview_draws_a_manual_break_as_sheetpeek_marks_it():
    # 3 screen points of #0000D0, centered on the gridline.
    rows = _captured()
    states = [r["state"] for r in rows if r["group"] == "page_break" and r["kind"] == "window"]
    assert states, "page_break was not captured"
    found = _edges(rows, "page_break", ("row break", "column break"))
    assert len(found) == 2
    for row in found:
        assert row.get("line_px"), row
        assert float(row["line_spt"]) == pytest.approx(
            freeze_view.EXPECTED["break_width"], abs=WIDTH_TOLERANCE
        ), row
        assert _near(row["line_color"], freeze_view.EXPECTED["break_color"], COLOR_TOLERANCE), row
        width = int(row["line_px"])
        assert -width <= int(row["line_offset"]) * 2 + width <= width, row  # about centered


@pytest.mark.excel
def test_the_frozen_window_is_frozen_at_c3():
    rows = _captured()
    states = [r["state"] for r in rows if r["kind"] == "window" and r["group"] in ("normal",)]
    assert states
    for state in states:
        fields = dict(part.partition("=")[::2] for part in state.split(";"))
        assert fields.get("freeze_panes") == "true", state
        assert (fields.get("split_row"), fields.get("split_column")) == ("2", "2"), state


@pytest.mark.excel
def test_a_zoom_excel_does_not_allow_opens_at_sheetpeeks_zoom():
    # 85.5, "abc" and an empty zoomScale open at 100; 5 and 401 at their value, and Excel draws
    # them: at 5 its window shows more rows than any zoom of 10 or more could (72 at 100), at
    # 401 no more than 400 shows. A file Excel refuses is left out (the refusal is the answer).
    rows = _captured()
    refused = {r["group"] for r in rows if r["kind"] == "refused"}
    assert refused <= set(freeze_view.REFUSALS)
    states = {r["item"]: r["state"] for r in rows if r["kind"] == "zoom"}
    zooms = {
        name: float(r["zoom"]) for name, r in ((r["item"], r) for r in rows if r["kind"] == "zoom")
    }
    assert zooms.get("85") == 85, zooms
    for name, found in zooms.items():
        assert found == freeze_view.EXPECTED_ZOOM[name], (name, found)
    at_100 = freeze_view.visible_rows(states["abc"])
    assert at_100 == 72, states["abc"]
    assert freeze_view.visible_rows(states["5"]) > at_100 * 100 / 10, states["5"]
    assert freeze_view.visible_rows(states["401"]) <= at_100 * 100 / 400, states["401"]
