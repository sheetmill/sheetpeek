"""The screen-capture suites' workbooks and measuring, without Excel.

Each suite's workbook is built and read back as built. Its measuring runs on a synthetic capture:
a white image with cells laid out from the worksheet's own column widths and row heights at 2
pixels a point (zoom 100 on a 2x display), with a few known marks drawn in it.
"""

import openpyxl
from capture import borders, fills, text_style
from capture.screen import png
from capture.screen.run import Capture
from capture.screen.window import Grid
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek

SCALE = 2


def _geometry(ws, cells):
    """Each cell's left, top, width and height in points, from the worksheet's sizes."""

    def width(column):
        found = ws.column_dimensions[get_column_letter(column)].width
        return round(found * 7 + 5) * 0.75 if found else 48.0

    def height(row):
        return ws.row_dimensions[row].height or 15.0

    geometry = {}
    for cell in cells:
        letters, row = coordinate_from_string(cell)
        column = column_index_from_string(letters)
        geometry[cell] = {
            "left": sum(width(c) for c in range(1, column)),
            "top": sum(height(r) for r in range(1, row)),
            "width": width(column),
            "height": height(row),
        }
    return geometry


class Canvas:
    """A white image with cells at known places, and a ``Capture`` of it."""

    def __init__(self, ws, cells, rows=()):
        self.geometry = _geometry(ws, cells)
        right = max(g["left"] + g["width"] for g in self.geometry.values())
        bottom = max(g["top"] + g["height"] for g in self.geometry.values())
        self.width, self.height = round(right * SCALE) + 40, round(bottom * SCALE) + 40
        self.rows = [bytearray(b"\xff" * (self.width * 3)) for _ in range(self.height)]
        self.reads = list(rows)

    def paint(self, x0, y0, x1, y1, rgb):
        for y in range(y0, y1):
            for x in range(x0, x1):
                self.rows[y][3 * x : 3 * x + 3] = bytes(rgb)

    def capture(self):
        image = png.Image(self.width, self.height, 3, self.rows, "sRGB")
        grid = Grid(0.0, 0.0, float(SCALE))
        return Capture(100, image, grid, {}, self.geometry, self.reads)

    def box(self, cell):
        g = self.geometry[cell]
        return (
            round(g["left"] * SCALE),
            round(g["top"] * SCALE),
            round((g["left"] + g["width"]) * SCALE),
            round((g["top"] + g["height"]) * SCALE),
        )


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


def test_every_case_names_a_distinct_cell_in_its_group():
    for group in text_style.GROUPS:
        found = text_style.cases(group)
        assert len({case.cell for case in found}) == len(found), group


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
