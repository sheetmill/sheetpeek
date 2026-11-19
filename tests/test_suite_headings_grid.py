"""The screen suite ``headings_grid``: its workbooks, its measuring, and the comparison.

ARCHITECTURE.md §4.14 (the grid color and the headings). The workbooks are built and read back
with sheetpeek; the measuring runs on synthetic captures with the headings painted at negative
sheet points; the comparison with ``tests/captures/headings_grid.tsv`` is skipped until that
capture exists.
"""

import re
from pathlib import Path

import openpyxl
import pytest
from capture import headings_grid
from capture.screen import png
from capture.screen.run import Capture
from capture.screen.window import GRIDLINE, Grid

import sheetpeek
from sheetpeek.backends.html._css import CSS
from sheetpeek.backends.html._headings import bands

CAPTURED = Path(__file__).parent / "captures" / "headings_grid.tsv"
DARK_PAPER, SEPARATOR, EDGE_LINE = (22, 22, 22), (86, 86, 86), (188, 187, 186)


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("headings_grid")
    headings_grid.build(folder)
    return {
        group.name: sheetpeek.read(folder / f"headings_grid_{group.name}.xlsx")
        for group in headings_grid.GROUPS
    }


# --- the workbooks


def test_every_group_reads_back_as_built(built):
    assert set(built) == {g.name for g in headings_grid.GROUPS}
    for group in headings_grid.GROUPS:
        sheet = built[group.name]
        font = sheet.normal_font
        assert (font.name, font.size) == group.font, group.name
        if group.grid:
            for cell, color in headings_grid.FILLS.items():
                letters, row = openpyxl.utils.cell.coordinate_from_string(cell)
                column = openpyxl.utils.column_index_from_string(letters)
                spec = sheet.cell(row, column).style.fill_spec
                assert (spec.kind, spec.fg) == ("solid", color), (group.name, cell)


def test_the_expected_grid_colors_are_sheetpeeks(built):
    for name, expected in headings_grid.EXPECTED_GRID.items():
        assert built[name].grid_color == expected, name
        assert built[name].show_gridlines


def test_the_crafted_color_id_is_in_the_file_as_written(tmp_path):
    import zipfile

    headings_grid.build(tmp_path)
    with zipfile.ZipFile(tmp_path / "headings_grid_grid_x.xlsx") as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert '<sheetView defaultGridColor="0" colorId="x" ' in sheet
    assert headings_grid.REFUSALS == ("grid_65", "grid_99999", "grid_x")  # 16.115 refused 65, 99999
    assert headings_grid.GROUPS[-1].name == "grid_x"  # runs on its own, after the others
    with zipfile.ZipFile(tmp_path / "headings_grid_font_calibri_8.xlsx") as archive:
        assert "<sheetFormatPr" not in archive.read("xl/worksheets/sheet1.xml").decode()


def test_the_expected_bands_are_sheetpeeks():
    for (_, size, digits), expected in headings_grid.EXPECTED_BANDS.items():
        assert tuple(round(v, 2) for v in bands(size, digits)) == expected
    for group in headings_grid.GROUPS:
        digits = headings_grid.EXPECTED_DIGITS.get(group.name, 2)
        assert (*group.font, digits) in headings_grid.EXPECTED_BANDS, group.name


def _css_colors():
    dark = CSS.index('.sheetpeek[data-color-scheme="dark"]')

    def take(text):
        return {
            key: re.search(rf"--sheetpeek-heading-{name}: #(\w{{6}})", text)[1].upper()
            for key, name in (("paper", "paper"), ("text", "ink"), ("separator", "grid"))
        }

    return {"light": take(CSS[:dark]), "dark": take(CSS[dark:])}


def test_the_expected_heading_colors_are_sheetpeeks():
    assert _css_colors() == headings_grid.EXPECTED_HEADING_COLORS


def test_no_case_draws_the_marker_color():
    colors = {*headings_grid.FILLS.values(), headings_grid.TEAL, "0000FF"}
    for color in colors:
        r, g, b = (int(color[i : i + 2], 16) for i in (0, 2, 4))
        assert not (r > 150 and b > 150 and g < 120), color


def test_the_markers_keep_off_the_first_row_and_inside_a_small_window():
    # The first marker below the top row (a marker touching the headings starts a pixel late),
    # the second in another row and column, within the 35 rows a window shows at zoom 200.
    for group in headings_grid.GROUPS:
        first, second = (openpyxl.utils.cell.coordinate_from_string(m) for m in group.markers)
        assert first[1] == group.top_row + 1
        assert first[0] != second[0]
        assert first[1] < second[1] < group.top_row + 30


def test_the_last_row_of_a_visible_range():
    assert headings_grid.last_row("$A$9990:$AQ$10059") == 10059
    assert headings_grid.last_row("$A$1:$AQ$70") == 70
    assert headings_grid.last_row("") is None


# --- the measuring, on a synthetic capture

ORIGIN_X, ORIGIN_Y, CHROME_BOTTOM = 44, 62, 20
"""The cells' first boundary in image pixels at 2 a point: a 22 pt row band and a 21 pt
column band below 20 pixels of other chrome."""


class Window:
    """A white window with Excel's dark headings left of and above the cells."""

    def __init__(self, group, columns=12, rows=12, visible="$A$1:$AQ$70"):
        self.group = headings_grid.GROUP[group]
        self.width = ORIGIN_X + columns * 96 + 20
        self.height = ORIGIN_Y + rows * 30 + 20
        self.rows = [bytearray(b"\xff" * (self.width * 3)) for _ in range(self.height)]
        self.paint(0, 0, self.width, CHROME_BOTTOM, (38, 38, 38))
        self.paint(0, CHROME_BOTTOM, self.width, ORIGIN_Y, DARK_PAPER)
        self.paint(0, 0, ORIGIN_X, self.height, DARK_PAPER)
        self.paint(0, 0, ORIGIN_X, CHROME_BOTTOM, (38, 38, 38))
        for k in range(rows):  # row separators at each row's bottom, in the row band
            y = ORIGIN_Y + 30 * k + 28
            self.paint(0, y, ORIGIN_X, y + 2, SEPARATOR)
        for k in range(columns):
            x = ORIGIN_X + 96 * k + 94
            self.paint(x, CHROME_BOTTOM, x + 2, ORIGIN_Y, SEPARATOR)
        self.paint(ORIGIN_X, ORIGIN_Y, ORIGIN_X + 2, self.height, EDGE_LINE)
        self.paint(0, ORIGIN_Y, self.width, ORIGIN_Y + 2, EDGE_LINE)
        self.paint(18, ORIGIN_Y + 40, 24, ORIGIN_Y + 50, (255, 255, 255))  # "2"
        self.paint(ORIGIN_X + 140, 34, ORIGIN_X + 146, 48, (255, 255, 255))  # "B"
        self.reads = [("window", "visible.100", visible)]

    def paint(self, x0, y0, x1, y1, rgb):
        for y in range(y0, y1):
            for x in range(x0, x1):
                self.rows[y][3 * x : 3 * x + 3] = bytes(rgb)

    def x(self, points):
        return ORIGIN_X + round(points * 2)

    def y(self, points):
        return ORIGIN_Y + round(points * 2)

    def capture(self):
        image = png.Image(self.width, self.height, 3, self.rows, "sRGB")
        first, second = self.group.markers
        geometry = {}
        for cell in (first, second):
            letters, row = openpyxl.utils.cell.coordinate_from_string(cell)
            column = openpyxl.utils.column_index_from_string(letters)
            geometry[cell] = {
                "left": 48.0 * (column - 1),
                "top": 15.0 * (row - 1),
                "width": 48.0,
                "height": 15.0,
            }
        top = 15.0 * (self.group.top_row - 1)
        grid = Grid(float(ORIGIN_X), ORIGIN_Y - top * 2, 2.0)
        return Capture(100, image, grid, {}, geometry, self.reads)


def test_the_bands_and_heading_colors_read_from_the_window():
    (record, *_) = headings_grid.measure(Window("grid_default").capture(), "grid_default")
    assert record["kind"] == "headings"
    assert (record["row_band"], record["col_band"]) == ("22", "21")
    assert (record["row_band_px"], record["col_band_px"], record["band_left_px"]) == (
        "44",
        "42",
        "0",
    )
    assert record["paper"] == "161616"
    assert (record["separator_row"], record["separator_col"]) == ("565656", "565656")
    assert (record["text_row"], record["text_col"]) == ("FFFFFF", "FFFFFF")
    assert record["edge_row"] == "161616x6 BCBBBAx2 FFFFFFx4"
    assert record["edge_col"] == "161616x6 BCBBBAx2 FFFFFFx4"
    assert record["visible"] == "$A$1:$AQ$70"
    assert set(record) <= set(headings_grid.COLUMNS)


def gray(v):
    return (v, v, v)


def test_the_light_bands_are_measured_across_their_gradient():
    # The light appearance's paper steps by 1 from 253 to 247 toward the cells, below white
    # chrome; its separators fade in from the paper to 221 next to the cells.
    window = Window("grid_default")
    for y in range(CHROME_BOTTOM):
        window.paint(0, y, window.width, y + 1, gray(255))
    for y in range(CHROME_BOTTOM, ORIGIN_Y):
        window.paint(ORIGIN_X + 2, y, window.width, y + 1, gray(253 - (y - CHROME_BOTTOM) // 7))
    for x in range(ORIGIN_X):
        window.paint(x, CHROME_BOTTOM, x + 1, window.height, gray(253 - x // 7))
    for k in range(12):
        y = ORIGIN_Y + 30 * k + 28
        for x in range(ORIGIN_X):
            window.paint(x, y, x + 1, y + 2, gray(252 - 31 * x // (ORIGIN_X - 1)))
        x = ORIGIN_X + 96 * k + 94
        for y in range(CHROME_BOTTOM, ORIGIN_Y):
            shade = 252 - 31 * (y - CHROME_BOTTOM) // (ORIGIN_Y - CHROME_BOTTOM - 1)
            window.paint(x, y, x + 2, y + 1, gray(shade))
    window.paint(18, ORIGIN_Y + 40, 24, ORIGIN_Y + 50, gray(94))  # "2"
    window.paint(ORIGIN_X + 140, 34, ORIGIN_X + 146, 48, gray(94))  # "B"
    (record, *_) = headings_grid.measure(window.capture(), "grid_default", "light")
    assert (record["row_band"], record["col_band"], record["band_left_px"]) == ("22", "21", "0")
    assert record["paper"] == "F7F7F7"
    assert (record["separator_row"], record["separator_col"]) == ("DEDEDE", "DEDEDE")
    assert (record["text_row"], record["text_col"]) == ("5E5E5E", "5E5E5E")


def test_the_column_band_ends_at_row_1_when_the_rows_differ_in_height():
    # A font group's empty rows keep the file's default height, its marker row the font's:
    # Excel's geometry puts row 1's top at 0 pt, whatever the marker row's height.
    window = Window("font_calibri_20")
    capture = window.capture()
    first = window.group.markers[0]
    capture.geometry[first] = capture.geometry[first] | {"top": 15.0, "height": 26.0}
    (record,) = headings_grid.measure(capture, "font_calibri_20")
    assert (record["row_band"], record["col_band"]) == ("22", "21")


def test_a_scrolled_window_measures_from_its_top_row():
    window = Window("rows_4_5", visible="$A$9990:$AQ$10059")
    (record,) = headings_grid.measure(window.capture(), "rows_4_5")
    assert (record["top_row"], record["row_band"], record["col_band"]) == ("9990", "22", "21")
    assert headings_grid.last_row(record["visible"]) == 10059


def test_a_colored_gridline_reads_as_one_device_pixel():
    window = Window("grid_10_teal")
    # C8's right boundary at 144 pt; Excel draws a colored line on the second of the gray
    # line's two pixels (ARCHITECTURE §4.14).
    x = window.x(144) + 1
    window.paint(x, window.y(105), x + 1, window.y(120), (0, 128, 128))
    # The yellow E4 covers its boundaries.
    window.paint(window.x(192), window.y(45), window.x(240), window.y(60), (255, 255, 0))
    records = headings_grid.measure(window.capture(), "grid_10_teal")
    edges = {(r["cell"], r["side"]): r["strip"] for r in records if r["kind"] == "edge"}
    assert len(edges) == len(headings_grid.EDGES)
    assert edges["C8", "right"] == "FFFFFFx7 008080x1 FFFFFFx4"
    assert edges["C8", "bottom"] == "FFFFFFx12"
    assert edges["E4", "left"] == "FFFFFFx6 FFFF00x6"
    assert edges["E4", "bottom"] == "FFFF00x6 FFFFFFx6"
    assert headings_grid.strip(edges["C8", "right"]) == ["FFFFFF"] * 7 + ["008080"] + ["FFFFFF"] * 4


# --- the comparison with Excel's window


def _captured():
    if not CAPTURED.exists():
        pytest.skip("not captured yet: tests/captures/headings_grid.tsv")
    lines = [line.split("\t") for line in CAPTURED.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return [dict(zip(header, row, strict=False)) for row in body]


def _rgb(text):
    return tuple(int(text[i : i + 2], 16) for i in (0, 2, 4))


def _near(a, b, tolerance):
    return max(abs(x - y) for x, y in zip(_rgb(a), _rgb(b), strict=True)) <= tolerance


COLOR_TOLERANCE = 10
"""Per channel, for a line's color against the file's: the captures are converted to sRGB, and
flat colors come back within a few units (the gray gridline 212, the headings' 22 and 86, the
line between headings and cells 188,187,186, exact in the earlier suites' windows)."""
HEADING_TOLERANCE = 4
"""Per channel, for heading colors against sheetpeek's and between groups: flat colors."""


def _line(strip_text):
    """The pixels of a strip on white that are not white: the gridline's."""
    return [c for c in headings_grid.strip(strip_text) if not _near(c, "FFFFFF", 12)]


def _captured_groups(rows):
    refused = {r["group"] for r in rows if r["kind"] == "refused"}
    return refused, [r for r in rows if r["kind"] != "refused"]


def _share(pixel, color):
    """How much of ``color`` over white ``pixel`` holds, by its channels farthest from white,
    or ``None`` where the channels disagree by more than the conversion's rounding."""
    shares = [
        (255 - p) / (255 - c)
        for p, c in zip(_rgb(pixel), _rgb(color), strict=True)
        if 255 - c >= 64
    ]
    if max(shares) - min(shares) > SHARE_TOLERANCE:
        return None
    return sum(shares) / len(shares)


SHARE_TOLERANCE = 0.05
"""Of a pixel's share of the color, per channel and of a line's total: the captured lines' two
pixels are 0.40 and 0.60 of teal (008080) and of blue (0000FF) to within 0.01."""


@pytest.mark.excel
def test_the_grid_color_is_drawn_in_the_palette_entry_one_device_pixel_wide(built):
    # colorId draws the palette's entry, one device pixel's worth: Excel anti-aliases it over two
    # pixels, 0.4 and 0.6 of the color (the heavier second at zoom 100, first at 200), where
    # sheetpeek draws one pixel in the full color. 64, past the palette and malformed are the
    # automatic gray, 2 pixels wide.
    refused, rows = _captured_groups(_captured())
    gray = "".join(f"{c:02X}" for c in GRIDLINE["mac"])
    for row in rows:
        if row["kind"] != "edge" or row["item"] != "open" or row["appearance"] != "dark":
            continue
        expected = built[row["group"]].grid_color
        line = _line(row["strip"])
        if expected is None:
            assert len(line) == 2, row
            assert all(_near(c, gray, COLOR_TOLERANCE) for c in line), row
        else:
            shares = [_share(c, expected) for c in line]
            assert 1 <= len(line) <= 2, row
            assert None not in shares, row
            assert abs(sum(shares) - 1) <= SHARE_TOLERANCE, row
    assert refused <= set(headings_grid.REFUSALS)


@pytest.mark.excel
def test_a_fill_covers_a_colored_gridline_as_it_covers_the_gray_one(built):
    _, rows = _captured_groups(_captured())
    edges = {
        (r["group"], r["zoom"], r["cell"], r["side"]): r["strip"]
        for r in rows
        if r["kind"] == "edge" and r["appearance"] == "dark"
    }
    gray = "".join(f"{c:02X}" for c in GRIDLINE["mac"])
    compared = 0
    for (group, zoom, cell, side), found in edges.items():
        color = built[group].grid_color
        if color is None or cell == "C8":
            continue
        default = edges.get(("grid_default", zoom, cell, side))
        if default is None:
            continue
        colored = any(_near(c, color, COLOR_TOLERANCE) for c in headings_grid.strip(found))
        plain = any(_near(c, gray, COLOR_TOLERANCE) for c in headings_grid.strip(default))
        assert colored == plain, (group, zoom, cell, side, found, default)
        compared += 1
    assert compared


HEADING_KEYS = ("paper", "separator_row", "separator_col", "text_row", "text_col")


@pytest.mark.excel
def test_the_headings_keep_their_colors_under_a_grid_color():
    _, rows = _captured_groups(_captured())
    heads = {(r["group"], r["zoom"], r["appearance"]): r for r in rows if r["kind"] == "headings"}
    compared = 0
    for (group, zoom, appearance), row in heads.items():
        if not group.startswith("grid_") or group == "grid_default":
            continue
        default = heads.get(("grid_default", zoom, appearance))
        if default is None:
            continue
        for key in HEADING_KEYS:
            assert _near(row[key], default[key], HEADING_TOLERANCE), (group, zoom, key)
        compared += 1
    assert compared


@pytest.mark.excel
def test_the_heading_colors_are_sheetpeeks():
    # Dark paper 22, text 255, separators 86; light paper 247, separators 224,223,221 next to
    # the cells, text 94 (the pixel of the glyphs farthest from the paper).
    _, rows = _captured_groups(_captured())
    expected = _css_colors()
    found = set()
    for row in rows:
        if row["kind"] != "headings" or row["group"] != "grid_default" or row["zoom"] != "100":
            continue
        colors = expected[row["appearance"]]
        assert _near(row["paper"], colors["paper"], HEADING_TOLERANCE), row
        for key in ("separator_row", "separator_col"):
            assert _near(row[key], colors["separator"], HEADING_TOLERANCE), (key, row)
        for key in ("text_row", "text_col"):
            assert _near(row[key], colors["text"], HEADING_TOLERANCE), (key, row)
        found.add(row["appearance"])
    assert "dark" in found


def _band_rows(rows, families):
    for row in rows:
        if row["kind"] != "headings" or row["appearance"] != "dark":
            continue
        family, _, size = row["font"].rpartition(" ")
        if family in families:
            yield row, family, float(size)


def _tolerance(zoom):
    """One device pixel in sheet points: the bands' ends fall on device pixels."""
    return 100 / (2 * int(zoom))


SCALE_ZOOMS = ("50", "100", "150")
"""Zooms where Excel's bands are sheetpeek's, in sheet units."""
NARROWER_ZOOMS = {"75": 6.0, "200": 4.5, "400": 4.5}
"""Zooms where Excel's row band is narrower than sheetpeek's, by at most this many sheet
points (7 digits; 0.5 at 2 and 3): Excel sizes the heading digits at the zoom, an advance that
is not a fixed share of the zoom (3, 4, 6, 9, 11 and 22 screen points at 50, 75, 100, 150, 200
and 400), where sheetpeek scales the zoom-100 band. A known deviation."""
BAND_PIXELS = 1
"""Device pixels between a captured band and sheetpeek's: the bands end on device pixels (the
column band at zoom 150 is 64 where 21 pt makes 63)."""
COLUMN_OUTLIER = ("rows_7", 3)
"""The column band of the 7-digit window is 2 device pixels off the other windows' at zoom 150,
200 and 400 (66, 84, 176 against 64, 86, 174), cause not found; its row band is not."""


def _pixels(row, points):
    """Sheet points as the device pixels of a capture at its zoom on a 2x display."""
    return points * int(row["zoom"]) / 100 * 2


@pytest.mark.excel
def test_the_bands_follow_the_last_row_in_the_window_at_calibri_11():
    # The column band 21 pt; the row band max(22, 4 + 6 x digits) pt, the digits of the last row
    # in Excel's visible range, at zoom 50, 100 and 150.
    _, rows = _captured_groups(_captured())
    compared = set()
    for row, _, size in _band_rows(rows, ("Calibri",)):
        if size != 11 or row["zoom"] not in SCALE_ZOOMS:
            continue
        digits = len(str(headings_grid.last_row(row["visible"])))
        expected = bands(size, digits)
        assert abs(int(row["row_band_px"]) - _pixels(row, expected[0])) <= BAND_PIXELS, row
        column = COLUMN_OUTLIER[1] if row["group"] == COLUMN_OUTLIER[0] else BAND_PIXELS
        assert abs(int(row["col_band_px"]) - _pixels(row, expected[1])) <= column, row
        compared.add((row["zoom"], digits))
    assert {d for _, d in compared} == {2, 3, 4, 5, 7}
    assert {z for z, _ in compared} == set(SCALE_ZOOMS)


@pytest.mark.excel
def test_at_other_zooms_the_row_band_is_narrower_within_the_known_deviation():
    _, rows = _captured_groups(_captured())
    compared = set()
    for row, _, size in _band_rows(rows, ("Calibri",)):
        if size != 11 or row["zoom"] not in NARROWER_ZOOMS:
            continue
        digits = len(str(headings_grid.last_row(row["visible"])))
        expected = _pixels(row, bands(size, digits)[0])
        short = (expected - int(row["row_band_px"])) / _pixels(row, 1)
        assert 0 <= short <= NARROWER_ZOOMS[row["zoom"]], (row, short)
        compared.add(row["zoom"])
    assert compared == set(NARROWER_ZOOMS)


@pytest.mark.excel
def test_the_row_band_takes_the_last_rows_digits_not_the_first_rows():
    _, rows = _captured_groups(_captured())
    for row in rows:
        if row["group"] == "rows_4_5" and row["kind"] == "headings" and row["zoom"] in SCALE_ZOOMS:
            assert len(str(headings_grid.last_row(row["visible"]))) == 5, row
            assert abs(float(row["row_band"]) - bands(11.0, 5)[0]) < abs(
                float(row["row_band"]) - bands(11.0, 4)[0]
            ), row


@pytest.mark.excel
def test_the_bands_grow_with_the_normal_font():
    # Calibri 8 to 36 with 2 digits, sheetpeek's table; Aptos Narrow 11 as Calibri 11.
    _, rows = _captured_groups(_captured())
    compared = 0
    for row, _, size in _band_rows(rows, ("Calibri", "Aptos Narrow")):
        if not row["group"].startswith("font_"):
            continue
        digits = len(str(headings_grid.last_row(row["visible"])))
        expected = bands(size, digits)
        tolerance = _tolerance(row["zoom"])
        assert float(row["row_band"]) == pytest.approx(expected[0], abs=tolerance), row
        assert float(row["col_band"]) == pytest.approx(expected[1], abs=tolerance), row
        compared += 1
    assert compared
