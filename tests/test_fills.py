"""Fills: pattern and gradient fills read, their one color, the census, and the HTML fill layer.

The expected tiles, shares and colors come from the screen suite ``fills``; the gradient angles
are unverified (ARCHITECTURE.md §4.12).
"""

import math
import re

import openpyxl
import pytest
from openpyxl.styles import Color, GradientFill, PatternFill

import sheetpeek
from sheetpeek import Gradient, Pattern
from sheetpeek._fill_border_model import TILES, blend, share
from sheetpeek.backends.html._patterns import gradient as gradient_css
from tests.workbooks import rewrite, save

STYLES = "xl/styles.xml"

# The captured foreground shares, in percent (suite fills).
SHARES = {
    "mediumGray": 50,
    "darkGray": 75,
    "lightGray": 25,
    "darkHorizontal": 50,
    "lightHorizontal": 25,
    "darkVertical": 50,
    "lightVertical": 25,
    "darkDown": 50,
    "darkUp": 50,
    "lightDown": 25,
    "lightUp": 25,
    "darkGrid": 50,
    "darkTrellis": 75,
    "lightGrid": 43.75,
    "lightTrellis": 37.5,
    "gray125": 12.5,
    "gray0625": 6.25,
}


def filled(tmp_path, *fills, name="fills.xlsx"):
    """A sheet with ``fills[k]`` on row ``k + 1`` of column B, and text in A."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for row, fill in enumerate(fills, 1):
        ws.cell(row, 1, "x")
        ws.cell(row, 2).fill = fill
    return save(wb, tmp_path / name)


def spec(tmp_path, fill):
    sheet = sheetpeek.read(filled(tmp_path, fill))
    style = sheet.cell(1, 2).style
    return style.fill_spec, style.fill


# --- reading


def test_every_pattern_has_its_captured_tile_and_share():
    assert set(TILES) == set(SHARES)
    for kind, percent in SHARES.items():
        assert share(kind) * 100 == percent, kind
        assert len({len(row) for row in TILES[kind]}) == 1, kind


@pytest.mark.parametrize("kind", sorted(SHARES))
def test_a_pattern_is_read_with_its_colors(tmp_path, kind):
    found, color = spec(tmp_path, PatternFill(kind, fgColor="FF0000FF", bgColor="FFFFFF00"))
    assert found == Pattern(kind, "0000FF", "FFFF00")
    assert color == blend("0000FF", "FFFF00", SHARES[kind] / 100)


def test_gray0625_is_shown_as_its_blend_not_solid_black(tmp_path):
    # The most common pattern: 6.25 % black dots on white, once shown as solid black.
    found, color = spec(tmp_path, PatternFill("gray0625", fgColor="FF000000"))
    assert found == Pattern("gray0625", "000000", None)
    assert color == "EFEFEF"


def test_solid_draws_its_foreground(tmp_path):
    assert spec(tmp_path, PatternFill("solid", fgColor="FFFF0000", bgColor="FF0000FF")) == (
        Pattern("solid", "FF0000"),
        "FF0000",
    )


def test_solid_with_only_a_background_draws_black(tmp_path):
    # A missing fgColor is black (captured, auto_colors).
    assert spec(tmp_path, PatternFill("solid", bgColor="FFFF0000"))[1] == "000000"


def test_the_alpha_byte_is_ignored(tmp_path):
    assert spec(tmp_path, PatternFill("solid", fgColor="00FF0000"))[1] == "FF0000"


@pytest.mark.parametrize(
    ("fill", "expected"),
    [
        # Indexed 64, the system foreground, draws a solid fill black, and an automatic
        # foreground white (captured, auto_colors).
        (PatternFill("solid", fgColor=Color(indexed=64)), "000000"),
        (PatternFill("solid", fgColor=Color(auto=True)), "FFFFFF"),
        (PatternFill("solid", fgColor=Color(indexed=65)), "FFFFFF"),
        # An index past the palette draws black (captured: 81).
        (PatternFill("solid", fgColor=Color(indexed=81)), "000000"),
        (PatternFill("solid", fgColor=Color(indexed=10)), "FF0000"),
        (PatternFill("none", fgColor="FFFF0000"), None),
        (PatternFill(None, fgColor="FFFF0000"), None),
    ],
)
def test_automatic_and_system_colors_of_a_solid_fill(tmp_path, fill, expected):
    assert spec(tmp_path, fill)[1] == expected


@pytest.mark.parametrize(
    "fill",
    [
        PatternFill("lightUp"),
        PatternFill("lightUp", fgColor=Color(auto=True), bgColor=Color(auto=True)),
        PatternFill("lightUp", fgColor=Color(indexed=64), bgColor=Color(indexed=65)),
    ],
)
def test_a_pattern_without_colors_is_black_on_white(tmp_path, fill):
    found, color = spec(tmp_path, fill)
    assert (found.foreground, found.background) == ("000000", "FFFFFF")
    assert color == blend("000000", "FFFFFF", 0.25)


def test_one_pattern_color_given_the_other_takes_its_default(tmp_path):
    assert spec(tmp_path, PatternFill("lightGray", fgColor="FFFF0000"))[0].background == "FFFFFF"
    assert spec(tmp_path, PatternFill("lightGray", bgColor="FFFFFF00"))[0].foreground == "000000"


def test_a_linear_gradient_is_read(tmp_path):
    found, color = spec(tmp_path, GradientFill(degree=90, stop=("FFFFFFFF", "FF0000FF")))
    assert found == Gradient("linear", ((0.0, "FFFFFF"), (1.0, "0000FF")), degree=90.0)
    assert color == "8080FF"  # the midpoint


def test_a_path_gradient_is_read(tmp_path):
    fill = GradientFill(
        "path", left=0.2, right=0.8, top=0.2, bottom=0.8, stop=("FFFF0000", "FF0000FF")
    )
    found = spec(tmp_path, fill)[0]
    assert (found.kind, found.left, found.right, found.top, found.bottom) == (
        "path",
        0.2,
        0.8,
        0.2,
        0.8,
    )


def test_stops_inside_0_and_1_run_from_and_to_black():
    # Captured (gradients; stops white at 0.25, blue at 0.75): black at 0 and at 1.
    found = Gradient("linear", ((0.25, "FFFFFF"), (0.75, "0000FF")))
    assert found.color_at(0) == "000000"
    assert found.color_at(0.25) == "FFFFFF"
    assert found.color_at(0.75) == "0000FF"
    assert found.color_at(1) == "000000"


def test_stops_at_one_position_end_the_stretch_before_in_the_last_and_start_after_in_the_first():
    # Captured (gradients): white 0, red 0.5, blue 0.5 runs white to blue, then red to black.
    found = Gradient("linear", ((0.0, "FFFFFF"), (0.5, "FF0000"), (0.5, "0000FF")))
    assert found.color_at(0.4999) == blend("0000FF", "FFFFFF", 0.4999 / 0.5)
    assert found.color_at(0.5001) == blend("000000", "FF0000", 0.0001 / 0.5)
    assert found.drawn_stops == ((0.0, "FFFFFF"), (0.5, "0000FF"), (0.5, "FF0000"), (1.0, "000000"))


def test_stops_are_taken_in_order_of_position(tmp_path):
    path = filled(tmp_path, GradientFill(stop=("FFFFFFFF", "FF0000FF")))
    rewrite(
        path,
        STYLES,
        lambda t: (
            t.replace('position="0"', 'position="T"')
            .replace('position="1"', 'position="0"')
            .replace('position="T"', 'position="1"')
        ),
    )
    assert sheetpeek.read(path).cell(1, 2).style.fill_spec.stops == (
        (0.0, "0000FF"),
        (1.0, "FFFFFF"),
    )


def test_a_gradient_without_stops_is_black(tmp_path):
    # Captured (gradients): Excel draws it solid black, linear or path.
    assert spec(tmp_path, GradientFill(stop=())) == (Pattern("solid", "000000"), "000000")


def test_duplicate_stop_positions_are_read_in_order(tmp_path):
    # openpyxl refused the whole stylesheet; xlio reads it. Kept in the file's order, which
    # decides what Excel draws (Gradient.drawn_stops, captured).
    path = filled(tmp_path, GradientFill(stop=("FFFFFFFF", "FF0000FF")))
    rewrite(path, STYLES, lambda t: t.replace('<stop position="1"', '<stop position="0"', 1))
    stops = sheetpeek.read(path).cell(1, 2).style.fill_spec.stops
    # In the file's order, white then blue, not sorted by color.
    assert stops == ((0.0, "FFFFFF"), (0.0, "0000FF"))


@pytest.mark.parametrize("value", ["nan", "inf", "-inf"])
def test_an_odd_degree_is_read_as_0(tmp_path, value):
    path = filled(tmp_path, GradientFill(degree=45, stop=("FFFFFFFF", "FF0000FF")))
    rewrite(path, STYLES, lambda t: t.replace('degree="45"', f'degree="{value}"'))
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 2).style.fill_spec.degree == 0
    page = sheetpeek.render(sheet, "html")
    assert "nan" not in page.lower().replace("nanum", "")
    assert "inf" not in re.sub(r"[a-z]inf|inf[a-z]", "", page.lower())


def test_a_path_rectangle_outside_the_cell_is_clamped(tmp_path):
    fill = GradientFill(
        "path", left=-3, right=7, top=0.5, bottom=0.5, stop=("FFFFFFFF", "FF0000FF")
    )
    found = spec(tmp_path, fill)[0]
    assert (found.left, found.right) == (0, 1)


# --- the census and the outputs


def test_the_census_tells_solid_pattern_and_gradient_fills_apart(tmp_path):
    path = filled(
        tmp_path,
        PatternFill("solid", fgColor="FFFF0000"),
        PatternFill("gray125", fgColor="FF000000"),
        GradientFill(stop=("FFFFFFFF", "FF0000FF")),
    )
    features = sheetpeek.read(path).layout("excel").features
    assert {code: [str(r) for r in p.cells] for code, p in features.items()} == {
        "fill": ["B1"],
        "fill.pattern": ["B2"],
        "fill.gradient": ["B3"],
    }


def test_rich_shows_a_pattern_as_its_blend_and_reports_it(tmp_path):
    path = filled(tmp_path, PatternFill("gray0625", fgColor="FF000000"))
    rendered = sheetpeek.read(path).layout("characters")
    assert rendered.cell(1, 2).fill == "EFEFEF"
    found = {n.code: n for n in sheetpeek.notices(rendered, "rich")}
    assert found["fill.pattern"].outcome == "approximated"
    assert "fill.pattern" not in {n.code for n in sheetpeek.notices(path, "html")}
    for name in ("text", "markdown", "csv", "tsv"):
        assert {n.code: n.outcome for n in sheetpeek.notices(path, name)}["fill.pattern"] == (
            "omitted"
        ), name


def boxes(page):
    return re.findall(r'<i class="(sp-f-[a-z]+)" style="([^"]*)"></i>', page)


def rule(page, name):
    return re.search(rf"\.sheetpeek \.{name} \{{ ([^}}]*) \}}", page)[1]


def test_a_pattern_is_drawn_from_its_tile_anchored_at_the_sheet(tmp_path):
    path = filled(tmp_path, PatternFill("gray0625", fgColor="FF0000FF", bgColor="FFFFFF00"))
    page = sheetpeek.render(path, "html")
    ((name, style),) = boxes(page)
    # B1: x 71 px (53 pt), 71 px wide to C's gridline plus the far gridline; 20 px plus one.
    assert style == "left:71px;top:0px;width:71px;height:21px;background-position:1px 0px"
    body = rule(page, name)
    assert "background-size:8px 4px" in body  # 8 x 4 screen points, one to a CSS px
    assert "fill='%230000FF'" in body
    assert "fill='%23FFFF00'" in body
    assert "M2 0h1v1h-1zM6 2h1v1h-1z" in body  # the two dots of the captured tile (16.115)
    assert "image-rendering: pixelated" in page
    # The cell itself draws no background: the layer does.
    assert "<td style=" not in page.split("<tbody>")[1].split("</tr>")[0]


def test_a_fill_covers_its_far_gridlines_unless_the_neighbor_is_filled(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"].fill = PatternFill("gray125", fgColor="FF000000")
    ws["A2"].fill = PatternFill("gray125", fgColor="FF000000")
    ws["B2"].fill = PatternFill("solid", fgColor="FFFF0000")
    ws["A3"].fill = PatternFill("solid", fgColor="FF00FF00")
    page = sheetpeek.render(save(wb, tmp_path / "far.xlsx"), "html")
    first, second = (style for _, style in boxes(page))
    # A1: the cell below is filled, so it paints the gridline between them.
    assert first == "left:0px;top:0px;width:72px;height:20px"
    # A2: B2 is filled to the right, A3 below.
    assert second == "left:0px;top:20px;width:71px;height:20px"


def test_a_merge_is_one_box(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B2"].fill = PatternFill("darkGrid", fgColor="FF000000")
    ws["C3"].fill = PatternFill("solid", fgColor="FFFF0000")  # covered: ignored
    ws.merge_cells("B2:C3")
    page = sheetpeek.render(save(wb, tmp_path / "merged.xlsx"), "html")
    ((_, style),) = boxes(page)
    # The range starts at B2; B and C are 71 and 70 px, plus the far gridlines.
    assert style.startswith("left:0px;top:0px;width:142px;height:41px")


@pytest.mark.parametrize("degree", [0, 30, 45, 90, 135, 180, 270, 315])
def test_a_linear_gradient_takes_its_angle_in_the_unit_cell(degree):
    # Captured on a 240 x 90 px cell: 30 degrees measures 56.7 in pixels.
    found = gradient_css(Gradient("linear", ((0, "FFFFFF"), (1, "0000FF")), degree), 120, 45)
    angle = float(re.search(r"linear-gradient\(([-0-9.]+)deg", found)[1])
    d = math.radians(degree)
    assert angle == pytest.approx(
        (90 + math.degrees(math.atan2(120 * math.sin(d), 45 * math.cos(d)))) % 360, abs=1e-3
    )
    assert "#FFFFFF 0%,#0000FF 100%" in found


def test_a_linear_gradient_inside_0_and_1_ramps_from_and_to_black():
    found = gradient_css(Gradient("linear", ((0.25, "FFFFFF"), (0.75, "0000FF"))), 120, 45)
    assert "#000000 0%,#FFFFFF 25%,#0000FF 75%,#000000 100%" in found


def test_a_path_gradient_is_four_bands(tmp_path):
    fill = GradientFill(
        "path", left=0.2, right=0.8, top=0.2, bottom=0.8, stop=("FFFF0000", "FF0000FF")
    )
    page = sheetpeek.render(filled(tmp_path, fill), "html")
    ((name, _),) = boxes(page)
    body = rule(page, name)
    assert body.count("%3Cpolygon") == 4
    assert body.count("%3ClinearGradient") == 4
    # The first stop's rectangle, 0.2 to 0.8 of the 71 x 21 px box, in the first color.
    assert "%3Crect x='14.2' y='4.2' width='42.6' height='12.6' fill='%23FF0000'" in body


def test_a_path_rectangle_given_backwards_is_taken_in_order():
    forward = Gradient(
        "path", ((0, "FFFFFF"), (1, "0000FF")), left=0.2, right=0.8, top=0.1, bottom=0.9
    )
    backward = Gradient(
        "path", ((0, "FFFFFF"), (1, "0000FF")), left=0.8, right=0.2, top=0.9, bottom=0.1
    )
    assert gradient_css(backward, 100, 50) == gradient_css(forward, 100, 50)
