"""Borders: sides read, conflicts resolved in the layout, and the HTML border layer.

The line geometry comes from the screen suite ``borders``; the order, ties, merges and spill are
unverified (ARCHITECTURE.md §4.12).
A default column here is 71 or 70 px wide (53 pt) and a row 20 px (15 pt).
"""

import itertools
import re

import openpyxl
import pytest
from openpyxl.styles import Border, Side

import sheetpeek
from sheetpeek import Diagonal
from sheetpeek import Side as Line
from sheetpeek._fill_border_model import STYLES, winner
from sheetpeek._pass_borders import rendered_borders
from tests.workbooks import save

RED, BLUE = "FFFF0000", "FF0000FF"


def bordered(tmp_path, cells, *, merges=(), hidden=(), name="borders.xlsx"):
    """A sheet with ``cells``: coordinate -> (value, Border)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for coordinate, (value, _) in cells.items():
        ws[coordinate] = value
    # Merged first: openpyxl copies an anchor's border to the merge's edges when it merges.
    for merge in merges:
        ws.merge_cells(merge)
    for coordinate, (_, border) in cells.items():
        if border is not None:
            ws[coordinate].border = border
    for letter in hidden:
        ws.column_dimensions[letter].hidden = True
    return save(wb, tmp_path / name)


def resolved(path):
    found = rendered_borders(sheetpeek.read(path).layout("excel"))
    assert found is not None
    return found


# --- reading


def test_sides_are_read_with_their_styles_and_colors(tmp_path):
    path = bordered(
        tmp_path, {"A1": ("x", Border(left=Side("thin", color=RED), bottom=Side("double")))}
    )
    borders = sheetpeek.read(path).cell(1, 1).style.borders
    assert borders.left == Line("thin", "FF0000")
    assert borders.bottom == Line("double", None)
    assert borders.bottom.drawn_color == "000000"  # automatic is black (unverified)
    assert (borders.right, borders.top, borders.diagonal) == (None, None, None)


def test_start_and_end_are_read_as_left_and_right(tmp_path):
    path = bordered(tmp_path, {"A1": ("x", Border(start=Side("thick"), end=Side("thin")))})
    borders = sheetpeek.read(path).cell(1, 1).style.borders
    assert (borders.left, borders.right) == (Line("thick"), Line("thin"))


def test_a_diagonal_needs_a_style_and_a_direction(tmp_path):
    path = bordered(
        tmp_path,
        {
            "A1": ("x", Border(diagonal=Side("thin"))),
            "A2": ("x", Border(diagonalUp=True)),
            "A3": ("x", Border(diagonal=Side("thin"), diagonalDown=True)),
        },
    )
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 1).style.borders is None
    assert sheet.cell(2, 1).style.borders is None
    borders = sheet.cell(3, 1).style.borders
    assert (borders.diagonal, borders.diagonal_up, borders.diagonal_down) == (
        Line("thin"),
        False,
        True,
    )


# --- the conflict rule


def test_the_stronger_style_wins_in_either_order():
    for stronger, weaker in itertools.combinations(STYLES, 2):
        a, b = Line(stronger, "00FF00"), Line(weaker, "000000")
        assert winner(a, b) == a == winner(b, a), (stronger, weaker)


@pytest.mark.parametrize(
    ("darker", "lighter"),
    [  # 11 pairs (unverified)
        ("FF0000", "00FF00"),
        ("0000FF", "00FF00"),
        ("000000", "0000FF"),
        ("000000", "FFFFFF"),
        ("800000", "FF0000"),
        ("808080", "C0C0C0"),
        ("00FFFF", "FFFF00"),
        ("0080FF", "FF8000"),
        ("FF0000", "00FFFF"),
        ("800080", "008000"),
        ("0000FF", "FF0000"),
    ],
)
def test_a_tie_goes_to_the_darker_color(darker, lighter):
    for style in ("thin", "thick", "dashed", "double", "hair"):
        a, b = Line(style, darker), Line(style, lighter)
        assert winner(a, b) == a == winner(b, a)


def test_equal_luma_goes_to_the_larger_rgb():
    # Unverified: no capture has two colors of equal luma.
    a, b = Line("thin", "003223"), Line("thin", "1400F0")
    assert winner(a, b) == b == winner(b, a)


def test_automatic_competes_as_black():
    assert winner(Line("thin", "000080"), Line("thin")) == Line("thin")


# --- the layout


def test_neighbors_sides_compete_on_their_shared_edge(tmp_path):
    path = bordered(
        tmp_path,
        {
            "A1": ("x", Border(right=Side("thin", color=BLUE))),
            "B1": ("x", Border(left=Side("thick", color=RED))),
            "A2": ("x", Border(bottom=Side("thin", color=RED))),
            "A3": ("x", Border(top=Side("thin", color=BLUE))),
        },
    )
    found = resolved(path)
    assert found.vertical[0][1] == Line("thick", "FF0000")
    assert found.horizontal[2][0] == Line("thin", "0000FF")  # blue is darker than red


def test_a_merge_keeps_its_cells_sides_and_drops_the_inner_ones(tmp_path):
    thick = Side("thick", color=RED)
    path = bordered(
        tmp_path,
        {
            "A1": ("", None),
            "B2": ("x", Border(left=thick, right=thick, top=thick, bottom=thick)),
            "C3": ("", Border(right=Side("thin"), left=Side("thin"))),
            "C2": ("", Border(diagonal=Side("thin"), diagonalUp=True)),
        },
        merges=["B2:C3"],
    )
    found = resolved(path)
    red = Line("thick", "FF0000")
    # The anchor's top and left segments only; its right and bottom are inside the merge.
    assert [found.horizontal[1][p] for p in (1, 2)] == [red, None]
    assert [found.vertical[1][q] for q in (1, 2, 3)] == [red, None, None]
    # The covered C3 keeps its outer right side; its inner left side is dropped.
    assert [found.vertical[2][q] for q in (1, 2, 3)] == [None, None, Line("thin")]
    assert found.horizontal[2][1] is None
    assert found.diagonals == ()  # a covered cell's diagonal is ignored


def test_a_neighbor_competes_with_a_merge_s_segment(tmp_path):
    path = bordered(
        tmp_path,
        {
            "A1": ("x", Border(right=Side("thin", color=RED))),
            "B1": ("", None),
            "A2": ("x", Border(right=Side("thick", color=BLUE))),
            "B2": ("", Border(left=Side("thin", color=RED))),
        },
        merges=["B1:C2"],
    )
    found = resolved(path)
    assert [found.vertical[i][1] for i in (0, 1)] == [
        Line("thin", "FF0000"),
        Line("thick", "0000FF"),
    ]


def test_a_merge_s_diagonal_spans_the_whole_merge(tmp_path):
    path = bordered(
        tmp_path,
        {"B2": ("x", Border(diagonal=Side("medium"), diagonalUp=True, diagonalDown=True))},
        merges=["B2:D3"],
    )
    found = resolved(path)
    assert found.diagonals == (Diagonal(0, 0, 1, 2, Line("medium"), True, True),)


def test_spilled_text_removes_the_vertical_sides_it_crosses(tmp_path):
    thick = Side("thick")
    path = bordered(
        tmp_path,
        {
            "A1": ("a long text that spills two columns", Border(right=thick, bottom=thick)),
            "B1": ("", Border(right=Side("thin"), top=Side("thin"))),
            "D1": ("stop", Border(left=Side("thick", color=BLUE))),
            "A2": ("short", Border(right=thick)),
        },
    )
    found = resolved(path)
    # A1's text crosses A|B and B|C and stops at D1's value: C|D keeps D1's left side.
    assert found.vertical[0][1:4] == (None, None, Line("thick", "0000FF"))
    # Horizontal sides stay.
    assert found.horizontal[1][0] == Line("thick")
    assert found.horizontal[0][1] == Line("thin")
    # A text that does not spill keeps its border.
    assert found.vertical[1][1] == Line("thick")


def test_a_hidden_column_s_sides_do_not_compete(tmp_path):
    path = bordered(
        tmp_path,
        {
            "A1": ("x", Border(right=Side("thick", color=RED))),
            "B1": ("x", Border(left=Side("double"), right=Side("double"))),
            "C1": ("x", Border(left=Side("thick", color=BLUE))),
        },
        hidden=["B"],
    )
    # A and C meet on one edge: their thick sides compete, blue being darker (unverified).
    assert resolved(path).vertical[0][1] == Line("thick", "0000FF")


def test_only_excel_mode_resolves_borders(tmp_path):
    path = bordered(tmp_path, {"A1": ("x", Border(left=Side("thin")))})
    sheet = sheetpeek.read(path)
    assert rendered_borders(sheet.layout("characters")) is None
    plain = bordered(tmp_path, {"A1": ("x", None)}, name="plain.xlsx")
    assert sheetpeek.read(plain).layout("excel").borders is None


# --- the outputs


def test_borders_are_shown_in_html_and_reported_elsewhere(tmp_path):
    path = bordered(
        tmp_path,
        {"A1": ("x", Border(left=Side("thin"), diagonal=Side("thin"), diagonalUp=True))},
    )
    assert sheetpeek.notices(path, "html") == ()
    for name, reason in (
        ("rich", "not_implemented"),
        ("text", "not_implemented"),
        ("markdown", "not_representable"),
        ("csv", "not_representable"),
    ):
        found = {n.code: n for n in sheetpeek.notices(path, name)}
        for code in ("border", "border.diagonal"):
            assert (found[code].outcome, found[code].reason) == ("omitted", reason), (name, code)


def test_slant_dash_dot_is_drawn_as_medium_dash_dot_and_reported(tmp_path):
    path = bordered(
        tmp_path,
        {
            "A1": ("x", Border(bottom=Side("slantDashDot"))),
            "C1": ("x", Border(bottom=Side("mediumDashDot"))),
        },
    )
    (notice,) = sheetpeek.notices(path, "html")
    assert (notice.code, notice.outcome) == ("border.slant_dash_dot", "approximated")
    page = sheetpeek.render(path, "html", on_notice="ignore")
    assert len(rules(page)) == 1  # one class for both


def lines(page):
    return re.findall(r'<i class="(sp-b-[a-z]+)" style="([^"]*)"></i>', page)


def rules(page):
    return dict(re.findall(r"\.sheetpeek \.(sp-b-[a-z]+) \{ ([^}]*) \}", page))


def geometry(style):
    return tuple(int(n) for n in re.findall(r"(-?\d+)px", style))


@pytest.mark.parametrize(
    ("style", "top", "height"),
    [  # the band around A1's bottom gridline at y 20 px: screen points are CSS px (captured)
        ("thin", 20, 1),
        ("hair", 20, 1),
        ("dotted", 20, 1),
        ("medium", 19, 2),
        ("mediumDashed", 19, 2),
        ("thick", 19, 3),
        ("double", 19, 3),
    ],
)
def test_each_style_sits_on_its_gridline_in_css_px(tmp_path, style, top, height):
    path = bordered(tmp_path, {"A1": ("x", Border(bottom=Side(style)))})
    ((_, geometry_style),) = lines(sheetpeek.render(path, "html"))
    # Across A, 71 px, and over the far gridline.
    assert geometry(geometry_style) == (0, top, 72, height)


def test_a_vertical_line_runs_down_its_rows_as_one_box(tmp_path):
    path = bordered(
        tmp_path,
        {f"B{row}": ("x", Border(left=Side("thin", color=RED))) for row in (1, 2, 3)},
    )
    ((_, style),) = lines(sheetpeek.render(path, "html"))
    assert geometry(style) == (0, 0, 1, 61)  # the range starts at B


def test_corners_are_square(tmp_path):
    thick = Side("thick")
    path = bordered(
        tmp_path,
        {"A1": ("", None), "B2": ("x", Border(left=thick, right=thick, top=thick, bottom=thick))},
    )
    found = sorted(geometry(style) for _, style in lines(sheetpeek.render(path, "html")))
    # Each line reaches over the bands of the lines across its ends.
    assert found == [
        (70, 19, 3, 23),  # left of B2: y 20 to 40, from 1 above to 2 below
        (70, 19, 73, 3),  # top: B is 70 px, from 1 before to 2 past
        (70, 39, 73, 3),  # bottom
        (140, 19, 3, 23),  # right
    ]


def test_dashes_are_tiles_with_white_gaps(tmp_path):
    path = bordered(
        tmp_path,
        {
            "A1": ("x", Border(bottom=Side("dashed", color=RED))),
            "C1": ("x", Border(left=Side("dotted"))),
        },
    )
    page = sheetpeek.render(path, "html")
    bodies = list(rules(page).values())
    dashed = next(b for b in bodies if "FF0000" in b)
    assert dashed.startswith("background:#FFFFFF url(")
    assert "background-size:4px 1px" in dashed  # 8 x 2 device px
    assert "fill='%23FFFFFF'" in dashed
    # Excel 16.115 (screen suite borders): on 3 off 1 screen point, both device rows alike.
    on = "".join(f"M{x} {y}h1v1h-1z" for y in (0, 1) for x in range(6))
    assert on in dashed
    dotted = next(b for b in bodies if "FF0000" not in b)
    assert "background-size:1px 4px" in dotted  # 2 on 2 off, turned upright


def test_a_diagonal_is_a_line_from_corner_to_corner(tmp_path):
    path = bordered(
        tmp_path,
        {"A1": ("x", Border(diagonal=Side("hair", color=RED), diagonalDown=True))},
    )
    page = sheetpeek.render(path, "html")
    ((name, style),) = lines(page)
    assert geometry(style) == (0, 0, 72, 21)
    body = rules(page)[name]
    # Hair is drawn as thin (unverified).
    assert "%3Cline x1='0' y1='0' x2='72' y2='21' stroke='%23FF0000' stroke-width='1'/%3E" in body


def test_no_border_layer_without_borders(tmp_path):
    page = sheetpeek.render(bordered(tmp_path, {"A1": ("x", None)}), "html")
    assert "sp-borders" not in page
    assert "sp-fills" not in page
