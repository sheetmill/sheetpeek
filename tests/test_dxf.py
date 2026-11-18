"""The differential-format layer: a cell's own style over the partial formats beneath it.

The rules are ARCHITECTURE.md §4.20's, unverified until the screen suite ``tables`` captures them.
"""

import pytest

from sheetpeek import CellStyle, Font
from sheetpeek._dxf import stack
from sheetpeek._fill_border_model import Borders, Pattern, Side
from sheetpeek._table_model import DifferentialFormat

PLAIN = CellStyle("General", Font("Calibri", 11.0))
WHITE_BOLD = DifferentialFormat(bold=True, color="FFFFFF")
BLUE = DifferentialFormat(fill=Pattern("solid", "4472C4"))
THIN = Side("thin", "FFFFFF")


def test_nothing_beneath_returns_the_style_itself():
    assert stack(PLAIN, (), own_color=False) is PLAIN
    assert stack(PLAIN, (DifferentialFormat(),), own_color=False) is PLAIN
    own = CellStyle("General", PLAIN.font, borders=Borders(left=THIN))
    assert stack(own, (), own_color=False, sides=Borders(left=Side("thick"))) is own


def test_a_plain_cell_takes_the_formats():
    drawn = stack(PLAIN, (BLUE, WHITE_BOLD), own_color=False)
    assert drawn.fill == "4472C4"
    assert drawn.fill_spec == Pattern("solid", "4472C4")
    assert drawn.font == Font("Calibri", 11.0, bold=True, color="FFFFFF")
    assert drawn.number_format == "General"


def test_a_later_format_wins_property_by_property():
    first = DifferentialFormat(fill=Pattern("solid", "FF0000"), color="000000", italic=True)
    second = DifferentialFormat(fill=Pattern("solid", "00FF00"))
    drawn = stack(PLAIN, (first, second), own_color=False)
    assert drawn.fill == "00FF00"
    assert drawn.font.color == "000000"  # the later format sets no color
    assert drawn.font.italic


def test_a_later_format_can_turn_bold_off():
    drawn = stack(PLAIN, (WHITE_BOLD, DifferentialFormat(bold=False)), own_color=False)
    assert not drawn.font.bold


@pytest.mark.parametrize("own", [Pattern("solid", "FF0000"), Pattern("solid", "FFFFFF")])
def test_an_own_fill_wins(own):
    style = CellStyle("General", PLAIN.font, fill=own.color, fill_spec=own)
    drawn = stack(style, (BLUE,), own_color=False)
    assert drawn.fill_spec == own
    assert drawn.fill == own.color


def test_an_own_pattern_fill_wins():
    own = Pattern("gray125", None, None)
    style = CellStyle("General", PLAIN.font, fill=own.color, fill_spec=own)
    assert stack(style, (BLUE,), own_color=False).fill_spec == own


def test_an_own_font_color_wins():
    style = CellStyle("General", Font("Calibri", 11.0, color="FF0000"))
    assert stack(style, (WHITE_BOLD,), own_color=True).font.color == "FF0000"


def test_an_own_automatic_color_wins_and_draws_black():
    # Automatic, or a font without a color: own, so the table's white does not apply.
    assert stack(PLAIN, (WHITE_BOLD,), own_color=True).font.color is None


def test_an_explicit_theme_1_lets_the_format_color_through():
    assert stack(PLAIN, (WHITE_BOLD,), own_color=False).font.color == "FFFFFF"


def test_own_bold_and_italic_apply_where_the_formats_have_neither():
    style = CellStyle("General", Font("Calibri", 11.0, bold=True, italic=True))
    drawn = stack(style, (BLUE,), own_color=False)
    assert drawn.font.bold
    assert drawn.font.italic


def test_own_bold_wins_over_a_format_turning_it_off():
    style = CellStyle("General", Font("Calibri", 11.0, bold=True))
    assert stack(style, (DifferentialFormat(bold=False),), own_color=False).font.bold


def test_the_own_font_name_and_size_stay_and_the_format_bold_applies_to_them():
    style = CellStyle("General", Font("Arial", 14.0))
    drawn = stack(style, (WHITE_BOLD,), own_color=False)
    assert (drawn.font.name, drawn.font.size, drawn.font.bold) == ("Arial", 14.0, True)


def test_underline_and_strike_stack_like_bold():
    drawn = stack(PLAIN, (DifferentialFormat(underline="double", strike=True),), own_color=False)
    assert (drawn.font.underline, drawn.font.strike) == ("double", True)
    cancelled = (DifferentialFormat(underline="single"), DifferentialFormat(underline="none"))
    assert stack(PLAIN, cancelled, own_color=False).font.underline is None
    own = CellStyle("General", Font("Calibri", 11.0, underline="single"))
    assert stack(
        own, (DifferentialFormat(underline="double"),), own_color=False
    ).font.underline == ("single")


def test_the_formats_sides_apply_where_the_cell_has_none():
    beneath = Borders(top=THIN, bottom=Side("thick", "FFFFFF"))
    drawn = stack(PLAIN, (), own_color=False, sides=beneath)
    assert drawn.borders == Borders(top=THIN, bottom=Side("thick", "FFFFFF"))


def test_own_sides_replace_the_formats_side_by_side():
    own = Borders(top=Side("thin", "FF0000"), diagonal=Side("thin"), diagonal_down=True)
    style = CellStyle("General", PLAIN.font, borders=own)
    beneath = Borders(left=THIN, top=Side("thick", "FFFFFF"))
    drawn = stack(style, (), own_color=False, sides=beneath)
    assert drawn.borders == Borders(
        left=THIN, top=Side("thin", "FF0000"), diagonal=Side("thin"), diagonal_down=True
    )


def test_the_cell_style_is_otherwise_kept():
    style = CellStyle("0.00", PLAIN.font, h_align="center", wrap=True, indent=2)
    drawn = stack(style, (BLUE, WHITE_BOLD), own_color=False)
    assert (drawn.number_format, drawn.h_align, drawn.wrap, drawn.indent) == (
        "0.00",
        "center",
        True,
        2,
    )
