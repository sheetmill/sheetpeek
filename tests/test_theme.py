"""Theme fonts, the workbook palette, the grid color and the headings (ARCHITECTURE.md §4.14).

Expected values come from captures (Excel 16.114.4 for Mac, dark appearance), to be
recaptured on Microsoft 365; the light heading colors were measured on Microsoft 365.
"""

import io
import re

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color

import sheetpeek
from sheetpeek import _structure
from sheetpeek._colors import (
    DEFAULT_PALETTE,
    grid_color,
    recolored_sections,
    workbook_palette,
)
from sheetpeek._model import Font as ModelFont
from sheetpeek.backends import HTMLBackend, TextBackend
from sheetpeek.backends.html._headings import bands
from tests.workbooks import NS, SHEET1, replace_once, rewrite, save

THEME = "xl/theme/theme1.xml"
STYLES = "xl/styles.xml"
TEAL, BROWN = "008080", "804000"


def themed(path, minor="Courier New", major="Georgia"):
    """Set the theme's Latin typefaces (openpyxl's theme has Calibri and Cambria)."""
    rewrite(
        path,
        THEME,
        lambda t: re.sub(
            r'(<a:minorFont>\s*<a:latin typeface=)"[^"]*"',
            rf'\1"{minor}"',
            re.sub(r'(<a:majorFont>\s*<a:latin typeface=)"[^"]*"', rf'\1"{major}"', t),
        ),
    )
    return path


def build(tmp_path, cells, name="book.xlsx", palette=None):
    wb = openpyxl.Workbook()
    for ref, (value, style) in cells.items():
        wb.active[ref] = value
        for attribute, setting in style.items():
            setattr(wb.active[ref], attribute, setting)
    if palette is not None:
        wb._colors = palette
    return save(wb, tmp_path / name)


def custom_palette():
    """The capture's palette: index 10 ([Red]'s and [Color3]'s slot) teal, 17 ([Color10]) brown."""
    colors = [f"00{c}" for c in DEFAULT_PALETTE]
    colors[10], colors[17] = f"00{TEAL}", f"00{BROWN}"
    return colors


# --- theme fonts


def test_a_scheme_font_is_drawn_in_the_theme_typeface(tmp_path):
    # Excel 16.114 capture: Calibri-minor and Times-New-Roman-minor both draw the theme's Courier
    # New; Calibri-major draws Georgia; Calibri without a scheme stays Calibri.
    path = build(
        tmp_path,
        {
            "A1": ("x", {"font": Font(name="Calibri", scheme="minor")}),
            "A2": ("x", {"font": Font(name="Times New Roman", scheme="minor")}),
            "A3": ("x", {"font": Font(name="Calibri", scheme="major")}),
            "A4": ("x", {"font": Font(name="Calibri")}),
        },
    )
    sheet = sheetpeek.read(themed(path))
    a1, a2, a3, a4 = (sheet.cell(r, 1).font for r in range(1, 5))
    assert (a1.name, a1.scheme, a1.stored_name) == ("Courier New", "minor", "Calibri")
    assert (a2.name, a2.scheme, a2.stored_name) == ("Courier New", "minor", "Times New Roman")
    assert (a3.name, a3.scheme, a3.stored_name) == ("Georgia", "major", "Calibri")
    assert (a4.name, a4.scheme, a4.stored_name) == ("Calibri", None, None)
    assert not [n for n in sheet.notices if n.code.startswith("font")]


def test_the_normal_font_follows_the_theme(tmp_path):
    # Excel 16.114 capture: the Normal style's Calibri (minor) is replaced as well, so the default
    # column takes the theme font's digits (61 pt for Courier New). xlformat has no glyph table
    # for Courier New, so it is measured as Calibri and reported.
    path = themed(build(tmp_path, {"A1": (1, {})}))
    sheet = sheetpeek.read(path)
    assert sheet.normal_font == ModelFont(
        "Courier New", 11.0, scheme="minor", stored_name="Calibri"
    )
    assert sheet.cell(1, 1).font.name == "Courier New"
    rendered = sheet.layout("excel")
    assert [n.context for n in rendered.notices if n.code == "font.no_metrics"] == [
        {"font": "Courier New"}
    ]
    plain = sheetpeek.read(build(tmp_path, {"A1": (1, {})}, "plain.xlsx")).layout("excel")
    assert rendered.columns[0].width_pt == plain.columns[0].width_pt  # measured as Calibri


def test_a_theme_font_with_glyph_tables_sets_the_widths(tmp_path):
    path = themed(build(tmp_path, {"A1": (1, {})}), minor="Aptos Narrow")
    rendered = sheetpeek.read(path).layout("excel")
    assert rendered.sheet.normal_font.name == "Aptos Narrow"
    assert not [n for n in rendered.notices if n.code == "font.no_metrics"]


def test_without_a_theme_font_the_stored_name_is_drawn(tmp_path):
    path = build(tmp_path, {"A1": ("x", {"font": Font(name="Arial", scheme="minor")})})
    rewrite(path, THEME, lambda t: re.sub(r"<a:minorFont>.*?</a:minorFont>", "", t, flags=re.S))
    font = sheetpeek.read(path).cell(1, 1).font
    assert (font.name, font.scheme, font.stored_name) == ("Arial", "minor", None)


def test_an_empty_theme_typeface_keeps_the_stored_name(tmp_path):
    path = build(tmp_path, {"A1": ("x", {"font": Font(name="Arial", scheme="minor")})})
    assert sheetpeek.read(themed(path, minor="")).cell(1, 1).font.name == "Arial"


# --- the palette


def test_indexed_colors_follow_the_workbook_palette(tmp_path):
    # Excel 16.114 capture: fill and font indexed 10 and font 17 follow the palette; fill indexed 2
    # stays red (0 to 7 do not mirror 8 to 15 once a palette is written); fill indexed 64 is no
    # fill; fill 81, past the palette, is black; font indexed 64 is the automatic black.
    path = build(
        tmp_path,
        {
            "A1": ("x", {"fill": PatternFill("solid", fgColor=Color(indexed=10))}),
            "A2": ("x", {"font": Font(color=Color(indexed=10))}),
            "A3": ("x", {"font": Font(color=Color(indexed=17))}),
            "A4": ("x", {"fill": PatternFill("solid", fgColor=Color(indexed=2))}),
            "A5": ("x", {"fill": PatternFill("solid", fgColor=Color(indexed=64))}),
            "A6": ("x", {"fill": PatternFill("solid", fgColor=Color(indexed=81))}),
            "A7": ("x", {"font": Font(color=Color(indexed=64))}),
        },
        palette=custom_palette(),
    )
    sheet = sheetpeek.read(path)
    cells = [sheet.cell(r, 1) for r in range(1, 8)]
    assert cells[0].fill == TEAL
    assert cells[1].font.color == TEAL
    assert cells[2].font.color == BROWN
    assert cells[3].fill == "FF0000"
    assert cells[4].fill is None
    assert cells[5].fill == "000000"
    assert cells[6].font.color is None


def test_the_palette_takes_the_default_where_the_file_gives_none():
    assert workbook_palette(None) == DEFAULT_PALETTE
    assert workbook_palette([]) == DEFAULT_PALETTE
    short = workbook_palette(["00123456"])
    assert short[0] == "123456"
    assert short[1:] == DEFAULT_PALETTE[1:]
    long = workbook_palette(["00ABCDEF"] * 100)
    assert len(long) == 64
    assert set(long) == {"ABCDEF"}
    assert workbook_palette(["not a color", 5, "00FF00FF"])[:3] == ("000000", "FFFFFF", "FF00FF")


def test_a_colors_element_without_indexed_colors_keeps_the_default_palette(tmp_path):
    # openpyxl reads <colors> with only mruColors as an empty palette, which would make every
    # indexed color black or automatic.
    path = build(tmp_path, {"A1": ("x", {"font": Font(color=Color(indexed=10))})})
    rewrite(
        path,
        STYLES,
        lambda t: re.sub(
            r"<colors>.*?</colors>",
            '<colors><mruColors><color rgb="FF00FF00" /></mruColors></colors>',
            t,
            flags=re.S,
        ),
    )
    assert sheetpeek.read(path).cell(1, 1).font.color == "FF0000"


# --- number-format colors


def test_recolored_sections():
    palette = workbook_palette(custom_palette())
    assert recolored_sections("[Color10]0", palette) == (10,)
    assert recolored_sections("[color 3]0;[COLOR10]-0", palette) == (3, 10)
    assert recolored_sections("[Red]0;[Color11]-0", palette) == ()  # named and unchanged
    assert recolored_sections('"[Color10]"0;\\[Color10]', palette) == ()  # literal text
    assert recolored_sections("[Color10]0", DEFAULT_PALETTE) == ()
    assert recolored_sections("[Color57]0;[Color999]0;[Color0]0", palette) == ()


def test_a_section_color_the_palette_changes_is_reported(tmp_path):
    # Excel 16.114 capture: [Color10] follows the palette (128,64,0 here); [Red] keeps red.
    path = build(
        tmp_path,
        {
            "A1": (-1, {"number_format": "[Color10]0;[Red]-0"}),
            "A2": (1, {"number_format": "[Red]0"}),
            "A3": (None, {"number_format": "[Color10]0"}),  # no value: nothing drawn
        },
        palette=custom_palette(),
    )
    sheet = sheetpeek.read(path)
    (notice,) = [n for n in sheet.notices if n.code == "format.palette"]
    assert (notice.outcome, notice.reason, notice.count) == ("approximated", "not_implemented", 1)
    assert notice.cells == (sheetpeek.CellRange(1, 1, 1, 1),)
    assert notice.context == {"colors": "[Color10]"}
    assert "[Color10]" in notice.message
    # Drawn from the default palette until the palette reaches the layout.
    assert sheet.layout("excel").cell(1, 1).format_color == "FF0000"  # [Red] for -1


def test_a_default_palette_reports_no_section_color(tmp_path):
    path = build(tmp_path, {"A1": (1, {"number_format": "[Color10]0"})})
    assert not [n for n in sheetpeek.read(path).notices if n.code == "format.palette"]


# --- the grid color


def grid_book(tmp_path, color_id="10", palette=None, default="0", gridlines=True, fill=False):
    cells = {"A1": (1, {}), "B2": (2, {})}
    if fill:
        cells["A1"] = (1, {"fill": PatternFill("solid", fgColor="FFFFFF")})
    path = build(tmp_path, cells, palette=palette)
    attributes = f'defaultGridColor="{default}" colorId="{color_id}"'
    if not gridlines:
        attributes += ' showGridLines="0"'
    return rewrite(path, SHEET1, replace_once("<sheetView ", f"<sheetView {attributes} "))


def test_the_grid_color_is_the_workbook_palette_entry(tmp_path):
    # Excel 16.114 capture: colorId 10 with index 10 teal draws teal gridlines.
    assert sheetpeek.read(grid_book(tmp_path, palette=custom_palette())).grid_color == TEAL
    assert sheetpeek.read(grid_book(tmp_path, color_id="12")).grid_color == "0000FF"


@pytest.mark.parametrize(
    ("color_id", "default"), [("64", "0"), ("10", "1"), ("65", "0"), ("99999", "0")]
)
def test_an_automatic_grid_color_is_none(tmp_path, color_id, default):
    # 64 is the automatic gray (capture on Excel 16.114); past 64: unverified, automatic.
    assert sheetpeek.read(grid_book(tmp_path, color_id, default=default)).grid_color is None


@pytest.mark.parametrize("color_id", ["9" * 5000, "x", "-1", "1.5", "\u0661", ""])
def test_a_malformed_color_id_is_automatic_in_the_scan(color_id):
    # openpyxl refuses most of these when it reads the sheet; the scan runs first.
    xml = (
        f'<worksheet {NS}><sheetViews><sheetView defaultGridColor="0" colorId="{color_id}" '
        'workbookViewId="0"/></sheetViews><sheetData/></worksheet>'
    ).encode()
    assert _structure.scan(io.BytesIO(xml)).grid_color_id == 64


def test_grid_color_bounds():
    assert grid_color(None, DEFAULT_PALETTE) is None
    assert grid_color(-1, DEFAULT_PALETTE) is None
    assert grid_color(64, DEFAULT_PALETTE) is None
    assert grid_color(63, DEFAULT_PALETTE) == "333333"
    assert grid_color(5, DEFAULT_PALETTE[:5]) is None


def test_html_draws_the_grid_color_one_device_pixel_wide(tmp_path):
    text = sheetpeek.render(sheetpeek.read(grid_book(tmp_path, palette=custom_palette())), "html")
    assert f'<div class="sheetpeek sp-g-{TEAL}"' in text
    assert f".sheetpeek.sp-g-{TEAL} {{ --sheetpeek-grid: #{TEAL}; }}" in text
    widths = "border-left-width: 0.5px; border-top-width: 0.5px;"
    assert f".sheetpeek.sp-g-{TEAL} td {{ {widths} }}" in text


def test_a_fill_covers_a_colored_gridline(tmp_path):
    # Excel 16.114 capture: a white fill hides the colored gridline, as it hides the gray one.
    text = sheetpeek.render(sheetpeek.read(grid_book(tmp_path, fill=True)), "html")
    assert '<td style="background:#FFFFFF;border-color:#FFFFFF">' in text


def test_hidden_gridlines_draw_no_grid_color(tmp_path):
    sheet = sheetpeek.read(grid_book(tmp_path, gridlines=False))
    assert sheet.grid_color == "FF0000"
    text = sheetpeek.render(sheet, "html")
    assert "sp-g-" not in text
    assert "sp-nogrid" in text
    assert "grid.color" not in sheet.layout("excel").features


@pytest.mark.parametrize(
    ("name", "outcome"),
    [
        ("html", None),
        ("text", "omitted"),
        ("markdown", "omitted"),
        ("csv", "omitted"),
        ("tsv", "omitted"),
        ("rich", "approximated"),
    ],
)
def test_each_output_declares_the_grid_color(tmp_path, name, outcome):
    sheet = sheetpeek.read(grid_book(tmp_path))
    assert not [n for n in sheet.notices if n.code == "grid.color"]  # no reader notice
    found = [n for n in sheetpeek.notices(sheet, name) if n.code == "grid.color"]
    assert [n.outcome for n in found] == ([outcome] if outcome else [])


def test_rich_draws_the_separators_in_the_grid_color(tmp_path):
    pytest.importorskip("rich")
    from sheetpeek.backends import RichBackend

    sheet = sheetpeek.read(grid_book(tmp_path, palette=custom_palette()))
    backend = RichBackend("|", color_system="truecolor", force_terminal=True)
    text = sheetpeek.render(sheet, backend)
    assert "\x1b[38;2;0;128;128m|" in text
    plain = sheetpeek.render(sheetpeek.read(grid_book(tmp_path, color_id="64")), backend)
    assert "38;2" not in plain


def test_text_draws_no_grid_color(tmp_path):
    sheet = sheetpeek.read(grid_book(tmp_path))
    assert "\x1b" not in sheetpeek.render(sheet, TextBackend("|"))


# --- the headings


def test_heading_bands_at_calibri_11():
    # Excel 16.114 capture at zoom 100 and 200: the column band 21 pt; the row band 16 + 3 pt a
    # digit of the last visible row (22, 25 and 28 pt with 2, 3 and 4 digits).
    assert bands(11.0, 2) == (22.0, 21.0)
    assert bands(11.0, 3) == (25.0, 21.0)
    assert bands(11.0, 4) == (28.0, 21.0)


def test_heading_bands_grow_with_the_normal_font():
    # Excel 16.114 capture at Calibri 20: column band 25 pt, row band 36 pt with 2 digits. Other
    # sizes follow a line through the two (unverified), positive at every size Excel accepts.
    assert bands(20.0, 2) == pytest.approx((36.0, 25.0))
    for size in (1.0, 409.0):
        row, column = bands(size, 7)
        assert row > 0
        assert column > 0


def test_html_headings_take_the_captured_sizes(tmp_path):
    path = build(tmp_path, {"A1": (1, {}), "A10": (10, {})})
    text = sheetpeek.render(
        sheetpeek.read(path), HTMLBackend(column_headings=True, row_headings=True)
    )
    # 22 pt (2 digits) and 21 pt, in whole CSS px at 4/3.
    assert '<col style="width:29px">' in text
    assert '<thead><tr style="height:28px">' in text


@pytest.mark.parametrize(
    ("scheme", "colors"),
    [
        ("light", ("#f7f7f7", "#5e5e5e", "#e0dfdd")),
        ("dark", ("#161616", "#ffffff", "#565656")),
    ],
)
def test_heading_colors(tmp_path, scheme, colors):
    text = sheetpeek.render(sheetpeek.read(build(tmp_path, {"A1": (1, {})})), "html")
    selector = ".sheetpeek {" if scheme == "light" else '.sheetpeek[data-color-scheme="dark"] {'
    block = text[text.index(selector) :].split("}")[0]
    for name, color in zip(("paper", "ink", "grid"), colors, strict=True):
        assert f"--sheetpeek-heading-{name}: {color};" in block
