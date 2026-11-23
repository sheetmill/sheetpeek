"""Text styles: superscript and subscript, underlines, strikethrough, shrink to fit, vertical
justify and distributed, horizontal distributed with indent, and reading order (ARCHITECTURE.md
§4.13).

The expected values come from the screen suite ``text_style``; each test says which.
"""

import io

import openpyxl
import pytest
from openpyxl.styles import Alignment, Font

import sheetpeek
from sheetpeek import TextRule
from sheetpeek._model import Font as ModelFont
from sheetpeek._rows import script_parts, script_size
from sheetpeek._text_styles import shrink_scale
from sheetpeek.backends.html import HTMLBackend
from tests.workbooks import rewrite, save

HEBREW = "אבג!"


def book(tmp_path, build):
    wb = openpyxl.Workbook()
    build(wb.active)
    return save(wb, tmp_path / "book.xlsx")


def lay_out(tmp_path, build, mode="excel"):
    return sheetpeek.read(book(tmp_path, build)).layout(mode)


def runs_of(rendered, row, column=None):
    found = next(r for r in rendered.rows if r.index == row).runs
    return [r for r in found if column is None or r.source[1] == column]


def html(rendered):
    out = io.StringIO()
    HTMLBackend(fragment=True).render(rendered, out)
    return out.getvalue()


# Reading.


def test_the_styles_are_read(tmp_path):
    def build(ws):
        ws["A1"] = "x"
        ws["A1"].font = Font(vertAlign="superscript")
        ws["A1"].alignment = Alignment(
            shrink_to_fit=True, readingOrder=2, horizontal="distributed", justifyLastLine=True
        )
        ws["A2"] = "x"
        ws["A2"].font = Font(vertAlign="subscript")
        ws["A3"] = "x"
        ws["A3"].font = Font(vertAlign="baseline")

    sheet = sheetpeek.read(book(tmp_path, build))
    one = sheet.cell(1, 1).style
    assert (one.font.vert_align, one.shrink, one.reading_order, one.justify_last_line) == (
        "superscript",
        True,
        2,
        True,
    )
    assert sheet.cell(2, 1).font.vert_align == "subscript"
    assert sheet.cell(3, 1).font.vert_align is None
    assert sheet.notices == ()  # nothing of it is reported by the reader any more


@pytest.mark.parametrize("value", ["3", "1e308", "nan", "inf"])
def test_a_reading_order_no_interface_writes_reads_as_context(tmp_path, value):
    # The guard: anything but 1 and 2 is 0, so no other value reaches the layout or the census.
    def build(ws):
        ws["A1"] = HEBREW
        ws["A1"].alignment = Alignment(readingOrder=1)

    path = book(tmp_path, build)
    rewrite(
        path, "xl/styles.xml", lambda t: t.replace('readingOrder="1"', f'readingOrder="{value}"')
    )
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 1).style.reading_order == 0
    rendered = sheet.layout("excel")
    assert "text.reading_order" not in rendered.features
    assert runs_of(rendered, 1)[0].direction == "rtl"


# Superscript and subscript (suite text_style).


def test_superscript_and_subscript_are_drawn_smaller_raised_and_lowered(tmp_path):
    def build(ws):
        for column, position in (("A", None), ("B", "superscript"), ("C", "subscript")):
            ws[f"{column}1"] = "HHH"
            ws[f"{column}1"].font = Font(vertAlign=position)

    rendered = lay_out(tmp_path, build)
    plain, sup, sub = runs_of(rendered, 1)
    # drawn at round(2 * 11 / 3) = 7 pt
    assert (sup.scale, sub.scale) == (pytest.approx(7 / 11), pytest.approx(7 / 11))
    assert plain.scale == 1
    # Raised 5 pt and lowered 2 pt at 11 pt; a row with both is fitted to 18 pt.
    assert (plain.y - sup.y, sub.y - plain.y) == (5, 2)
    assert rendered.rows[0].height_pt == 18
    # Measured at the drawn size: an H is 5 pt at 7 pt, 7 at 11.
    assert (plain.width, sup.width) == (21, 15)


@pytest.mark.parametrize("position", ["superscript", "subscript"])
def test_one_kind_grows_a_fitted_row_by_two_points(tmp_path, position):
    def build(ws):
        ws["A1"] = "Hxg"
        ws["A1"].font = Font(vertAlign=position)

    assert lay_out(tmp_path, build).rows[0].height_pt == 17  # 15 plain


def test_the_raise_and_lowering_at_24_points(tmp_path):
    # Captured (row_bottoms, scripts groups): raised floor(24/2) = 12 and lowered
    # floor(24/5) = 4 from the row's baseline; drawn at round(2 * 24 / 3) = 16.
    assert script_size(ModelFont("Calibri", 24.0, vert_align="superscript")) == (24, 16, 12, 4)
    # a half size acts at the whole size above it
    assert script_size(ModelFont("Calibri", 23.5, vert_align="subscript")) == (24, 16, 12, 4)

    def build(ws):
        ws["A1"], ws["B1"] = "HHH", "HHH"
        ws["A1"].font = Font(size=24)
        ws["B1"].font = Font(size=24, vertAlign="superscript")
        ws["C1"] = "HHH"
        ws["C1"].font = Font(size=24, vertAlign="subscript")

    plain, sup, sub = runs_of(lay_out(tmp_path, build), 1)
    assert (plain.y - sup.y, sub.y - plain.y) == (12, 4)


def test_a_subscript_lowers_the_rows_shared_bottom(tmp_path):
    # Captured (row_bottoms, "subscript 24 probe"): beside a Calibri 24 subscript, plain 11 pt
    # text in a 64 pt row sits on l + B(d) = 4 + 4 = 8 above the bottom; beside a superscript,
    # on the 24 pt font's own B, 6.
    assert script_parts(ModelFont("Calibri", 24.0, vert_align="subscript"))[1] == 8
    assert script_parts(ModelFont("Calibri", 24.0, vert_align="superscript"))[1] == 6

    def build(ws):
        ws.row_dimensions[1].height = 64
        ws["A1"], ws["B1"] = "HHH", "HHH"
        ws["B1"].font = Font(size=24, vertAlign="subscript")

    rendered = lay_out(tmp_path, build)
    row = rendered.rows[0]
    (probe,) = runs_of(rendered, 1, 1)
    assert row.y_pt + row.height_pt - probe.y == 8


def test_a_tiny_superscript_never_lowers_a_row(tmp_path):
    # A 1 pt superscript's parts stay under the Normal font's: the row is 15 pt.
    def build(ws):
        ws["A1"] = "HHH"
        ws["A1"].font = Font(size=1, vertAlign="superscript")

    assert lay_out(tmp_path, build).rows[0].height_pt == 15


# Underlines and strikethrough (suite text_style).


def test_underline_offsets(tmp_path):
    def build(ws):
        for column, kind in zip(
            "ABCD", ["single", "double", "singleAccounting", "doubleAccounting"], strict=False
        ):
            ws[f"{column}1"] = "HHH"
            ws[f"{column}1"].font = Font(underline=kind)

    single, double, accounting, double_accounting = runs_of(lay_out(tmp_path, build), 1)
    # Single 1 pt below the baseline; double 0 pt (and a second line 2 screen pt lower);
    # accounting 3 pt.
    assert [(r.rules[0].y - r.y, r.rules[0].lines) for r in (single, double)] == [(1, 1), (0, 2)]
    assert [(r.rules[0].y - r.y, r.rules[0].lines) for r in (accounting, double_accounting)] == [
        (3, 1),
        (3, 2),
    ]
    # Single and double span the text.
    # Excel 16.115 runs a single or double underline 1 pt past the text (text_style suite).
    assert (single.rules[0].x0, single.rules[0].x1) == (single.x, single.x + single.width + 1)


def test_accounting_underlines_span_text_cells_and_numbers(tmp_path):
    def build(ws):
        ws.column_dimensions["A"].width = 30
        ws["A1"] = "HH"
        ws["B1"] = 123
        for cell in ("A1", "B1"):
            ws[cell].font = Font(underline="singleAccounting")

    rendered = lay_out(tmp_path, build)
    text, number = runs_of(rendered, 1)
    a = rendered.columns[0]
    # Under text, the cell less 3 pt and 2 pt; under a number, the number.
    assert (text.rules[0].x0, text.rules[0].x1) == (a.x_pt + 3, a.x_pt + a.width_pt - 2)
    assert (number.rules[0].x0, number.rules[0].x1) == (number.x, number.x + number.width)
    # The row is fitted to 18 pt, the baseline 6 pt above its bottom.
    row = rendered.rows[0]
    assert row.height_pt == 18
    assert row.y_pt + row.height_pt - text.y == 6


def test_accounting_is_one_rule_per_line_under_distributed_words(tmp_path):
    def build(ws):
        ws.column_dimensions["A"].width = 30
        ws["A1"] = "aa bb cc"
        ws["A1"].alignment = Alignment(horizontal="distributed")
        ws["A1"].font = Font(underline="singleAccounting")

    words = runs_of(lay_out(tmp_path, build), 1)
    assert [len(r.rules) for r in words] == [1, 0, 0]


def test_strike_is_three_points_above_the_baseline(tmp_path):
    def build(ws):
        ws["A1"] = "Hxg"
        ws["A1"].font = Font(strike=True)

    (run,) = runs_of(lay_out(tmp_path, build), 1)
    assert run.rules == (TextRule("strike", run.x, run.x + run.width, run.y - 3),)


def test_rules_are_one_pixel_at_whole_pixels_in_html(tmp_path):
    def build(ws):
        ws["A1"] = "HHH"
        ws["A1"].font = Font(underline="double", strike=True, color="FF0000")

    page = html(lay_out(tmp_path, build))
    # Not the browser's decorations, which scale with the page.
    assert "text-decoration:underline" not in page
    assert "line-through" not in page
    # Baseline 12 pt is 16 px; the double underline's top is at the baseline, the strike's top
    # 3 pt (4 px) above it. The underline runs 1 pt (1 px here) past the text, the strike not.
    assert (
        '<i class="sp-rule sp-double" style="left:4px;top:16px;width:29px;color:#FF0000">' in page
    )
    assert '<i class="sp-rule" style="left:4px;top:12px;width:28px;color:#FF0000">' in page
    assert ".sheetpeek .sp-rule { position: absolute; height: 1px;" in page


def test_rules_at_36_points_are_two_points_thick(tmp_path):
    # Captured (text_style_more, group rules, zoom 100): at Calibri 36 a single underline 4 pt
    # below the baseline and a strikethrough 9 pt above it, each 2 pt thick; a double's two
    # lines 2 pt thick and 4 pt apart, the first 4 pt below; at 11 pt every rule is 1 pt.
    def build(ws):
        ws["A1"], ws["B1"], ws["C1"] = "HHH", "HHH", "HHH"
        ws["A1"].font = Font(size=36, underline="single", strike=True)
        ws["B1"].font = Font(size=36, underline="double")
        ws["C1"].font = Font(underline="single")

    single, double, small = runs_of(lay_out(tmp_path, build), 1)
    assert [(r.kind, r.y - single.y, r.thickness) for r in single.rules] == [
        ("underline", 4, 2),
        ("strike", -9, 2),
    ]
    (rule,) = double.rules
    assert (rule.y - double.y, rule.lines, rule.thickness) == (4, 2, 2)
    assert small.rules[0].thickness == 1


def test_a_double_underlines_gap_is_xlformats(tmp_path):
    # xlformat's rule_metrics: Times New Roman 38 draws a double underline's 1 pt lines 3 pt
    # apart, not two thicknesses; Calibri 36's 2 pt lines are 4 apart.
    def build(ws):
        ws["A1"], ws["B1"] = "HHH", "HHH"
        ws["A1"].font = Font(name="Times New Roman", size=38, underline="double")
        ws["B1"].font = Font(size=36, underline="double")

    times, calibri = runs_of(lay_out(tmp_path, build), 1)
    assert (times.rules[0].thickness, times.rules[0].gap) == (1, 3)
    assert (calibri.rules[0].thickness, calibri.rules[0].gap) == (2, 4)
    out = io.StringIO()
    HTMLBackend(fragment=True, zoom=100).render(lay_out(tmp_path, build), out)
    assert ';height:4px;border-width:1px"' in out.getvalue()


def test_a_thick_rule_is_drawn_at_its_thickness_in_html(tmp_path):
    def build(ws):
        ws["A1"], ws["B1"] = "HHH", "HHH"
        ws["A1"].font = Font(size=36, underline="single")
        ws["B1"].font = Font(size=36, underline="double")

    out = io.StringIO()
    HTMLBackend(fragment=True, zoom=100).render(lay_out(tmp_path, build), out)
    page = out.getvalue()
    # 2 pt is 2 px at zoom 100; a double's lines 2 px each in a box of three thicknesses
    assert ';height:2px"' in page
    assert ';height:6px;border-width:2px"' in page


def test_superscript_is_drawn_smaller_in_html(tmp_path):
    def build(ws):
        ws["A1"] = "HHH"
        ws["A1"].font = Font(vertAlign="superscript")

    assert "font:7pt 'sheetpeek-calibri'" in html(lay_out(tmp_path, build))


# Shrink to fit (suite text_style).

CAPTURED_STEPS = {
    **dict.fromkeys(range(4, 7), 7),
    **dict.fromkeys((7, 8), 6),
    **dict.fromkeys((9, 10), 5),
    **dict.fromkeys((11, 12), 4),
    **dict.fromkeys(range(13, 17), 3),
    **dict.fromkeys(range(17, 23), 2),
}


def test_shrink_to_fit_takes_the_captured_levels(tmp_path):
    def build(ws):
        for n in CAPTURED_STEPS:
            ws.cell(n, 1, "H" * n).alignment = Alignment(horizontal="left", shrink_to_fit=True)

    rendered = lay_out(tmp_path, build)
    found = {n: round(runs_of(rendered, n)[0].scale * 7, 6) for n in CAPTURED_STEPS}
    assert found == CAPTURED_STEPS
    for n in CAPTURED_STEPS:
        (run,) = runs_of(rendered, n)
        assert run.clip[2] == rendered.columns[0].width_pt  # never spills


def test_shrink_scale_takes_the_largest_level_that_fits():
    # Calibri 11 in a 53 pt cell (budget 48), "H" 7 pt: 7 Hs (49) take level 6, 6 Hs fit.
    assert shrink_scale(49, 7, 53, 48) == 6 / 7
    assert shrink_scale(42, 7, 53, 48) == 1.0
    # Calibri 16 in 69 pt, "H" 10 pt: 11 Hs are refused at 6 ((11 + 0.375)·6 > 68), take 5;
    # 22 Hs take 3, the same 66 pt of advances that a single budget could not tell apart.
    assert shrink_scale(110, 10, 69, 64) == 5 / 10
    assert shrink_scale(220, 10, 69, 64) == 3 / 10


def test_shrink_scale_is_bounded():
    assert shrink_scale(7000, 7, 53, 48) == 1 / 7  # nothing fits: the smallest level, clipped
    assert shrink_scale(7, 7, -5, -10) == 1 / 7
    assert shrink_scale(100, 1, 53, 48) == 1.0  # no level below a 1 pt advance
    assert shrink_scale(100, 0, 53, 48) == 1.0


def test_shrink_replaces_hashes_and_wrap_wins(tmp_path):
    def build(ws):
        ws["A1"] = 123456789012
        ws["A1"].number_format = "0"
        ws["A1"].alignment = Alignment(shrink_to_fit=True)
        ws["A2"] = "shrink and wrap both set here"
        ws["A2"].alignment = Alignment(wrap_text=True, shrink_to_fit=True)
        ws["A3"] = "H" * 12
        ws["A3"].alignment = Alignment(horizontal="left", shrink_to_fit=True)
        ws["A3"].font = Font(size=20)

    rendered = lay_out(tmp_path, build)
    cell = rendered.cell(1, 1)
    assert (cell.text, cell.overflow) == ("123456789012", False)
    (number,) = runs_of(rendered, 1)
    assert number.scale < 1
    assert number.align == "right"
    assert all(r.scale == 1 for r in runs_of(rendered, 2))
    assert len(runs_of(rendered, 2)) > 1
    # The row is fitted to the unshrunk size.
    assert rendered.rows[2].height_pt > 15
    assert runs_of(rendered, 3)[0].scale < 1


def test_shrink_fits_what_the_indent_leaves(tmp_path):
    def build(ws):
        ws["A1"] = "H" * 12
        ws["A1"].alignment = Alignment(horizontal="left", indent=2, shrink_to_fit=True)

    (run,) = runs_of(lay_out(tmp_path, build), 1)
    assert run.x == 3 + 12
    assert run.x + run.width <= 53 - 2


def test_a_merge_is_not_shrunk(tmp_path):
    # Captured with the flag on the anchor only (doubtful): clipped at the merge's edge.
    def build(ws):
        ws["A1"] = "a long text merged and shrunk"
        ws.merge_cells("A1:B1")
        ws["A1"].alignment = Alignment(shrink_to_fit=True)

    assert {r.scale for r in runs_of(lay_out(tmp_path, build), 1)} == {1}


def test_characters_mode_clips_shrunk_text_and_keeps_hashes(tmp_path):
    def build(ws):
        ws["A1"] = "a text too long for its cell"
        ws["A1"].alignment = Alignment(shrink_to_fit=True)
        ws["A2"] = 123456789012
        ws["A2"].number_format = "0"
        ws["A2"].alignment = Alignment(shrink_to_fit=True)

    rendered = lay_out(tmp_path, build, "characters")
    assert rendered.cell(1, 1).spill == (0, 0)
    assert rendered.cell(1, 1).lines == ("a text t",)
    assert rendered.cell(2, 1).text == "########"
    found = sheetpeek.notices(rendered, "text")
    assert [n.outcome for n in found if n.code == "text.shrink"] == ["approximated"]


# Vertical justify and distributed (suite text_style).


@pytest.mark.parametrize(("v_align", "baseline"), [("justify", 12), ("distributed", 34)])
def test_one_line_is_at_top_or_center(tmp_path, v_align, baseline):
    def build(ws):
        ws.row_dimensions[1].height = 60
        ws["A1"] = "Hxg"
        ws["A1"].alignment = Alignment(vertical=v_align)

    (run,) = runs_of(lay_out(tmp_path, build), 1)
    assert run.y == baseline  # top: 12 below the top; center: floor((60 - 3 - 12) / 2) + 12


@pytest.mark.parametrize("v_align", ["justify", "distributed"])
def test_lines_spread_from_the_top_position_to_the_bottom_one(tmp_path, v_align):
    def build(ws):
        ws.row_dimensions[1].height = 60
        ws["A1"] = "aaa\nbbb\nccc"
        ws["A1"].alignment = Alignment(vertical=v_align, wrap_text=True)
        ws["B1"] = "aaaa bbbb cccc dddd"  # wrap off: justify and distributed wrap it
        ws["B1"].alignment = Alignment(vertical=v_align)

    rendered = lay_out(tmp_path, build)
    assert [r.y for r in runs_of(rendered, 1, 1)] == [12, 34, 57]  # top 12, bottom 60 - 3
    wrapped = runs_of(rendered, 1, 2)
    assert len(wrapped) > 1
    assert (wrapped[0].y, wrapped[-1].y) == (12, 57)
    assert rendered.notices == ()


def test_characters_mode_spreads_in_whole_lines(tmp_path):
    def build(ws):
        ws["A1"] = "a\nb"
        ws["A1"].alignment = Alignment(vertical="distributed", wrap_text=True)
        ws["B1"] = "one\ntwo\nthree\nfour"
        ws["B1"].alignment = Alignment(wrap_text=True)
        ws["C1"] = "x"
        ws["C1"].alignment = Alignment(vertical="distributed")
        ws["D1"] = "y"
        ws["D1"].alignment = Alignment(vertical="justify")

    rendered = lay_out(tmp_path, build, "characters")
    assert [line.strip() for line in rendered.cell(1, 1).lines] == ["a", "", "", "b"]
    assert [line.strip() for line in rendered.cell(1, 3).lines] == ["", "x", "", ""]
    assert [line.strip() for line in rendered.cell(1, 4).lines] == ["y", "", "", ""]


# Horizontal distributed and indent (suite text_style).


def test_distributed_indent_moves_both_ends(tmp_path):
    def build(ws):
        ws.column_dimensions["A"].width = 26
        for row, indent in ((1, 0), (2, 2)):
            ws.cell(row, 1, "aa bb cc").alignment = Alignment(
                horizontal="distributed", indent=indent
            )

    rendered = lay_out(tmp_path, build)
    plain, indented = runs_of(rendered, 1), runs_of(rendered, 2)
    assert indented[0].x - plain[0].x == 12
    assert (plain[-1].x + plain[-1].width) - (indented[-1].x + indented[-1].width) == 12


def test_right_indent_goes_on_the_side_general_resolves_to(tmp_path):
    def build(ws):
        ws["A1"] = 123
        ws["A1"].alignment = Alignment(indent=2)
        ws["B1"] = "HHH"
        ws["B1"].alignment = Alignment(horizontal="right", indent=1)

    number, text = runs_of(lay_out(tmp_path, build), 1)
    assert number.x + number.width == 53 - 2 - 12
    assert text.x + text.width == 106 - 2 - 6


# Reading order (suite text_style).


def test_general_right_to_left_text_aligns_right_whatever_the_order(tmp_path):
    def build(ws):
        ws["A1"] = HEBREW
        ws["B1"] = HEBREW
        ws["B1"].alignment = Alignment(readingOrder=1)
        ws["C1"] = "abc!"
        ws["C1"].alignment = Alignment(readingOrder=2)
        ws["D1"] = HEBREW
        ws["D1"].alignment = Alignment(horizontal="left")

    rendered = lay_out(tmp_path, build)
    context, ltr, latin, left = runs_of(rendered, 1)
    assert (context.align, context.direction) == ("right", "rtl")
    assert (ltr.align, ltr.direction) == ("right", "ltr")  # order 1 changes the bidi only
    assert (latin.align, latin.direction) == ("left", "ltr")  # order 2 on Latin: no change
    assert (left.align, left.direction) == ("left", "rtl")
    page = html(rendered)
    assert page.count("direction:rtl") == 2
    assert lay_out(tmp_path, build, "characters").cell(1, 1).align == "right"


# Notices.


def test_what_each_output_reports(tmp_path):
    def build(ws):
        ws["A1"] = "x"
        ws["A1"].font = Font(vertAlign="superscript", underline="doubleAccounting")
        ws["A2"] = "y"
        ws["A2"].alignment = Alignment(shrink_to_fit=True, readingOrder=2, vertical="justify")

    path = book(tmp_path, build)
    codes = {"font.vert_align", "font.underline_accounting", "text.shrink", "text.reading_order"}

    def found(to):
        return {n.code: n.outcome for n in sheetpeek.notices(path, to) if n.code in codes}

    assert found("html") == {}
    assert found("rich") == dict.fromkeys(codes, "approximated")
    assert found("text") == {
        "font.vert_align": "omitted",
        "font.underline_accounting": "omitted",
        "text.shrink": "approximated",
        "text.reading_order": "approximated",
    }
    assert found("markdown") == {
        "font.vert_align": "approximated",
        "font.underline_accounting": "omitted",
        "text.reading_order": "omitted",
    }
    assert found("csv") == {
        "font.vert_align": "omitted",
        "font.underline_accounting": "omitted",
        "text.reading_order": "omitted",
    }


def test_a_row_fits_the_lines_of_distributed_text(tmp_path):
    # Excel 16.115 fits the row to the lines distributed text draws, wrap flag or not: two lines
    # of Calibri 11 are 32 pt (screen suite text_style, G14).
    wb = openpyxl.Workbook()
    wb.active["A1"] = "spilled text that is much wider than one column"
    wb.active["A1"].alignment = Alignment(horizontal="distributed")
    wb.active.row_dimensions[1]  # listed, so Excel fits it
    path = tmp_path / "distributed.xlsx"
    wb.save(path)
    rendered = sheetpeek.read(path).layout("excel")
    lines = {round(run.y) for run in rendered.rows[0].runs}
    assert rendered.rows[0].height_pt == 16.0 * len(lines)
    assert len(lines) > 1
