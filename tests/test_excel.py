"""excel mode: geometry in points, text runs, row heights (ARCHITECTURE.md §2.1, §4.7)."""

import openpyxl
import pytest
import xlformat
from openpyxl.styles import Alignment, Font

import sheetpeek
from sheetpeek import _platform
from sheetpeek._rows import row_height
from tests.workbooks import SHEET1, raw_sheet, replace_once, rewrite, save

LONG = "a long text value that spills"


def lay_out(tmp_path, build, ref=None, **options):
    wb = openpyxl.Workbook()
    build(wb.active)
    return sheetpeek.read(save(wb, tmp_path / "book.xlsx"), ref=ref, platform="mac").layout(
        "excel", **options
    )


def test_geometry_in_points(tmp_path):
    def build(ws):
        ws["A1"] = 1
        ws["B2"] = "x"
        ws.column_dimensions["C"].width = 20  # 120 pt
        ws["C2"] = 2

    rendered = lay_out(tmp_path, build)
    assert [(c.x_pt, c.width_pt) for c in rendered.columns] == [(0, 53), (53, 53), (106, 120)]
    assert [(r.y_pt, r.height_pt) for r in rendered.rows] == [(0, 15), (15, 15)]
    assert rendered.width_pt == 226
    assert all(c.width is None for c in rendered.columns)
    (one,) = rendered.rows[0].runs
    # Right-aligned: ends 2 pt before the next gridline; baseline 3 pt above it.
    assert (one.text, one.x + one.width, one.y, one.align) == ("1", 51, 12, "right")
    (x,) = [r for r in rendered.rows[1].runs if r.text == "x"]
    assert (x.x, x.clip) == (56, (54, 16, 106, 30))
    assert rendered.cell(1, 1).lines == ()


def test_spill_both_ways_and_competing(tmp_path):
    def build(ws):
        ws["A1"] = LONG  # left: spills right over B and C
        ws["D2"] = LONG  # right-aligned: spills left
        ws["D2"].alignment = Alignment(horizontal="right")
        ws["A3"] = LONG
        ws["C3"] = LONG  # competing for B
        ws["C3"].alignment = Alignment(horizontal="right")

    rendered = lay_out(tmp_path, build, ref="A1:D3")
    width = xlformat.text_width(LONG, font=xlformat.Font("Calibri", 11), platform="mac")
    assert rendered.cell(1, 1).spill[1] >= 2
    assert rendered.rows[0].crossed[:2] == (True, True)
    (run,) = rendered.rows[1].runs
    assert run.x + run.width == 4 * 53 - 2
    assert run.width == width
    assert rendered.cell(2, 4).spill[0] >= 2
    a, c = rendered.rows[2].runs  # decided: both spills are recorded, as Excel draws both
    assert a.clip[2] == 106  # stops at C, which holds text
    assert c.clip[0] == 54  # stops at A
    assert rendered.rows[2].crossed == (True, True, False)


def test_merged_width_formats_the_anchor(tmp_path):
    # Suite merge_width: a number too wide for one column shows in full merged over three (§4.4).
    def build(ws):
        ws["A1"] = ws["A2"] = 123456789
        ws["A1"].number_format = ws["A2"].number_format = "#,##0.00"
        ws.merge_cells("A2:C2")

    rendered = lay_out(tmp_path, build, ref="A1:C2")
    assert rendered.rows[0].runs[0].text == "#########"
    (merged,) = rendered.rows[1].runs
    assert merged.text == "123,456,789.00"
    assert merged.clip[:3:2] == (1, 159)
    assert rendered.rows[1].crossed == (True, True)


def test_indent_wrap_and_vertical_clip(tmp_path):
    def build(ws):
        ws["A1"] = "a b"
        ws["A1"].alignment = Alignment(indent=2)
        ws["A2"] = "one two\nthree"
        ws["A2"].alignment = Alignment(wrap_text=True)

    rendered = lay_out(tmp_path, build, ref="A1:B2")
    assert rendered.rows[0].runs[0].x == 3 + 2 * 6  # one digit per indent level (§4.6)
    lines = rendered.rows[1].runs
    assert [r.text for r in lines] == ["one two", "three"]
    # Excel fits the row to two wrapped lines, 16 pt each (captured, suite row_height), and
    # draws them 15 pt apart (captured, suite text_style).
    assert rendered.rows[1].height_pt == 32
    assert lines[1].y - lines[0].y == 15
    assert lines[1].y == 15 + 32 - 3  # bottom-aligned
    assert lines[0].clip[1:4:2] == (16, 47)
    assert rendered.rows[1].height == 2


def test_wrapped_rows_follow_xlformats_row_model(tmp_path):
    # xlformat's capture (its CAPTURING.md section 2 item 44): k wrapped lines are k(F + 1), one
    # included; beside a larger font, plus the difference of the faces' insets p (Calibri 11: 1,
    # Calibri 20: 3).
    def build(ws):
        ws["A1"] = "one"
        ws["A1"].alignment = Alignment(wrap_text=True)
        ws["A2"] = "one\ntwo\nthree"
        ws["A2"].alignment = Alignment(wrap_text=True)
        ws["B2"] = 1
        ws["B2"].font = Font(size=20)
        ws["A3"] = "one"
        ws["A3"].alignment = Alignment(wrap_text=True)
        ws["A3"].font = Font(name="Verdana", size=11)  # no row metrics: Calibri's (§6.2)

    rendered = lay_out(tmp_path, build, ref="A1:B3")
    assert [r.height_pt for r in rendered.rows] == [16, 3 * 16 + 2, 16]
    assert {n.context["function"] for n in rendered.notices} == {"row_metrics"}


def test_rows_fit_their_cells(tmp_path):
    # Captured (suite row_height): a 14 pt cell makes its row 19 pt; a 9 pt one does not shrink
    # it; a stored height wins.
    def build(ws):
        ws["A1"] = ws["A2"] = ws["A3"] = 1
        ws["A1"].font = Font(size=14)
        ws["A2"].font = Font(size=9)
        ws["A3"].font = Font(size=14)
        ws.row_dimensions[3].height = 15

    rendered = lay_out(tmp_path, build, ref="A1:A3")
    assert [r.height_pt for r in rendered.rows] == [19, 15, 15]
    # In a row too short for its text, the first line is placed as a block of its own, here at
    # the bottom: 4 pt above it at Calibri 14 (ARCHITECTURE.md §3.6; captured, text_style_more
    # group overflow).
    assert rendered.rows[2].runs[0].y == 19 + 15 + 15 - 4


def test_hidden_column_is_crossed(tmp_path):
    def build(ws):
        ws["A1"] = LONG
        ws["B1"] = "x"
        ws.column_dimensions["B"].hidden = True

    rendered = lay_out(tmp_path, build, ref="A1:C1")
    assert [c.letter for c in rendered.columns][:2] == ["A", "C"]
    assert rendered.rows[0].crossed[0]


def test_unmeasured_cell_font_warns_in_excel_mode(tmp_path):
    def build(ws):
        ws["A1"] = "x"
        ws["A1"].font = Font(name="Segoe UI", size=10)

    rendered = lay_out(tmp_path, build)
    assert rendered.warnings == (
        "font 'Segoe UI' has no glyph table in xlformat; measured as Calibri at the same size",
    )


def test_empty_sheet(tmp_path):
    rendered = sheetpeek.read(
        save(openpyxl.Workbook(), tmp_path / "e.xlsx"), platform="mac"
    ).layout("excel")
    assert rendered.rows == ()


# Row heights (captured, suite row_height; ARCHITECTURE.md §4.7) ------------------------------


def sheet_with(tmp_path, sheet_format, font_size=11):
    """A sheet of a number in A1, an empty <row r="3"/>, and <sheetFormatPr> as given."""
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "rows.xlsx")
    element = "" if sheet_format is None else f"<sheetFormatPr {sheet_format}/>"
    rewrite(
        path,
        SHEET1,
        replace_once('<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />', element),
    )
    rewrite(path, SHEET1, replace_once("</sheetData>", '<row r="3"/></sheetData>'))
    if font_size != 11:
        rewrite(
            path,
            "xl/styles.xml",
            lambda x: x.replace('<sz val="11" />', f'<sz val="{font_size}" />'),
        )
    return sheetpeek.read(path, ref="A1:A4", platform="mac")


@pytest.mark.parametrize(
    ("sheet_format", "size", "heights"),
    [
        # rows 1 (a value), 2 (no element), 3 (an empty element), 4 (no element)
        ('baseColWidth="8" defaultRowHeight="15"', 14, [19, 19, 19, 19]),  # stored 15 ignored
        ('baseColWidth="8" defaultRowHeight="30"', 11, [15, 15, 15, 15]),
        ('baseColWidth="8" defaultRowHeight="20" customHeight="1"', 14, [20, 20, 20, 20]),
        (None, 9, [12, 16, 12, 16]),  # without <sheetFormatPr>, unlisted rows are 16 pt
        (
            'baseColWidth="8" defaultRowHeight="15" thickTop="1" thickBottom="1"',
            11,
            [15, 17, 15, 17],
        ),
        (
            'baseColWidth="8" defaultRowHeight="20" customHeight="1" zeroHeight="1"',
            11,
            [15, 0, 15, 0],
        ),
    ],
)
def test_row_heights(tmp_path, sheet_format, size, heights):
    sheet = sheet_with(tmp_path, sheet_format, size)
    with _platform.measuring(sheet.platform):
        assert [row_height(row, sheet) for row in sheet.rows] == heights
    assert [r.listed for r in sheet.rows] == [True, False, True, False]


@pytest.mark.parametrize(
    ("family", "size", "height"),
    [
        ("Calibri", 18, 24),
        ("Calibri", 24, 31),
        ("Aptos Narrow", 20, 27),
        ("Calibri", 13, 17),
        ("Calibri", 19, 25),
        ("Calibri", 12.5, 17),
        ("Times New Roman", 14, 18),
    ],
)
def test_row_heights_of_larger_normal_fonts(tmp_path, family, size, height):
    # xlformat's row model at the font's own size: 19 and 12.5 as Excel fits them as a cell font
    # (row_height_fitted, "plain 19" and "plain 12.5"), and Times New Roman 14's own row
    # ("Times New Roman 14").
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "rows.xlsx")
    rewrite(
        path,
        "xl/styles.xml",
        lambda x: (
            x.replace('<sz val="11" />', f'<sz val="{size}" />', 1)
            .replace('<name val="Calibri" />', f'<name val="{family}" />', 1)
            .replace('<scheme val="minor" />', "", 1)
        ),  # else the theme's minor font names it
    )
    sheet = sheetpeek.read(path, ref="A1:A2", platform="mac")
    with _platform.measuring(sheet.platform):
        assert [row_height(row, sheet) for row in sheet.rows] == [height, height]


def test_zero_height_rows_are_hidden(tmp_path):
    sheet = sheet_with(tmp_path, 'baseColWidth="8" defaultRowHeight="15" zeroHeight="1"')
    assert [r.index for r in sheet.layout("excel").rows] == [1, 3]
    assert [r.index for r in sheet.layout("characters").rows] == [1, 3]
    assert [r.index for r in sheet.layout("excel", show_hidden=True).rows] == [1, 2, 3, 4]


def test_stored_row_height_wins(tmp_path):
    path = raw_sheet(
        tmp_path / "ht.xlsx",
        '<sheetFormatPr defaultRowHeight="15"/>'
        "<sheetData><row r='1' ht='30' customHeight='1'><c r='A1'><v>1</v></c></row></sheetData>",
    )
    rendered = sheetpeek.read(path, platform="mac").layout("excel")
    assert rendered.rows[0].height_pt == 30
    assert rendered.rows[0].runs[0].y == 27


def _rows_sheet(tmp_path, sheet_format, rows):
    return raw_sheet(
        tmp_path / "autofit.xlsx",
        f"<sheetFormatPr {sheet_format}/><sheetData>{rows}</sheetData>",
    )


AUTOFIT_ROWS = (
    "<row r='1' ht='30'><c r='A1'><v>1</v></c></row>"  # an ht without customHeight
    "<row r='2' ht='30' customHeight='1'><c r='A2'><v>1</v></c></row>"
    "<row r='3'><c r='A3'><v>1</v></c></row>"  # no ht
    "<row r='4' ht='0'/>"
    "<row r='5' ht=''/>"
    "<row r='6' ht='0' customHeight='1'/>"
)


def test_a_height_without_custom_height_reads_as_autofit(tmp_path):
    path = _rows_sheet(tmp_path, 'baseColWidth="8" defaultRowHeight="15"', AUTOFIT_ROWS)
    rows = sheetpeek.read(path, ref="A1:A6", platform="mac").rows
    assert [(r.height, r.autofit) for r in rows] == [
        (None, True),
        (30.0, False),
        (None, False),
        (None, True),  # a height of 0 without customHeight does not hide the row
        (None, True),
        (0.0, False),
    ]


@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_an_autofit_row_is_fitted_under_a_custom_default_height(tmp_path, platform):
    # Captured (refit_rows, group customsheet): Excel fits a row whose ht has no customHeight,
    # where a row with no ht takes the custom default of 20.
    wb = openpyxl.Workbook()
    for row in (1, 3):
        wb.active.cell(row, 1, 1).font = Font(name="Calibri", sz=28, family=2)
    wb.active.row_dimensions[1].height = 30
    wb.active.row_dimensions[2].height = 30
    path = save(wb, tmp_path / "custom_default.xlsx")
    rewrite(
        path,
        SHEET1,
        lambda x: replace_once(
            '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />',
            '<sheetFormatPr baseColWidth="8" defaultRowHeight="20" customHeight="1" />',
        )(x).replace(' customHeight="1">', ">"),
    )
    rendered = sheetpeek.read(path, ref="A1:A3", platform=platform).layout("excel")
    fit = 37.0 if platform == "mac" else 36.0
    default = 20.0 if platform == "mac" else 19.5
    assert [r.height_pt for r in rendered.rows] == [fit, 15.0, default]


def test_an_autofit_row_has_no_height():
    assert sheetpeek.Row(1, None, autofit=True).autofit
    with pytest.raises(ValueError, match="autofit"):
        sheetpeek.Row(1, 30.0, autofit=True)
