"""Layout in characters and none modes: what each column draws.

Rules verified against Excel cite ARCHITECTURE.md §4; characters mode's own choices are named as
such.
"""

import dataclasses
import math

import openpyxl
import pytest
import xlformat
from openpyxl.styles import Alignment, Font, PatternFill

import sheetpeek
from sheetpeek._layout import _Fonts
from sheetpeek._notices import Collector
from tests.workbooks import SHEET1, raw_sheet, replace_once, rewrite, save

LONG = "a long text value"  # 17 characters; default columns are 8 wide


def drawn(rendered, separator="|"):
    """The sheet as text, with ``separator`` where no text crosses between columns."""
    lines = []
    for row in rendered.rows:
        for line in range(row.height):
            parts = []
            for p, cell in enumerate(row.cells):
                parts.append(cell.lines[line])
                if p < len(row.gaps):
                    crossing = row.gaps[p].lines[line]
                    parts.append(separator if crossing is None else crossing)
            lines.append("".join(parts).rstrip())
    return lines


def lay_out(tmp_path, build, ref=None, mode="characters", **options):
    wb = openpyxl.Workbook()
    build(wb.active)
    sheet = sheetpeek.read(save(wb, tmp_path / "book.xlsx"), ref=ref)
    return sheet.layout(mode, **options)


def test_default_width_is_eight_characters(tmp_path):
    rendered = lay_out(tmp_path, lambda ws: ws.cell(1, 1, 1), ref="A1:C1")
    assert [c.width for c in rendered.columns] == [8, 8, 8]
    assert rendered.width == 26


def test_custom_widths_drop_the_padding(tmp_path):
    def build(ws):
        ws["A1"] = 1
        ws.column_dimensions["A"].width = 20  # 120 pt; Excel for Mac's dialog shows 19.17
        ws.column_dimensions["B"].width = 9.140625  # 55 pt, 8.33 (Windows Excel's 8.43)

    rendered = lay_out(tmp_path, build, ref="A1:B1")
    assert [c.width for c in rendered.columns] == [19, 8]
    assert [c.width for c in lay_out(tmp_path, build, ref="A1:B1", scale=0.5).columns] == [10, 4]
    assert [c.width for c in lay_out(tmp_path, build, ref="A1:B1", max_width=5).columns] == [5, 5]


@pytest.mark.parametrize(
    ("file_width", "characters"),
    # file width -> xlformat's dialog units -> characters (ARCHITECTURE.md §2.2)
    [(None, 8), (20, 19), (9.140625, 8), (12.71, 12), (14.71, 14), (3.71, 3)],
)
def test_widths_from_xlformat(tmp_path, file_width, characters):
    def build(ws):
        ws["A1"] = 1
        if file_width is not None:
            ws.column_dimensions["A"].width = file_width

    assert lay_out(tmp_path, build, ref="A1").columns[0].width == characters


def test_default_column_width_from_the_file(tmp_path):
    # <sheetFormatPr defaultColWidth> is a file width: 12.71 is 12.17 units, 12 characters.
    path = save(openpyxl.Workbook(), tmp_path / "book.xlsx")
    rewrite(
        path,
        SHEET1,
        replace_once("<sheetFormatPr ", '<sheetFormatPr defaultColWidth="12.71" '),
    )
    sheet = sheetpeek.read(path, ref="A1:B1")
    assert [c.width for c in sheet.layout().columns] == [12, 12]


def test_unmeasured_normal_font(tmp_path):
    # Decided (ARCHITECTURE.md §6.2): column widths use Calibri at the same size, with a warning
    # in characters mode, which measures the Normal font; none mode measures nothing.
    path = save(openpyxl.Workbook(), tmp_path / "book.xlsx")
    rewrite(path, "xl/styles.xml", replace_once('<name val="Calibri"', '<name val="Arial"'))
    rewrite(path, "xl/styles.xml", replace_once('<scheme val="minor" />', ""))  # §4.14
    sheet = sheetpeek.read(path, ref="A1")
    assert sheet.normal_font.name == "Arial"
    rendered = sheet.layout()
    assert rendered.columns[0].width == 8
    assert rendered.warnings == (
        "font 'Arial' has no glyph table in xlformat; measured as Calibri at the same size",
    )
    assert sheet.layout("none").warnings == ()


def test_spill_right_over_empty_and_fill_only_neighbors(tmp_path):
    # A neighbor holding only a fill does not stop spilled text (checked by eye, §4.1).
    def build(ws):
        ws["A1"] = LONG
        ws["A2"] = LONG
        ws["B2"].fill = PatternFill("solid", fgColor="FFFFFF00")

    rendered = lay_out(tmp_path, build, ref="A1:C2")
    # 17 characters fill two 8-character columns and the gap between them.
    assert drawn(rendered) == [f"{LONG}|", f"{LONG}|"]
    a2, b2, _ = rendered.rows[1].cells
    assert a2.spill == (0, 1)
    assert (b2.source, b2.fill, b2.lines) == ((2, 1), "FFFF00", ("xt value",))
    assert a2.fill is None  # the spilling cell's fill stays in its own column
    assert rendered.rows[1].gaps[0].lines == ("e",)
    assert rendered.rows[1].gaps[1].lines == (None,)


@pytest.mark.parametrize(
    ("neighbor", "label"),
    [(5, "a number"), (" ", "a space"), ("x", "text")],
)
def test_spill_stops_at_a_value(tmp_path, neighbor, label):
    def build(ws):
        ws["A1"] = LONG
        ws["B1"] = neighbor

    first = drawn(lay_out(tmp_path, build, ref="A1:C1"))[0]
    assert first.startswith("a long t|")


def test_spill_stops_at_a_formula_returning_the_empty_string(tmp_path):
    path = raw_sheet(
        tmp_path / "f.xlsx",
        "<sheetData><row r='1'><c r='A1' t='inlineStr'><is><t>a long text value</t></is></c>"
        "<c r='B1' t='str'><f>\"\"</f><v></v></c><c r='C1'/></row></sheetData>",
    )
    assert drawn(sheetpeek.read(path).layout())[0] == "a long t|        |"


def test_spill_left_and_both_ways(tmp_path):
    # Directions checked by eye (§4.1); how far each reaches is not yet measured.
    def build(ws):
        ws["C1"] = LONG
        ws["C1"].alignment = Alignment(horizontal="right")
        ws["C2"] = LONG
        ws["C2"].alignment = Alignment(horizontal="center")

    rendered = lay_out(tmp_path, build, ref="A1:E2")
    assert drawn(rendered) == [
        "        |a long text value|        |",
        "        |    a long text value     |",
    ]
    assert rendered.cell(1, 3).spill == (1, 0)
    assert rendered.cell(2, 3).spill == (1, 1)


def test_spill_crosses_hidden_columns(tmp_path):
    def build(ws):
        ws["A1"] = LONG
        ws.column_dimensions["B"].hidden = True

    rendered = lay_out(tmp_path, build, ref="A1:D1")
    assert [c.index for c in rendered.columns] == [1, 3, 4]
    assert drawn(rendered) == [f"{LONG}|"]


def test_a_hidden_value_does_not_stop_spill(tmp_path):
    # Captured (layout suite, "hidden K holds a value"): hidden columns are crossed whatever
    # they hold.
    def build(ws):
        ws["A1"] = LONG
        ws["B1"] = "hidden"
        ws.column_dimensions["B"].hidden = True

    assert drawn(lay_out(tmp_path, build, ref="A1:C1")) == [LONG]


def test_centered_spill_takes_both_sides(tmp_path):
    # Captured: centered text takes a column on each side at once, even when only one side's
    # characters reach it (9 zeros in an 8-character column).
    def build(ws):
        ws["B1"] = "0" * 9
        ws["B1"].alignment = Alignment(horizontal="center")
        ws["B2"] = "0" * 9
        ws["B2"].alignment = Alignment(horizontal="center")
        ws["A2"] = "x"  # the left side is blocked

    rendered = lay_out(tmp_path, build, ref="A1:E2")
    assert rendered.cell(1, 2).spill == (1, 1)
    assert drawn(rendered)[0] == "        000000000         |        |"
    # Blocked on the left: the text is clipped there, and the right takes only what it needs.
    assert rendered.cell(2, 2).spill == (0, 0)
    assert drawn(rendered)[1] == "x       |00000000|        |        |"


def test_first_spill_wins_a_contested_cell(tmp_path):
    # Characters mode's choice: left to right, the first spill placed takes a free cell (§4.1).
    def build(ws):
        ws["A1"] = LONG
        ws["C1"] = "right side text"
        ws["C1"].alignment = Alignment(horizontal="right")

    rendered = lay_out(tmp_path, build, ref="A1:C1")
    assert drawn(rendered) == ["a long text value|ide text"]
    assert rendered.cell(1, 3).spill == (0, 0)


def test_spill_past_the_used_range_adds_columns(tmp_path):
    # range_edge="extend": empty columns at the default width, as far as the text reaches.
    text = LONG + " and more!"  # 27 characters: three 8-character columns and two gaps hold 26
    rendered = lay_out(tmp_path, lambda ws: ws.cell(1, 1, text))
    assert drawn(rendered) == [text]
    assert [(c.letter, c.width, c.added) for c in rendered.columns] == [
        ("A", 8, False),
        ("B", 8, True),
        ("C", 8, True),
        ("D", 8, True),
    ]
    # The last character falls in the gap after C; D is kept so the gap can be drawn.
    assert rendered.rows[0].gaps[2].lines == ("!",)
    assert rendered.cell(1, 4).source is None
    assert rendered.cell(1, 1).spill == (0, 3)
    fits = lay_out(tmp_path, lambda ws: ws.cell(1, 1, text[:-1]))
    assert [c.letter for c in fits.columns] == ["A", "B", "C"]


def test_no_columns_added_without_spill(tmp_path):
    rendered = lay_out(tmp_path, lambda ws: ws.cell(1, 1, "short"))
    assert [c.letter for c in rendered.columns] == ["A"]


def test_spill_stops_where_an_explicit_range_cuts_the_data(tmp_path):
    # Columns past the range were not read and may hold values, so nothing is added there.
    def build(ws):
        ws["A1"] = LONG
        ws["C2"] = "data beyond the range"

    rendered = lay_out(tmp_path, build, ref="A1:A2")
    assert drawn(rendered) == ["a long t", ""]
    assert [c.added for c in rendered.columns] == [False]


def test_spill_left_past_the_range(tmp_path):
    def build(ws):
        ws["C1"] = LONG
        ws["C1"].alignment = Alignment(horizontal="right")

    rendered = lay_out(tmp_path, build)
    assert [(c.letter, c.added) for c in rendered.columns] == [("B", True), ("C", False)]
    assert drawn(rendered) == [LONG]


def test_unknown_range_edge(tmp_path):
    sheet = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "e.xlsx"))
    with pytest.raises(ValueError, match="only 'extend' is available"):
        sheet.layout(range_edge="clip")


def test_spill_edge_is_range_edge_under_its_former_name(tmp_path):
    sheet = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "e.xlsx"))
    with pytest.warns(DeprecationWarning, match="range_edge"):
        assert sheet.layout(spill_edge="extend") == sheet.layout()
    with (
        pytest.warns(DeprecationWarning, match="spill_edge"),
        pytest.raises(ValueError, match="only 'extend' is available"),
    ):
        sheet.layout(spill_edge="clip")


def test_numbers_overflow_and_do_not_spill(tmp_path):
    def build(ws):
        ws["A1"] = 123456789
        ws["A1"].number_format = "#,##0.00"
        ws["A2"] = 1234567890123.0  # General reduces before it overflows
        ws["A3"] = True

    rendered = lay_out(tmp_path, build, ref="A1:B3")
    assert drawn(rendered) == ["########|", "1.23E+12|", "  TRUE  |"]
    assert rendered.cell(1, 1).overflow
    assert rendered.cell(1, 1).full_text == "123,456,789.00"


def test_errors_overflow_to_hashes(tmp_path):
    # Route A: at width 3 an error shows ##-style overflow, like a number (§4.2).
    path = raw_sheet(
        tmp_path / "e.xlsx",
        "<cols><col min='1' max='1' width='3.7109375' customWidth='1'/></cols>"
        "<sheetData><row r='1'><c r='A1' t='e'><v>#DIV/0!</v></c>"
        "<c r='B1' t='e'><v>#N/A</v></c></row></sheetData>",
    )
    rendered = sheetpeek.read(path).layout()
    a1, b1 = rendered.rows[0].cells
    assert (a1.text, a1.overflow, a1.align) == ("###", True, "center")
    assert (b1.text, b1.overflow, b1.lines) == ("#N/A", False, ("  #N/A  ",))
    none = sheetpeek.read(path).layout("none")
    assert none.cell(1, 1).lines == ("#DIV/0!",)


def test_default_alignment(tmp_path):
    def build(ws):
        ws["A1"] = 42
        ws["B1"] = "text"
        ws["C1"] = True
        ws["D1"] = 12345
        ws["D1"].number_format = "@"
        ws["E1"] = 12345
        ws["E1"].number_format = '@" units"'

    rendered = lay_out(tmp_path, build, ref="A1:E1")
    aligns = [c.align for c in rendered.rows[0].cells]
    # D1 and E1: a number under a format with only a text section is left-aligned (captured for
    # "@", §4.3; xlformat's general rule) and shows as General.
    assert aligns == ["right", "left", "center", "left", "left"]
    assert drawn(rendered) == ["      42|text    |  TRUE  |12345   |12345"]


def test_explicit_alignment_and_indent(tmp_path):
    def build(ws):
        ws["A1"] = 42
        ws["A1"].alignment = Alignment(horizontal="left")
        ws["B1"] = "ab"
        ws["B1"].alignment = Alignment(horizontal="right", indent=1)
        ws["C1"] = "ab"
        ws["C1"].alignment = Alignment(horizontal="left", indent=2)
        ws["D1"] = "ab"
        ws["D1"].alignment = Alignment(horizontal="fill")

    rendered = lay_out(tmp_path, build, ref="A1:D1")
    assert drawn(rendered) == ["42      |     ab |  ab    |abababab"]


def test_section_color(tmp_path):
    def build(ws):
        ws["A1"] = -5
        ws["A1"].number_format = "[Red]0;[Blue]-0"
        ws["A1"].font = Font(color="FF00FF00")
        ws["B1"] = 5
        ws["B1"].font = Font(color="FF00FF00")

    rendered = lay_out(tmp_path, build, ref="A1:B1")
    a1, b1 = rendered.rows[0].cells
    assert (a1.format_color, a1.color) == ("0000FF", "0000FF")
    assert (b1.format_color, b1.color) == (None, "00FF00")


def test_refused_format_falls_back_to_general(tmp_path):
    def build(ws):
        ws["A1"] = 1.5
        ws["A2"] = 2.5
        ws["A3"] = 3

    wb = openpyxl.Workbook()
    build(wb.active)
    for cell in ("A1", "A2"):
        wb.active[cell].number_format = "0;0;0;0;0"  # five sections: Excel refuses the file
    sheet = sheetpeek.read(save(wb, tmp_path / "bad.xlsx"))
    rendered = sheet.layout()
    assert [r.cells[0].text for r in rendered.rows] == ["1.5", "2.5", "3"]
    assert rendered.warnings == (
        "number format '0;0;0;0;0' is one Excel refuses to open; shown as General",
    )


def test_merge_uses_the_merged_width(tmp_path):
    # Route A: a number too wide for one column shows in full merged over three (§4.4).
    def build(ws):
        ws["A1"] = 123456789
        ws["A1"].number_format = "#,##0.00"
        ws.merge_cells("A1:C1")
        ws["A1"].fill = PatternFill("solid", fgColor="FFFFFF00")

    rendered = lay_out(tmp_path, build, ref="A1:D1")
    assert drawn(rendered) == ["            123,456,789.00|"]
    a1, b1, c1, d1 = rendered.rows[0].cells
    assert a1.is_merge_anchor
    assert not b1.is_merge_anchor
    assert a1.text == "123,456,789.00"
    assert (b1.fill, c1.fill, d1.fill) == ("FFFF00", "FFFF00", None)
    assert rendered.rows[0].gaps[0].fill == "FFFF00"
    assert rendered.rows[0].gaps[2].fill is None


def test_merge_does_not_spill(tmp_path):
    def build(ws):
        ws["A1"] = "a very long text inside a merge"
        ws.merge_cells("A1:B1")

    assert drawn(lay_out(tmp_path, build, ref="A1:D1")) == ["a very long text |        |"]


def test_merge_across_a_hidden_column(tmp_path):
    def build(ws):
        ws["A1"] = "merged across"
        ws.merge_cells("A1:C1")
        ws.column_dimensions["B"].hidden = True

    rendered = lay_out(tmp_path, build, ref="A1:C1")
    assert drawn(rendered) == ["merged across"]
    assert rendered.cell(1, 3).source == (1, 1)


def test_vertical_merge_places_text_by_vertical_alignment(tmp_path):
    def build(ws):
        for col, v_align in zip("ABC", ["top", "center", "bottom"], strict=True):
            ws[f"{col}1"] = v_align
            ws[f"{col}1"].alignment = Alignment(vertical=v_align)
            ws.merge_cells(f"{col}1:{col}3")

    assert drawn(lay_out(tmp_path, build, ref="A1:C3")) == [
        "top     |        |",
        "        |center  |",
        "        |        |bottom",
    ]


def test_center_across_selection(tmp_path):
    def build(ws):
        ws["A1"] = "centered"
        for col in "ABC":
            ws[f"{col}1"].alignment = Alignment(horizontal="centerContinuous")

    rendered = lay_out(tmp_path, build, ref="A1:D1")
    assert drawn(rendered) == ["         centered         |"]
    assert rendered.cell(1, 2).source == (1, 1)


def test_wrapped_text_grows_the_row(tmp_path):
    def build(ws):
        ws["A1"] = "wrap this text"
        ws["A1"].alignment = Alignment(wrap_text=True)
        ws["B1"] = 7
        ws["C1"] = "top"
        ws["C1"].alignment = Alignment(vertical="top")

    rendered = lay_out(tmp_path, build, ref="A1:C1")
    assert rendered.rows[0].height == 3
    assert drawn(rendered) == ["wrap    |        |top", "this    |        |", "text    |       7|"]
    assert rendered.cell(1, 1).spill == (0, 0)


def test_hidden_rows_and_columns(tmp_path):
    def build(ws):
        ws["A1"] = 1
        ws["B2"] = 2
        ws["C3"] = 3
        ws.row_dimensions[2].hidden = True
        ws.column_dimensions["B"].hidden = True

    rendered = lay_out(tmp_path, build)
    assert [r.index for r in rendered.rows] == [1, 3]
    assert [c.index for c in rendered.columns] == [1, 3]
    shown = lay_out(tmp_path, build, show_hidden=True)
    assert [r.index for r in shown.rows] == [1, 2, 3]
    assert [(c.index, c.hidden) for c in shown.columns] == [(1, False), (2, True), (3, False)]


def test_zero_width_column_is_hidden(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active["B1"] = 2
    wb.active.column_dimensions["B"].width = 5
    path = save(wb, tmp_path / "zero.xlsx")
    # openpyxl does not write a width of 0, so the file is edited.
    rewrite(path, SHEET1, replace_once('width="5"', 'width="0"'))
    rendered = sheetpeek.read(path).layout()
    assert [c.index for c in rendered.columns] == [1]


def test_none_mode(tmp_path):
    def build(ws):
        ws["A1"] = 123456789
        ws["A1"].number_format = "#,##0.00"
        ws["B1"] = LONG
        ws["A2"] = "merged"
        ws.merge_cells("A2:B2")

    rendered = lay_out(tmp_path, build, mode="none")
    assert rendered.gap == 0
    assert [c.width for c in rendered.columns] == [None, None]
    assert [[c.lines for c in row.cells] for row in rendered.rows] == [
        [("123,456,789.00",), (LONG,)],
        [("merged",), ("",)],
    ]
    assert all(row.gaps == () for row in rendered.rows)
    assert rendered.cell(2, 2).merge == sheetpeek.CellRange.parse("A2:B2")


def test_empty_sheet_layout(tmp_path):
    rendered = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "e.xlsx")).layout()
    assert rendered.rows == ()
    assert rendered.columns == ()


def test_layout_rejects_unknown_modes(tmp_path):
    sheet = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "e.xlsx"))
    with pytest.raises(ValueError, match="unknown width mode"):
        sheet.layout("points")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="scale"):
        sheet.layout(scale=0)
    # decided: excel mode has no character widths to scale or cap
    with pytest.raises(ValueError, match="characters mode only"):
        sheet.layout("excel", scale=2)
    with pytest.raises(ValueError, match="characters mode only"):
        sheet.layout("excel", max_width=5)


def test_fonts_for_xlformat():
    # Decided (ARCHITECTURE.md §6.2): a family without a glyph table is measured as Calibri at
    # the same size and style, with one warning per family, only where the font is measured.
    notes = Collector("Sheet")
    fonts = _Fonts(notes)
    bold = sheetpeek.Font("Calibri", 11.0, bold=True, color="FF0000")
    assert fonts.get(bold) == xlformat.Font("Calibri", 11.0, bold=True)
    assert fonts.get(bold) is fonts.get(bold)
    assert fonts.get(sheetpeek.Font("aptos narrow", 10.0)) == xlformat.Font("aptos narrow", 10.0)
    assert notes.notices() == ()
    arial = sheetpeek.Font("Arial", 10.0, italic=True)
    assert fonts.get(arial, measured=False) == xlformat.Font("Calibri", 10.0, italic=True)
    assert notes.notices() == ()
    fonts.get(arial)
    fonts.get(sheetpeek.Font("Arial", 12.0))
    assert [n.message for n in notes.notices()] == [
        "font 'Arial' has no glyph table in xlformat; measured as Calibri at the same size"
    ]


def _with_value(sheet, value):
    """``sheet`` with A1's value replaced: infinity reaches the model only from code, since a
    file's number too large for a double reads as #NULL!."""
    cell = dataclasses.replace(sheet.cells[0][0], value=value)
    return dataclasses.replace(sheet, cells=((cell, *sheet.cells[0][1:]), *sheet.cells[1:]))


def test_a_number_too_large_for_a_double_reads_as_null(tmp_path):
    # As Excel reads it (xlio's capture value_forms).
    path = raw_sheet(
        tmp_path / "inf.xlsx",
        "<sheetData><row r='1'>"
        "<c r='A1'><v>1E400</v></c><c r='B1'><v>-1E400</v></c><c r='C1'><v>1</v></c>"
        "</row></sheetData>",
    )
    sheet = sheetpeek.read(path)
    assert [(c.value, c.kind) for c in sheet.cells[0]] == [
        ("#NULL!", "error"),
        ("#NULL!", "error"),
        (1, "number"),
    ]


def test_values_xlformat_cannot_format(tmp_path):
    # ARCHITECTURE.md §6.2: shown as their text, with one warning per format string.
    path = raw_sheet(
        tmp_path / "one.xlsx", "<sheetData><row r='1'><c r='A1'><v>1</v></c></row></sheetData>"
    )
    rendered = _with_value(sheetpeek.read(path), math.inf).layout()
    assert [c.text for c in rendered.rows[0].cells] == ["inf"]
    assert rendered.warnings == (
        "number format 'General' cannot show the value inf (first: A1; cannot format "
        "non-finite number inf); shown as its text",
    )


def test_refused_format_with_a_value_general_cannot_format(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active["A1"].number_format = "0;0;0;0;0"  # five sections: Excel refuses the file
    sheet = sheetpeek.read(save(wb, tmp_path / "bad.xlsx"))
    rendered = _with_value(sheet, math.inf).layout()
    assert rendered.cell(1, 1).text == "inf"
    assert rendered.warnings == (
        "number format '0;0;0;0;0' is one Excel refuses to open; shown as General",
        "number format 'General' cannot show the value inf (first: A1; cannot format "
        "non-finite number inf); shown as its text",
    )


def test_the_font_cache_still_warns_once_a_font_is_measured():
    from sheetpeek._layout import _Fonts
    from sheetpeek._model import Font
    from sheetpeek._notices import Collector

    notes = Collector("Sheet")
    fonts = _Fonts(notes)
    font = Font("No Such Typeface", 11.0)
    first = fonts.get(font, measured=False)  # cached without the check
    assert not [n for n in notes.notices() if n.code == "font.no_metrics"]
    assert fonts.get(font) is first
    assert [n.code for n in notes.notices()].count("font.no_metrics") == 1
    fonts.get(font)
    assert [n.code for n in notes.notices()].count("font.no_metrics") == 1
    assert fonts.row_height(font) == fonts.row_height(Font("No Such Typeface", 11.0))
