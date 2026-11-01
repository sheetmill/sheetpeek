"""Reading cells, styles and geometry into a Sheet."""

import io
from pathlib import Path

import openpyxl
import pytest
import xlformat
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils.datetime import CALENDAR_MAC_1904

import sheetpeek
import sheetpeek._read
from sheetpeek import CellRange
from sheetpeek._read import DATE_FALLBACK_WARNING, disable_date_conversion
from tests.workbooks import (
    SHEET1,
    raw_sheet,
    replace_once,
    replace_sheet_data,
    rewrite,
    save,
    set_cached,
)

# A double whose hh:mm:ss display changes after openpyxl's millisecond date conversion
# (ARCHITECTURE.md §5.3): 09:04:32 from the serial, 09:04:33 from the datetime.
CLOCK_SERIAL = 28139.378153931324


@pytest.fixture
def dated(tmp_path):
    """Date-formatted cells whose stored numbers are exact serials."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for row, fmt in enumerate(["yyyy-mm-dd", "hh:mm:ss", "[h]:mm:ss", "General"], 1):
        ws.cell(row, 1, 0).number_format = fmt
    path = save(wb, tmp_path / "dated.xlsx")
    for row, serial in enumerate(["60", repr(CLOCK_SERIAL), "1.5", "60"], 1):
        set_cached(path, f"A{row}", serial)
    return path


def test_fast_path_is_available():
    # Pins the private openpyxl detail sheetpeek relies on; if this fails on a new openpyxl,
    # dates go through the fallback and the reader warns.
    wb = openpyxl.Workbook()
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    assert disable_date_conversion(openpyxl.load_workbook(stream, read_only=True)) is True


def test_dates_stay_serials(dated):
    sheet = sheetpeek.read(dated)
    assert [sheet.cell(r, 1).value for r in range(1, 5)] == [60, CLOCK_SERIAL, 1.5, 60]
    assert {sheet.cell(r, 1).kind for r in range(1, 5)} == {"number"}
    assert sheet.warnings == ()


def test_date_fallback(dated, monkeypatch):
    real = sheetpeek._read.load_workbook

    def changed_private_attributes(*args, **kwargs):
        # As if a later openpyxl kept its date formats in another type: still used by openpyxl
        # itself, no longer recognized by sheetpeek.
        workbook = real(*args, **kwargs)
        workbook._date_formats = dict.fromkeys(workbook._date_formats)
        return workbook

    monkeypatch.setattr(sheetpeek._read, "load_workbook", changed_private_attributes)
    sheet = sheetpeek.read(dated)
    # Serial 60 comes back as 59 and the time to the millisecond: the loss the fast path avoids.
    assert sheet.cell(1, 1).value == 59
    assert sheet.cell(2, 1).value == pytest.approx(CLOCK_SERIAL, abs=0.0005 / 86400)
    assert sheet.cell(2, 1).value != CLOCK_SERIAL
    assert sheet.cell(3, 1).value == 1.5
    assert sheet.cell(4, 1).value == 60  # General: never a date
    assert sheet.warnings == (DATE_FALLBACK_WARNING,)


def test_disable_date_conversion_needs_both_sets():
    class Partial:
        def __init__(self):
            self._date_formats = {1}

    partial = Partial()
    assert disable_date_conversion(partial) is False
    assert partial._date_formats == {1}


def test_value_kinds(tmp_path):
    path = raw_sheet(
        tmp_path / "kinds.xlsx",
        "<sheetData><row r='1'>"
        "<c r='A1'><v>42</v></c>"
        "<c r='B1'><v>0.5</v></c>"
        "<c r='C1' t='b'><v>1</v></c>"
        "<c r='D1' t='e'><f>1/0</f><v>#DIV/0!</v></c>"
        "<c r='E1' t='inlineStr'><is><t>inline</t></is></c>"
        "<c r='F1' t='str'><f>\"a\"&amp;\"b\"</f><v>ab</v></c>"
        "<c r='G1' t='d'><v>2023-01-01T06:00:00</v></c>"
        "<c r='H1'/>"
        "<c r='I1' t='e'><v>#N/A</v></c>"
        "<c r='J1' t='str'><f>\"\"</f><v></v></c>"
        "</row></sheetData>",
    )
    cells = sheetpeek.read(path).cells[0]
    assert [(c.value, c.kind) for c in cells] == [
        (42, "number"),
        (0.5, "number"),
        (True, "bool"),
        ("#DIV/0!", "error"),
        ("inline", "text"),
        ("ab", "text"),
        (44927.25, "number"),
        (None, "empty"),
        ("#N/A", "error"),
        ("", "text"),  # a formula returning "": openpyxl says empty, Excel treats it as text
    ]
    assert cells[3].value == xlformat.ExcelError.DIV0
    assert isinstance(cells[3].value, xlformat.ExcelError)
    assert type(cells[5].value) is str


def test_error_values(tmp_path):
    # A code this xlformat does not list still reads as an error; text not shaped like an error
    # code (a malformed file) reads as text rather than failing.
    path = raw_sheet(
        tmp_path / "errors.xlsx",
        "<sheetData><row r='1'>"
        "<c r='A1' t='e'><v>#SPILL!</v></c>"
        "<c r='B1' t='e'><v>#NEWCODE!</v></c>"
        "<c r='C1' t='e'><v>oops</v></c>"
        "</row></sheetData>",
    )
    cells = sheetpeek.read(path).cells[0]
    assert [(c.value, c.kind) for c in cells] == [
        ("#SPILL!", "error"),
        ("#NEWCODE!", "error"),
        ("oops", "text"),
    ]
    assert cells[0].value == xlformat.ExcelError.SPILL
    assert isinstance(cells[1].value, xlformat.ExcelError)
    assert type(cells[2].value) is str


@pytest.mark.parametrize(
    ("dimension", "label"),
    [
        ("", "missing"),
        ('<dimension ref="A1"/>', "too small"),
        ('<dimension ref="A1:Z99"/>', "too large"),
    ],
)
def test_used_range_ignores_dimension(tmp_path, dimension, label):
    path = raw_sheet(
        tmp_path / f"{label}.xlsx",
        f"{dimension}<sheetData>"
        "<row r='2'><c r='B2'><v>1</v></c></row>"
        "<row r='5'><c r='D5'><v>2</v></c></row>"
        "</sheetData>",
    )
    sheet = sheetpeek.read(path)
    assert sheet.used_range == sheet.ref == CellRange.parse("B2:D5")
    assert len(sheet.cells) == 4
    assert {len(row) for row in sheet.cells} == {3}
    assert sheet.cell(2, 2).value == 1
    assert sheet.cell(5, 4).value == 2


def test_ref_beyond_the_data(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["B2"] = "x"
    sheet = sheetpeek.read(save(wb, tmp_path / "s.xlsx"), ref="A1:D6")
    assert (len(sheet.cells), len(sheet.cells[0])) == (6, 4)
    assert sheet.cell(2, 2).value == "x"
    assert sheet.cell(6, 4).kind == "empty"
    assert sheet.cells[0][0].style.number_format == "General"
    with pytest.raises(IndexError):
        sheet.cell(7, 1)


def test_empty_sheet(tmp_path):
    sheet = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "empty.xlsx"))
    assert sheet.ref is None
    assert sheet.used_range is None
    assert sheet.cells == ()
    assert sheet.columns == ()
    assert sheet.rows == ()


def test_geometry(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["E6"] = 2
    ws.merge_cells("A2:C3")
    ws.merge_cells("H1:H2")  # outside the range read below
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["D"].hidden = True
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[5].hidden = True
    sheet = sheetpeek.read(save(wb, tmp_path / "g.xlsx"), ref="A1:E6")
    assert sheet.merges == (CellRange.parse("A2:C3"),)
    assert [(c.letter, c.width, c.hidden) for c in sheet.columns] == [
        ("A", None, False),
        ("B", 20.0, False),
        ("C", None, False),
        ("D", 13.0, True),  # openpyxl writes a width with every hidden column
        ("E", None, False),
    ]
    assert [(r.height, r.hidden) for r in sheet.rows] == [
        (None, False),
        (None, False),
        (30.0, False),
        (None, False),
        (None, True),
        (None, False),
    ]
    assert sheet.base_column_width == 8
    assert sheet.default_row_height == 15


def test_styles(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "styled"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, italic=True, underline="double",
                         strike=True, color="FF00FF00")  # fmt: skip
    ws["A1"].fill = PatternFill("solid", fgColor="FFFFFF00")
    ws["A1"].alignment = Alignment(horizontal="centerContinuous", vertical="top", wrap_text=True,
                                   indent=2)  # fmt: skip
    ws["B1"] = "theme"
    ws["B1"].font = Font(color=Color(theme=4, tint=0.4))  # no name or size: inherits Normal
    ws["C1"] = "indexed"
    ws["C1"].font = Font(color=Color(indexed=10))
    ws["D1"] = "plain"
    ws["E1"] = 1234.5
    ws["E1"].number_format = "#,##0.00"
    sheet = sheetpeek.read(save(wb, tmp_path / "styles.xlsx"))
    a1, b1, c1, d1, e1 = sheet.cells[0]
    assert a1.font == sheetpeek.Font("Arial", 14.0, True, True, "double", True, "00FF00")
    assert a1.fill == "FFFF00"
    assert (a1.style.h_align, a1.style.v_align, a1.style.wrap, a1.style.indent) == (
        "center_across",
        "top",
        True,
        2,
    )
    assert (b1.font.name, b1.font.size, b1.font.color) == ("Calibri", 11.0, "95B3D7")
    assert c1.font.color == "FF0000"
    # The default font's color is theme 1 ("Text 1"): the default foreground, like automatic.
    assert d1.font.color is None
    assert (d1.fill, d1.style.h_align, d1.style.v_align) == (None, "general", "bottom")
    assert e1.number_format == "#,##0.00"
    assert sheet.normal_font == sheetpeek.Font("Calibri", 11.0)


def test_styles_are_shared(tmp_path):
    wb = openpyxl.Workbook()
    for row in range(1, 4):
        wb.active.cell(row, 1, row).font = Font(bold=True)
    sheet = sheetpeek.read(save(wb, tmp_path / "shared.xlsx"))
    assert len({id(row[0].style) for row in sheet.cells}) == 1


def test_normal_font_from_the_workbook(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "arial.xlsx")
    rewrite(path, "xl/styles.xml", replace_once('<name val="Calibri"', '<name val="Arial"'))
    rewrite(path, "xl/styles.xml", replace_once('<sz val="11"', '<sz val="10"'))
    sheet = sheetpeek.read(path)
    assert sheet.normal_font == sheetpeek.Font("Arial", 10.0)
    assert sheet.cell(1, 1).font.name == "Arial"


def test_hyperlinks_reach_cells(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "site"
    ws["A1"].hyperlink = "https://example.com/"
    ws["B1"] = "plain"
    sheet = sheetpeek.read(save(wb, tmp_path / "links.xlsx"))
    assert sheet.cell(1, 1).hyperlink == "https://example.com/"
    assert sheet.cell(1, 2).hyperlink is None


def test_uncached_formulas_warn(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "=1+1"
    sheet = sheetpeek.read(save(wb, tmp_path / "formula.xlsx"))
    assert sheet.cell(1, 1).kind == "empty"
    assert len(sheet.warnings) == 1
    assert "1 formula cell(s) have no cached value" in sheet.warnings[0]
    assert "first: A1" in sheet.warnings[0]


def test_sheet_selection(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "Hidden"
    wb.active.sheet_state = "hidden"
    wb.create_sheet("Shown")["A1"] = "here"
    wb.create_sheet("Secret").sheet_state = "veryHidden"
    path = save(wb, tmp_path / "sheets.xlsx")
    with sheetpeek.open(path) as book:
        assert book.sheet_names == ["Hidden", "Shown", "Secret"]
        assert book.sheet().name == "Shown"  # the active sheet is hidden
        assert book.sheet("Hidden").state == "hidden"
        assert book.sheet("Secret").state == "very_hidden"
        with pytest.raises(KeyError, match="no sheet named 'Missing'"):
            book.sheet("Missing")


def test_file_object_and_1904(tmp_path):
    wb = openpyxl.Workbook()
    wb.epoch = CALENDAR_MAC_1904
    wb.active["A1"] = 0
    wb.active["A1"].number_format = "yyyy-mm-dd"
    path = save(wb, tmp_path / "mac.xlsx")
    with path.open("rb") as handle:
        sheet = sheetpeek.read(handle)
    assert sheet.date1904 is True
    assert sheet.cell(1, 1).value == 0


def test_sheet_data_replacement_helper(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "h.xlsx")
    rewrite(
        path,
        SHEET1,
        replace_sheet_data("<sheetData><row r='3'><c r='C3'><v>7</v></c></row></sheetData>"),
    )
    assert sheetpeek.read(path).cell(3, 3).value == 7


def test_a_workbook_saved_by_excel():
    # Saved by Excel 16.114 for Mac from formulas (capture/excel_saved.py), so every value is
    # Excel's own cached result and the XML is Excel's own.
    path = Path(__file__).with_name("files") / "excel_formulas.xlsx"
    sheet = sheetpeek.read(path)
    column = [(c.value, c.kind) for c in (row[0] for row in sheet.cells)]
    assert column == [
        ("", "text"),  # ="" is written <c t="str"><f>""</f><v/></c>
        (2, "number"),
        ("ab", "text"),
        ("#DIV/0!", "error"),
        (0, "number"),
        (True, "bool"),
        (" ", "text"),
        ("#N/A", "error"),
        (60, "number"),  # =DATE(1900,2,29), the serial openpyxl would turn into 1900-02-28
    ]
    assert sheet.warnings == ()
    assert sheet.layout("none").cell(9, 1).text == "1900-02-29"
