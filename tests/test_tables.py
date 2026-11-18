"""Tables and table styles: what the reader keeps (ARCHITECTURE.md §4.20).

The styles' colors resolve through the workbook's theme, here openpyxl's default (Office 2007's),
as Excel draws them (the screen suite ``tables``, ``tests/test_table_captures.py``).
"""

from __future__ import annotations

import pytest

import sheetpeek
from sheetpeek import CellRange, DifferentialFormat, Pattern, Side, Table, TableStyleInfo
from tests.workbooks import replace_once, rewrite, save

openpyxl = pytest.importorskip("openpyxl")
tables = pytest.importorskip("openpyxl.worksheet.table")

STYLES = "xl/styles.xml"
TABLE = "xl/tables/table1.xml"


def book(tmp_path, style="TableStyleMedium2", ref="A1:C4", **flags):
    wb = openpyxl.Workbook()
    ws = wb.active
    first = CellRange.parse(ref)
    for r in range(first.min_row, first.max_row + 1):
        for c in range(first.min_col, first.max_col + 1):
            ws.cell(r, c, f"h{c}" if r == first.min_row else r * c)
    table = tables.Table(displayName="Sales", ref=ref)
    if style is not None:
        table.tableStyleInfo = tables.TableStyleInfo(name=style, **flags)
    ws.add_table(table)
    return save(wb, tmp_path / "tables.xlsx")


def test_a_preset_is_read_through_the_theme(tmp_path):
    sheet = sheetpeek.read(book(tmp_path, showRowStripes=True))
    assert sheet.tables == (
        Table(
            "Sales",
            CellRange.parse("A1:C4"),
            1,
            0,
            TableStyleInfo("TableStyleMedium2", show_row_stripes=True),
        ),
    )
    style = sheet.table_style("TableStyleMedium2")
    assert style is not None
    header = style.element("headerRow")
    assert header is not None
    # Bold white on accent 1; what the <dxf> leaves out stays None.
    assert header.format == DifferentialFormat(
        bold=True, color="FFFFFF", fill=Pattern("solid", "4F81BD")
    )
    stripe = style.element("firstRowStripe")
    assert stripe is not None
    assert stripe.format.fill == Pattern("solid", "DCE6F1")  # accent 1 at tint 0.8
    whole = style.element("wholeTable")
    assert whole is not None
    assert whole.format.horizontal == whole.format.top == Side("thin", "95B3D7")
    assert whole.format.vertical is None


def test_the_workbooks_own_style_and_a_partial_font(tmp_path):
    styles = pytest.importorskip("openpyxl.styles")
    differential = pytest.importorskip("openpyxl.styles.differential")
    table_styles = pytest.importorskip("openpyxl.styles.table")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"], ws["A2"] = "a", "b", 1
    dxf = differential.DifferentialStyle(
        font=styles.Font(b=True, i=True), fill=styles.PatternFill(bgColor="FF00FF00")
    )
    index = wb._differential_styles.add(dxf)
    element = table_styles.TableStyleElement(type="headerRow", dxfId=index, size=3)
    own = table_styles.TableStyle(name="Mine", pivot=False, tableStyleElement=[element])
    wb._table_styles.tableStyle.append(own)
    table = tables.Table(displayName="Sales", ref="A1:B2")
    table.tableStyleInfo = tables.TableStyleInfo(name="Mine")
    ws.add_table(table)
    path = save(wb, tmp_path / "own.xlsx")
    rewrite(path, STYLES, replace_once('<b val="1" />', '<b val="0" />'))
    style = sheetpeek.read(path).table_style("Mine")
    assert style is not None
    (found,) = style.elements
    # <b val="0"/> sets bold off and <i/> italic on; nothing else is set. A patternFill without
    # a type is solid, in its bgColor. A size counts only for stripes; the header keeps 3.
    assert found.format == DifferentialFormat(
        bold=False, italic=True, fill=Pattern("solid", "00FF00")
    )
    assert found.size == 3


def test_no_style_info_no_header_and_a_table_outside_the_range(tmp_path):
    sheet = sheetpeek.read(book(tmp_path, style=None))
    (table,) = sheet.tables
    assert table.style is None
    assert sheet.table_styles == ()
    path = book(tmp_path)
    rewrite(path, TABLE, lambda text: text.replace('headerRowCount="1"', 'headerRowCount="0"'))
    (table,) = sheetpeek.read(path).tables
    assert (table.header_row_count, table.data_rows) == (0, (1, 4))
    assert sheetpeek.read(book(tmp_path), ref="F6:G7").tables == ()


def test_a_style_name_neither_preset_nor_own_is_kept_on_the_table_only(tmp_path):
    sheet = sheetpeek.read(book(tmp_path, style="Nowhere"))
    (table,) = sheet.tables
    assert table.style is not None
    assert table.style.name == "Nowhere"
    assert sheet.table_styles == ()
