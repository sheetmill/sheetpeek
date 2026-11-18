"""Tables in the layout: the style each table cell is displayed in (ARCHITECTURE.md §4.20).

The regions, the stacking and the own-format rule are unverified until the screen suite
``tables`` captures them. TableStyleMedium2 under openpyxl's default theme: the whole table thin
95B3D7 lines with black text, the header bold white on 4F81BD, the first row stripe and first
column stripe DCE6F1, the first and last column bold, the totals row bold over a double 4F81BD top.
"""

from __future__ import annotations

import pytest

import sheetpeek
from sheetpeek import Side
from tests.workbooks import rewrite, save

openpyxl = pytest.importorskip("openpyxl")
tables = pytest.importorskip("openpyxl.worksheet.table")

TABLE = "xl/tables/table1.xml"
BLUE, STRIPE, LINE = "4F81BD", "DCE6F1", Side("thin", "95B3D7")


def book(tmp_path, ref="A1:C5", style="TableStyleMedium2", edit=None, **flags):
    wb = openpyxl.Workbook()
    ws = wb.active
    rows = int(ref.split(":")[1][1:])
    for r in range(1, rows + 1):
        for c in range(1, 4):
            ws.cell(r, c, f"h{c}" if r == 1 else r * c)
    table = tables.Table(displayName="Sales", ref=ref)
    table.tableStyleInfo = tables.TableStyleInfo(name=style, **flags)
    ws.add_table(table)
    if edit is not None:
        edit(ws)
    return save(wb, tmp_path / "table.xlsx")


def displayed(path, mode="excel"):
    """{(row, column): display style} of every laid-out cell."""
    rendered = sheetpeek.read(path).layout(mode)
    return {(c.row, c.column): c.display_style for row in rendered.rows for c in row.cells}


def test_a_sheet_without_tables_displays_every_cell_in_its_own_style(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "a"
    rendered = sheetpeek.read(save(wb, tmp_path / "plain.xlsx")).layout("excel")
    assert all(c.display_style is c.cell.style for row in rendered.rows for c in row.cells)


def test_header_stripes_and_first_column(tmp_path):
    path = book(tmp_path, showRowStripes=True, showFirstColumn=True)
    found = displayed(path)
    header = found[(1, 2)]
    assert (header.fill, header.font.color, header.font.bold) == (BLUE, "FFFFFF", True)
    assert [found[(r, 2)].fill for r in range(2, 6)] == [STRIPE, None, STRIPE, None]
    assert found[(3, 1)].font.bold  # first column
    assert not found[(3, 2)].font.bold
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 2).style.fill is None  # the file's own style is kept
    characters = sheetpeek.read(path).layout()
    assert characters.rows[0].cells[1].fill == BLUE


def test_the_flags_turn_their_elements_off(tmp_path):
    found = displayed(book(tmp_path))
    assert [found[(r, 2)].fill for r in range(2, 6)] == [None] * 4
    assert not found[(3, 1)].font.bold


def test_row_stripes_count_visible_rows_only(tmp_path):
    def hide(ws):
        ws.row_dimensions[3].hidden = True

    found = displayed(book(tmp_path, ref="A1:C6", edit=hide, showRowStripes=True))
    # Data rows 2, 4, 5, 6 are visible: stripe, none, stripe, none. Counting hidden row 3
    # would make row 5 the second band.
    assert [found[(r, 2)].fill for r in (2, 4, 5, 6)] == [STRIPE, None, STRIPE, None]


def test_column_stripes_band_from_the_first_column_over_the_data_rows(tmp_path):
    found = displayed(book(tmp_path, showColumnStripes=True))
    assert [found[(3, c)].fill for c in (1, 2, 3)] == [STRIPE, None, STRIPE]
    assert found[(1, 1)].fill == BLUE  # the header is not striped


def test_the_cells_own_format_wins_property_by_property(tmp_path):
    styles = pytest.importorskip("openpyxl.styles")

    def own(ws):
        ws["A1"].fill = styles.PatternFill("solid", fgColor="FFFF0000")
        ws["B1"].font = styles.Font(color="FF00B050")
        ws["C1"].font = styles.Font(name="Calibri", size=11)  # no color: automatic

    found = displayed(book(tmp_path, edit=own))
    assert (found[(1, 1)].fill, found[(1, 1)].font.color) == ("FF0000", "FFFFFF")
    assert (found[(1, 2)].fill, found[(1, 2)].font.color) == (BLUE, "00B050")
    # Automatic is the cell's own and draws black; the default font's theme 1 lets white through.
    assert found[(1, 3)].font.color is None
    assert found[(1, 3)].font.bold


def test_an_unknown_style_draws_nothing(tmp_path):
    rendered = sheetpeek.read(book(tmp_path, style="Nowhere")).layout("excel")
    assert all(c.display_style is c.cell.style for row in rendered.rows for c in row.cells)


def totals(path):
    rewrite(path, TABLE, lambda text: text.replace('ref="A1:C5"', 'ref="A1:C5" totalsRowCount="1"'))
    return path


def test_a_later_elements_line_wins_whatever_its_weight(tmp_path):
    borders = sheetpeek.read(totals(book(tmp_path))).layout("excel").borders
    assert borders is not None
    # Above the totals row (row 5): the totals' double top over the whole table's thin line.
    assert borders.horizontal[4] == (Side("double", BLUE),) * 3
    assert borders.horizontal[3] == (LINE,) * 3
    assert borders.vertical[2][0] == LINE  # the whole table's outline
    assert borders.vertical[2][1] is None  # no vertical line inside


def test_an_own_side_from_either_cell_replaces_the_tables_line(tmp_path):
    styles = pytest.importorskip("openpyxl.styles")
    red = styles.Side("thin", color="FFFF0000")

    def own(ws):
        ws["B4"].border = styles.Border(bottom=red)  # above the totals' double top
        ws["C5"].border = styles.Border(top=red)  # the totals cell's own

    borders = sheetpeek.read(totals(book(tmp_path, edit=own))).layout("excel").borders
    assert borders is not None
    assert borders.horizontal[4] == (
        Side("double", BLUE),
        Side("thin", "FF0000"),
        Side("thin", "FF0000"),
    )


def test_table_cells_share_their_displayed_styles(tmp_path):
    rendered = sheetpeek.read(book(tmp_path, showRowStripes=True)).layout("excel")
    styles = [rendered.rows[1].cells[1].display_style, rendered.rows[3].cells[1].display_style]
    assert styles[0] is styles[1]


def test_the_rendered_cell_keeps_the_files_own_style(tmp_path):
    path = book(tmp_path)
    own = sheetpeek.read(path)
    for mode in ("excel", "characters", "none"):
        rendered = sheetpeek.read(path).layout(mode)
        header = rendered.rows[0].cells[1]
        assert header.cell.style == own.cell(1, 2).style
        assert header.display_style.fill == BLUE
        assert header.display_style is not header.cell.style


def own_style(tmp_path, **sides):
    """A table A1:C4 in the workbook's own style: one wholeTable element with these sides."""
    styles = pytest.importorskip("openpyxl.styles")
    differential = pytest.importorskip("openpyxl.styles.differential")
    table_styles = pytest.importorskip("openpyxl.styles.table")
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, 5):
        for c in range(1, 4):
            ws.cell(r, c, f"h{c}" if r == 1 else r * c)
    border = styles.Border(**{k: styles.Side(v, color="FF000000") for k, v in sides.items()})
    index = wb._differential_styles.add(differential.DifferentialStyle(border=border))
    element = table_styles.TableStyleElement(type="wholeTable", dxfId=index)
    wb._table_styles.tableStyle.append(
        table_styles.TableStyle(name="Lines", pivot=False, tableStyleElement=[element])
    )
    table = tables.Table(displayName="Lines", ref="A1:C4")
    table.tableStyleInfo = tables.TableStyleInfo(name="Lines")
    ws.add_table(table)
    return save(wb, tmp_path / "lines.xlsx")


def test_outer_sides_on_the_outline_and_inner_lines_inside(tmp_path):
    path = own_style(
        tmp_path, top="thick", bottom="double", left="medium", right="dashed",
        horizontal="thin", vertical="dotted",
    )  # fmt: skip
    borders = sheetpeek.read(path).layout("excel").borders
    assert borders is not None
    line = {name: Side(name, "000000") for name in ("thick", "double", "medium", "dashed", "thin")}
    dotted = Side("dotted", "000000")
    assert borders.horizontal[0] == (line["thick"],) * 3
    assert borders.horizontal[2] == (line["thin"],) * 3
    assert borders.horizontal[4] == (line["double"],) * 3
    assert borders.vertical[1] == (line["medium"], dotted, dotted, line["dashed"])


def test_a_later_element_replaces_an_earlier_bottom_line(tmp_path):
    # The header's bottom (later) over the whole table's inner line (earlier), whatever the
    # weights: a thin header bottom replaces a thick inner line.
    styles = pytest.importorskip("openpyxl.styles")
    differential = pytest.importorskip("openpyxl.styles.differential")
    table_styles = pytest.importorskip("openpyxl.styles.table")
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, 5):
        for c in range(1, 4):
            ws.cell(r, c, f"h{c}" if r == 1 else r * c)
    thick = differential.DifferentialStyle(
        border=styles.Border(horizontal=styles.Side("thick", color="FF000000"))
    )
    thin = differential.DifferentialStyle(
        border=styles.Border(bottom=styles.Side("thin", color="FF0000FF"))
    )
    elements = [
        table_styles.TableStyleElement(type="wholeTable", dxfId=wb._differential_styles.add(thick)),
        table_styles.TableStyleElement(type="headerRow", dxfId=wb._differential_styles.add(thin)),
    ]
    wb._table_styles.tableStyle.append(
        table_styles.TableStyle(name="Order", pivot=False, tableStyleElement=elements)
    )
    table = tables.Table(displayName="Order", ref="A1:C4")
    table.tableStyleInfo = tables.TableStyleInfo(name="Order")
    ws.add_table(table)
    borders = sheetpeek.read(save(wb, tmp_path / "order.xlsx")).layout("excel").borders
    assert borders is not None
    assert borders.horizontal[1] == (Side("thin", "0000FF"),) * 3
    assert borders.horizontal[2] == (Side("thick", "000000"),) * 3


def test_an_uncolored_style_equal_to_a_theme_1_style_keeps_its_own_color(tmp_path):
    # B1 keeps the default font (theme 1); C1's font is the same but without a color. Their
    # styles compare equal, so the record of theme-1 styles must go by the style's identity:
    # were equal styles ever shared, C1 would take the header's white as well.
    styles = pytest.importorskip("openpyxl.styles")

    def own(ws):
        ws["C1"].font = styles.Font(name="Calibri", sz=11, family=2, scheme="minor")

    path = book(tmp_path, edit=own)
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 3).style == sheet.cell(1, 2).style
    found = displayed(path)
    assert found[(1, 2)].font.color == "FFFFFF"
    assert found[(1, 3)].font.color is None
