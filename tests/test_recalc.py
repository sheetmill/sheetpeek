"""What a formula cell shows on opening: ``recalc_on_load``, through xlrecalc."""

import io
import re

import openpyxl
import pytest
import xlformat
import xlrecalc

import sheetpeek
from sheetpeek import cli
from sheetpeek._recalc import shown, why_not
from sheetpeek._serve import make_server
from tests.workbooks import SHEET1, replace_sheet_data, rewrite, save

SHEET2 = "xl/worksheets/sheet2.xml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XM = "http://schemas.microsoft.com/office/excel/2006/main"


def excel_saved(text):
    """The workbook part as Excel saves it: its own ``calcId`` and ``fileVersion``."""
    text = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029"/>', text)
    return text.replace("<workbookPr", '<fileVersion appName="xl"/><workbookPr', 1)


def book(tmp_path, *rows, other=None, by_excel=True):
    """Sheet ``S`` holding ``rows`` (``<row>`` elements), and sheet ``T`` holding ``other``."""
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.create_sheet("T")["A1"] = 1
    path = save(wb, tmp_path / "book.xlsx")
    rewrite(path, SHEET1, replace_sheet_data(f"<sheetData>{''.join(rows)}</sheetData>"))
    if other is not None:
        rewrite(path, SHEET2, replace_sheet_data(f"<sheetData>{other}</sheetData>"))
    if by_excel:
        rewrite(path, "xl/workbook.xml", excel_saved)
    return path


def row(number, *cells):
    return f'<row r="{number}">{"".join(cells)}</row>'


def values(sheet):
    return [[(c.value, c.kind) for c in line] for line in sheet.cells]


A1 = '<c r="A1"><v>2</v></c>'
STALE = '<c r="B1"><f>A1*3</f><v>99</v></c>'


def notices(sheet, code="value.not_computed"):
    return [n for n in sheet.notices if n.code == code]


@pytest.mark.parametrize(
    ("by_excel", "mode", "shown_value"),
    [
        (True, "excel", 99),  # Excel keeps the values of a file it saved
        (True, "always", 6),
        (True, "never", 99),
        (False, "excel", 6),  # and recalculates a file another program wrote
        (False, "always", 6),
        (False, "never", 99),
    ],
)
def test_each_mode_follows_xlrecalc(tmp_path, by_excel, mode, shown_value):
    path = book(tmp_path, row(1, A1, STALE), by_excel=by_excel)
    sheet = sheetpeek.read(path, recalc_on_load=mode)
    assert sheet.cell(1, 2).value == shown_value
    assert sheet.notices == ()


def test_excel_recalculates_the_cells_it_marked(tmp_path):
    marked = '<c r="B1"><f ca="1">A1*3</f><v>99</v></c>'
    unmarked = '<c r="C1"><f>A1*4</f><v>98</v></c>'
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, marked, unmarked)))
    assert [sheet.cell(1, 2).value, sheet.cell(1, 3).value] == [6, 98]


def test_an_unmarked_volatile_function_on_another_sheet_recalculates_every_formula(tmp_path):
    # Another program's file that looks like Excel's: the other sheet decides the answer.
    path = book(tmp_path, row(1, A1, STALE), other=row(1, '<c r="A1"><f>NOW()</f><v>1</v></c>'))
    assert sheetpeek.read(path, "S").cell(1, 2).value == 6
    assert sheetpeek.read(path, "S", recalc_on_load="never").cell(1, 2).value == 99


def counted_reads(monkeypatch):
    """Each reading of a sheet through xlio, as ``(kind, sheet name)``."""
    import xlio

    reads = []
    for kind in ("cells", "formulas"):
        original = getattr(xlio.Sheet, kind)

        def counted(self, *args, _original=original, _kind=kind, **kwargs):
            reads.append((_kind, self.name))
            return _original(self, *args, **kwargs)

        monkeypatch.setattr(xlio.Sheet, kind, counted)
    return reads


def test_the_other_sheets_answer_from_their_formulas_alone(tmp_path, monkeypatch):
    # Excel's file: whether T holds a formula Excel recomputes decides the answer; T's formulas
    # are read for it, not its values, and with every cached value standing, never its cells.
    reads = counted_reads(monkeypatch)
    later = "".join(f'<c r="{c}1"><f>A1+1</f><v>2</v></c>' for c in "BCD")
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE), other=row(1, A1, later)), "S")
    assert sheet.cell(1, 2).value == 99
    assert reads == [("cells", "S"), ("formulas", "T")]


def test_a_sheet_without_formulas_asks_nothing(tmp_path, monkeypatch):
    asked, loads = [], []
    opening = xlrecalc.Loader.opening

    def counted(self):
        asked.append(self)
        return opening.fget(self)

    monkeypatch.setattr(xlrecalc.Loader, "opening", property(counted))
    monkeypatch.setattr(xlrecalc.Loader, "finish", lambda self: loads.append(self))
    reads = counted_reads(monkeypatch)
    path = book(tmp_path, row(1, A1), other=row(1, '<c r="A1"><f>NOW()</f></c>'))
    assert sheetpeek.read(path, "S").notices == ()
    assert (asked, loads, reads) == ([], [], [("cells", "S")])


# Each case: sheet S (A1 = 2, B1 = A1*3 cached 99, and the cells given), sheet T, Excel's file
# or not, and what S shows in each mode (B1, then each cell given). The answers are xlrecalc's
# RecalcCheck's, as sheetpeek asked them before it took them from the Loader.
OPENING_CASES = {
    "cached": ([], None, True, {"excel": [99], "always": [6], "never": [99]}),
    "marked": (
        ['<c r="C1"><f ca="1">A1*4</f><v>98</v></c>'],
        None,
        True,
        {"excel": [99, 8], "always": [6, 8], "never": [99, 98]},
    ),
    "uncached": (
        ['<c r="C1"><f>A1*4</f></c>'],
        None,
        True,
        {"excel": [99, 8], "always": [6, 8], "never": [99, 8]},
    ),
    "volatile elsewhere": (
        [],
        '<c r="A1"><f>NOW()</f><v>1</v></c>',
        True,
        {"excel": [6], "always": [6], "never": [99]},
    ),
    "another program": ([], None, False, {"excel": [6], "always": [6], "never": [99]}),
}


@pytest.mark.parametrize("mode", ["excel", "always", "never"])
@pytest.mark.parametrize("case", list(OPENING_CASES))
def test_recalc_on_load_answers_as_the_separate_check_did(tmp_path, case, mode):
    import xlio

    cells, other, by_excel, shown_values = OPENING_CASES[case]
    path = book(
        tmp_path, row(1, A1, STALE, *cells), other=other and row(1, other), by_excel=by_excel
    )
    check = xlrecalc.RecalcCheck(io.BytesIO(path.read_bytes()), recalc_on_load=mode)
    with xlio.open(path) as opened:
        for each in opened:
            for cell in each.cells():
                if check.result is None and cell.formula is not None:
                    ca = cell.formula.attributes.get("ca") in ("1", "true")
                    check.add(cell.formula.text, cached=cell.type != "none", ca=ca)
        expected = check.finish()
    with xlio.open(path) as opened:
        assert xlrecalc.Loader(opened, recalc_on_load=mode).opening == expected
    sheet = sheetpeek.read(path, "S", recalc_on_load=mode)
    assert [c.value for c in sheet.cells[0][1:]] == shown_values[mode]


def test_no_engine_when_every_cached_value_stands(tmp_path, monkeypatch):
    loads = []
    monkeypatch.setattr(xlrecalc.Loader, "finish", lambda self: loads.append(self))
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE)))
    assert (sheet.cell(1, 2).value, loads) == (99, [])


def test_values_of_every_kind(tmp_path):
    cells = [
        A1,
        '<c r="B1"><f>A1/4</f></c>',
        '<c r="C1"><f>1/0</f></c>',
        '<c r="D1"><f>A1&amp;"x"</f></c>',
        '<c r="E1"><f>A1&gt;1</f></c>',
        '<c r="F1"><f>""</f></c>',
        '<c r="G1"><f>A1*1E+300</f></c>',
    ]
    sheet = sheetpeek.read(book(tmp_path, row(1, *cells), by_excel=False))
    assert values(sheet) == [
        [
            (2, "number"),
            (0.5, "number"),
            (xlformat.ExcelError("#DIV/0!"), "error"),
            ("2x", "text"),
            (True, "bool"),
            ("", "text"),
            (2e300, "number"),
        ]
    ]


def test_shared_formulas_are_computed(tmp_path):
    rows = [
        row(1, A1, '<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f></c>'),
        row(2, '<c r="A2"><v>5</v></c>', '<c r="B2"><f t="shared" si="0"/></c>'),
    ]
    sheet = sheetpeek.read(book(tmp_path, *rows))
    assert [sheet.cell(1, 2).value, sheet.cell(2, 2).value] == [4, 10]


def test_an_array_formula_fills_its_range(tmp_path):
    # The cells after the first have no <c> here (openpyxl writes none); Excel writes them.
    rows = [
        row(1, '<c r="A1"><v>1</v></c>', '<c r="C1"><f t="array" ref="C1:C3">A1:A3*10</f></c>'),
        row(2, '<c r="A2"><v>2</v></c>'),
        row(3, '<c r="A3"><v>3</v></c>', '<c r="C3"><v>0</v></c>'),
    ]
    path = book(tmp_path, *rows)
    sheet = sheetpeek.read(path)
    assert [sheet.cell(r, 3).value for r in (1, 2, 3)] == [10, 20, 30]
    # Its first cell outside the range read.
    part = sheetpeek.read(path, ref="C2:C3")
    assert [part.cell(r, 3).value for r in (2, 3)] == [20, 30]


def test_an_array_xlrecalc_does_not_compute_keeps_its_values(tmp_path):
    rows = [
        row(1, '<c r="C1"><f t="array" ref="C1:C2">CUBEVALUE("x")</f><v>5</v></c>'),
        row(2, '<c r="C2"><v>6</v></c>'),
    ]
    path = book(tmp_path, *rows, by_excel=False)
    assert [sheetpeek.read(path).cell(r, 3).value for r in (1, 2)] == [5, 6]
    assert sheetpeek.read(path, ref="C2").cell(2, 3).value == 6


def test_a_malformed_array_range_is_ignored(tmp_path):
    cell = '<c r="B1"><f t="array" ref="nonsense">A1*3</f></c>'
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, cell)))
    assert sheet.cell(1, 2).value in (6, None)


def test_a_function_xlrecalc_does_not_compute_keeps_its_value(tmp_path):
    cells = [
        '<c r="A1" t="str"><f>CUBEVALUE("x")</f><v>kept</v></c>',
        '<c r="B1" t="str"><f>A1&amp;"!"</f><v>kept!</v></c>',
        '<c r="C1"><f>CUBEVALUE("y")</f></c>',
    ]
    sheet = sheetpeek.read(book(tmp_path, row(1, *cells), by_excel=False))
    assert values(sheet) == [[("kept", "text"), ("kept!", "text"), (None, "empty")]]
    found = {n.message.split(": ", 1)[-1]: [str(r) for r in n.cells] for n in notices(sheet)}
    assert found == {
        "CUBEVALUE": ["A1", "C1"],
        "depends on a formula xlrecalc does not compute": ["B1"],
    }
    (empty,) = notices(sheet, "value.no_cached_value")
    assert [str(r) for r in empty.cells] == ["C1"]
    assert "first: C1" in empty.message


def test_a_failure_inside_xlrecalc_keeps_the_value(tmp_path, monkeypatch):
    def fail(self, ref):
        raise RuntimeError("inside")

    monkeypatch.setattr(xlrecalc.Workbook, "value", fail)
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE), by_excel=False))
    assert sheet.cell(1, 2).value == 99
    (notice,) = notices(sheet)
    assert "an error inside xlrecalc (RuntimeError)" in notice.message


def test_a_value_sheetpeek_does_not_show_keeps_the_cached_one(tmp_path, monkeypatch):
    monkeypatch.setattr(xlrecalc.Workbook, "value", lambda self, ref: [[1]])
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE), by_excel=False))
    assert sheet.cell(1, 2).value == 99
    (notice,) = notices(sheet)
    assert "a value sheetpeek does not show" in notice.message


def test_the_reason_takes_the_error_then_the_labels_then_the_precedent():
    status = xlrecalc.CellStatus
    kept = "uncomputed"
    error = status(kept, labels=("CUBEVALUE",), error="Boom\nTraceback")
    assert why_not(error) == "an error inside xlrecalc (Boom)"
    labels = status(kept, labels=("external or 3-D reference", "_XLL.HEAT", "CUBEVALUE"))
    assert why_not(labels) == "CUBEVALUE; _XLL.HEAT; external or 3-D reference"
    inherited = status("kept", inherited=True)
    assert why_not(inherited) == "depends on a formula xlrecalc does not compute"


def test_shown_values():
    assert shown(None) == (None, "empty")
    assert shown(3.0) == (3, "number")
    assert shown(2.0**53) == (2.0**53, "number")
    assert shown(xlrecalc.ErrorValue("#N/A")) == (xlformat.ExcelError("#N/A"), "error")
    assert shown(object()) is None
    assert shown([[1]]) is None


@pytest.mark.parametrize("refused", ["Loader", "opening"])
def test_a_workbook_xlrecalc_refuses_keeps_every_cached_value(tmp_path, monkeypatch, refused):
    def refuse(*args, **kwargs):
        raise xlrecalc.PackageError("no")

    if refused == "Loader":
        monkeypatch.setattr(xlrecalc, "Loader", refuse)
    else:
        monkeypatch.setattr(xlrecalc.Loader, "opening", property(refuse))
    uncached = '<c r="C1"><f>A1*4</f></c>'
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE, uncached), by_excel=False))
    assert values(sheet) == [[(2, "number"), (99, "number"), (None, "empty")]]
    (notice,) = notices(sheet)
    assert "xlrecalc could not read the workbook: no" in notice.message
    assert [str(r) for r in notice.cells] == ["B1:C1"]
    assert len(notices(sheet, "value.no_cached_value")) == 1


def test_another_sheet_that_cannot_be_read_is_reported(tmp_path, monkeypatch):
    import xlio

    def unreadable(self, *args, **kwargs):
        raise RuntimeError("a broken part")

    monkeypatch.setattr(xlio.Sheet, "formulas", unreadable)
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE)))
    assert sheet.cell(1, 2).value == 99
    (notice,) = notices(sheet)
    assert "xlrecalc could not read the workbook" in notice.message


def test_a_macro_sheet_keeps_its_cached_values(tmp_path):
    # xlrecalc computes worksheets only. A macro sheet in Excel's form, as test_workbook builds it.
    path = book(tmp_path, row(1, A1), other=row(1, A1, STALE), by_excel=False)
    rewrite(
        path,
        SHEET2,
        lambda t: re.sub(
            r"<worksheet ([^>]*)>(.*)</worksheet>",
            rf'<xm:macrosheet \1 xmlns:xm="{XM}">\2</xm:macrosheet>',
            t,
            flags=re.S,
        ),
    )
    rewrite(
        path,
        "xl/_rels/workbook.xml.rels",
        lambda t: t.replace(
            f'{REL}/worksheet" Target="/xl/worksheets/sheet2.xml"',
            'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" '
            'Target="/xl/worksheets/sheet2.xml"',
        ),
    )
    assert sheetpeek.read(path, "T").cell(1, 2).value == 99


def test_one_engine_serves_every_sheet_of_a_book(tmp_path, monkeypatch):
    loads = []
    workbook = xlrecalc.Loader.workbook

    def counted(self):
        loads.append(self)
        return workbook(self)

    monkeypatch.setattr(xlrecalc.Loader, "workbook", counted)
    path = book(tmp_path, row(1, A1, STALE), other=row(1, A1, STALE), by_excel=False)
    with sheetpeek.open(path) as opened:
        assert opened.sheet("S").cell(1, 2).value == 6
        assert opened.sheet("T").cell(1, 2).value == 6
        assert opened.sheet("S").cell(1, 1).value == 2  # the book's own reads are unmoved
    assert len(loads) == 1


@pytest.mark.parametrize("by_excel", [False, True])
def test_a_sheet_no_value_reaches_is_not_read(tmp_path, monkeypatch, by_excel):
    # The engine is built from the reader's own stream (xlrecalc's Loader) and computes a cell
    # when it is asked for: reading S, whose formula refers to S alone, reads no cell of T
    # (whose formulas alone are read for the opening answer).
    import xlio

    reads = []
    cells = xlio.Sheet.cells

    def counted(self, *args, **kwargs):
        reads.append(self.name)
        return cells(self, *args, **kwargs)

    monkeypatch.setattr(xlio.Sheet, "cells", counted)
    volatile = '<c r="C1"><f>NOW()</f><v>1</v></c>'  # Excel's file: the check reads T for it
    path = book(tmp_path, row(1, A1, STALE), other=row(1, A1, volatile), by_excel=by_excel)
    assert sheetpeek.read(path, "S").cell(1, 2).value == 6
    assert reads == ["S"]


def test_a_sheet_a_value_reaches_is_read_by_the_engine(tmp_path, monkeypatch):
    import xlio

    reads = []
    cells = xlio.Sheet.cells

    def counted(self, *args, **kwargs):
        reads.append(self.name)
        return cells(self, *args, **kwargs)

    monkeypatch.setattr(xlio.Sheet, "cells", counted)
    across = '<c r="B1"><f>T!A1*3</f><v>99</v></c>'
    path = book(
        tmp_path, row(1, A1, across), other=row(1, '<c r="A1"><v>5</v></c>'), by_excel=False
    )
    with sheetpeek.open(path) as opened:
        assert opened.sheet("S").cell(1, 2).value == 15
        assert opened.sheet("T").cell(1, 1).value == 5  # read again, from the file
    assert reads == ["S", "T", "T"]


def test_a_failure_while_recording_keeps_cached_values_and_reads_on(tmp_path, monkeypatch):
    def failing(self, sheet, *, rich=False):
        for k, cell in enumerate(sheet.cells(rich=rich)):
            if k == 1:
                raise RuntimeError("recording failed")
            yield cell

    monkeypatch.setattr(xlrecalc.Loader, "cells", failing)
    uncached = '<c r="C1"><f>A1*4</f></c>'
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE, uncached), by_excel=False))
    assert values(sheet) == [[(2, "number"), (99, "number"), (None, "empty")]]
    (notice,) = notices(sheet)
    assert "xlrecalc could not read the workbook: recording failed" in notice.message


def test_an_error_of_the_files_own_is_raised_through_the_loader(tmp_path, monkeypatch):
    import xlio

    def broken(self, sheet, *, rich=False):
        raise xlio.PackageError("a broken sheet")
        yield  # a generator, as Loader.cells is

    monkeypatch.setattr(xlrecalc.Loader, "cells", broken)
    with pytest.raises(sheetpeek.PackageError, match="a broken sheet"):
        sheetpeek.read(book(tmp_path, row(1, A1, STALE), by_excel=False))


def test_a_file_object_is_computed_too(tmp_path):
    data = book(tmp_path, row(1, A1, STALE), by_excel=False).read_bytes()
    assert sheetpeek.read(io.BytesIO(data)).cell(1, 2).value == 6


@pytest.mark.parametrize("mode", ["sometimes", None, "EXCEL", 1])
def test_an_unknown_mode_is_refused(tmp_path, mode):
    path = book(tmp_path, row(1, A1))
    with pytest.raises(ValueError, match="recalc_on_load must be one of"):
        sheetpeek.read(path, recalc_on_load=mode)
    with pytest.raises(ValueError, match="recalc_on_load must be one of"):
        sheetpeek.open(path, recalc_on_load=mode)
    with pytest.raises(ValueError, match="recalc_on_load must be one of"):
        make_server(tmp_path, port=0, recalc_on_load=mode)


def test_the_command_line_takes_the_mode(tmp_path):
    path = book(tmp_path, row(1, A1, STALE), by_excel=False)
    parser = cli._parser()
    for mode, expected in (("excel", "6"), ("never", "99")):
        out, err = io.StringIO(), io.StringIO()
        out.isatty = lambda: False
        args = parser.parse_args(["view", str(path), "--recalc-on-load", mode])
        assert cli._view(args, out, err) == 0
        assert out.getvalue().split() == ["2", expected]
    with pytest.raises(SystemExit):
        parser.parse_args(["view", str(path), "--recalc-on-load", "sometimes"])
    assert parser.parse_args(["serve", ".", "--recalc-on-load", "never"]).recalc_on_load == "never"


def test_the_engine_computes_cell_by_cell_without_a_whole_recalculation(tmp_path, monkeypatch):
    runs = []
    recalculate = xlrecalc.Workbook.recalculate

    def counted(self, *args, **kwargs):
        runs.append(args)
        return recalculate(self, *args, **kwargs)

    monkeypatch.setattr(xlrecalc.Workbook, "recalculate", counted)
    cells = [A1, '<c r="B1"><f>CUBEVALUE("x")</f><v>1</v></c>']
    cells += [f'<c r="{c}1"><f>B1+A1</f><v>0</v></c>' for c in "CDEFG"]
    cells += [f'<c r="{c}1"><f>A1*2</f><v>0</v></c>' for c in "HIJKL"]
    sheet = sheetpeek.read(book(tmp_path, row(1, *cells), by_excel=False))
    # Resting on B1, which xlrecalc keeps: their cached values; the others computed.
    assert [sheet.cell(1, c).value for c in range(3, 13)] == [0] * 5 + [4] * 5
    assert runs == []


def test_a_failure_to_build_the_engine_keeps_every_cached_value(tmp_path, monkeypatch):
    def fail(self, *args, **kwargs):
        raise RuntimeError("inside")

    monkeypatch.setattr(xlrecalc.Loader, "workbook", fail)
    sheet = sheetpeek.read(book(tmp_path, row(1, A1, STALE), by_excel=False))
    assert sheet.cell(1, 2).value == 99
    (notice,) = notices(sheet)
    assert "xlrecalc could not read the workbook: inside" in notice.message


def test_the_reason_names_an_external_reference_and_an_add_in_function(tmp_path):
    cells = [
        '<c r="A1"><f>[1]Other!A1*2</f><v>4</v></c>',
        '<c r="B1"><f>Sheet1:Sheet3!A1+1</f><v>5</v></c>',
        '<c r="C1"><f>_xll.HEAT(1)</f><v>6</v></c>',
        '<c r="D1"><f>A1+1</f><v>7</v></c>',
    ]
    sheet = sheetpeek.read(book(tmp_path, row(1, *cells), by_excel=False))
    assert [c.value for c in sheet.cells[0]] == [4, 5, 6, 7]
    found = {n.message.split(": ", 1)[-1]: [str(r) for r in n.cells] for n in notices(sheet)}
    assert found == {
        "external or 3-D reference": ["A1:B1"],
        "_XLL.HEAT": ["C1"],
        "depends on a formula xlrecalc does not compute": ["D1"],
    }
