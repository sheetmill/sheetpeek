"""Untrusted files: what is refused before anything parses them, and what must still be read.

Excel's own behavior with these packages is in ``tests/captures/package.tsv`` (ARCHITECTURE.md
§5.5). The crafted files are built here, small, and never committed. A test whose guard could fail
into a huge allocation is ``bounded``: if it regresses, the test fails and the machine is fine.
"""

import io
import re
import warnings
import zipfile
import zlib
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import Alignment

import sheetpeek
from sheetpeek import LimitError, PackageError, ReadLimits
from tests.watchdog import bounded
from tests.workbooks import (
    NS,
    SHEET1,
    replace_part,
    replace_sheet_data,
    rewrite,
    save,
    shared_strings,
    state_size,
)

STYLES = "xl/styles.xml"
WORKBOOK = "xl/workbook.xml"
RELS = "xl/_rels/workbook.xml.rels"
DTD = '<!DOCTYPE x [<!ENTITY e "boom">]>'
BODY = (
    f'<worksheet {NS}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>'
    "</sheetData></worksheet>"
)


@pytest.fixture
def book(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 111
    wb.active["B1"] = "text"
    return save(wb, tmp_path / "book.xlsx")


def part_bytes(path: Path, part: str) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read(part)


# --- DTDs


@pytest.mark.parametrize(
    "form",
    [
        pytest.param((DTD + BODY).encode(), id="plain"),
        pytest.param(('<?xml version="1.0"?>' + DTD + BODY).encode(), id="declaration"),
        pytest.param(b"\xef\xbb\xbf" + (DTD + BODY).encode(), id="utf-8 bom"),
        pytest.param(("\n  " + DTD + BODY).encode(), id="whitespace"),
        pytest.param(("<!--" + "c" * 200_000 + "-->" + DTD + BODY).encode(), id="after a comment"),
        pytest.param(("<?p " + "c" * 900_000 + "?>" + DTD + BODY).encode(), id="after a pi"),
        pytest.param((DTD + BODY).encode("utf-16"), id="utf-16 bom"),
        pytest.param(b"\xfe\xff" + (DTD + BODY).encode("utf-16-be"), id="utf-16-be bom"),
        pytest.param(
            ('<?xml version="1.0" encoding="utf-16"?>' + DTD + BODY).encode("utf-16"),
            id="utf-16 declared",
        ),
        pytest.param(
            ('<?xml version="1.0" encoding="iso-8859-1"?>' + DTD + BODY).encode("latin-1"),
            id="latin-1 declared",
        ),
        pytest.param(
            b"<!--" + b"c" * 65_530 + b"-->" + b"<!DOCTYPE x>" + BODY.encode(), id="split"
        ),
    ],
)
@pytest.mark.parametrize("method", [zipfile.ZIP_DEFLATED, zipfile.ZIP_STORED])
def test_a_dtd_is_refused_in_any_form(book, form, method):
    replace_part(book, SHEET1, form, method)
    with pytest.raises(PackageError, match="declares a DTD"):
        sheetpeek.read(book)


@pytest.mark.parametrize(
    "part",
    [
        "[Content_Types].xml",
        "_rels/.rels",
        WORKBOOK,
        RELS,
        STYLES,
        "xl/theme/theme1.xml",  # sheetpeek swallows a theme it cannot parse: the check must not
        "docProps/core.xml",
        "docProps/app.xml",
        "xl/never-read.bin",  # no parser opens it, and the check does not care
    ],
)
def test_a_dtd_is_refused_in_any_part(book, part):
    replace_part(book, part, b"<!DOCTYPE x><x/>")
    with pytest.raises(PackageError, match="declares a DTD"):
        sheetpeek.read(book)


def test_a_dtd_is_refused_under_a_name_without_xml(book):
    # The relationships decide what a worksheet is called, so the name decides nothing.
    rels = part_bytes(book, RELS).replace(b"worksheets/sheet1.xml", b"worksheets/data.bin")
    replace_part(book, RELS, rels)
    replace_part(book, "xl/worksheets/data.bin", (DTD + BODY).encode())
    with pytest.raises(PackageError, match=r"data\.bin: declares a DTD"):
        sheetpeek.read(book)


def test_the_entity_is_never_expanded(book):
    use = BODY.replace("<t>x</t>", "<t>&e;</t>")
    replace_part(book, SHEET1, (DTD + use).encode())
    with pytest.raises(PackageError, match="declares a DTD"):
        sheetpeek.read(book)


def test_a_billion_laughs_is_refused_whatever_expat_does(book):
    entities = '<!ENTITY a "aaaaaaaaaa">' + "".join(
        f'<!ENTITY {chr(98 + i)} "{("&" + chr(97 + i) + ";") * 10}">' for i in range(9)
    )
    use = BODY.replace("<t>x</t>", "<t>&j;</t>")
    replace_part(book, SHEET1, f"<!DOCTYPE x [{entities}]>{use}".encode())
    with pytest.raises(PackageError, match="declares a DTD"):
        sheetpeek.read(book)


@pytest.mark.parametrize(
    "form",
    [
        pytest.param(('<?xml version="1.0"?>' + BODY).encode(), id="declaration"),
        pytest.param(BODY.encode("utf-16"), id="utf-16"),
        pytest.param(("<!-- <!DOCTYPE x> -->" + BODY).encode(), id="in a comment"),
        pytest.param(("<?p <!DOCTYPE x>?>" + BODY).encode(), id="in a pi"),
        pytest.param(BODY.replace("<t>x</t>", "<t>&lt;!DOCTYPE x&gt;</t>").encode(), id="in text"),
        pytest.param(
            BODY.replace("<t>x</t>", "<t><![CDATA[<!DOCTYPE x>]]></t>").encode(), id="in cdata"
        ),
    ],
)
def test_the_word_alone_is_not_a_dtd(book, form):
    replace_part(book, SHEET1, form)
    assert sheetpeek.read(book).used_range is not None


@pytest.mark.parametrize("body", [b"\x89PNG\r\n\x1a\n" + b"\0" * 2000, b"", b"plain text"])
def test_a_part_that_is_not_xml_passes(book, body):
    replace_part(book, "xl/media/image1.png", body)
    assert sheetpeek.read(book).cell(1, 1).value == 111


def test_a_dtd_cannot_hide_after_the_root(book):
    # The check stops at the root: past it, the parser itself rejects a DTD.
    replace_part(book, SHEET1, (BODY + DTD).encode())
    with pytest.raises(PackageError, match="cannot read this workbook"):
        sheetpeek.read(book)


def test_a_long_prolog_is_refused(book):
    replace_part(book, SHEET1, b"<!--" + b"c" * (2 << 20) + b"-->" + BODY.encode())
    with pytest.raises(PackageError, match="no root element"):
        sheetpeek.read(book)


@pytest.mark.parametrize("encoding", ["gb2312", "shift_jis", "utf-7", "bogus"])
def test_an_encoding_expat_cannot_read_is_refused(book, encoding):
    # lxml, when installed, reads some of these with openpyxl: the check cannot see into them.
    replace_part(book, SHEET1, f'<?xml version="1.0" encoding="{encoding}"?>{BODY}'.encode())
    with pytest.raises(PackageError):
        sheetpeek.read(book)


# --- sizes and methods


def test_a_part_larger_than_stated_is_refused_without_inflating_it(book):
    # The real part, then spaces, with a CRC of the part alone: zipfile reads that without an
    # error, and ZipFile.read inflates all of it first.
    xml = part_bytes(book, STYLES)
    replace_part(book, STYLES, xml + b" " * (64 << 20))
    state_size(book, STYLES, len(xml), zlib.crc32(xml))
    with pytest.raises(PackageError, match=r"styles\.xml: the archive states"):
        sheetpeek.read(book)


@pytest.mark.parametrize("change", [-100, 1000])
def test_a_false_size_is_refused(book, change):
    size = len(part_bytes(book, SHEET1))
    state_size(book, SHEET1, size + change)
    with pytest.raises(PackageError, match="the archive states"):
        sheetpeek.read(book)


@pytest.mark.parametrize("method", [zipfile.ZIP_BZIP2, zipfile.ZIP_LZMA])
def test_other_compression_methods_are_refused(book, method):
    # ZipFile.read gives these no limit on the output at all.
    replace_part(book, SHEET1, part_bytes(book, SHEET1), method)
    with pytest.raises(PackageError, match="compression method"):
        sheetpeek.read(book)


def test_a_part_named_twice_is_refused(book):
    with zipfile.ZipFile(book, "a") as archive, warnings.catch_warnings():
        warnings.simplefilter("ignore")
        archive.writestr(SHEET1, part_bytes(book, SHEET1))
    with pytest.raises(PackageError, match="names this part twice"):
        sheetpeek.read(book)


def test_an_encrypted_part_is_refused(book):
    data = bytearray(book.read_bytes())
    with zipfile.ZipFile(book) as archive:
        data[archive.start_dir + 8] |= 1
    book.write_bytes(data)
    with pytest.raises(PackageError, match="encrypted"):
        sheetpeek.read(book)


def test_damaged_data_is_refused(book):
    with zipfile.ZipFile(book) as archive:
        info = archive.getinfo(SHEET1)
    data = bytearray(book.read_bytes())
    start = info.header_offset + 30 + len(info.filename)
    data[start : start + 8] = b"\xff" * 8
    book.write_bytes(data)
    with pytest.raises(PackageError, match="damaged"):
        sheetpeek.read(book)


# --- the limits


def test_the_number_of_parts(book):
    with zipfile.ZipFile(book) as archive:
        parts = len(archive.infolist())
    assert sheetpeek.read(book, limits=ReadLimits(max_parts=parts)).cell(1, 1).value == 111
    with pytest.raises(LimitError, match=f"{parts} parts"):
        sheetpeek.read(book, limits=ReadLimits(max_parts=parts - 1))


def test_the_size_of_a_part(book):
    with zipfile.ZipFile(book) as archive:
        largest = max(i.file_size for i in archive.infolist())
    assert sheetpeek.read(book, limits=ReadLimits(max_part_bytes=largest)).cell(1, 1).value == 111
    with pytest.raises(LimitError, match="more than the limit"):
        sheetpeek.read(book, limits=ReadLimits(max_part_bytes=largest - 1))


def test_the_total_size(book):
    with zipfile.ZipFile(book) as archive:
        total = sum(i.file_size for i in archive.infolist())
    assert sheetpeek.read(book, limits=ReadLimits(max_total_bytes=total)).cell(1, 1).value == 111
    with pytest.raises(LimitError, match="in all"):
        sheetpeek.read(book, limits=ReadLimits(max_total_bytes=total - 1))


def test_limit_errors_are_value_errors():
    assert issubclass(LimitError, PackageError)
    assert issubclass(PackageError, ValueError)


def test_limits_reach_every_entry_point(book, tmp_path):
    tight = ReadLimits(max_parts=1)
    with pytest.raises(LimitError):
        sheetpeek.open(book, limits=tight)
    with pytest.raises(LimitError):
        sheetpeek.read(book, limits=tight)
    with pytest.raises(LimitError):
        sheetpeek.render(book, limits=tight)
    with pytest.raises(LimitError), book.open("rb") as handle:
        sheetpeek.read(handle, limits=tight)


# --- the grid


def cells_limit(n):
    return ReadLimits(max_cells=n)


@bounded
def test_max_cells_counts_the_area_not_what_is_populated(book):
    # A1:B1 holds two values; A1:T5 holds the same two values in a hundred positions.
    assert sheetpeek.read(book, ref="A1:T5", limits=cells_limit(100)).cell(1, 1).value == 111
    with pytest.raises(LimitError, match="A1:T5 is 100 cells"):
        sheetpeek.read(book, ref="A1:T5", limits=cells_limit(99))


@pytest.mark.parametrize(
    "tail",
    [
        pytest.param(
            '<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>', id="merge"
        ),
        pytest.param("", id="stray cell"),
    ],
)
@bounded
def test_a_tiny_file_cannot_ask_for_a_huge_grid(book, tail):
    stray = "" if tail else '<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>'
    data = f'<sheetData><row r="1"><c r="A1"><v>1</v></c></row>{stray}</sheetData>{tail}'
    rewrite(book, SHEET1, replace_sheet_data(data))
    with pytest.raises(LimitError, match="more than the limit"):
        sheetpeek.read(book)
    # The rest of the sheet can still be read.
    assert sheetpeek.read(book, ref="A1:B2").cell(1, 1).value == 1


@bounded
def test_a_row_past_excel_cannot_ask_for_a_huge_grid(book):
    data = '<sheetData><row r="99999999999"><c r="A99999999999"><v>1</v></c></row></sheetData>'
    rewrite(book, SHEET1, replace_sheet_data(data))
    with pytest.raises(LimitError):
        sheetpeek.read(book)


@bounded
def test_a_huge_merge_costs_nothing_past_the_range_read(book):
    # The layout's map of merges walked the whole merge, 17 billion positions.
    data = (
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
        '<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>'
    )
    rewrite(book, SHEET1, replace_sheet_data(data))
    for mode in ("excel", "characters", "none"):
        rendered = sheetpeek.read(book, ref="A1:B2").layout(mode)
        assert rendered.cell(1, 1).merge == sheetpeek.CellRange(1, 1, 1048576, 16384)


ODD_NUMBERS = [
    "nan",
    "inf",
    "-inf",
    "-1",
    "-0.5",
    "0",
    "1e4",
    "1e6",
    "99999999999",
    "1e30",
    "1e308",
]
ODD_FIELDS = {
    "width": ("<sheetData>", '<cols><col min="1" max="1" width="{}"/></cols><sheetData>'),
    "height": ('<row r="1"', '<row r="1" ht="{}" customHeight="1"'),
    "default width": ("<sheetFormatPr ", '<sheetFormatPr defaultColWidth="{}" '),
    "default height": ('defaultRowHeight="15"', 'defaultRowHeight="{}" customHeight="1"'),
    "row outline level": ('<row r="1"', '<row r="1" outlineLevel="{}"'),
    "column outline level": (
        "<sheetData>",
        '<cols><col min="1" max="16384" outlineLevel="{}"/></cols><sheetData>',
    ),
    "zoom": ('<sheetView workbookViewId="0">', '<sheetView workbookViewId="0" zoomScale="{}">'),
    "frozen rows": (
        '<sheetView workbookViewId="0">',
        '<sheetView workbookViewId="0"><pane ySplit="{0}" xSplit="{0}" state="frozen"/>',
    ),
    "page break": ("</worksheet>", '<rowBreaks><brk id="{}" man="1"/></rowBreaks></worksheet>'),
}


@pytest.mark.parametrize("field", ODD_FIELDS)
@pytest.mark.parametrize("value", ODD_NUMBERS)
@bounded
def test_no_mode_allocates_or_fails_by_an_odd_number(book, field, value):
    # What Excel does with each is compared in test_captures.py; here, none costs anything.
    old, new = ODD_FIELDS[field]
    rewrite(book, SHEET1, lambda text: text.replace(old, new.format(value), 1))
    try:
        sheet = sheetpeek.read(book)
    except PackageError:
        return
    for mode in ("excel", "characters", "none"):
        sheetpeek.render(
            sheet.layout(mode), {"excel": "html", "characters": "text"}.get(mode, "csv")
        )


# Numbers of fills and borders, in the stylesheet: a gradient's angle, rectangle and stop
# positions, and the indexes and tints of fill and border colors.
ODD_STYLE_FIELDS = {
    "gradient degree": 'degree="45"',
    "gradient left": 'left="0.25"',
    "gradient right": 'right="0.75"',
    "gradient top": 'top="0.125"',
    "gradient bottom": 'bottom="0.875"',
    "stop position": '<stop position="1"',
    "fill index": 'fgColor indexed="10"',
    "fill tint": 'bgColor theme="4" tint="0.5"',
    "border index": 'color indexed="12"',
    "border tint": 'color theme="5" tint="-0.25"',
}


def styled_book(path: Path) -> Path:
    from openpyxl.styles import Border, Color, GradientFill, PatternFill, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws["A1"].fill = GradientFill(
        "path", 45, 0.25, 0.75, 0.125, 0.875, stop=("FFFFFFFF", "FF0000FF")
    )
    ws["B1"].fill = PatternFill(
        "gray125", fgColor=Color(indexed=10), bgColor=Color(theme=4, tint=0.5)
    )
    ws["C1"].border = Border(
        left=Side("thick", color=Color(indexed=12)),
        diagonal=Side("thin", color=Color(theme=5, tint=-0.25)),
        diagonalUp=True,
    )
    return save(wb, path)


@pytest.fixture(scope="module")
def styled(tmp_path_factory):
    return styled_book(tmp_path_factory.mktemp("styled") / "styled.xlsx").read_bytes()


@pytest.mark.parametrize("field", ODD_STYLE_FIELDS)
@bounded
def test_no_mode_allocates_or_fails_by_an_odd_style_number(styled, tmp_path, field):
    # One bounded child for each field, trying every odd number: a child for each pair cost
    # most of the suite's time, and the memory bound holds for the whole loop.
    old = ODD_STYLE_FIELDS[field]
    # The number replaces the first one in the attribute (an index or a tint, by field).
    attribute = re.findall(r'(\w+)="[^"]*"$', old)[0]
    for value in ODD_NUMBERS:
        book = tmp_path / "styled.xlsx"
        book.write_bytes(styled)
        new = re.sub(rf'{attribute}="[^"]*"$', f'{attribute}="{value}"', old)
        rewrite(book, STYLES, lambda text, new=new: text.replace(old, new, 1))
        try:
            sheet = sheetpeek.read(book)
        except PackageError:
            continue
        for mode in ("excel", "characters", "none"):
            page = sheetpeek.render(
                sheet.layout(mode),
                {"excel": "html", "characters": "text"}.get(mode, "csv"),
                on_notice="ignore",
            )
            assert "nan" not in page.lower().replace("sans", ""), (field, value, mode)


ODD_ALIGNMENT_FIELDS = ("readingOrder", "indent", "textRotation")


@pytest.mark.parametrize("field", ODD_ALIGNMENT_FIELDS)
@bounded
def test_no_mode_allocates_or_fails_by_an_odd_alignment_number(tmp_path, field):
    # One bounded child for each field, trying every odd number (as above).
    wb = openpyxl.Workbook()
    wb.active["A1"] = "text that is too long for its cell"
    wb.active["A1"].alignment = Alignment(
        readingOrder=1, indent=1, horizontal="distributed", shrink_to_fit=True, textRotation=1
    )
    wb.active["B2"] = 1
    plain = save(wb, tmp_path / "plain.xlsx").read_bytes()
    for value in ODD_NUMBERS:
        path = tmp_path / "odd.xlsx"
        path.write_bytes(plain)
        rewrite(
            path, STYLES, lambda text, v=value: text.replace(f'{field}="1"', f'{field}="{v}"', 1)
        )
        try:
            sheet = sheetpeek.read(path)
        except PackageError:
            continue
        style = sheet.cell(1, 1).style
        assert style.reading_order in (0, 1, 2), (field, value)
        assert -90 <= style.rotation <= 90, (field, value)
        for mode in ("excel", "characters", "none"):
            sheetpeek.render(
                sheet.layout(mode),
                {"excel": "html", "characters": "text"}.get(mode, "csv"),
                on_notice="ignore",
            )


# Numbers that pick a palette entry: the grid color's index, and a number format's [ColorN].
ODD_PALETTE_FIELDS = {
    "grid color id": (SHEET1, "<sheetView ", '<sheetView defaultGridColor="0" colorId="{}" '),
    "section color": (STYLES, '<numFmts count="0" />', '<numFmts count="1"><numFmt '
                      'numFmtId="164" formatCode="[Color{}]0" /></numFmts>'),
}  # fmt: skip


@pytest.mark.parametrize("field", ODD_PALETTE_FIELDS)
@bounded
def test_no_mode_allocates_or_fails_by_an_odd_palette_number(tmp_path, field):
    # One bounded child for each field, trying every odd number (as above).
    part, old, new = ODD_PALETTE_FIELDS[field]
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb._colors = ["00804000"] * 64  # a custom palette: every entry differs from the default
    plain = save(wb, tmp_path / "plain.xlsx").read_bytes()
    for value in [*ODD_NUMBERS, "9" * 5000, "64", "65", "56", "57", "999"]:
        path = tmp_path / "odd.xlsx"
        path.write_bytes(plain)
        rewrite(path, part, lambda text, v=value: text.replace(old, new.format(v), 1))
        if part == STYLES:  # the cell takes the format
            rewrite(path, STYLES, lambda t: t.replace('<xf numFmtId="0"', '<xf numFmtId="164"', 2))
        try:
            sheet = sheetpeek.read(path)
        except PackageError:
            continue
        assert sheet.grid_color in (None, "804000"), (field, value)
        for mode in ("excel", "characters", "none"):
            sheetpeek.render(
                sheet.layout(mode),
                {"excel": "html", "characters": "text"}.get(mode, "csv"),
                on_notice="ignore",
            )


# Numbers of the sheet list (ARCHITECTURE.md §4.15): a tab color's index, theme and tint, and the
# workbook's active tab.
ODD_SHEET_FIELDS = {
    "tab index": (SHEET1, '<tabColor indexed="12" />', '<tabColor indexed="{}" />'),
    "tab theme": (SHEET1, '<tabColor indexed="12" />', '<tabColor theme="{}" />'),
    "tab tint": (SHEET1, '<tabColor indexed="12" />', '<tabColor theme="4" tint="{}" />'),
    "active tab": (WORKBOOK, 'activeTab="0"', 'activeTab="{}"'),
}


def tabbed_book(path: Path) -> Path:
    from openpyxl.styles.colors import Color

    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active.sheet_properties.tabColor = Color(indexed=12)
    wb.create_sheet("Two")["A1"] = 2
    return save(wb, path)


@pytest.mark.parametrize("field", ODD_SHEET_FIELDS)
@bounded
def test_the_sheet_list_takes_any_odd_number(tmp_path, field):
    # One bounded child for each field, trying every odd number (as above).
    part, old, new = ODD_SHEET_FIELDS[field]
    plain = tabbed_book(tmp_path / "plain.xlsx").read_bytes()
    for value in [*ODD_NUMBERS, "12", "4", "0.5", "1", "65", " 3", "+3", "3_0", "0x7", "9" * 5000]:
        path = tmp_path / "odd.xlsx"
        path.write_bytes(plain)
        if new.format(value) != old:
            rewrite(path, part, lambda t, v=value: t.replace(old, new.format(v), 1))
        try:
            book = sheetpeek.open(path)
        except PackageError:
            continue
        with book:
            # sheetpeek's own reading first: openpyxl refuses some of these on the sheet.
            colors = [s.tab_color for s in book.sheets]
            assert all(c is None or re.fullmatch("[0-9A-F]{6}", c) for c in colors), (field, value)
            try:
                assert book.sheet().name in ("Sheet", "Two"), (field, value)
                markup = sheetpeek.render_workbook(
                    book, hidden_sheets=True, show_hidden=True, on_notice="ignore"
                )
            except PackageError:
                continue
        assert markup is not None
        assert "nan" not in markup.lower().replace("sans", ""), (field, value)


@pytest.mark.parametrize("value", ["9" * 5000, "0x7", " 3", "", "1.5"])
@bounded
def test_an_odd_outline_level_is_read_as_a_level(book, value):
    rewrite(
        book, SHEET1, lambda text: text.replace('<row r="1"', f'<row r="1" outlineLevel="{value}"')
    )
    sheet = sheetpeek.read(book)
    assert 0 <= sheet.rows[0].outline_level <= 7
    for name in ("html", "text", "csv"):
        sheetpeek.render(sheet, name, on_notice="ignore")


@bounded
def test_a_grouped_column_run_costs_only_the_range_read(book):
    cols = '<cols><col min="1" max="2147483647" outlineLevel="1" collapsed="1"/></cols><sheetData>'
    rewrite(book, SHEET1, lambda text: text.replace("<sheetData>", cols, 1))
    sheet = sheetpeek.read(book)
    assert [c.outline_level for c in sheet.columns] == [1, 1]
    assert sheet.outline is not None
    assert sheet.outline.columns_beyond == (0, 1)
    sheetpeek.render(sheet, "html", on_notice="ignore")


@pytest.mark.parametrize("value", ["300", "99999999999", "-1"])
@bounded
def test_no_mode_allocates_or_fails_by_a_base_width(book, value):
    rewrite(book, SHEET1, lambda text: text.replace('baseColWidth="8"', f'baseColWidth="{value}"'))
    try:
        sheet = sheetpeek.read(book)
    except PackageError:
        return
    for mode in ("excel", "characters", "none"):
        sheet.layout(mode)


@pytest.mark.parametrize("size", ["0", "0.5", "-0.5", "410", "inf", "nan", "1e400"])
@bounded
def test_a_font_size_excel_refuses_is_refused(book, size):
    rewrite(book, "xl/styles.xml", lambda text: text.replace('<sz val="11"', f'<sz val="{size}"'))
    with pytest.raises(PackageError, match="font size"):
        sheetpeek.read(book)


@pytest.mark.parametrize("value", ["-1", "1e4", "1e6", "256", "255"])
@bounded
def test_a_column_is_at_most_255_characters(book, value):
    col = f'<cols><col min="1" max="1" width="{value}"/></cols><sheetData>'
    rewrite(book, SHEET1, lambda text: text.replace("<sheetData>", col, 1))
    rendered = sheetpeek.read(book, ref="A1:A2000").layout("characters")
    assert rendered.columns[0].width <= 255
    assert rendered.columns[0].width == 255 or value == "255"
    assert sheetpeek.read(book, ref="A1").layout("characters", scale=9).columns[0].width == 255


@bounded
def test_a_huge_ref_on_an_ordinary_file_is_refused(book):
    with pytest.raises(LimitError, match="A1:XFD1048576"):
        sheetpeek.read(book, ref="A1:XFD1048576")


@bounded
def test_max_cells_never_uses_dimension(book):
    # A false <dimension> says one cell; the cells say otherwise.
    data = (
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row>'
        '<row r="9"><c r="J9"><v>1</v></c></row></sheetData>'
    )
    rewrite(book, SHEET1, replace_sheet_data(data))
    rewrite(
        book, SHEET1, lambda text: text.replace('<dimension ref="A1:B1"', '<dimension ref="A1"')
    )
    with pytest.raises(LimitError, match="A1:J9 is 90 cells"):
        sheetpeek.read(book, limits=cells_limit(89))
    assert sheetpeek.read(book, limits=cells_limit(90)).ref == sheetpeek.CellRange(1, 1, 9, 10)


@pytest.mark.parametrize("mode", ["excel", "characters"])
@bounded
def test_columns_added_for_spill_count_against_max_cells(tmp_path, mode):
    # 300 digits in characters mode reach 34 columns. The read is 100 cells; the layout is not.
    wb = openpyxl.Workbook()
    wb.active["A1"] = "1" * 300
    path = save(wb, tmp_path / "spill.xlsx")
    sheet = sheetpeek.read(path, ref="A1:A100", limits=cells_limit(100))
    with pytest.raises(LimitError, match="columns spilled text reaches"):
        sheet.layout(mode)
    roomy = sheetpeek.read(path, ref="A1:A100", limits=cells_limit(100_000)).layout(mode)
    assert len(roomy.columns) > 1


def test_a_sheet_built_by_hand_has_no_cell_limit(book):
    sheet = sheetpeek.read(book)
    assert sheet.max_cells == ReadLimits().max_cells
    assert sheet == sheetpeek.read(book, limits=cells_limit(50))  # not part of what a sheet says


# --- files that are not xlsx


def docx_like() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        archive.writestr("word/document.xml", "<w/>")
    return buffer.getvalue()


def ods_like() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        archive.writestr("content.xml", "<office/>")
    return buffer.getvalue()


def xlsb_like() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/workbook.bin", b"\x83\x01\x00")
    return buffer.getvalue()


COMPOUND = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 1024


@pytest.mark.parametrize(
    ("data", "message"),
    [
        pytest.param(COMPOUND, "old .xls workbook or a password-protected", id="compound"),
        pytest.param(b"a,b\n1,2\n", "not an xlsx file", id="csv"),
        pytest.param(docx_like(), "cannot read this workbook", id="docx"),
        pytest.param(ods_like(), "cannot read this workbook", id="ods"),
        pytest.param(xlsb_like(), r"\.xlsb", id="xlsb content"),
    ],
)
def test_other_files_through_a_file_object(data, message):
    with pytest.raises(PackageError, match=message):
        sheetpeek.read(io.BytesIO(data))


@pytest.mark.parametrize(
    ("name", "data", "message"),
    [
        pytest.param("old.xls", COMPOUND, r"old \.xls format", id="compound xls"),
        pytest.param("old.xls", b"<html>", r"old \.xls format", id="xls by name alone"),
        pytest.param("sheet.xlsb", b"junk", r"binary \.xlsb", id="xlsb by name alone"),
        pytest.param("sheet.xlsb", xlsb_like(), r"binary \.xlsb", id="xlsb"),
        pytest.param("data.csv", b"a,b\n", "not an xlsx file", id="csv"),
        pytest.param("letter.docx", docx_like(), "cannot read this workbook", id="docx"),
        pytest.param("sums.ods", ods_like(), "cannot read this workbook", id="ods"),
    ],
)
def test_other_files_by_path(tmp_path, name, data, message):
    path = tmp_path / name
    path.write_bytes(data)
    with pytest.raises(PackageError, match=message) as refused:
        sheetpeek.read(path)
    assert str(tmp_path) not in str(refused.value)


def test_a_path_that_cannot_be_opened_keeps_its_own_error(tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetpeek.read(tmp_path / "absent.xlsx")
    with pytest.raises(IsADirectoryError):
        sheetpeek.read(tmp_path)


def test_the_suffix_is_not_required(book, tmp_path):
    named = book.rename(tmp_path / "book.dat")
    assert sheetpeek.read(named).cell(1, 1).value == 111


@pytest.mark.parametrize("name", ["book.xlsb", "book.xls"])
def test_the_content_decides_before_the_suffix(book, tmp_path, name):
    named = book.rename(tmp_path / name)
    assert sheetpeek.read(named).cell(1, 1).value == 111


def test_a_refusal_closes_the_file(book, recwarn):
    replace_part(book, SHEET1, (DTD + BODY).encode())
    with pytest.raises(PackageError):
        sheetpeek.read(book)
    assert not [w for w in recwarn if issubclass(w.category, ResourceWarning)]


def test_an_unreadable_sheet_part_names_no_path(book):
    replace_part(book, SHEET1, b"<worksheet><sheetData>")
    with pytest.raises(PackageError) as refused:
        sheetpeek.read(book)
    assert str(book.parent) not in str(refused.value)


# --- what must not be refused


class _Unseekable(io.RawIOBase):
    """A stream that cannot seek: zipfile then sets the data descriptor flag."""

    def __init__(self):
        self.written = bytearray()

    def writable(self):
        return True

    def write(self, data):
        self.written.extend(data)
        return len(data)

    def seekable(self):
        return False


def rewritten(path: Path, stream, **options):
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(stream, "w", **options) as target:
        for info in source.infolist():
            with target.open(info.filename, "w") as part:
                part.write(source.read(info.filename))


def test_a_data_descriptor_is_read(book):
    stream = _Unseekable()
    rewritten(book, stream, compression=zipfile.ZIP_DEFLATED)
    data = bytes(stream.written)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert all(i.flag_bits & 0x08 for i in archive.infolist())
    assert sheetpeek.read(io.BytesIO(data)).cell(1, 1).value == 111


def test_stored_parts_are_read(book):
    buffer = io.BytesIO()
    rewritten(book, buffer, compression=zipfile.ZIP_STORED)
    assert sheetpeek.read(io.BytesIO(buffer.getvalue())).cell(1, 1).value == 111


def test_zip64_is_read(book, monkeypatch):
    # Zip64 fields for every size and offset, without writing gigabytes.
    monkeypatch.setattr(zipfile, "ZIP64_LIMIT", 16)
    buffer = io.BytesIO()
    with zipfile.ZipFile(book) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            with out.open(info.filename, "w", force_zip64=True) as part:
                part.write(source.read(info.filename))
    monkeypatch.undo()
    data = buffer.getvalue()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert all(i.extra for i in archive.infolist() if i.file_size > 16)
    assert sheetpeek.read(io.BytesIO(data)).cell(1, 1).value == 111


def test_a_local_zip64_extra_is_read(book):
    buffer = io.BytesIO()
    with zipfile.ZipFile(book) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            with out.open(info.filename, "w", force_zip64=True) as part:
                part.write(source.read(info.filename))
    assert sheetpeek.read(io.BytesIO(buffer.getvalue())).cell(1, 1).value == 111


def test_a_prefixed_archive_is_read(book):
    data = b"#!/bin/sh\nexit 0\n" + book.read_bytes()
    assert sheetpeek.read(io.BytesIO(data)).cell(1, 1).value == 111


def test_an_empty_deflated_part_is_read(book):
    replace_part(book, "xl/empty.xml", b"")
    assert sheetpeek.read(book).cell(1, 1).value == 111


def test_a_part_that_compresses_very_well_is_read(book):
    # Its output comes out in many calls from little input, the last with none left.
    replace_part(book, "xl/media/blank.bin", b"\0" * (16 << 20))
    assert sheetpeek.read(book).cell(1, 1).value == 111


def test_every_committed_workbook_passes():
    files = sorted((Path(__file__).parent / "files").glob("*.xlsx"))
    assert files
    for path in files:
        with sheetpeek.open(path) as book:
            assert book.sheet_names


# Numbers of rich-text runs, in the shared strings: a run's size, and its color's index and tint.
ODD_RUN_FIELDS = {
    "size": '<sz val="{}"/>',
    "color index": '<color indexed="{}"/>',
    "color theme": '<color theme="{}"/>',
    "color tint": '<color theme="4" tint="{}"/>',
}


@pytest.mark.parametrize("field", ODD_RUN_FIELDS)
@bounded
def test_no_mode_allocates_or_fails_by_an_odd_run_number(tmp_path, field):
    # One bounded child for each field, trying every odd number (as above).
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    wb.active["A1"].alignment = Alignment(wrap_text=True, shrink_to_fit=True)
    wb.active["A2"] = "x"
    plain = save(wb, tmp_path / "plain.xlsx").read_bytes()
    for value in ODD_NUMBERS:
        path = tmp_path / "odd.xlsx"
        path.write_bytes(plain)
        properties = ODD_RUN_FIELDS[field].format(value)
        runs = (
            f"<r><rPr>{properties}</rPr><t>a run of words long enough to wrap</t></r>"
            '<r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>'
        )
        shared_strings(path, {"A1": runs, "A2": runs})
        try:
            sheet = sheetpeek.read(path)
        except PackageError:
            continue
        for mode in ("excel", "characters", "none"):
            page = sheetpeek.render(
                sheet.layout(mode),
                {"excel": "html", "characters": "text"}.get(mode, "csv"),
                on_notice="ignore",
            )
            assert "nan" not in page.lower().replace("sans", ""), (field, value, mode)
        sheetpeek.render(sheet, "markdown", on_notice="ignore")
