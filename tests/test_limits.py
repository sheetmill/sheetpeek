"""Untrusted files: what is refused before anything parses them, and what must still be read.

Excel's own behavior with these packages is in ``tests/captures/package.tsv`` (ARCHITECTURE.md
§5.5). The crafted files are built here, small, and never committed. A test whose guard could fail
into a huge allocation is ``bounded``: if it regresses, the test fails and the machine is fine.
"""

import io
import warnings
import zipfile
import zlib
from pathlib import Path

import openpyxl
import pytest

import sheetpeek
from sheetpeek import LimitError, PackageError, ReadLimits
from tests.watchdog import bounded
from tests.workbooks import (
    NS,
    SHEET1,
    replace_once,
    replace_part,
    replace_sheet_data,
    rewrite,
    save,
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


@pytest.mark.parametrize("value", ["99999999999", "1e30", "-inf", "-0.5"])
@bounded
def test_a_column_width_excel_was_not_shown_is_still_bounded(book, value):
    # Not captured: the nearest captured behavior, so that no mode allocates by it.
    col = f'<cols><col min="1" max="1" width="{value}" customWidth="1"/></cols><sheetData>'
    rewrite(book, SHEET1, lambda text: text.replace("<sheetData>", col, 1))
    sheet = sheetpeek.read(book)
    assert sheet.columns[0].width in (0.0, 10_922.0)
    for mode in ("excel", "characters", "none"):
        sheet.layout(mode)


@pytest.mark.parametrize(
    "edit",
    [
        pytest.param(('baseColWidth="8"', 'baseColWidth="99999999999"'), id="base width"),
        pytest.param(('baseColWidth="8"', 'baseColWidth="-1"'), id="negative base width"),
        pytest.param(("<sheetFormatPr ", '<sheetFormatPr defaultColWidth="1e30" '), id="width"),
        pytest.param(
            ('defaultRowHeight="15"', 'defaultRowHeight="-1" customHeight="1"'), id="height"
        ),
    ],
)
@bounded
def test_a_sheet_default_excel_was_not_shown_is_still_bounded(book, edit):
    rewrite(book, SHEET1, replace_once(*edit))
    sheet = sheetpeek.read(book)
    for mode in ("excel", "characters", "none"):
        sheet.layout(mode)


@pytest.mark.parametrize("size", ["-0.5", "0.5", "inf", "1e400"])
@bounded
def test_a_font_size_excel_was_not_shown_gives_no_raw_error(book, size):
    rewrite(book, "xl/styles.xml", lambda text: text.replace('<sz val="11"', f'<sz val="{size}"'))
    try:
        sheet = sheetpeek.read(book)
    except PackageError:
        return
    for mode in ("excel", "characters", "none"):
        sheet.layout(mode)


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
