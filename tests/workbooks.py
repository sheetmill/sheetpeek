"""Build test workbooks with openpyxl, and rewrite their XML parts where openpyxl cannot.

openpyxl writes no cached values for formulas, no error cells, no raw serials and always a
correct ``<dimension>``; the tests that need those edit the written XML directly.
"""

from __future__ import annotations

import re
import struct
import zipfile
from collections.abc import Callable
from pathlib import Path

import openpyxl

SHEET1 = "xl/worksheets/sheet1.xml"
NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def save(workbook: openpyxl.Workbook, path: Path) -> Path:
    workbook.save(path)
    return path


def rewrite(path: Path, part: str, edit: Callable[[str], str]) -> Path:
    """Replace one part of the package with ``edit(text)``; the edit must change something."""
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == part:
                # lxml, when installed, writes <v/> or <v></v> where the standard library writes
                # <v />, and the edits name the standard library's form.
                text = re.sub(r"(?<! )/>", " />", data.decode()).replace("<v></v>", "<v />")
                changed = edit(text)
                assert changed != text, f"edit of {part} changed nothing"
                data = changed.encode()
            archive.writestr(info, data)
    return path


def replace_once(old: str, new: str) -> Callable[[str], str]:
    """An edit replacing exactly one occurrence of ``old``."""

    def edit(text: str) -> str:
        assert text.count(old) == 1, f"{old!r} occurs {text.count(old)} times"
        return text.replace(old, new)

    return edit


def replace_sheet_data(sheet_data: str) -> Callable[[str], str]:
    """An edit replacing the whole ``<sheetData>`` element."""

    def edit(text: str) -> str:
        new, count = re.subn(
            r"<sheetData>.*</sheetData>|<sheetData\s*/>", sheet_data, text, flags=re.S
        )
        assert count == 1
        return new

    return edit


def set_cached(path: Path, coordinate: str, text: str) -> Path:
    """Write ``text`` as the cached value of one cell of the first sheet, exactly as given."""

    def edit(xml: str) -> str:
        pattern = rf'(<c r="{coordinate}"[^>]*>)<v>[^<]*</v>'
        new, count = re.subn(pattern, lambda m: f"{m[1]}<v>{text}</v>", xml)
        assert count == 1, f"{coordinate} has no cached value to replace"
        return new

    return rewrite(path, SHEET1, edit)


def raw_sheet(path: Path, body: str) -> Path:
    """A one-sheet workbook whose worksheet part is exactly ``body`` inside ``<worksheet>``."""
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    save(wb, path)
    xml = f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS} {R_NS}>{body}</worksheet>'
    return rewrite(path, SHEET1, lambda _: xml)


def replace_part(path: Path, part: str, data: bytes, method: int = zipfile.ZIP_DEFLATED) -> Path:
    """Replace one part with ``data`` as it stands, in any encoding, under the given method."""
    with zipfile.ZipFile(path) as archive:
        items = [(i.filename, archive.read(i.filename)) for i in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, body in items:
            if name != part:
                archive.writestr(name, body)
        archive.writestr(part, data, compress_type=method)
    return path


def state_size(path: Path, part: str, size: int, crc: int | None = None) -> Path:
    """Change a part's size, and its CRC, in the central directory alone."""
    data = bytearray(path.read_bytes())
    with zipfile.ZipFile(path) as archive:
        position = archive.start_dir
    while data[position : position + 4] == b"PK\x01\x02":
        name_length, extra_length, comment_length = struct.unpack_from("<HHH", data, position + 28)
        if bytes(data[position + 46 : position + 46 + name_length]) == part.encode():
            struct.pack_into("<I", data, position + 24, size)
            if crc is not None:
                struct.pack_into("<I", data, position + 16, crc)
            path.write_bytes(data)
            return path
        position += 46 + name_length + extra_length + comment_length
    raise KeyError(part)


def shared_strings(path: Path, strings: dict[str, str]) -> Path:
    """Make cells of the first sheet shared strings, as Excel writes them; openpyxl writes inline.

    ``strings`` maps a coordinate to the inner XML of its ``<si>``, such as rich-text runs
    (``<r><rPr><b/></rPr><t>bold</t></r>``). Each cell must hold an inline string already.
    """
    order = list(strings)

    def cells(xml: str) -> str:
        for index, coordinate in enumerate(order):
            pattern = rf'<c r="{coordinate}"([^>]*?) t="inlineStr"([^>]*)><is>.*?</is></c>'
            xml, count = re.subn(pattern, rf'<c r="{coordinate}"\1 t="s"\2><v>{index}</v></c>', xml)
            assert count == 1, f"{coordinate} holds no inline string"
        return xml

    rewrite(path, SHEET1, cells)
    items = "".join(f"<si>{strings[c]}</si>" for c in order)
    table = f'<sst {NS} count="{len(order)}" uniqueCount="{len(order)}">{items}</sst>'
    replace_part(path, "xl/sharedStrings.xml", table.encode())
    rewrite(
        path,
        "[Content_Types].xml",
        replace_once(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" /></Types>',
        ),
    )
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    return rewrite(
        path,
        "xl/_rels/workbook.xml.rels",
        replace_once(
            "</Relationships>",
            f'<Relationship Id="rIdSst" Type="{rel}" Target="sharedStrings.xml" /></Relationships>',
        ),
    )
