"""Build test workbooks with openpyxl, and rewrite their XML parts where openpyxl cannot.

openpyxl writes no cached values for formulas, no error cells, no raw serials and always a
correct ``<dimension>``; the tests that need those edit the written XML directly.
"""

from __future__ import annotations

import re
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
                text = data.decode()
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
