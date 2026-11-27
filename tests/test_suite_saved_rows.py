"""The capture suite ``saved_rows`` (the row heights Excel writes when it saves, §4.7), without
Excel: the workbook as described."""

import re
import zipfile

import pytest
from capture import saved_rows as suite

pytest.importorskip("openpyxl")


def test_each_row_is_written_as_rows_says(tmp_path):
    archive = zipfile.ZipFile(suite.build(tmp_path))
    sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    rows = {int(m.group(1)): m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*>', sheet)}
    assert sorted(rows) == list(range(1, len(suite.ROWS) + 1))
    for row, (name, _, _, _, height, custom) in enumerate(suite.ROWS, 1):
        assert ("customHeight" in rows[row]) is custom, name
        if height is None:
            assert "ht=" not in rows[row], name
        else:
            assert f'ht="{height:g}"' in rows[row], name
    fonts = re.findall(r'<name val="([^"]+)"', archive.read("xl/styles.xml").decode())
    assert "arial" in fonts
    assert "Arial" in fonts
