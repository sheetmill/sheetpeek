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


@pytest.mark.excel
def test_excel_saves_a_fitted_row_as_its_fit_without_custom_height():
    # Excel for Mac writes the height it fitted a row to as the row's ht, without customHeight
    # (none for a row at the default height); keeps a custom row's ht and customHeight; and fits
    # a cell in "arial" as Arial (25), not as an unknown face (Calibri 20: 26). sheetpeek's rows
    # are Excel's, but the Tahoma row: sheetpeek keeps the 15 it stores (xlformat has no Tahoma
    # rows), where Excel for Mac fits 16.
    import tempfile
    from pathlib import Path

    import sheetpeek

    lines = suite.OUTPUT.read_text(encoding="utf-8").splitlines()
    lines = [x for x in lines if not x.startswith("# ")]
    header = lines[0].split("\t")
    rows = [dict(zip(header, x.split("\t"), strict=True)) for x in lines[1:]]
    assert [(r["saved_ht"], r["saved_custom"]) for r in rows] == [
        ("", ""), ("26", ""), ("37", ""), ("25", ""), ("25", ""), ("16", ""), ("30", "1"),
    ]  # fmt: skip
    with tempfile.TemporaryDirectory() as tmp:
        layout = sheetpeek.read(suite.build(Path(tmp)), platform="mac").layout("excel")
    ours = [r.height_pt for r in layout.rows]
    excel = [float(r["saved_ht"] or 15) for r in rows]
    assert [k for k, (a, b) in enumerate(zip(ours, excel, strict=True), 1) if a != b] == [6]
    assert ours[5] == 15.0  # the stored height, kept
