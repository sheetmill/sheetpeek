"""The capture suite ``color_kinds`` (ARCHITECTURE.md §3.3): which kind Excel takes from a color
that states two, in a fill, a font and a border. The workbook is built and read back with
sheetpeek; its colors are compared with ``tests/captures/color_kinds.tsv``, skipped until it
exists."""

import pytest
from capture import color_kinds as suite

import sheetpeek


def _sheetpeek(tmp_path):
    path = tmp_path / "color_kinds.xlsx"
    suite.build(path)
    sheet = sheetpeek.read(path)
    found = {}
    for cell in suite.CASES:
        column = ord(cell[0]) - 64
        model = sheet.cell(int(cell[1:]), column)
        style = model.style
        bottom = style.borders.bottom if style.borders is not None else None
        found[cell] = (
            style.fill,
            style.font.color,
            bottom.color if bottom is not None else None,
        )
    return found


def test_every_case_is_written_with_both_kinds(tmp_path):
    import zipfile

    path = tmp_path / "color_kinds.xlsx"
    suite.build(path)
    styles = zipfile.ZipFile(path).read("xl/styles.xml").decode()
    for attributes in suite.CASES.values():
        assert f"<fgColor {attributes} />" in styles, attributes
        assert styles.count(f"<color {attributes} />") == 2, attributes  # font and border


@pytest.mark.excel
def test_sheetpeek_takes_the_kind_excel_draws(tmp_path):
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.color_kinds run")
    lines = [ln for ln in suite.OUTPUT.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    rows = [dict(zip(names, ln.split("\t"), strict=True)) for ln in lines[1:]]
    captured = {r["cell"]: r for r in rows}
    ours = _sheetpeek(tmp_path)
    for cell in suite.CASES:
        record = captured[cell]
        excel = (record["fill"], record["font_color"], record["border_color"])
        assert tuple(c.upper() if c else c for c in ours[cell]) == excel, cell
