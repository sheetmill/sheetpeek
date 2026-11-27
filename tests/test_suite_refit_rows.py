"""The capture suite ``refit_rows`` (a row's stored height without ``customHeight``, §4.7),
without Excel: the workbooks as described, and an answer read into records."""

import re
import tempfile
import zipfile
from pathlib import Path

import pytest
from capture import refit_rows as suite

pytest.importorskip("openpyxl")


def test_every_row_stores_its_height_as_its_case_says(tmp_path):
    for path in suite.build(tmp_path):
        sheet = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode("utf-8")
        name = path.stem.removeprefix("refit_rows_")
        cases = suite.CASES[name]
        rows = re.findall(r"<row [^>]*>", sheet)
        assert sum("customHeight" in r for r in rows) == sum(c.custom for c in cases), path.name
        excel = name in suite.EXCEL_NORMAL
        for case in cases:
            if case.written is not None:
                custom = ' customHeight="1"' if case.custom else ""
                descent = suite.EXCELFORM_DESCENT.get(case.size or 10, "0.2")
                spans = max(1, len(case.cells))
                before = f" {case.row_format}" if case.row_format else ""
                if "{s}" in before:
                    style = re.search(rf'<c r="A{case.row}" s="(\d+)"', sheet)
                    assert style, case.name
                    before = before.replace("{s}", style.group(1))
                element = (
                    f'<row r="{case.row}" spans="1:{spans}"{before} ht="{case.written}" '
                    f'x14ac:dyDescent="{descent}">'
                    if excel
                    else f'<row r="{case.row}" ht="{case.written}"{custom}>'
                )
                assert element in rows, case.name
            elif case.size is None:
                assert f'<row r="{case.row}"' not in sheet, case.name
        sheet_format = re.findall(r"<sheetFormatPr [^>]*>", sheet)
        expected = suite.SHEET_FORMATS.get(name, 'baseColWidth="8" defaultRowHeight="15"')
        if excel:
            expected = f'defaultRowHeight="{suite.EXCEL_NORMAL[name][2]}" x14ac:dyDescent="0.2"'
        assert sheet_format == [f"<sheetFormatPr {expected} />"], path.name


def test_the_excel_form_carries_its_attributes_and_arial_10(tmp_path):
    archive = zipfile.ZipFile(suite.build_one("excelform", tmp_path))
    assert "<AppVersion>15.0300</AppVersion>" in archive.read("docProps/app.xml").decode()
    assert 'rupBuild="14420"' in archive.read("xl/workbook.xml").decode()
    fonts = archive.read("xl/styles.xml").decode()
    assert '<font><name val="Arial" /><family val="2" /><color theme="1" /><sz val="10" />' in fonts
    assert 'mc:Ignorable="x14ac"' in archive.read("xl/worksheets/sheet1.xml").decode()


def test_case_names_are_unique():
    names = [c.name for cases in suite.CASES.values() for c in cases]
    assert len(names) == len(set(names))
    assert 'zero empty ht ""' in names
    assert "zerosheet calibri 11 no ht" in names


def test_an_answer_reads_into_records(tmp_path):
    lines = ["# excel: 16.115 (test), en_US"]
    for name, cases in suite.CASES.items():
        lines += [
            f"{suite.workbook_name(name)}\t{c.row}\tCalibri\t11\t15.0\t8.43\t48.0" for c in cases
        ]
    found = suite.records("\n".join(lines))
    assert [r["name"] for r in found] == [c.name for cs in suite.CASES.values() for c in cs]
    assert {r["row_height"] for r in found} == {"15"}
    with pytest.raises(ValueError, match="expected"):
        suite.records("\n".join(lines[:-1]))


@pytest.mark.excel
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_excel_refits_a_row_without_custom_height(tmp_path, platform):
    # Excel's height is the row's fitted height whatever it stores (15 for Calibri 11 and an
    # empty row; Calibri 20 and 28: 26 and 37 on the Mac, 26.25 and 36 on Windows), which is
    # sheetpeek's for the same row without the ht, and sheetpeek's for the row as stored.
    import openpyxl

    import sheetpeek

    capture = (
        suite.OUTPUT if platform == "mac" else suite.ROOT / "tests/captures/windows/refit_rows.tsv"
    )
    lines = [x for x in capture.read_text(encoding="utf-8").splitlines() if not x.startswith("# ")]
    header = lines[0].split("\t")
    excel = {
        r["name"]: float(r["row_height"])
        for r in (dict(zip(header, x.split("\t"), strict=True)) for x in lines[1:])
    }
    wrong, refitted = {}, {}
    for name, cases in suite.CASES.items():
        if name not in ("refit", "macsaved"):
            continue
        path = suite.build_one(name, tmp_path / name)
        stored = sheetpeek.read(path, platform=platform, ref=f"A1:A{len(cases)}").layout("excel")
        wb = openpyxl.load_workbook(path)
        for case in cases:
            wb.active.row_dimensions[case.row].height = None
        wb.save(tmp_path / f"{name}_fitted.xlsx")
        fitted = sheetpeek.read(tmp_path / f"{name}_fitted.xlsx", platform=platform).layout("excel")
        for case in cases:
            ours = next(r for r in stored.rows if r.index == case.row).height_pt
            fit = next((r for r in fitted.rows if r.index == case.row), None)
            refitted[case.name] = fit.height_pt if fit else 15.0
            if ours != excel[case.name]:
                wrong[case.name] = (excel[case.name], ours)
    assert refitted == {n: excel[n] for n in refitted}  # the rule: Excel's row is the fitted one
    assert len(refitted) == 17
    assert wrong == {}


@pytest.mark.excel
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_excel_fits_a_row_without_custom_height_in_every_sheet_format(tmp_path, platform):
    # Excel fits every row whose ht has no customHeight: a height of 0 does not hide it
    # (ht="0" and ht=""; with customHeight it does), and under a custom default height of 20 or
    # zeroHeight the row takes its fit (15 beside Calibri 11 or empty, Calibri 20 26 on the Mac
    # and 26.25 on Windows, Calibri 28 37 and 36), where a row with no ht there takes 20 or is
    # hidden; sheetpeek's rows are Excel's. The Windows rows compare by the height Excel lays
    # out (``height``).
    import sheetpeek

    capture = (
        suite.OUTPUT if platform == "mac" else suite.ROOT / "tests/captures/windows/refit_rows.tsv"
    )
    lines = [x for x in capture.read_text(encoding="utf-8").splitlines() if not x.startswith("# ")]
    header = lines[0].split("\t")
    excel = {
        r["name"]: float(r["height"] or r["row_height"])
        for r in (dict(zip(header, x.split("\t"), strict=True)) for x in lines[1:])
    }
    assert len(excel) == (38 if platform == "mac" else 33)
    fits = {"empty": 15.0, "calibri 11": 15.0}
    fits |= {"calibri 20": 26.0, "calibri 28": 37.0} if platform == "mac" else {}
    fits |= {"calibri 20": 26.25, "calibri 28": 36.0} if platform == "windows" else {}
    wrong = {}
    for name in ("zero", "customsheet", "zerosheet"):
        cases = suite.CASES[name]
        path = suite.build_one(name, tmp_path / name)
        layout = sheetpeek.read(path, platform=platform, ref=f"A1:A{len(cases)}").layout("excel")
        ours = {r.index: r.height_pt for r in layout.rows}
        for case in cases:
            cell = "empty" if case.size is None else f"calibri {case.size}"
            if case.written is not None and not case.custom:
                assert excel[case.name] == fits[cell], case.name  # fitted, whatever it stores
            if ours.get(case.row, 0.0) != excel[case.name]:
                wrong[case.name] = (excel[case.name], ours.get(case.row))
    assert wrong == {}
    assert excel["zero empty ht 0 custom"] == excel['zero calibri 20 ht "" custom'] == 0.0
    assert excel["customsheet empty unlisted"] == (20.0 if platform == "mac" else 19.5)
    assert excel["zerosheet empty unlisted"] == 0.0


@pytest.mark.excel
def test_excel_refits_rows_with_spans_and_dydescent():
    # Excel for Mac fits rows with spans and x14ac:dyDescent (AppVersion 15.0300) under an Arial 10
    # Normal font whatever they store (13.5, 12.75, 30): 13 pt, and 16 beside Arial 12, as
    # sheetpeek fits them.
    import sheetpeek

    lines = suite.OUTPUT.read_text(encoding="utf-8").splitlines()
    lines = [x for x in lines if not x.startswith("# ")]
    header = lines[0].split("\t")
    excel = {
        r["name"]: float(r["row_height"])
        for r in (dict(zip(header, x.split("\t"), strict=True)) for x in lines[1:])
        if r["workbook"] == "excelform"
    }
    assert list(excel.values()) == [13.0, 13.0, 13.0, 16.0, 13.0]
    with tempfile.TemporaryDirectory() as tmp:
        path = suite.build_one("excelform", Path(tmp))
        layout = sheetpeek.read(path, platform="mac", ref="A1:A5").layout("excel")
    cases = suite.CASES["excelform"]
    assert [r.height_pt for r in layout.rows] == [excel[c.name] for c in cases]


def test_excel8_styles_columns_b_on_in_arial_10(tmp_path):
    archive = zipfile.ZipFile(suite.build_one("excel8", tmp_path))
    sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    style = re.search(r'<cols><col min="2" max="16384" width="9.140625" style="(\d+)" />', sheet)
    assert style
    styles = archive.read("xl/styles.xml").decode()
    xfs = re.findall(r"<xf [^>]*>", styles.split("<cellXfs")[1])
    font_id = int(re.search(r'fontId="(\d+)"', xfs[int(style.group(1))]).group(1))
    assert re.findall(r"<font>.*?</font>", styles)[font_id] == (
        '<font><name val="Arial" /><family val="2" /><sz val="10" /></font>'
    )
    assert 'r="AX1"' not in sheet  # the spare cell (row 1, column 50) is gone
