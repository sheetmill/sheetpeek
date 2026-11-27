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
    assert len(excel) == 58
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


NO_ROW_METRICS = "xlformat has no row heights for Tahoma or Book Antiqua (font.no_metrics)"
"""Why sheetpeek's fit differs from Excel's in some Excel-form rows: it measures those faces as
Calibri, a pixel or two off."""


@pytest.mark.excel
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_excel_refits_rows_with_spans_and_dydescent(platform):
    # Excel fits rows with spans and x14ac:dyDescent (AppVersion 15.0300) under Arial 10, 11 and 8
    # Normal fonts whatever they store, as sheetpeek fits them: an Arial 10 row stored at 13.5 is
    # 13 on the Mac and 12.75 on Windows, and a row with customFormat and no s takes style 0
    # before an Arial 10 column style (Arial 8: 11 and 11.25). Rows in Tahoma or Book Antiqua
    # are known deviations.
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
    assert len(excel) == 58
    wrong, pinned = {}, set()
    with tempfile.TemporaryDirectory() as tmp:
        for name in suite.EXCEL_NORMAL:
            cases = suite.CASES[name]
            path = suite.build_one(name, Path(tmp))
            layout = sheetpeek.read(path, platform=platform, ref=f"A1:C{len(cases)}")
            ours = {r.index: r.height_pt for r in layout.layout("excel").rows}
            for case in cases:
                if ours[case.row] == excel[case.name]:
                    continue
                faces = {cell[0] for cell in case.cells}
                if faces & {"Tahoma", "Book Antiqua"}:
                    pinned.add(case.name)
                else:
                    wrong[case.name] = (excel[case.name], ours[case.row])
    assert wrong == {}
    assert len(pinned) == (5 if platform == "mac" else 7), NO_ROW_METRICS
