"""The capture suite ``refit_rows`` (a row's stored height without ``customHeight``, §4.7),
without Excel: the workbooks as described, and an answer read into records."""

import re
import zipfile

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
        for case in cases:
            if case.written is not None:
                custom = ' customHeight="1"' if case.custom else ""
                assert f'<row r="{case.row}" ht="{case.written}"{custom}>' in rows, case.name
            elif case.size is None:
                assert f'<row r="{case.row}"' not in sheet, case.name
        sheet_format = re.findall(r"<sheetFormatPr [^>]*>", sheet)
        expected = suite.SHEET_FORMATS.get(name, 'baseColWidth="8" defaultRowHeight="15"')
        assert sheet_format == [f"<sheetFormatPr {expected} />"], path.name


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


REFIT = "Excel refits a row whose ht has no customHeight; sheetpeek keeps the stored height"
"""Why sheetpeek's row differs: every such row is a known deviation until the rule is modeled."""


@pytest.mark.excel
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_excel_refits_a_row_without_custom_height(tmp_path, platform):
    # Excel's height is the row's fitted height whatever it stores (15 for Calibri 11 and an
    # empty row; Calibri 20 and 28: 26 and 37 on the Mac, 26.25 and 36 on Windows), which is
    # sheetpeek's for the same row without the ht; sheetpeek keeps the stored ht instead.
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
    kept, refitted = set(), {}
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
                kept.add(case.name)
    assert refitted == excel  # the rule: Excel's row is the fitted one
    assert len(kept) == (9 if platform == "mac" else 11), REFIT
