"""The capture suite ``refit_rows`` (a row's stored height without ``customHeight``, §4.7),
without Excel: the workbooks as described, and an answer read into records."""

import zipfile

import pytest
from capture import refit_rows as suite

pytest.importorskip("openpyxl")


def test_every_row_stores_its_height_without_custom_height(tmp_path):
    for path in suite.build(tmp_path):
        sheet = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "customHeight" not in sheet, path.name
        name = path.stem.removeprefix("refit_rows_")
        for case in suite.CASES[name]:
            if case.stored is not None:
                assert f'<row r="{case.row}" ht="{case.stored:g}"' in sheet, case.name


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
