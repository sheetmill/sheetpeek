"""``capture/job.py``: the jobs it writes, and the suite files a result gives, without Excel.

The reader runs on a result made from the Mac's committed captures, so that its files must hold
the Mac's rows exactly: the grid a job returns and the text the Mac's AppleScript returns are
written by the same code.
"""

import hashlib
import json

import pytest

pytest.importorskip("openpyxl")

from capture import cell_text, job

CAPTURES = job.ROOT / "tests" / "captures"


def _sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_writing_a_job_again_gives_the_same_bytes(tmp_path):
    first = job.write("cells", "sheetpeek-test-cells", tmp_path / "a")
    second = job.write("cells", "sheetpeek-test-cells", tmp_path / "b")
    for name in ("job.json", "files/cell_text.xlsx"):
        assert _sha(first / name) == _sha(second / name), name
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    assert doc["package"] == "sheetpeek"
    assert doc["purpose"].startswith("sheetpeek capture group cells")
    (workbook,) = doc["workbooks"]
    assert workbook["steps"][0] == {"op": "open"}
    assert [s["range"] for s in workbook["steps"][1:]] == cell_text.specs()


def _mac_grids():
    """The text the Mac read, as the grids a job's read steps return, from the committed files."""
    rows = {}
    for suite in ("merge_width", "error_width"):
        lines = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
        rows[suite] = [line.split("\t") for line in lines if not line.startswith("# ")][1:]
    # The Mac's file goes width by width: each width's merges, then its hidden-column rows.
    by_sheet: dict[str, list[list[str]]] = {}
    for r in rows["merge_width"]:
        sheet = ("h" if r[1] == "A:C, B hidden" else "w") + r[0]
        by_sheet.setdefault(sheet, []).append([r[4]])
    grids = {}
    for k, spec in enumerate(cell_text.specs()):
        if spec.startswith("errors!"):
            width = len(cell_text.ERROR_WIDTHS)
            texts = [r[2] for r in rows["error_width"]]
            grids[f"text_{k}"] = [texts[i : i + width] for i in range(0, len(texts), width)]
        else:
            grids[f"text_{k}"] = by_sheet[spec.split("!")[0]]
    return grids


def _result(folder, job_folder, grids, *, manifest_sha=None):
    steps = [{"op": "open", "status": "ok"}]
    steps += [
        {"op": "read", "id": key, "status": "ok", "result": {"text": grid}}
        for key, grid in grids.items()
    ]
    header = {
        "manifest_sha256": manifest_sha or _sha(job_folder / "job.json"),
        "inputs": {"cell_text.xlsx": {"sha256": _sha(job_folder / "files" / "cell_text.xlsx")}},
        "environment": {"excel": {"build": "16.0.0.0"}},
    }
    outcome = {
        "status": "ok",
        "started": "2026-11-23T00:00:00Z",
        "header": header,
        "workbooks": [{"file": "cell_text.xlsx", "status": "ok", "steps": steps}],
    }
    folder.mkdir(parents=True)
    (folder / "result.json").write_text(json.dumps(outcome), encoding="utf-8")
    return folder


def test_a_result_reads_into_the_macs_schema_and_rows(tmp_path):
    job_folder = job.write("cells", "sheetpeek-test-cells", tmp_path / "jobs")
    result = _result(tmp_path / "result", job_folder, _mac_grids())
    out = tmp_path / "windows"
    written = job.read("cells", job_folder, result, out)
    assert [p.name for p in written] == ["merge_width.tsv", "error_width.tsv"]
    for path in written:
        ours = [x for x in path.read_text(encoding="utf-8").splitlines() if not x.startswith("# ")]
        theirs = (CAPTURES / path.name).read_text(encoding="utf-8").splitlines()
        assert ours == [x for x in theirs if not x.startswith("# ")], path.name
        assert "# platform: windows" in path.read_text(encoding="utf-8")
        assert "# excel_build: 16.0.0.0" in path.read_text(encoding="utf-8")
    assert (out / "cells.job.json").read_bytes() == (job_folder / "job.json").read_bytes()


def test_a_result_of_another_job_is_refused(tmp_path):
    job_folder = job.write("cells", "sheetpeek-test-cells", tmp_path / "jobs")
    result = _result(tmp_path / "result", job_folder, _mac_grids(), manifest_sha="0" * 64)
    with pytest.raises(ValueError, match="manifest_sha256"):
        job.read("cells", job_folder, result, tmp_path / "windows")
