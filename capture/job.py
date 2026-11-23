"""Capture suites as Excel capture jobs, and the suite files their results give.

    uv run python -m capture.job write GROUP JOB_ID OUT_DIR
    uv run python -m capture.job read GROUP JOB_DIR RESULT_DIR

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/e3a5207c251929a8a19d713777b70ace3506fb95/docs/capture-jobs.md
A group (``GROUPS``) is one workbook and the suites read from it. ``write`` makes the job folder:
the manifest (``job.json``) and the workbook in ``files/``, built as the suite's module builds it
for Excel for Mac, with its zip entries' times and document dates fixed, so that writing it again
gives the same bytes and the SHA-256 the result records. ``read`` checks the manifest's and the
input's SHA-256 against the result's and writes each suite's file in
``tests/captures/windows/`` in the schema of the Mac's, under a header naming the platform, the
Excel build and the job, with the result's ``header`` whole as one line of JSON
(``job_header``); the manifest is copied beside them, byte for byte, as ``<group>.job.json``.

Groups:

- ``cells``: ``capture/cell_text.py``'s workbook (suites ``merge_width`` and ``error_width``),
  opened, and the text each cell shows at its width read (``read`` of ``text``, one cell at a
  time), as the Mac's ``string value`` reads it.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from capture import cell_text

ROOT = Path(__file__).resolve().parents[1]
WINDOWS = ROOT / "tests" / "captures" / "windows"
FIXED_TIME = "2026-01-01T00:00:00Z"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
GROUPS = {"cells": ("merge_width", "error_width")}
PURPOSES = {
    "cells": "sheetpeek capture group cells: the text Excel shows in merged cells and for "
    "error values at set column widths (suites merge_width and error_width)",
}


def fixed_copy(raw: Path, path: Path) -> None:
    """``raw`` rewritten to ``path`` with fixed zip times and document dates."""
    with zipfile.ZipFile(raw) as source, zipfile.ZipFile(path, "w") as out:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "docProps/core.xml":
                text = re.sub(
                    r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(<)",
                    rf"\g<1>{FIXED_TIME}\2",
                    data.decode("utf-8"),
                )
                data = text.encode("utf-8")
            entry = zipfile.ZipInfo(info.filename, date_time=ZIP_TIME)
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr = 0o600 << 16
            out.writestr(entry, data)


def _cells_workbook(files: Path) -> dict[str, Any]:
    with tempfile.TemporaryDirectory() as tmp:
        raw = cell_text.build(Path(tmp) / "raw.xlsx")
        fixed_copy(raw, files / "cell_text.xlsx")
    steps: list[dict[str, Any]] = [{"op": "open"}]
    for k, spec in enumerate(cell_text.specs()):
        steps.append({"op": "read", "id": f"text_{k}", "range": spec, "props": ["text"]})
    return {"file": "cell_text.xlsx", "steps": steps}


def manifest(group: str, job_id: str, workbooks: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "sheetpeek",
        "purpose": PURPOSES[group],
        "limits": {"job_seconds": 1800, "step_seconds": 600},
        "audit": False,
        "workbooks": workbooks,
    }


def write(group: str, job_id: str, out: Path) -> Path:
    """The job folder ``out/<job_id>``."""
    if group not in GROUPS:
        sys.exit(f"unknown group {group!r}; expected one of {', '.join(GROUPS)}")
    folder = out / job_id
    files = folder / "files"
    files.mkdir(parents=True)
    doc = manifest(group, job_id, [_cells_workbook(files)])
    (folder / "job.json").write_text(json.dumps(doc, indent=1) + "\n", encoding="utf-8")
    return folder


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _results(workbook: dict[str, Any]) -> dict[str, Any]:
    if workbook.get("status") != "ok":
        raise ValueError(f"{workbook.get('file')}: workbook status {workbook.get('status')!r}")
    found = {}
    for step in workbook["steps"]:
        if step.get("status") != "ok":
            raise ValueError(f"{workbook.get('file')}: step {step.get('op')}: {step.get('error')}")
        if "id" in step:
            found[step["id"]] = step.get("result")
    return found


def read(group: str, job: Path, result: Path, folder: Path = WINDOWS) -> list[Path]:
    """The suite files from a job folder and its result folder (``result.json``)."""
    outcome = json.loads((result / "result.json").read_text(encoding="utf-8"))
    if outcome.get("status") != "ok":
        raise ValueError(f"the job did not end ok: {outcome.get('status')!r}")
    header = outcome["header"]
    if header["manifest_sha256"] != _sha256(job / "job.json"):
        raise ValueError("the result's manifest_sha256 is not this job's manifest")
    for name, entry in header["inputs"].items():
        if entry["sha256"] != _sha256(job / "files" / name):
            raise ValueError(f"the result's input {name} is not this job's")
    manifest_doc = json.loads((job / "job.json").read_text(encoding="utf-8"))
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    results = _results(by_file["cell_text.xlsx"])
    shown: dict[tuple[str, int, int], str] = {}
    for k, spec in enumerate(cell_text.specs()):
        for r, row in enumerate(results[f"text_{k}"]["text"], 1):
            for c, text in enumerate(row, 1):
                shown[(spec, r, c)] = str(text)
    excel = header.get("excel", {})
    method = "\n".join(
        (
            "# platform: windows",
            f"# job: {group}.job.json",
            f"# job_id: {manifest_doc['job_id']}",
            f"# captured: {str(outcome.get('started', ''))[:10]}; method: read text (capture job)",
            f"# job_header: {json.dumps(header, sort_keys=True, ensure_ascii=False)}",
        )
    )
    folder.mkdir(parents=True, exist_ok=True)
    written = cell_text.write_tsvs(
        shown, f"Excel {excel.get('build', '')} for Windows", method, folder
    )
    shutil.copyfile(job / "job.json", folder / f"{group}.job.json")
    return written


def main(argv: list[str]) -> int:
    if len(argv) == 4 and argv[0] == "write":
        print(write(argv[1], argv[2], Path(argv[3])))
        return 0
    if len(argv) == 4 and argv[0] == "read":
        for path in read(argv[1], Path(argv[2]), Path(argv[3])):
            print(path)
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
