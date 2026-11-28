"""The headers of the captures Excel for Windows gave by a job (``tests/captures/windows/``)."""

import gzip
import hashlib
import json
from pathlib import Path

WINDOWS = Path(__file__).resolve().parent / "captures" / "windows"


def _header(path: Path) -> list[str]:
    opener = gzip.open if path.suffix == ".gz" else open
    with opener(path, "rt", encoding="utf-8") as text:
        lines = []
        for line in text:
            if not line.startswith("# "):
                break
            lines.append(line.rstrip("\n"))
    return lines


def test_every_job_capture_attests_the_manifest_committed_beside_it():
    # A capture's job_header records the sha256 of the manifest the job ran; the manifest it names
    # (# job:) must be that file, byte for byte, or the capture attests a job no commit holds.
    captures = sorted(p for p in WINDOWS.iterdir() if p.suffix in (".tsv", ".gz"))
    assert captures
    for path in captures:
        header = _header(path)
        jobs = [line.removeprefix("# job: ") for line in header if line.startswith("# job: ")]
        records = [line for line in header if line.startswith("# job_header: ")]
        assert len(jobs) == 1, path.name
        assert len(records) == 1, path.name
        recorded = json.loads(records[0].removeprefix("# job_header: "))["manifest_sha256"]
        manifest = (WINDOWS / jobs[0]).read_bytes()
        assert hashlib.sha256(manifest).hexdigest() == recorded, path.name
