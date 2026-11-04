"""What Excel does with a package sheetpeek refuses or doubts: which refusals cost anything.

    uv run --group capture python -m capture.package_probe build   # write the workbooks only
    uv run --group capture python -m capture.package_probe run     # build, open in Excel, write TSV

Each case is one small workbook holding 111 in A1 of its only sheet, with one thing changed in
its worksheet part: a minimal ``<!DOCTYPE>``, with no entity and no external reference; another
compression method; the part stored twice under one name (the second holding 222); or sizes in
the central directory that are not the part's. ``package_open.applescript`` opens each one in its
own run, under a time limit, and reads A1 back. A control opens first and last: an outcome counts
only between two that open. The output is ``tests/captures/package.tsv``. New cases go before the
last control.
"""

from __future__ import annotations

import datetime as dt
import io
import struct
import subprocess
import sys
import warnings
import zipfile
import zlib
from collections.abc import Callable
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "package.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "package"
)
SHEET = "xl/worksheets/sheet1.xml"
TIME_LIMIT = "60"


def base() -> bytes:
    """A workbook holding 111 in A1, as openpyxl writes it."""
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = 111
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _rewrite(data: bytes, write_sheet: Callable[[zipfile.ZipFile, bytes], None]) -> bytes:
    """A copy of the package with its worksheet part written by ``write_sheet``."""
    out = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as source,
        zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target,
    ):
        for info in source.infolist():
            if info.filename == SHEET:
                write_sheet(target, source.read(SHEET))
            else:
                target.writestr(info.filename, source.read(info.filename))
    return out.getvalue()


def _directory_sizes(data: bytes, *, size: int, crc: int | None = None) -> bytes:
    """Rewrite the worksheet part's uncompressed size, and its CRC, in the central directory."""
    patched = bytearray(data)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        directory = archive.start_dir
    position = directory
    while patched[position : position + 4] == b"PK\x01\x02":
        name_length, extra_length, comment_length = struct.unpack_from(
            "<HHH", patched, position + 28
        )
        name = bytes(patched[position + 46 : position + 46 + name_length])
        if name == SHEET.encode():
            struct.pack_into("<I", patched, position + 24, size)
            if crc is not None:
                struct.pack_into("<I", patched, position + 16, crc)
            return bytes(patched)
        position += 46 + name_length + extra_length + comment_length
    raise KeyError(SHEET)


def control(data: bytes) -> bytes:
    return data


def doctype(data: bytes) -> bytes:
    def write(target: zipfile.ZipFile, xml: bytes) -> None:
        at = xml.index(b"<worksheet")
        target.writestr(SHEET, xml[:at] + b"<!DOCTYPE worksheet>" + xml[at:])

    return _rewrite(data, write)


def _method(method: int) -> Callable[[bytes], bytes]:
    def case(data: bytes) -> bytes:
        return _rewrite(data, lambda target, xml: target.writestr(SHEET, xml, compress_type=method))

    return case


def duplicate(data: bytes) -> bytes:
    def write(target: zipfile.ZipFile, xml: bytes) -> None:
        target.writestr(SHEET, xml)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # zipfile warns of the duplicate name
            target.writestr(SHEET, xml.replace(b"<v>111</v>", b"<v>222</v>"))

    return _rewrite(data, write)


def false_smaller(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        size = archive.getinfo(SHEET).file_size
    return _directory_sizes(data, size=size - 100)


def false_larger(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        size = archive.getinfo(SHEET).file_size
    return _directory_sizes(data, size=size + 1000)


def false_prefix(data: bytes) -> bytes:
    """The directory describes the whole part as far as its XML; 100 KB of spaces follow it."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        xml = archive.read(SHEET)
    padded = _rewrite(data, lambda target, _: target.writestr(SHEET, xml + b" " * 100_000))
    return _directory_sizes(padded, size=len(xml), crc=zlib.crc32(xml))


CASES: tuple[tuple[str, Callable[[bytes], bytes]], ...] = (
    ("control", control),
    ("doctype", doctype),
    ("bzip2", _method(zipfile.ZIP_BZIP2)),
    ("lzma", _method(zipfile.ZIP_LZMA)),
    ("duplicate_name", duplicate),
    ("size_smaller", false_smaller),
    ("size_larger", false_larger),
    ("size_prefix", false_prefix),
    ("control_again", control),
)


def build() -> list[Path]:
    """Write one workbook per case into Excel's container."""
    CONTAINER_DIR.mkdir(parents=True, exist_ok=True)
    data = base()
    paths = []
    for name, case in CASES:
        path = CONTAINER_DIR / f"package_{name}.xlsx"
        path.write_bytes(case(data))
        paths.append(path)
    return paths


def _version() -> str:
    result = subprocess.run(
        ["gtimeout", "20", "osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.stdout.strip() or "unknown"


def _open(path: Path) -> str:
    """One case in its own ``osascript``; alerts are restored even when the limit stops it."""
    script = Path(__file__).with_name("package_open.applescript")
    result = subprocess.run(
        ["gtimeout", TIME_LIMIT, "osascript", str(script), str(path)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode == 124:
        restore = Path(__file__).with_name("restore_alerts.applescript")
        subprocess.run(["gtimeout", "20", "osascript", str(restore)], check=False)
        return f"timed out after {TIME_LIMIT} s"
    if result.returncode != 0:
        raise SystemExit(f"{path.name}: {result.stderr.strip()}")
    return result.stdout.strip()


def run() -> None:
    """Open every case, print the outcome and write the TSV."""
    paths = build()
    version = _version()
    rows = []
    for (name, _), path in zip(CASES, paths, strict=True):
        outcome = _open(path)
        print(f"{name}\t{outcome}")
        rows.append(f"{name}\t{outcome}\n")
    with OUTPUT.open("w", encoding="utf-8") as out:
        out.write(f"# suite: package (capture.package_probe); Excel {version} for Mac, en_US\n")
        out.write(f"# captured: {dt.date.today().isoformat()}; method: open with alerts off, ")
        out.write("read A1 (111 unless the case changes it) and the number of sheets\n")
        out.write("case\toutcome\n")
        out.writelines(rows)
    print(f"{len(rows)} rows written to {OUTPUT}")


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
