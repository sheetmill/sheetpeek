"""Time the excel-mode layout of every sheet of a workbook, read time apart, and its peak memory.

    python bench/layout.py WORKBOOK [LABEL]

Each sheet is read, laid out in ``excel`` mode and dropped before the next, in one process; a
sheet that cannot be read or laid out is skipped and named. Peak memory is the process's maximum
resident size (bytes on macOS, kilobytes on Linux). ``ARCHITECTURE.md`` §5.2 gives the figures
for the workbook ``large_workbook.py`` writes; run it in a fresh process for each measurement,
best of three.
"""

from __future__ import annotations

import resource
import sys
import time

import sheetpeek


def main(path: str, label: str) -> None:
    rows = sheets = 0
    read_s = layout_s = 0.0
    with sheetpeek.open(path) as book:
        for name in book.sheet_names:
            start = time.perf_counter()
            try:
                sheet = book.sheet(name)
            except Exception as error:
                print(f"skipped reading {name!r}: {type(error).__name__}", file=sys.stderr)
                continue
            read = time.perf_counter()
            try:
                rendered = sheet.layout("excel")
            except Exception as error:
                print(f"skipped laying out {name!r}: {type(error).__name__}", file=sys.stderr)
                continue
            done = time.perf_counter()
            read_s += read - start
            layout_s += done - read
            rows += len(rendered.rows)
            sheets += 1
            del sheet, rendered
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    peak_mib = peak / 2**20 if sys.platform == "darwin" else peak / 2**10
    print(
        f"{label} sheetpeek {sheetpeek.__version__}: {sheets} sheets, {rows:,} rows, "
        f"read {read_s:.2f} s, layout {layout_s:.2f} s, peak {peak_mib:.0f} MiB"
    )


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "")
