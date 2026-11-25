"""Time reading every sheet of a workbook, and its peak memory, in a recalculation mode.

    python bench/read.py WORKBOOK [excel|always|never] [--trace]

Each sheet is read as a ``Sheet`` and dropped before the next, in one process, with
``recalc_on_load`` as given (``excel`` by default); a sheet that cannot be read is skipped and
named. Peak memory is the process's maximum resident size. ``--trace`` also gives, from
``tracemalloc``, the largest peak of Python's live allocations while a sheet is read, and that
sheet (tracing slows reading several times, so its time is not comparable). ``ARCHITECTURE.md``
§5.2, §5.4 and §5.7 give the figures for the workbook ``large_workbook.py`` writes; run it in a
fresh process for each measurement, best of three.
"""

from __future__ import annotations

import resource
import sys
import time
import tracemalloc

import sheetpeek


def main(path: str, mode: str, trace: bool) -> None:
    cells = sheets = 0
    largest, traced = "", 0
    start = time.perf_counter()
    with sheetpeek.open(path, recalc_on_load=mode) as book:
        for name in book.sheet_names:
            if trace:
                tracemalloc.start()
            try:
                sheet = book.sheet(name)
            except Exception as error:
                print(f"skipped reading {name!r}: {type(error).__name__}", file=sys.stderr)
                continue
            if trace:
                if tracemalloc.get_traced_memory()[1] > traced:
                    largest, traced = name, tracemalloc.get_traced_memory()[1]
                tracemalloc.stop()
            cells += sum(len(row) for row in sheet.cells)
            sheets += 1
            del sheet
    seconds = time.perf_counter() - start
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    peak_mib = peak / 2**20 if sys.platform == "darwin" else peak / 2**10
    line = (
        f"sheetpeek {sheetpeek.__version__} {mode}: {sheets} sheets, {cells:,} cells, "
        f"read {seconds:.2f} s, peak {peak_mib:.0f} MiB"
    )
    if trace:
        line += f"; live allocations peak {traced / 2**20:.0f} MiB, reading {largest!r}"
    print(line)


if __name__ == "__main__":
    arguments = [a for a in sys.argv[1:] if a != "--trace"]
    main(arguments[0], arguments[1] if len(arguments) > 1 else "excel", "--trace" in sys.argv)
