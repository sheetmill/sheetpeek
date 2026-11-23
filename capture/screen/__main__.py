"""Run screen suites one at a time, in the order given.

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run SUITE ...
    uv run --group capture python -m capture.screen read SUITE ...   # saved captures, no Excel
    uv run --group capture python -m capture.screen geometry show|record   # this machine's geometry

A suite is a module ``capture.<suite>`` with ``run()`` (build, capture, write its TSV) and
``write()`` (measure its saved captures again). One Excel, one suite at a time: a suite that
fails stops the run, and the next ones are not started.
"""

from __future__ import annotations

import importlib
import sys

SUITES = (
    "rotation",
    "borders",
    "fills",
    "text_style",
    "filters",
    "tables",
    "links",
    "outline",
    "headings_grid",
    "theme_fonts_palette",
    "tab_strip",
    "runs",
    "gradients",
    "auto_colors",
    "border_conflicts",
    "merges_spill",
    "diagonals",
    "freeze_view",
    "spill_neighbors",
    "text_style_more",
    "rotation_more",
    "rotation_placement",
    "filters_more",
    "extensions",
    "text_steps",
    "row_bottoms",
)


def _geometry(action: str) -> int:
    """Show this machine's geometry (Excel's window maximized) beside the recorded one, or
    record it."""
    import json
    import os

    from openpyxl import Workbook

    from capture.screen import geometry
    from capture.screen.session import CONTAINER, ExcelSession

    folder = CONTAINER / "sheetpeek-capture"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"geometry_{os.getpid()}.xlsx"
    Workbook().save(path)
    try:
        with ExcelSession(path, evidence=folder / "excel_timeouts") as session:
            found = geometry.normalized(session.geometry())
    finally:
        path.unlink(missing_ok=True)
    print("found:", json.dumps(found, sort_keys=True))
    if action == "record":
        print("recorded in", geometry.record(found))
        return 0
    try:
        wanted = geometry.expected()
    except geometry.GeometryError as e:
        print(e)
        return 1
    differ = geometry.differences(found, wanted)
    print("expected:", json.dumps(wanted, sort_keys=True))
    print("differences:", "; ".join(differ) or "none")
    return 1 if differ else 0


def main(argv: list[str]) -> int:
    if argv[:1] == ["geometry"] and argv[1:] in (["show"], ["record"]):
        return _geometry(argv[1])
    if len(argv) < 2 or argv[0] not in ("run", "read") or any(s not in SUITES for s in argv[1:]):
        print(__doc__.split("\n\n")[1], file=sys.stderr)
        print(f"suites: {', '.join(SUITES)}", file=sys.stderr)
        return 2
    for name in argv[1:]:
        suite = importlib.import_module(f"capture.{name}")
        print(f"{argv[0]}: {name}", flush=True)
        suite.run() if argv[0] == "run" else suite.write()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
