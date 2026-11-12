"""Run screen suites one at a time, in the order given.

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run SUITE ...
    uv run --group capture python -m capture.screen read SUITE ...   # saved captures, no Excel

A suite is a module ``capture.<suite>`` with ``run()`` (build, capture, write its TSV) and
``write()`` (measure its saved captures again). One Excel, one suite at a time: a suite that
fails stops the run, and the next ones are not started.
"""

from __future__ import annotations

import importlib
import sys

SUITES = ("rotation", "borders", "fills", "text_style")


def main(argv: list[str]) -> int:
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
