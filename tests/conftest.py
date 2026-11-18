import importlib.util
import re
from pathlib import Path

import pytest

from tests.capture_report import REPORT
from tests.watchdog import watch_session


def _needing_openpyxl() -> list[str]:
    """Test modules that build workbooks with openpyxl, directly or through another test module.

    openpyxl is a development dependency: the sdist's tests run without it, so these are left
    out there (``pytest_report_header`` says how many).
    """
    sources = {p.stem: p.read_text(encoding="utf-8") for p in Path(__file__).parent.glob("*.py")}
    uses = {
        name: set(re.findall(r"^\s*from tests\.(\w+) import", text, re.M))
        for name, text in sources.items()
    }
    needs = {
        name
        for name, text in sources.items()
        if re.search(r"^\s*(import|from) openpyxl", text, re.M)
    }
    while more := {name for name in sources if name not in needs and uses[name] & needs}:
        needs |= more
    return sorted(f"{name}.py" for name in needs if name.startswith("test_"))


collect_ignore = [] if importlib.util.find_spec("openpyxl") else _needing_openpyxl()


def pytest_report_header() -> str | None:
    if collect_ignore:
        return f"openpyxl is not installed: {len(collect_ignore)} test modules that use it left out"
    return None


def pytest_sessionstart(session: pytest.Session) -> None:
    """A guard that regresses must cost a failed run, not the machine's memory."""
    watch_session()


def pytest_terminal_summary(terminalreporter: pytest.TerminalReporter) -> None:
    """Report the rows compared with Excel and the known deviations, on every run."""
    if REPORT:
        terminalreporter.section("captures")
        for line in REPORT.values():
            terminalreporter.write_line(line)
