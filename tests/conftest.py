import pytest

from tests.capture_report import REPORT
from tests.watchdog import watch_session


def pytest_sessionstart(session: pytest.Session) -> None:
    """A guard that regresses must cost a failed run, not the machine's memory."""
    watch_session()


def pytest_terminal_summary(terminalreporter: pytest.TerminalReporter) -> None:
    """Report the rows compared with Excel and the known deviations, on every run."""
    if REPORT:
        terminalreporter.section("captures")
        for line in REPORT.values():
            terminalreporter.write_line(line)
