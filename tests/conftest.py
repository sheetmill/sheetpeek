import pytest

from tests.capture_report import REPORT


def pytest_terminal_summary(terminalreporter: pytest.TerminalReporter) -> None:
    """Report the rows compared with Excel and the known deviations, on every run."""
    if REPORT:
        terminalreporter.section("captures")
        for line in REPORT.values():
            terminalreporter.write_line(line)
