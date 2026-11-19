"""The screen session's start-up checks, without Excel."""

import pytest
from capture.screen import session


def test_a_window_still_closing_is_waited_for(monkeypatch):
    shown = [[{"w": 10, "h": 10, "title": "closing"}], []]
    monkeypatch.setattr(session, "request", lambda *a, **k: "0")
    monkeypatch.setattr(session, "excel_windows", lambda: shown.pop(0) if shown else [])
    monkeypatch.setattr(session.time, "sleep", lambda s: None)
    session.ExcelSession.preflight(object.__new__(session.ExcelSession))
    assert shown == []


def test_a_window_that_stays_is_refused(monkeypatch):
    clock = iter(range(0, 1000))
    monkeypatch.setattr(session, "request", lambda *a, **k: "0")
    monkeypatch.setattr(session, "excel_windows", lambda: [{"w": 10, "h": 10, "title": "x"}])
    monkeypatch.setattr(session.time, "sleep", lambda s: None)
    monkeypatch.setattr(session.time, "monotonic", lambda: float(next(clock)))
    with pytest.raises(session.ExcelError, match="windows with no workbook open"):
        session.ExcelSession.preflight(object.__new__(session.ExcelSession))
