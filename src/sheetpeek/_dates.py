"""Dates written as ISO 8601 text (``t="d"`` cells), as the date serials Excel reads them as.

Excel stores dates as serials: days since an epoch, with the time of day as a fraction. In the
1900 system day 1 is 1900-01-01 and day 60 is 1900-02-29, a day the calendar does not have but
Excel's 1900 system keeps, so dates from 1900-03-01 on are one more than their distance from
1899-12-30 suggests for dates before it. In the 1904 system day 0 is 1904-01-01. The forms read
are those ECMA-376 lists (18.17.4: a date, a time, a date and time, an optional ``Z``) and the
``PT..H..M..S`` durations Excel writes in strict files.
"""

from __future__ import annotations

import datetime as dt
import re

_ISO = re.compile(
    r"""
    (?P<date>(?P<year>[0-9]{4})-(?P<month>[0-9]{2})-(?P<day>[0-9]{2}))?T?
    (?P<time>(?P<hour>[0-9]{2}):(?P<minute>[0-9]{2})
        (:(?P<second>[0-9]{2})(?P<fraction>\.[0-9]{1,3})?)?)?Z?
    """,
    re.VERBOSE,
)
_DURATION = re.compile(
    r"PT((?P<hours>[0-9]+)H)?((?P<minutes>[0-9]+)M)?((?P<seconds>[0-9]+(\.[0-9]{1,3})?)S)?"
)
_SECONDS_PER_DAY = 86_400
_EPOCH_1900 = dt.date(1899, 12, 30)
_EPOCH_1904 = dt.date(1904, 1, 1)
_LEAP_DAY_1900 = 60
"""1900-02-29's serial in the 1900 system."""


def iso_serial(text: str, *, date1904: bool = False) -> float:
    """The date serial of ISO 8601 text, in the 1900 or the 1904 date system.

    Args:
        text: A date (``2024-01-02``), a time (``T10:30``, ``10:30:15.250``), both
            (``2024-01-02T10:30:00Z``), or a duration (``PT36H``).
        date1904: The workbook uses the 1904 date system.

    Raises:
        ValueError: The text is none of these, or names a date or time that does not exist
            (1900-02-29 exists only in the 1900 system).
    """
    match = _ISO.match(text)
    if match is not None and any(match.groups()):
        try:
            return _serial(match.groupdict(""), date1904)
        except ValueError as error:  # a field out of range: the date or time does not exist
            raise ValueError(f"not an ISO 8601 date or time that exists: {text[:40]!r}") from error
    match = _DURATION.match(text)
    if match is not None and any(match.groups()):
        parts = match.groupdict()
        seconds = (
            float(parts["hours"] or 0) * 3600
            + float(parts["minutes"] or 0) * 60
            + float(parts["seconds"] or 0)
        )
        return seconds / _SECONDS_PER_DAY
    raise ValueError(f"not an ISO 8601 date, time or duration: {text[:40]!r}")


def _serial(parts: dict[str, str], date1904: bool) -> float:
    """The serial of a matched date, time or both; ``ValueError`` for one that does not exist."""
    fraction = 0.0
    if parts["time"]:
        micro = int(float(parts["fraction"]) * 1_000_000) if parts["fraction"] else 0
        moment = dt.time(int(parts["hour"]), int(parts["minute"]), int(parts["second"] or 0), micro)
        fraction = (
            moment.hour * 3600 + moment.minute * 60 + moment.second + micro / 1_000_000
        ) / _SECONDS_PER_DAY
    if not parts["date"]:
        return fraction
    year, month, day = int(parts["year"]), int(parts["month"]), int(parts["day"])
    if (year, month, day) == (1900, 2, 29) and not date1904:
        return _LEAP_DAY_1900 + fraction
    days = (dt.date(year, month, day) - (_EPOCH_1904 if date1904 else _EPOCH_1900)).days
    if not date1904 and 0 < days <= _LEAP_DAY_1900:
        days -= 1  # before 1900-03-01, where the 1900 system's leap day is not counted
    return days + fraction
