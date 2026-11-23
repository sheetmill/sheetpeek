"""Every "unverified" label outside the design document is listed, so that none drifts.

A rule a committed capture pins is cited by its suite, not called unverified; a rule no capture
covers keeps the label. A new label needs an entry below, and a removed one fails until its entry
goes, so each change of the kind is deliberate. Entries are keyed by file and the line's text,
stripped, not by line number. ``ARCHITECTURE.md`` and ``CHANGELOG.md`` are left out: the first
marks unverified rules where it states them, the second records history.
"""

import re
import subprocess
from collections import Counter
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
WORD = re.compile("unverified", re.IGNORECASE)
LEFT_OUT = {"ARCHITECTURE.md", "CHANGELOG.md", "tests/test_unverified.py"}
DATA = (".tsv", ".tsv.gz", ".xlsx")

LABELS: dict[str, tuple[str, ...]] = {
    "capture/extensions.py": (
        "are matched ignoring case until captured (unverified), and ``_detect._URIS_ANY_CASE`` "
        "follows",
        "timelines need pivot or table caches; their URIs stay unverified.",
    ),
    "capture/package_more.py": (
        '"Numbers no column, row or font can have", unverified; this suite asks them in the same '
        "way.",
    ),
    "src/sheetpeek/_census.py": ("unverified, so any side counts.",),
    "src/sheetpeek/_detect.py": (
        "# are matched ignoring case (unverified), so that none is missed: the suite extensions "
        "settles",
    ),
    "src/sheetpeek/_excel.py": (
        "degrees at the sizes xlformat measures; the 7 held below 9 pt is unverified).",
        "else:  # bottom; justify as bottom (unverified); a box taller than the row (captured)",
        "# An accounting underline runs under the text's width, turned with it (unverified).",
        "also leaves out other cells' runs (unverified).",
        "(unverified).",
    ),
    "src/sheetpeek/_rows.py": (
        "if font.size > _LARGEST:  # 72's scaled with the size (unverified; 409 caps at "
        "TALLEST_ROW)",
        "measures it as Calibri. Past 72 points, 72's height scaled with the size (unverified), at",
    ),
    "src/sheetpeek/_style/font.py": ("taken alike (unverified).",),
    "src/sheetpeek/_text_styles.py": (
        "# size, 1 point thick (unverified).",
        'at Calibri 11 at zoom 100; the step at other sizes is unverified)."""',
        "never below 0 (unverified).",
    ),
    "src/sheetpeek/backends/html/_headings.py": (
        "# segment's (unverified).",
        "# digit. At other sizes the digit term scales with the size (unverified).",
    ),
    "src/sheetpeek/backends/html/_text.py": (
        "gap apart (unverified past zoom 100). The rule of an angled run turns with it, about",
    ),
    "src/sheetpeek/backends/html/_workbook.py": ('unverified)."""',),
    "tests/test_theme.py": (
        "# the line through the two nearest (unverified), positive at every size Excel accepts.",
    ),
}
"""The lines that say "unverified" (any case), by file, each line's text stripped."""


def tracked() -> list[str]:
    """The tracked files the guard reads, relative to the repository; skipped outside git."""
    try:
        listed = subprocess.run(
            ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True
        ).stdout.splitlines()
    except (OSError, subprocess.CalledProcessError):
        pytest.skip("not a git checkout: the guard reads the tracked files")
    return sorted(
        name
        for name in listed
        if name not in LEFT_OUT and not name.endswith(DATA) and (ROOT / name).is_file()
    )


def labels() -> dict[str, Counter[str]]:
    """Every line of a tracked text file that says "unverified", by file."""
    found: dict[str, Counter[str]] = {}
    for name in tracked():
        try:
            text = (ROOT / name).read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue  # a binary file
        lines = [line.strip() for line in text.splitlines() if WORD.search(line)]
        if lines:
            found[name] = Counter(lines)
    return found


def test_every_unverified_label_is_listed():
    found = labels()
    listed = {name: Counter(lines) for name, lines in LABELS.items()}
    new = [
        f"{name}: {line}"
        for name, lines in found.items()
        for line in lines - listed.get(name, Counter())
    ]
    gone = [
        f"{name}: {line}"
        for name, lines in listed.items()
        for line in lines - found.get(name, Counter())
    ]
    assert new == [], "a new unverified label: cite its capture, or list it here"
    assert gone == [], "a listed label is gone: take its entry out"


def test_each_listed_line_says_unverified():
    for name, lines in LABELS.items():
        assert all(WORD.search(line) for line in lines), name
        assert name not in LEFT_OUT, name
