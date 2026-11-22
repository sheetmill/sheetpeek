"""The custom properties' defaults in docs/html.md are the ones the stylesheets write."""

import re
from pathlib import Path

from sheetpeek.backends.html import _css, _outline, _panes

DOC = Path(__file__).resolve().parents[1] / "docs" / "html.md"
BLOCK = re.compile(r'(\.sheetpeek(\[data-color-scheme="dark"\])?) \{([^}]*)\}')
SET = re.compile(r"(--sheetpeek-[a-z-]+): *(#[0-9a-f]{6})")
ROW = re.compile(r"^\| `(--sheetpeek-[a-z-]+)` \|[^|]*\| (.*) \|$")
VALUE = re.compile(r"`([^`]+)`")


def written() -> dict[str, tuple[str, str]]:
    """Each property the stylesheets set: its value, and its value in the dark scheme."""
    light: dict[str, str] = {}
    dark: dict[str, str] = {}
    for css in (_css.CSS, _outline._CSS, _panes.CSS, _panes.BREAK_CSS):
        for block in BLOCK.finditer(css):
            for name, value in SET.findall(block[3]):
                (dark if block[2] else light)[name] = value
    # The zoom is set by the page; the stylesheet reads it with its default.
    assert "var(--sheetpeek-zoom, 1)" in _css.CSS
    light["--sheetpeek-zoom"] = "1"
    return {name: (value, dark.get(name, value)) for name, value in light.items()}


def documented() -> dict[str, tuple[str, str]]:
    """Each property in docs/html.md's table: its default, and its default in the dark scheme."""
    found = {}
    for line in DOC.read_text(encoding="utf-8").splitlines():
        row = ROW.match(line)
        if row:
            values = VALUE.findall(row[2])
            found[row[1]] = (values[0], values[1] if len(values) > 1 else values[0])
    return found


def test_every_documented_default_is_the_one_written():
    assert documented() == written()
