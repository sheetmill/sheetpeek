"""Nothing in the package imports openpyxl, and only the entry points import the reader."""

import ast
from pathlib import Path

PACKAGE = Path(__file__).resolve().parents[1] / "src" / "sheetpeek"
READER = PACKAGE / "_read.py"


def imported_modules(path: Path) -> set[str]:
    names = set()
    for node in ast.walk(ast.parse(path.read_text())):
        if isinstance(node, ast.Import):
            names.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            names.add(node.module)
    return names


def test_nothing_imports_openpyxl():
    # The package reads through xlio; openpyxl only builds workbooks in the tests.
    importers = {
        path.relative_to(PACKAGE).as_posix()
        for path in PACKAGE.rglob("*.py")
        if any(name.split(".")[0] == "openpyxl" for name in imported_modules(path))
    }
    assert importers == set()


def test_nothing_but_the_package_entry_points_import_the_reader():
    importers = {
        path.relative_to(PACKAGE).as_posix()
        for path in PACKAGE.rglob("*.py")
        if "sheetpeek._read" in imported_modules(path)
    }
    # the entry points: the public API, the command line, render(), render_workbook(), and the
    # server
    assert importers <= {
        "__init__.py",
        "cli.py",
        "backends/__init__.py",
        "backends/html/_workbook.py",
        "_serve.py",
    }
