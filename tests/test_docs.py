"""The design document's references resolve: section numbers and names in the package."""

import ast
import re
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "src" / "sheetpeek"
ARCHITECTURE = (ROOT / "ARCHITECTURE.md").read_text(encoding="utf-8")
# Two levels at most in ARCHITECTURE.md; a reference with three or more parts, such as
# ECMA-376's §18.8.19, cites a standard.
SECTION = re.compile(r"§(\d+(?:\.\d+)?[a-z]?)(?![\d.]*\d)")


def tracked() -> list[Path]:
    """The repository's text files: the tracked ones, or under an sdist the files it holds."""
    try:
        listed = subprocess.run(
            ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True
        ).stdout.split()
        files = [ROOT / name for name in listed]
    except (OSError, subprocess.CalledProcessError):
        files = [p for p in ROOT.rglob("*") if ".venv" not in p.parts and "_build" not in p.parts]
    suffixes = {".md", ".py", ".applescript", ".yml", ".yaml", ".toml", ".html"}
    return sorted(p for p in files if p.suffix in suffixes and p.is_file())


def headings() -> set[str]:
    found = set()
    for line in ARCHITECTURE.splitlines():
        if match := re.match(r"#{2,3} (\d+(?:\.\d+)?[a-z]?)\.? ", line):
            found.add(match[1])
    return found


def test_section_references_resolve():
    known = headings()
    assert {"3.6", "4.7", "7.3", "9", "11"} <= known
    dangling = [
        f"{path.relative_to(ROOT)}:{number}: §{ref}"
        for path in tracked()
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1)
        for ref in SECTION.findall(line)
        if ref not in known
    ]
    assert dangling == []


def names(module: Path) -> set[str]:
    """Top-level names of a module, and ``Class.member`` for each class member."""
    tree = ast.parse(module.read_text(encoding="utf-8"))
    found = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
            found.add(node.name)
        if isinstance(node, ast.ClassDef):
            found |= {
                f"{node.name}.{item.name}"
                for item in node.body
                if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef))
            }
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            found |= {t.id for t in targets if isinstance(t, ast.Name)}
    return found


@pytest.mark.parametrize(
    "reference", sorted(set(re.findall(r"`(_[a-z_]+)\.(\w+(?:\.\w+)?)`", ARCHITECTURE)))
)
def test_module_references_resolve(reference):
    module, name = reference
    if name == "py":
        assert (PACKAGE / f"{module}.py").is_file(), f"{module}.py"
    else:
        assert name in names(PACKAGE / f"{module}.py"), f"{module}.{name}"


def test_every_capture_is_listed():
    # Each file in tests/captures/ is a row of CAPTURING.md's Suites table and named in §4.9.
    capturing = (ROOT / "CAPTURING.md").read_text(encoding="utf-8")
    start = ARCHITECTURE.index("### 4.9 Captures")
    section = ARCHITECTURE[start : ARCHITECTURE.index("\n### ", start + 1)]
    for path in sorted((ROOT / "tests" / "captures").iterdir()):
        if path.is_dir():
            continue
        suite = path.name.split(".")[0]
        assert f"| `{suite}` |" in capturing, suite
        assert f"`{suite}`" in section, suite
    # Each capture from Excel for Windows and its job are rows of the Windows table.
    for path in sorted((ROOT / "tests" / "captures" / "windows").glob("*")):
        assert f"`windows/{path.name.split('.')[0]}`" in capturing, path.name


RETIRED = (
    "6 pt a level",
    "row-height table",
    "loaded_theme",
    "for now",
    "not yet rows",
    "measures the mac only so far",
    "until the platform setting",
    "as excel for mac lays",
    "captured from excel for mac by",
    "captured from excel 16 for mac, never",
)
"""Phrases for rules that have changed: an indent is three spaces of the Normal font, rows come from
xlformat's row model, the theme is read through xlio, and Excel for Windows is a platform
sheetpeek shows, not future work. The changelog keeps its history."""


def test_no_retired_phrase():
    # Whitespace is joined, so a phrase broken across lines is found too.
    left = {ROOT / "CHANGELOG.md", Path(__file__)}
    found = [
        f"{path.relative_to(ROOT)}: {phrase}"
        for path in tracked()
        if path not in left
        for text in [" ".join(path.read_text(encoding="utf-8").lower().split())]
        for phrase in RETIRED
        if phrase in text
    ]
    assert found == []


def test_a_release_requires_each_package_once():
    # One "Requires <package>" entry per package in a release's section of the changelog: two
    # would state two minimums for one release.
    import collections

    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    for section in re.split(r"^## ", text, flags=re.MULTILINE)[1:]:
        title = section.splitlines()[0]
        found = collections.Counter(re.findall(r"^- Requires (\w+)", section, flags=re.MULTILINE))
        assert [name for name, n in found.items() if n > 1] == [], title
