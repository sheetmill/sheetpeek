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
