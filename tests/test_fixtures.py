"""Committed documents carry no personal data that Office saves into a file."""

import re
import subprocess
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SUFFIXES = (".xlsx", ".docx", ".pptx")
# Excel saves the folder a workbook was saved in (x15ac:absPath in xl/workbook.xml) and the user
# who saved it (cp:lastModifiedBy in docProps/core.xml).
MARKERS = (b"absPath", b"lastModifiedBy")
# Elements of docProps/app.xml that name an organization or a person when filled in.
FILLED = re.compile(rb"<(Company|Manager)>[^<]+</")
# Parts holding comments name their authors: Excel's comments and threaded comments and their
# persons, Word's and PowerPoint's comments and comment authors.
COMMENTS = re.compile(r"(^|/)(comments?\d*|threadedComment\w*|persons?|commentAuthors)\.xml$", re.I)


def documents() -> list[Path]:
    """Every committed Office document: the tracked ones, or under an sdist the files it holds."""
    try:
        listed = subprocess.run(
            ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True
        ).stdout.split()
        files = [ROOT / name for name in listed]
    except (OSError, subprocess.CalledProcessError):
        files = [p for p in ROOT.rglob("*") if ".venv" not in p.parts]
    return sorted(p for p in files if p.suffix.lower() in SUFFIXES and p.is_file())


DOCUMENTS = documents()


def test_there_are_documents():
    assert DOCUMENTS


@pytest.mark.parametrize("path", DOCUMENTS, ids=lambda p: p.relative_to(ROOT).as_posix())
def test_no_personal_data(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert "docProps/custom.xml" not in names, "custom document properties"
        for name in names:
            assert not COMMENTS.search(name), f"{name}: comments name their authors"
            data = archive.read(name)
            for marker in MARKERS:
                assert marker not in data, f"{name} contains {marker.decode()}"
            assert not FILLED.search(data), f"{name} names a company or manager"


CAPTURES = sorted((ROOT / "tests" / "captures").glob("*.tsv*"))
CONTROL = re.compile(r"[\x00-\x08\x0b-\x1f\x7f-\x9f]")
"""Control characters other than tab and newline, C1 included: git and GitHub take a file
holding one for binary."""


def control_lines(text: str) -> list[tuple[int, str]]:
    """The lines of ``text`` that hold a control character, numbered from 1.

    Split on newline alone: ``str.splitlines`` also splits on carriage return, form feed, the
    file, group and record separators and others, which would never reach the pattern.
    """
    return [
        (number, line[:60])
        for number, line in enumerate(text.split("\n"), 1)
        if CONTROL.search(line)
    ]


@pytest.mark.parametrize("path", CAPTURES, ids=lambda p: p.name)
def test_a_text_capture_holds_no_control_character(path):
    import gzip

    data = gzip.decompress(path.read_bytes()) if path.suffix == ".gz" else path.read_bytes()
    found = control_lines(data.decode("utf-8"))
    assert not found, found


@pytest.mark.parametrize(
    "char",
    ["\x00", "\x0b", "\x0c", "\r", "\x1c", "\x1d", "\x1e", "\x1f", "\x7f", "\x80", "\x85", "\x9f"],
)
def test_every_control_character_is_found(char):
    text = f"case\tvalue\nrow\t1{char}2\tdone\n"
    assert control_lines(text) == [(2, f"row\t1{char}2\tdone")]


def test_tab_newline_and_other_text_pass():
    assert control_lines("a\tb\nc\u00a0d\u2028e\u2029f\n") == []
