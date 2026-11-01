"""``sheetpeek serve``: an index, one page per sheet, and nothing outside the folder."""

import http.client
import os
import threading

import openpyxl
import pytest

from sheetpeek import cli
from sheetpeek._serve import make_server, resolve
from tests.workbooks import save


@pytest.fixture
def folder(tmp_path):
    root = tmp_path / "books"
    (root / "sub").mkdir(parents=True)
    wb = openpyxl.Workbook()
    wb.active.title = "First"
    wb.active["A1"] = "hello <world>"
    wb.create_sheet("Second")["B2"] = 42
    save(wb, root / "book.xlsx")
    save(openpyxl.Workbook(), root / "sub" / "other.xlsx")
    (root / "notes.txt").write_text("not a workbook")
    (root / "broken.xlsx").write_text("not a zip")
    (root / ".hidden.xlsx").write_bytes((root / "book.xlsx").read_bytes())
    save(openpyxl.Workbook(), tmp_path / "secret.xlsx")  # outside the folder
    os.symlink(tmp_path / "secret.xlsx", root / "link.xlsx")
    return root


@pytest.fixture
def server(folder):
    server = make_server(folder, port=0)
    thread = threading.Thread(target=server.serve_forever, args=(0.01,), daemon=True)
    thread.start()
    yield server
    server.shutdown()
    server.server_close()


def get(server, path, method="GET"):
    connection = http.client.HTTPConnection(*server.server_address[:2], timeout=5)
    connection.request(method, path)  # sent as given: no client-side normalization
    response = connection.getresponse()
    body = response.read().decode()
    connection.close()
    return response.status, body, response


def test_index_lists_the_workbooks(server):
    status, body, response = get(server, "/")
    assert status == 200
    assert response.getheader("Content-Type") == "text/html; charset=utf-8"
    assert '<a href="/view/book.xlsx">book.xlsx</a>' in body
    assert '<a href="/view/sub/other.xlsx">sub/other.xlsx</a>' in body
    assert "notes.txt" not in body
    assert ".hidden" not in body


def test_sheet_pages(server):
    status, body, _ = get(server, "/view/book.xlsx")
    assert status == 200
    assert "hello &lt;world&gt;" in body  # the default sheet, escaped
    assert '<a href="/view/book.xlsx?sheet=Second">Second</a>' in body
    assert '<div class="sheetpeek"' in body
    assert "<script" not in body
    status, body, _ = get(server, "/view/book.xlsx?sheet=Second&ref=A1:B2")
    assert status == 200
    assert ">42</span>" in body
    assert get(server, "/view/book.xlsx", method="HEAD")[1] == ""


@pytest.mark.parametrize(
    "path",
    [
        "/view/../secret.xlsx",
        "/view/%2e%2e/secret.xlsx",
        "/view/sub/../../secret.xlsx",
        "/view//etc/passwd",
        "/view/link.xlsx",  # a symbolic link out of the folder
        "/view/notes.txt",
        "/view/.hidden.xlsx",
        "/view/missing.xlsx",
        "/view/book.xlsx%00.txt",
        "/elsewhere",
    ],
)
def test_nothing_outside_the_folder(server, path):
    status, body, _ = get(server, path)
    assert status == 404
    assert "secret" not in body.replace(path, "")


def test_errors(server):
    assert get(server, "/view/book.xlsx?sheet=Nope")[0] == 404
    assert get(server, "/view/book.xlsx?ref=nonsense")[0] == 404
    assert get(server, "/view/broken.xlsx")[0] == 422
    assert get(server, "/", method="POST")[0] == 501


def test_resolve(folder):
    root = folder.resolve()
    assert resolve(root, "book.xlsx") == root / "book.xlsx"
    assert resolve(root, "sub/other.xlsx") == root / "sub" / "other.xlsx"
    assert resolve(root, "") is None
    assert resolve(root, str(root / "book.xlsx")) is None  # absolute


def test_command_line_defaults(folder, capsys):
    args = cli._parser().parse_args(["serve", str(folder)])
    assert (args.host, args.port) == ("127.0.0.1", 8000)
    assert cli.main(["serve", str(folder / "notes.txt")]) == 1
    assert "not a folder" in capsys.readouterr().err
