"""``sheetpeek serve``: an index, one page per sheet, and nothing outside the folder."""

import http.client
import os
import threading

import openpyxl
import pytest

from sheetpeek import ReadLimits, _serve, cli
from sheetpeek._package import SERVE_LIMITS
from sheetpeek._serve import make_server, resolve
from tests.workbooks import SHEET1, replace_part, rewrite, save


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


def get(server, path, method="GET", host=None):
    connection = http.client.HTTPConnection(*server.server_address[:2], timeout=5)
    headers = {} if host is None else {"Host": host}
    connection.request(method, path, headers=headers)  # sent as given: no client normalization
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


# --- what a request may cost, and who may ask


@pytest.mark.parametrize(
    "host",
    [
        "localhost",
        "localhost:8000",
        "LOCALHOST:9",
        "127.0.0.1",
        "127.0.0.1:12345",
        "[::1]",
        "[::1]:80",
    ],
)
def test_loopback_names_are_answered(server, host):
    assert get(server, "/", host=host)[0] == 200


@pytest.mark.parametrize(
    "host",
    [
        "evil.example",
        "evil.example:8000",
        "localhost.evil.example",
        "127.0.0.1.evil.example",
        "localhost:80@evil.example",
        "sub.localhost",
        "[::2]",
        "",
    ],
)
def test_other_names_are_refused(server, host):
    for path in ("/", "/view/book.xlsx"):
        status, body, _ = get(server, path, host=host)
        assert status == 403
        assert "hello" not in body
        assert "book.xlsx" not in body


def test_a_request_without_one_host_is_refused(server):
    connection = http.client.HTTPConnection(*server.server_address[:2], timeout=5)
    connection.putrequest("GET", "/", skip_host=True)
    connection.endheaders()
    assert connection.getresponse().status == 403
    connection.close()
    connection = http.client.HTTPConnection(*server.server_address[:2], timeout=5)
    connection.putrequest("GET", "/", skip_host=True)
    connection.putheader("Host", "localhost")
    connection.putheader("Host", "evil.example")
    connection.endheaders()
    assert connection.getresponse().status == 403
    connection.close()


def test_the_loopback_address_it_is_bound_to_is_answered(folder):
    try:
        server = make_server(folder, "127.0.0.2", 0)
    except OSError:
        pytest.skip("this system has no 127.0.0.2")
    thread = threading.Thread(target=server.serve_forever, args=(0.01,), daemon=True)
    thread.start()
    try:
        assert get(server, "/", host=f"127.0.0.2:{server.server_address[1]}")[0] == 200
        assert get(server, "/", host="evil.example")[0] == 403
    finally:
        server.shutdown()
        server.server_close()


def test_any_name_is_answered_when_bound_elsewhere(folder):
    server = make_server(folder, "0.0.0.0", 0)
    thread = threading.Thread(target=server.serve_forever, args=(0.01,), daemon=True)
    thread.start()
    try:
        connection = http.client.HTTPConnection("127.0.0.1", server.server_address[1], timeout=5)
        connection.request("GET", "/", headers={"Host": "reports.example"})
        assert connection.getresponse().status == 200
        connection.close()
    finally:
        server.shutdown()
        server.server_close()


def test_the_index_does_not_show_where_the_folder_is(server, folder):
    status, body, _ = get(server, "/")
    assert status == 200
    assert folder.name in body
    assert str(folder.parent) not in body


def test_an_error_from_openpyxl_does_not_show_where_the_folder_is(server, folder):
    # openpyxl's ValueError names the file by its full path.
    rewrite(
        folder / "book.xlsx",
        "xl/workbook.xml",
        lambda text: text.replace('state="visible"', 'state="bogus"', 1),
    )
    status, body, _ = get(server, "/view/book.xlsx")
    assert status == 422
    assert "Cannot read this file" in body
    assert str(folder) not in body
    assert str(folder.parent) not in body


def test_a_refused_file_is_422_not_404(server, folder):
    replace_part(folder / "book.xlsx", SHEET1, b"<!DOCTYPE x><worksheet/>")
    status, body, _ = get(server, "/view/book.xlsx")
    assert status == 422
    assert "declares a DTD" in body


def test_a_huge_ref_is_refused_without_reading_it(server):
    status, body, _ = get(server, "/view/book.xlsx?ref=A1:XFD1048576")
    assert status == 422
    assert "more than the limit of 1048576" in body
    assert SERVE_LIMITS.max_cells == 1 << 20  # one full column is allowed


def test_the_limits_can_be_changed(folder):
    server = make_server(folder, port=0, limits=ReadLimits(max_parts=1))
    thread = threading.Thread(target=server.serve_forever, args=(0.01,), daemon=True)
    thread.start()
    try:
        assert get(server, "/view/book.xlsx")[0] == 422
    finally:
        server.shutdown()
        server.server_close()


def test_a_read_that_gets_no_slot_is_turned_away(server, monkeypatch):
    monkeypatch.setattr(_serve, "_SLOT_WAIT", 0.01)
    slots = server.RequestHandlerClass.slots
    assert slots.acquire(blocking=False)
    assert slots.acquire(blocking=False)
    try:
        status, body, response = get(server, "/view/book.xlsx")
        assert status == 503
        assert response.getheader("Retry-After") == "5"
        assert "hello" not in body
        assert get(server, "/")[0] == 200  # the index reads no workbook
    finally:
        slots.release()
        slots.release()
    assert get(server, "/view/book.xlsx")[0] == 200


def test_a_failed_read_gives_its_slot_back(server):
    for _ in range(5):
        assert get(server, "/view/broken.xlsx")[0] == 422
    assert get(server, "/view/book.xlsx")[0] == 200


def test_a_read_past_the_layout_limit_is_refused(folder, tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "1" * 32_767
    save(wb, folder / "note.xlsx")
    server = make_server(folder, port=0)
    thread = threading.Thread(target=server.serve_forever, args=(0.01,), daemon=True)
    thread.start()
    try:
        status, body, _ = get(server, "/view/note.xlsx?ref=A1:A1000")
        assert status == 422
        assert "columns spilled text reaches" in body
        assert get(server, "/view/note.xlsx")[0] == 200
    finally:
        server.shutdown()
        server.server_close()
