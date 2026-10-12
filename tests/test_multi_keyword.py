"""多關鍵字搜尋:以空白(含全形空白)拆開,每個關鍵字都出現在同一列才算符合。"""

from pathlib import Path

import pytest

from conftest import loaded_client
from fixtures.generate_fixtures import merged_cells


@pytest.fixture
def client(tmp_path: Path):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    merged_cells(data_dir / "02_合併儲存格.xlsx")  # 資訊部(合併儲存格):王小明、陳大文、林美玲
    return loaded_client({"data_dir": data_dir})


def file_rows(client, q):
    return client.get("/api/files/02_合併儲存格/rows", params={"q": q}).json()


def search_rows(client, q):
    body = client.get("/api/search", params={"q": q}).json()
    return [row for group in body["results"] for row in group["rows"]]


def test_file_search_two_keywords(client):
    body = file_rows(client, "資訊部 陳大文")
    assert body["matched"] == 1 and body["rows"] == [["資訊部", "陳大文", "102"]]


def test_cross_file_search_two_keywords(client):
    assert search_rows(client, "資訊部 陳大文") == [["資訊部", "陳大文", "102"]]


@pytest.mark.parametrize("q", ["資訊部　陳大文", "  陳大文   資訊部  ", "陳大文 資訊部 102"])
def test_fullwidth_space_extra_spaces_and_order(client, q):
    assert file_rows(client, q)["matched"] == 1
    assert len(search_rows(client, q)) == 1


def test_all_keywords_required(client):
    assert file_rows(client, "資訊部 李小華")["matched"] == 0  # 李小華在業務部
    assert search_rows(client, "資訊部 李小華") == []


def test_single_keyword_unchanged(client):
    assert file_rows(client, "資訊部")["matched"] == 3
