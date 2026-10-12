import hashlib
import re
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

import excel_reader
from conftest import search, loaded_client


def make_xlsx(path: Path, rows: list[list], sheet: str = "工作表1") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for row in rows:
        ws.append(row)
    wb.save(path)


def file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    make_xlsx(d / "員工.xlsx", [["姓名", "部門"], ["王小明", "資訊部"], ["李小華", "業務部"]])
    make_xlsx(d / "設備.xlsx", [["編號", "名稱"], ["A001", "筆電"], ["A002", "螢幕"]])
    (d / "~$員工.xlsx").write_bytes(b"lock")
    return d


@pytest.fixture
def client(data_dir: Path) -> TestClient:
    return loaded_client({"data_dir": data_dir})


def test_list_categories_skips_lock_files(data_dir):
    assert excel_reader.list_categories(data_dir) == ["員工", "設備"]


def test_read_does_not_modify_original(data_dir):
    path = data_dir / "員工.xlsx"
    before_hash, before_mtime = file_hash(path), path.stat().st_mtime
    excel_reader.read_excel_safely(path)
    assert file_hash(path) == before_hash
    assert path.stat().st_mtime == before_mtime
    assert sorted(p.name for p in data_dir.iterdir()) == ["~$員工.xlsx", "員工.xlsx", "設備.xlsx"]


def test_temp_copy_is_deleted(data_dir, tmp_path, monkeypatch):
    tmp_root = tmp_path / "tmp"
    tmp_root.mkdir()
    monkeypatch.setattr(excel_reader.tempfile, "tempdir", str(tmp_root))
    excel_reader.read_excel_safely(data_dir / "員工.xlsx")
    assert list(tmp_root.iterdir()) == []


def test_search_all_categories(data_dir):
    results = search(data_dir, "a00")["results"]
    assert len(results) == 1
    assert results[0]["category"] == "設備"
    assert results[0]["columns"] == ["編號", "名稱"]
    assert results[0]["rows"] == [["A001", "筆電"], ["A002", "螢幕"]]


def test_search_respects_max_results(data_dir):
    results = search(data_dir, "A00", max_results=1)["results"]
    assert sum(len(g["rows"]) for g in results) == 1


def test_api_categories(client):
    cats = client.get("/api/categories").json()["files"]
    assert [(c["name"], c["status"], c["error"], c["rows"], c["sheets"]) for c in cats] == [
        ("員工", "ok", None, 2, 1),
        ("設備", "ok", None, 2, 1),
    ]
    assert all(re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", c["updated"]) for c in cats)


def test_api_search_with_category(client):
    res = client.get("/api/search", params={"q": "資訊", "category": "員工"})
    assert res.status_code == 200
    assert res.json()["results"][0]["rows"] == [["王小明", "資訊部"]]


def test_api_rejects_unknown_category(client):
    res = client.get("/api/search", params={"q": "x", "category": "../secret"})
    assert res.status_code == 404


def test_api_empty_query(client):
    assert client.get("/api/search", params={"q": "  "}).json() == {"results": [], "errors": [], "loading": []}


def test_index_page(client):
    res = client.get("/")
    assert res.status_code == 200
    assert "Excel 查詢" in res.text


def test_config_with_bom_and_windows_paths(tmp_path, monkeypatch):
    import app

    (tmp_path / "config.json").write_text(
        '{"data_dir": "D:\\\\共用\\\\Excel", "port": 9000}', encoding="utf-8-sig"
    )
    monkeypatch.setattr(app, "BASE_DIR", tmp_path)
    config = app.load_config()
    assert config["port"] == 9000
    assert str(config["data_dir"]).endswith("Excel")


@pytest.mark.parametrize("path", ["/docs", "/redoc", "/openapi.json"])
def test_api_docs_disabled(client, path):
    assert client.get(path).status_code == 404
