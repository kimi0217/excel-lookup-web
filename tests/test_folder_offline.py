"""Excel 資料夾暫時連不到(網路斷線、共用資料夾離線)時保留已載入的資料。"""

import logging
import re
from pathlib import Path

import pytest
from openpyxl import Workbook

import excel_reader
from conftest import loaded_client

TIME = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}")


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    for name, value in (("員工", "王小明"), ("設備", "筆電")):
        wb = Workbook()
        wb.active.append(["內容"])
        wb.active.append([value])
        wb.save(d / f"{name}.xlsx")
    return d


@pytest.fixture
def client(data_dir):
    return loaded_client({"data_dir": data_dir})


def categories(client) -> dict:
    return client.get("/api/categories").json()


def found(client, q) -> int:
    return sum(len(g["rows"]) for g in client.get("/api/search", params={"q": q}).json()["results"])


def go_offline(data_dir: Path) -> Path:
    offline = data_dir.with_name("data_離線")
    data_dir.rename(offline)
    return offline


def test_normal_status(client):
    body = categories(client)
    assert body["folder"]["ok"] is True and TIME.fullmatch(body["folder"]["last_success"])
    assert [f["name"] for f in body["files"]] == ["員工", "設備"]


def test_offline_keeps_data_then_recovers(client, data_dir):
    before = categories(client)["folder"]["last_success"]
    offline = go_offline(data_dir)
    client.app.state.catalog.refresh()

    body = categories(client)
    assert body["folder"] == {"ok": False, "last_success": before}
    assert [f["name"] for f in body["files"]] == ["員工", "設備"]
    assert all(f["status"] == "ok" for f in body["files"])
    assert found(client, "王小明") == 1
    rows = client.get("/api/files/員工/rows").json()
    assert rows["rows"] == [["王小明"]]

    offline.rename(data_dir)
    client.app.state.catalog.refresh()
    body = categories(client)
    assert body["folder"]["ok"] is True
    assert found(client, "王小明") == 1


def test_files_changed_while_offline_picked_up_after_recovery(client, data_dir):
    offline = go_offline(data_dir)
    client.app.state.catalog.refresh()
    (offline / "設備.xlsx").unlink()
    offline.rename(data_dir)
    client.app.state.catalog.refresh()
    assert [f["name"] for f in categories(client)["files"]] == ["員工"]


def test_offline_at_startup(tmp_path):
    client = loaded_client({"data_dir": tmp_path / "連不到的資料夾"})
    body = categories(client)
    assert body == {"folder": {"ok": False, "last_success": None}, "files": []}


def test_requests_do_not_touch_folder(client, data_dir, monkeypatch):
    """網路資料夾斷線時,存取檔案系統可能卡住數秒;請求只用已載入的資料。"""
    folder = str(data_dir)

    def guard(real):
        def wrapper(self, *args, **kwargs):
            if str(self).startswith(folder):
                raise AssertionError(f"請求時不應存取 Excel 資料夾:{self}")
            return real(self, *args, **kwargs)
        return wrapper

    for method in ("iterdir", "stat", "is_dir", "glob", "exists", "is_file"):
        monkeypatch.setattr(Path, method, guard(getattr(Path, method)))
    assert len(categories(client)["files"]) == 2
    assert found(client, "王小明") == 1
    assert client.get("/api/files/員工").status_code == 200
    assert client.get("/api/files/員工/rows").json()["rows"] == [["王小明"]]
    assert client.get("/api/search", params={"q": "x", "category": "員工"}).status_code == 200


def test_offline_logged_once(client, data_dir, caplog):
    offline = go_offline(data_dir)
    with caplog.at_level(logging.INFO, logger="excel_reader"):
        for _ in range(3):
            client.app.state.catalog.refresh()
        offline.rename(data_dir)
        client.app.state.catalog.refresh()
    messages = [r.getMessage() for r in caplog.records]
    assert sum("無法連線" in m for m in messages) == 1
    assert sum("已恢復連線" in m for m in messages) == 1


def test_frontend_has_offline_banner(client):
    html = client.get("/").text
    assert "Excel 資料夾暫時無法連線" in html and "的資料" in html
