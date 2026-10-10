"""背景預先載入:啟動時與檔案變動後先讀好,搜尋時不讀檔。"""

import os
import threading
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

import excel_reader
from app import create_app


def make_xlsx(path: Path, value: str) -> None:
    wb = Workbook()
    wb.active.append(["內容"])
    wb.active.append([value])
    wb.save(path)


def touch_later(path: Path) -> None:
    """確保修改時間改變(部分檔案系統時間精度較低)。"""
    stat = path.stat()
    os.utime(path, (stat.st_atime, stat.st_mtime + 5))


def wait_until(condition, timeout: float = 5) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if condition():
            return True
        time.sleep(0.02)
    return False


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    make_xlsx(d / "小檔.xlsx", "關鍵字 小檔")
    make_xlsx(d / "大檔.xlsx", "關鍵字 大檔")
    return d


def test_categories_loading_before_refresh(data_dir):
    catalog = excel_reader.Catalog(data_dir)
    assert {c["status"] for c in catalog.categories()} == {"loading"}
    result = catalog.search("關鍵字")
    assert result == {"results": [], "errors": [], "loading": ["大檔", "小檔"]}


def test_search_never_reads_files(data_dir, monkeypatch):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()

    def fail(*args, **kwargs):
        raise AssertionError("搜尋時不應讀檔")

    monkeypatch.setattr(excel_reader, "read_excel_safely", fail)
    assert len(catalog.search("關鍵字")["results"]) == 2


def test_refresh_only_rereads_changed_files(data_dir, monkeypatch):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    read = []
    real = excel_reader.read_excel_safely
    monkeypatch.setattr(excel_reader, "read_excel_safely", lambda p, h: read.append(p.stem) or real(p, h))

    catalog.refresh()
    assert read == []

    make_xlsx(data_dir / "小檔.xlsx", "已更新")
    touch_later(data_dir / "小檔.xlsx")
    catalog.refresh()
    assert read == ["小檔"]
    assert catalog.search("已更新")["results"][0]["category"] == "小檔"


def test_refresh_handles_added_and_removed_files(data_dir):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    (data_dir / "大檔.xlsx").unlink()
    make_xlsx(data_dir / "新檔.xlsx", "關鍵字 新檔")
    catalog.refresh()
    assert [c["name"] for c in catalog.categories()] == ["小檔", "新檔"]
    assert sorted(g["category"] for g in catalog.search("關鍵字")["results"]) == ["小檔", "新檔"]


def make_big(path: Path) -> None:
    wb = Workbook()
    wb.active.append(["內容"])
    wb.active.append(["關鍵字 大檔"])
    for i in range(2000):
        wb.active.append([f"填充資料 {i}"])
    wb.save(path)


def test_small_files_load_first(data_dir, monkeypatch):
    make_big(data_dir / "大檔.xlsx")
    order = []
    real = excel_reader.read_excel_safely
    monkeypatch.setattr(excel_reader, "read_excel_safely", lambda p, h: order.append(p.stem) or real(p, h))
    excel_reader.Catalog(data_dir).refresh()
    assert order == ["小檔", "大檔"]


def test_stale_data_served_while_reloading(data_dir, monkeypatch):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    started, release = threading.Event(), threading.Event()
    real = excel_reader.read_excel_safely

    def slow(path, include_hidden):
        started.set()
        release.wait(5)
        return real(path, include_hidden)

    monkeypatch.setattr(excel_reader, "read_excel_safely", slow)
    make_xlsx(data_dir / "小檔.xlsx", "新內容")
    touch_later(data_dir / "小檔.xlsx")
    worker = threading.Thread(target=catalog.refresh)
    worker.start()
    try:
        assert started.wait(5)
        # 重新載入期間仍可用舊資料搜尋,不顯示為載入中
        result = catalog.search("關鍵字 小檔")
        assert result["loading"] == []
        assert result["results"][0]["category"] == "小檔"
    finally:
        release.set()
        worker.join()
    assert catalog.search("新內容")["results"][0]["category"] == "小檔"


def test_app_preloads_in_background_and_reports_loading(data_dir, monkeypatch):
    make_big(data_dir / "大檔.xlsx")
    release = threading.Event()
    real = excel_reader.read_excel_safely

    def slow(path, include_hidden):
        if path.stem == "大檔":
            release.wait(5)
        return real(path, include_hidden)

    monkeypatch.setattr(excel_reader, "read_excel_safely", slow)
    app = create_app({"data_dir": data_dir, "reload_interval_seconds": 0.05})
    with TestClient(app) as client:
        try:
            def statuses():
                return {c["name"]: c["status"] for c in client.get("/api/categories").json()}

            # 小檔已載入,大檔仍在載入中;搜尋立即回應,不必等待大檔
            assert wait_until(lambda: statuses() == {"小檔": "ok", "大檔": "loading"})
            body = client.get("/api/search", params={"q": "關鍵字"}).json()
            assert [g["category"] for g in body["results"]] == ["小檔"]
            assert body["loading"] == ["大檔"]
        finally:
            release.set()
        assert wait_until(lambda: statuses() == {"小檔": "ok", "大檔": "ok"})


def test_app_detects_file_changes(data_dir):
    app = create_app({"data_dir": data_dir, "reload_interval_seconds": 0.05})
    with TestClient(app) as client:
        def found(q):
            return client.get("/api/search", params={"q": q}).json()["results"]

        assert wait_until(lambda: found("關鍵字 大檔"))
        make_xlsx(data_dir / "大檔.xlsx", "背景偵測到的新內容")
        touch_later(data_dir / "大檔.xlsx")
        assert wait_until(lambda: found("背景偵測到的新內容"))
