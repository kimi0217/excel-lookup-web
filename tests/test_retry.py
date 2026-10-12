"""讀取失敗的檔案每次檢查都重讀,不必等到修改時間或大小改變。"""

import logging
from pathlib import Path

import pytest
from openpyxl import Workbook

import excel_reader


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    wb = Workbook()
    wb.active.append(["姓名"])
    wb.active.append(["王小明"])
    wb.save(d / "員工.xlsx")
    return d


def status(catalog, name="員工"):
    return {c["name"]: c["status"] for c in catalog.categories()}[name]


@pytest.fixture
def copy_fails_once(monkeypatch):
    real = excel_reader.shutil.copyfile
    calls = []

    def flaky(src, dst, *args, **kwargs):
        calls.append(src)
        if len(calls) == 1:
            raise PermissionError("模擬檔案被其他程式鎖定")
        return real(src, dst, *args, **kwargs)

    monkeypatch.setattr(excel_reader.shutil, "copyfile", flaky)
    return calls


def test_error_file_retried_on_next_refresh(data_dir, copy_fails_once):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    assert status(catalog) == "error"
    catalog.refresh()  # 檔案沒有變動,仍要重讀
    assert status(catalog) == "ok"
    assert catalog.search("王小明")["results"]
    assert len(copy_fails_once) == 2


def test_ok_file_not_reread_when_unchanged(data_dir, monkeypatch):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    reads = []
    real = excel_reader.read_excel_safely
    monkeypatch.setattr(excel_reader, "read_excel_safely", lambda *a, **k: reads.append(a) or real(*a, **k))
    catalog.refresh()
    assert reads == []


def test_repeated_failure_logged_once_and_recovery_logged(data_dir, monkeypatch, caplog):
    (data_dir / "損毀.xlsx").write_bytes(b"not a zip")
    catalog = excel_reader.Catalog(data_dir)
    with caplog.at_level(logging.INFO, logger="excel_reader"):
        for _ in range(3):
            catalog.refresh()
    warnings = [r for r in caplog.records if r.levelno == logging.WARNING and "損毀.xlsx" in r.getMessage()]
    assert len(warnings) == 1  # 每 10 秒重試一次,但同樣的錯誤只記錄一次
    caplog.clear()
    wb = Workbook()
    wb.active.append(["項目"])
    wb.active.append(["已修復"])
    wb.save(data_dir / "損毀.xlsx")
    with caplog.at_level(logging.INFO, logger="excel_reader"):
        catalog.refresh()
    assert status(catalog, "損毀") == "ok"
    assert any("已可讀取" in r.getMessage() and "損毀.xlsx" in r.getMessage() for r in caplog.records)
