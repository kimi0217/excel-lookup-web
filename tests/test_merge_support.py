"""防止合併儲存格悄悄失效:openpyxl 升級後若取不到合併範圍,這裡的測試會失敗。"""

import logging
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

import excel_reader
from app import create_app
from fixtures.generate_fixtures import merged_cells


@pytest.fixture
def merged_file(tmp_path: Path) -> Path:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    path = data_dir / "02_合併儲存格.xlsx"
    merged_cells(path)  # 與 tests/fixtures/generated/ 相同的產生方式
    return path


def test_merge_ranges_found_in_fixture(merged_file):
    """取不到合併範圍(例如 openpyxl 內部屬性改名)時這個測試會失敗。"""
    wb = load_workbook(merged_file, read_only=True)
    try:
        with zipfile.ZipFile(merged_file) as archive:
            merges = excel_reader._merge_ranges(archive, wb.worksheets[0])
    finally:
        wb.close()
    assert sorted(merges, key=lambda m: (m.min_row, m.min_col)) == [
        excel_reader.Merge(0, 0, 1, 0),  # A1:A2 表頭「部門」
        excel_reader.Merge(0, 1, 0, 2),  # B1:C1 表頭「人員」
        excel_reader.Merge(2, 0, 4, 0),  # A3:A5 資料「資訊部」
    ]


def test_merged_values_filled_in_fixture(merged_file):
    rows = excel_reader.read_excel_safely(merged_file)["部門人員"].rows
    assert [r[:2] for r in rows if r[1] in ("王小明", "陳大文", "林美玲")] == [
        ["資訊部", "王小明"], ["資訊部", "陳大文"], ["資訊部", "林美玲"],
    ]


def test_merge_support_ok():
    assert excel_reader.merge_support_ok()


@pytest.fixture
def attr_missing(monkeypatch):
    """模擬 openpyxl 升級後內部屬性改名。"""
    monkeypatch.setattr(excel_reader, "SHEET_PATH_ATTR", "_改名後的屬性")
    monkeypatch.setattr(excel_reader, "_merge_warning_logged", False)


def test_missing_attr_does_not_crash(attr_missing, merged_file, caplog):
    with caplog.at_level(logging.WARNING):
        rows = excel_reader.read_excel_safely(merged_file)["部門人員"].rows
    assert excel_reader.merge_support_ok() is False
    assert ["", "陳大文", "102"] in rows  # 合併格沒填值,但仍讀得到資料
    assert "取不到合併儲存格範圍" in caplog.text


def test_startup_logs_warning_when_unsupported(attr_missing, tmp_path, caplog):
    with caplog.at_level(logging.WARNING, logger="app"):
        with TestClient(create_app({"data_dir": tmp_path})):
            pass
    assert any("取不到合併儲存格範圍" in r.getMessage() and r.name == "app" for r in caplog.records)


def test_startup_no_warning_normally(tmp_path, caplog):
    with caplog.at_level(logging.WARNING, logger="app"):
        with TestClient(create_app({"data_dir": tmp_path})):
            pass
    assert not any("合併儲存格" in r.getMessage() for r in caplog.records)
