"""支援 .xlsm,副檔名不分大小寫;備份也包含這些檔案。"""

from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

import backup
import excel_reader
from conftest import search
from deploy import common


def make(path: Path, value: str) -> None:
    wb = Workbook()
    wb.active.append(["內容"])
    wb.active.append([value])
    wb.save(path)


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    make(d / "一般.xlsx", "一般內容")
    make(d / "巨集.xlsm", "巨集內容")
    make(d / "大寫.XLSX", "大寫內容")
    make(d / "混合.Xlsm", "混合內容")
    (d / "~$巨集.xlsm").write_bytes(b"lock")
    (d / "~$大寫.XLSX").write_bytes(b"lock")
    (d / "舊版.xls").write_bytes(b"old")
    (d / "資料.csv").write_text("a,b", encoding="utf-8")
    return d


def test_scan_includes_xlsm_and_any_case(data_dir):
    assert list(excel_reader.scan_folder(data_dir)) == ["一般", "大寫", "巨集", "混合"]


def test_all_types_readable_and_searchable(data_dir):
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    assert {c["name"]: c["status"] for c in catalog.categories()} == {
        "一般": "ok", "大寫": "ok", "巨集": "ok", "混合": "ok",
    }
    for value in ("巨集內容", "大寫內容", "混合內容"):
        assert search(data_dir, value)["results"], value


def test_same_name_different_extension_kept_apart(data_dir):
    make(data_dir / "一般.xlsm", "另一個檔案")
    names = list(excel_reader.scan_folder(data_dir))
    assert names == ["一般.xlsm", "一般.xlsx", "大寫", "巨集", "混合"]
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    assert catalog.search("另一個檔案")["results"][0]["category"] == "一般.xlsm"
    assert catalog.search("一般內容")["results"][0]["category"] == "一般.xlsx"


def test_backup_includes_all_types(data_dir, tmp_path):
    target = backup.backup_excels(data_dir, tmp_path / "backups", today=date(2026, 10, 12))
    assert sorted(p.name for p in target.iterdir()) == ["一般.xlsx", "大寫.XLSX", "巨集.xlsm", "混合.Xlsm"]


def test_backup_with_only_xlsm(tmp_path):
    d = tmp_path / "data"
    d.mkdir()
    make(d / "巨集.xlsm", "x")
    assert backup.backup_excels(d, tmp_path / "backups", today=date(2026, 10, 12)) is not None


def test_deploy_counts_all_types(data_dir):
    assert common.count_xlsx(data_dir) == 4
