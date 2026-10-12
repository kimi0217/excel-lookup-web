"""子資料夾:遞迴讀取,分類名稱包含相對路徑;略過 . 與 ~ 開頭的資料夾;排除位於資料夾內的備份資料夾。"""

import json
import logging
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

import backup
import excel_reader
from conftest import loaded_client
from deploy import common

ROOT = Path(__file__).resolve().parents[1]


def make(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.active.append(["內容"])
    wb.active.append([value])
    wb.save(path)


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    make(d / "總表.xlsx", "總表內容")
    make(d / "醫材" / "耗材清單.xlsx", "紗布")
    make(d / "醫材" / "儀器" / "保養紀錄.xlsm", "呼吸器")
    make(d / ".git" / "隱藏資料夾.xlsx", "不該出現")
    make(d / "~暫存" / "暫存資料夾.xlsx", "不該出現")
    (d / "醫材" / "~$耗材清單.xlsx").write_bytes(b"lock")
    return d


def test_scan_reads_subfolders_with_relative_names(data_dir):
    assert list(excel_reader.scan_folder(data_dir)) == ["總表", "醫材/儀器/保養紀錄", "醫材/耗材清單"]


def test_scan_without_subfolders(data_dir):
    assert list(excel_reader.scan_folder(data_dir, include_subfolders=False)) == ["總表"]


def test_same_name_different_extension_in_subfolder(data_dir):
    make(data_dir / "醫材" / "耗材清單.xlsm", "另一個")
    names = list(excel_reader.scan_folder(data_dir))
    assert "醫材/耗材清單.xlsm" in names and "醫材/耗材清單.xlsx" in names


def test_same_file_name_in_different_folders(data_dir):
    make(data_dir / "藥品" / "耗材清單.xlsx", "藥用")
    catalog = excel_reader.Catalog(data_dir)
    catalog.refresh()
    assert catalog.search("紗布")["results"][0]["category"] == "醫材/耗材清單"
    assert catalog.search("藥用")["results"][0]["category"] == "藥品/耗材清單"


def test_backup_dir_inside_data_dir_excluded(data_dir):
    backup_dir = data_dir / "備份"
    backup.backup_excels(data_dir, backup_dir, today=date(2026, 10, 12))
    assert (backup_dir / "2026-10-12" / "總表.xlsx").is_file()
    names = list(excel_reader.scan_folder(data_dir, exclude=[backup_dir]))
    assert names == ["總表", "醫材/儀器/保養紀錄", "醫材/耗材清單"]
    # 再備份一次(隔天)也不會把前一天的備份當成資料備份進去
    target = backup.backup_excels(data_dir, backup_dir, today=date(2026, 10, 13))
    assert not (target / "備份").exists()


def test_backup_dir_same_as_data_dir_skips_dated_folders(data_dir):
    backup.backup_excels(data_dir, data_dir, today=date(2026, 10, 12))
    assert (data_dir / "2026-10-12" / "總表.xlsx").is_file()
    assert list(excel_reader.scan_folder(data_dir, exclude=[data_dir])) == ["總表", "醫材/儀器/保養紀錄", "醫材/耗材清單"]


def test_app_excludes_backup_dir_inside_data_dir(data_dir):
    backup_dir = data_dir / "backups"
    backup.backup_excels(data_dir, backup_dir, today=date(2026, 10, 12))
    client = loaded_client({"data_dir": data_dir, "backup_dir": backup_dir})
    names = [c["name"] for c in client.get("/api/categories").json()["files"]]
    assert names == ["總表", "醫材/儀器/保養紀錄", "醫材/耗材清單"]
    assert len(client.get("/api/search", params={"q": "紗布"}).json()["results"]) == 1


def test_backup_keeps_folder_structure(data_dir, tmp_path):
    target = backup.backup_excels(data_dir, tmp_path / "backups", today=date(2026, 10, 12))
    files = sorted(p.relative_to(target).as_posix() for p in target.rglob("*") if p.is_file())
    assert files == ["總表.xlsx", "醫材/儀器/保養紀錄.xlsm", "醫材/耗材清單.xlsx"]


def test_backup_without_subfolders(data_dir, tmp_path):
    target = backup.backup_excels(data_dir, tmp_path / "backups", today=date(2026, 10, 12), include_subfolders=False)
    assert sorted(p.name for p in target.rglob("*")) == ["總表.xlsx"]


def test_api_with_folder_names(data_dir):
    client = loaded_client({"data_dir": data_dir})
    info = client.get("/api/files/醫材/耗材清單")
    assert info.status_code == 200 and info.json()["name"] == "醫材/耗材清單"
    rows = client.get("/api/files/醫材/耗材清單/rows").json()
    assert rows["rows"] == [["紗布"]]
    assert client.get("/api/files/醫材%2F耗材清單/rows").json()["rows"] == [["紗布"]]
    assert client.get("/api/search", params={"q": "呼吸器", "category": "醫材/儀器/保養紀錄"}).json()["results"]
    assert client.get("/api/files/醫材").status_code == 404


def test_file_named_rows_in_subfolder(tmp_path):
    """「資料夾/rows」與「資料夾 的 rows 頁面」網址相同時,資料夾不存在則視為檔案。"""
    d = tmp_path / "data"
    make(d / "報表" / "rows.xlsx", "特別檔名")
    client = loaded_client({"data_dir": d})
    assert client.get("/api/files/報表/rows").json()["name"] == "報表/rows"
    assert client.get("/api/files/報表/rows/rows").json()["rows"] == [["特別檔名"]]


def test_config_option(data_dir):
    client = loaded_client({"data_dir": data_dir, "include_subfolders": False})
    assert [c["name"] for c in client.get("/api/categories").json()["files"]] == ["總表"]


def test_example_config_default_true():
    config = json.loads((ROOT / "config.example.json").read_text(encoding="utf-8"))
    assert config["include_subfolders"] is True


def test_unreadable_subfolder_skipped(data_dir, monkeypatch, caplog):
    """沒有權限的子資料夾略過,不影響其他檔案。"""
    blocked = data_dir / "醫材" / "儀器"
    real_iterdir = Path.iterdir

    def iterdir(self):
        if self == blocked:
            raise PermissionError("拒絕存取")
        return real_iterdir(self)

    monkeypatch.setattr(Path, "iterdir", iterdir)
    with caplog.at_level(logging.WARNING):
        assert list(excel_reader.scan_folder(data_dir)) == ["總表", "醫材/耗材清單"]


def test_deploy_counts_subfolders(data_dir):
    assert common.count_xlsx(data_dir) == 3
