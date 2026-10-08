import hashlib
import time
from datetime import date
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

import backup
from app import create_app


def make_xlsx(path: Path, value: str) -> None:
    wb = Workbook()
    wb.active.append(["欄位"])
    wb.active.append([value])
    wb.save(path)


def file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.fixture
def dirs(tmp_path: Path) -> tuple[Path, Path]:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    make_xlsx(data_dir / "員工.xlsx", "王小明")
    make_xlsx(data_dir / "設備.xlsx", "筆電")
    (data_dir / "~$員工.xlsx").write_bytes(b"lock")
    return data_dir, tmp_path / "backups"


def test_backup_copies_excels_into_dated_folder(dirs):
    data_dir, backup_dir = dirs
    target = backup.backup_excels(data_dir, backup_dir, today=date(2026, 10, 8))
    assert target == backup_dir / "2026-10-08"
    assert sorted(p.name for p in target.iterdir()) == ["員工.xlsx", "設備.xlsx"]
    assert file_hash(target / "員工.xlsx") == file_hash(data_dir / "員工.xlsx")
    assert not any(p.name.endswith(".partial") for p in backup_dir.iterdir())


def test_backup_does_not_modify_originals(dirs):
    data_dir, backup_dir = dirs
    before = {p.name: (file_hash(p), p.stat().st_mtime) for p in data_dir.iterdir()}
    backup.backup_excels(data_dir, backup_dir, today=date(2026, 10, 8))
    after = {p.name: (file_hash(p), p.stat().st_mtime) for p in data_dir.iterdir()}
    assert before == after


def test_backup_runs_once_per_day(dirs):
    data_dir, backup_dir = dirs
    today = date(2026, 10, 8)
    backup.backup_excels(data_dir, backup_dir, today=today)
    make_xlsx(data_dir / "員工.xlsx", "已修改")
    backup.backup_excels(data_dir, backup_dir, today=today)
    assert file_hash(backup_dir / "2026-10-08" / "員工.xlsx") != file_hash(data_dir / "員工.xlsx")


def test_backup_skips_when_no_excel(tmp_path):
    (tmp_path / "data").mkdir()
    assert backup.backup_excels(tmp_path / "data", tmp_path / "backups") is None
    assert not (tmp_path / "backups").exists()


def test_prune_keeps_30_days(tmp_path):
    backup_dir = tmp_path / "backups"
    for name in ("2026-09-07", "2026-09-08", "2026-10-08", "手動備份"):
        (backup_dir / name).mkdir(parents=True)
    (backup_dir / "2026-01-01.txt").write_text("不是資料夾", encoding="utf-8")

    removed = backup.prune_backups(backup_dir, keep_days=30, today=date(2026, 10, 8))

    assert removed == [backup_dir / "2026-09-07"]
    assert sorted(p.name for p in backup_dir.iterdir()) == [
        "2026-01-01.txt", "2026-09-08", "2026-10-08", "手動備份",
    ]


def test_run_backup_backs_up_and_prunes(dirs):
    data_dir, backup_dir = dirs
    (backup_dir / "2026-08-01").mkdir(parents=True)
    backup.run_backup(data_dir, backup_dir, keep_days=30, today=date(2026, 10, 8))
    assert sorted(p.name for p in backup_dir.iterdir()) == ["2026-10-08"]


def test_app_startup_triggers_backup(dirs):
    data_dir, backup_dir = dirs
    app = create_app({"data_dir": data_dir, "backup_dir": backup_dir})
    expected = backup_dir / date.today().strftime("%Y-%m-%d") / "員工.xlsx"
    with TestClient(app):
        deadline = time.monotonic() + 5
        while not expected.is_file() and time.monotonic() < deadline:
            time.sleep(0.05)
    assert expected.is_file()
