"""每日備份原始 Excel 到 backups/YYYY-MM-DD/,並刪除超過保留天數的舊備份。

只讀取原始檔(複製),絕不修改或刪除資料夾中的 Excel。
"""

from __future__ import annotations

import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

from excel_reader import is_excel_file

DATE_FORMAT = "%Y-%m-%d"


def backup_excels(data_dir: Path, backup_dir: Path, today: date | None = None) -> Path | None:
    """把所有 Excel(.xlsx、.xlsm,不分大小寫)複製到當日備份資料夾;當日已備份過則略過。回傳備份資料夾。"""
    today = today or date.today()
    files = [
        p for p in data_dir.iterdir() if p.is_file() and is_excel_file(p)
    ] if data_dir.is_dir() else []
    if not files:
        return None
    target = backup_dir / today.strftime(DATE_FORMAT)
    if target.is_dir():
        return target
    tmp_target = backup_dir / f".{target.name}.partial"
    shutil.rmtree(tmp_target, ignore_errors=True)
    tmp_target.mkdir(parents=True)
    for src in files:
        shutil.copy2(src, tmp_target / src.name)
    tmp_target.rename(target)  # 全部複製完才改名,避免中斷時留下不完整的當日備份
    return target


def prune_backups(backup_dir: Path, keep_days: int = 30, today: date | None = None) -> list[Path]:
    """刪除超過 keep_days 天的備份資料夾;只處理名稱為日期格式的資料夾。"""
    today = today or date.today()
    cutoff = today - timedelta(days=keep_days)
    removed: list[Path] = []
    if not backup_dir.is_dir():
        return removed
    for folder in backup_dir.iterdir():
        if not folder.is_dir():
            continue
        try:
            folder_date = datetime.strptime(folder.name, DATE_FORMAT).date()
        except ValueError:
            continue
        if folder_date < cutoff:
            shutil.rmtree(folder)
            removed.append(folder)
    return sorted(removed)


def run_backup(data_dir: Path, backup_dir: Path, keep_days: int = 30, today: date | None = None) -> None:
    backup_excels(data_dir, backup_dir, today)
    prune_backups(backup_dir, keep_days, today)
