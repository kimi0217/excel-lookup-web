"""唯讀讀取 Excel:一律先複製到暫存,讀完立即刪除,絕不碰原始檔。"""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

import pandas as pd


def list_categories(data_dir: Path) -> list[str]:
    """回傳資料夾中所有 .xlsx 的分類名稱(檔名去掉副檔名)。"""
    if not data_dir.is_dir():
        return []
    return sorted(
        p.stem
        for p in data_dir.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")  # 排除 Excel 開啟中的鎖定檔
    )


def read_excel_safely(path: Path) -> dict[str, pd.DataFrame]:
    """複製到暫存資料夾後讀取所有工作表,讀完立即刪除暫存檔。"""
    with tempfile.TemporaryDirectory(prefix="excel-lookup-") as tmp:
        tmp_path = Path(tmp) / path.name
        shutil.copyfile(path, tmp_path)
        try:
            sheets = pd.read_excel(tmp_path, sheet_name=None, dtype=str, engine="openpyxl")
        finally:
            tmp_path.unlink(missing_ok=True)
    return {name: df.fillna("") for name, df in sheets.items()}


_cache: dict[Path, tuple[tuple[float, int], dict[str, pd.DataFrame]]] = {}


def load_category(data_dir: Path, category: str) -> dict[str, pd.DataFrame]:
    """讀取某分類;檔案未變動(修改時間與大小相同)時使用快取。"""
    if category not in list_categories(data_dir):
        raise KeyError(category)
    path = data_dir / f"{category}.xlsx"
    stat = path.stat()
    key = (stat.st_mtime, stat.st_size)
    cached = _cache.get(path)
    if cached and cached[0] == key:
        return cached[1]
    sheets = read_excel_safely(path)
    _cache[path] = (key, sheets)
    return sheets


def search(
    data_dir: Path,
    query: str,
    category: str | None = None,
    max_results: int = 500,
) -> list[dict]:
    """在指定分類(或全部分類)中搜尋包含關鍵字的列,不分大小寫。"""
    query = query.strip()
    if not query:
        return []
    categories = [category] if category else list_categories(data_dir)
    results: list[dict] = []
    remaining = max_results
    for cat in categories:
        for sheet_name, df in load_category(data_dir, cat).items():
            if df.empty or remaining <= 0:
                continue
            mask = df.apply(
                lambda col: col.astype(str).str.contains(query, case=False, regex=False)
            ).any(axis=1)
            matched = df[mask].head(remaining)
            if matched.empty:
                continue
            remaining -= len(matched)
            results.append(
                {
                    "category": cat,
                    "sheet": sheet_name,
                    "columns": [str(c) for c in df.columns],
                    "rows": matched.astype(str).values.tolist(),
                }
            )
    return results
