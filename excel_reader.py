"""唯讀讀取 Excel:一律先複製到暫存,讀完立即刪除,絕不碰原始檔。"""

from __future__ import annotations

import logging
import shutil
import tempfile
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

# 加密的 .xlsx 與舊版 .xls 都是 OLE 複合文件(不是 zip)
OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
# 加密檔的 OLE 目錄中會有這個串流名稱(UTF-16LE)
ENCRYPTED_MARK = "EncryptedPackage".encode("utf-16-le")

MSG_ENCRYPTED = "此檔案有密碼保護,無法讀取"
MSG_LEGACY_XLS = "此檔案是舊版 .xls 格式,請另存為 .xlsx"
MSG_CORRUPTED = "此檔案已損毀,無法讀取"
MSG_UNAVAILABLE = "此檔案目前無法存取,請稍後再試"


class ExcelReadError(Exception):
    """讀取失敗;訊息是給使用者看的簡短原因。"""


def list_categories(data_dir: Path) -> list[str]:
    """回傳資料夾中所有 .xlsx 的分類名稱(檔名去掉副檔名)。"""
    if not data_dir.is_dir():
        return []
    return sorted(
        p.stem
        for p in data_dir.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")  # 排除 Excel 開啟中的鎖定檔
    )


def _describe_failure(tmp_path: Path) -> str:
    """依暫存複本的內容判斷讀取失敗的原因。"""
    data = tmp_path.read_bytes()
    if data.startswith(OLE_MAGIC):
        return MSG_ENCRYPTED if ENCRYPTED_MARK in data else MSG_LEGACY_XLS
    return MSG_CORRUPTED


def read_excel_safely(path: Path) -> dict[str, pd.DataFrame]:
    """複製到暫存資料夾後讀取所有工作表,讀完立即刪除暫存檔。

    讀取失敗時拋出 ExcelReadError,訊息為簡短的中文原因。
    """
    with tempfile.TemporaryDirectory(prefix="excel-lookup-") as tmp:
        tmp_path = Path(tmp) / path.name
        try:
            shutil.copyfile(path, tmp_path)
        except OSError as exc:
            raise ExcelReadError(MSG_UNAVAILABLE) from exc
        try:
            sheets = pd.read_excel(tmp_path, sheet_name=None, dtype=str, engine="openpyxl")
        except Exception as exc:
            logger.warning("無法讀取 %s:%r", path.name, exc)
            raise ExcelReadError(_describe_failure(tmp_path)) from exc
        finally:
            tmp_path.unlink(missing_ok=True)
    return {name: df.fillna("") for name, df in sheets.items()}


# 依檔案路徑快取讀取結果(成功為工作表、失敗為 ExcelReadError),檔案變動才重讀
_cache: dict[Path, tuple[tuple[float, int], dict[str, pd.DataFrame] | ExcelReadError]] = {}


def load_category(data_dir: Path, category: str) -> dict[str, pd.DataFrame]:
    """讀取某分類;檔案未變動(修改時間與大小相同)時使用快取。"""
    if category not in list_categories(data_dir):
        raise KeyError(category)
    path = data_dir / f"{category}.xlsx"
    try:
        stat = path.stat()
    except OSError as exc:  # 列出後到讀取前被刪除或改名
        raise ExcelReadError(MSG_UNAVAILABLE) from exc
    key = (stat.st_mtime, stat.st_size)
    cached = _cache.get(path)
    if cached and cached[0] == key:
        result = cached[1]
    else:
        try:
            result = read_excel_safely(path)
        except ExcelReadError as exc:
            result = exc
        _cache[path] = (key, result)
    if isinstance(result, ExcelReadError):
        raise result
    return result


def category_errors(data_dir: Path) -> dict[str, str]:
    """回傳已知讀取失敗的分類與原因(只包含已嘗試讀取過的檔案)。"""
    errors = {}
    for category in list_categories(data_dir):
        path = data_dir / f"{category}.xlsx"
        cached = _cache.get(path)
        if not cached or not isinstance(cached[1], ExcelReadError):
            continue
        stat = path.stat()
        if cached[0] == (stat.st_mtime, stat.st_size):
            errors[category] = str(cached[1])
    return errors


def search(
    data_dir: Path,
    query: str,
    category: str | None = None,
    max_results: int = 500,
) -> dict:
    """在指定分類(或全部分類)中搜尋包含關鍵字的列,不分大小寫。

    讀不了的檔案會略過並列在 errors,不影響其他檔案。
    """
    query = query.strip()
    results: list[dict] = []
    errors: list[dict] = []
    if not query:
        return {"results": results, "errors": errors}
    categories = [category] if category else list_categories(data_dir)
    remaining = max_results
    for cat in categories:
        try:
            sheets = load_category(data_dir, cat)
        except ExcelReadError as exc:
            errors.append({"category": cat, "error": str(exc)})
            continue
        for sheet_name, df in sheets.items():
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
    return {"results": results, "errors": errors}
