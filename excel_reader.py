"""唯讀讀取 Excel:一律先複製到暫存,讀完立即刪除,絕不碰原始檔。"""

from __future__ import annotations

import logging
import shutil
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from cell_format import display_value, normalize

logger = logging.getLogger(__name__)

# 加密的 .xlsx 與舊版 .xls 都是 OLE 複合文件(不是 zip)
OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
# 加密檔的 OLE 目錄中會有這個串流名稱(UTF-16LE)
ENCRYPTED_MARK = "EncryptedPackage".encode("utf-16-le")

MSG_ENCRYPTED = "此檔案有密碼保護,無法讀取"
MSG_LEGACY_XLS = "此檔案是舊版 .xls 格式,請另存為 .xlsx"
MSG_CORRUPTED = "此檔案已損毀,無法讀取"
MSG_UNAVAILABLE = "此檔案目前無法存取,請稍後再試"


ROW_SEPARATOR = "\x1f"  # 組合搜尋字串時分隔儲存格,避免跨儲存格誤判為符合


class ExcelReadError(Exception):
    """讀取失敗;訊息是給使用者看的簡短原因。"""


@dataclass
class Sheet:
    """一個工作表:欄位名稱、顯示用的各列文字,以及每列正規化後的搜尋字串。"""

    columns: list[str]
    rows: list[list[str]]
    haystack: pd.Series

    @classmethod
    def build(cls, columns: list[str], rows: list[list[str]]) -> Sheet:
        haystack = pd.Series(
            [normalize(ROW_SEPARATOR.join(row)) for row in rows], dtype=object
        )
        return cls(columns, rows, haystack)


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


def _unique_columns(header: list[str]) -> list[str]:
    """空白欄名改為「Unnamed: n」,重複欄名加上 .1、.2(與 pandas 相同)。"""
    columns, seen = [], {}
    for i, name in enumerate(header):
        name = name or f"Unnamed: {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        columns.append(name)
    return columns


def _read_sheet(ws) -> Sheet:
    """第一列為欄位名稱;完全空白的列略過,尾端的空白欄去掉。"""
    ws.reset_dimensions()  # 有些程式產生的檔案記錄的範圍不正確,改為實際讀到哪算到哪
    grid = [
        [display_value(cell.value, cell.number_format) for cell in row]
        for row in ws.iter_rows()
    ]
    if not grid:
        return Sheet.build([], [])
    width = max((max((i + 1 for i, v in enumerate(r) if v), default=0) for r in grid), default=0)
    grid = [(r + [""] * width)[:width] for r in grid]
    header, data = grid[0], [r for r in grid[1:] if any(r)]
    return Sheet.build(_unique_columns(header), data)


def _read_workbook(path: Path, include_hidden: bool) -> dict[str, Sheet]:
    wb = load_workbook(path, read_only=True, data_only=True)  # data_only:公式讀取 Excel 存下的計算結果
    try:
        return {
            ws.title: _read_sheet(ws)
            for ws in wb.worksheets
            if include_hidden or ws.sheet_state == "visible"  # 隱藏、極度隱藏的工作表預設不讀
        }
    finally:
        wb.close()  # 先釋放檔案,Windows 才能刪除暫存檔


def read_excel_safely(path: Path, include_hidden: bool = False) -> dict[str, Sheet]:
    """複製到暫存資料夾後讀取工作表,讀完立即刪除暫存檔。

    include_hidden 為 False 時不讀取隱藏與極度隱藏的工作表。

    讀取失敗時拋出 ExcelReadError,訊息為簡短的中文原因。
    """
    with tempfile.TemporaryDirectory(prefix="excel-lookup-") as tmp:
        tmp_path = Path(tmp) / path.name
        try:
            shutil.copyfile(path, tmp_path)
        except OSError as exc:
            raise ExcelReadError(MSG_UNAVAILABLE) from exc
        try:
            return _read_workbook(tmp_path, include_hidden)
        except Exception as exc:
            logger.warning("無法讀取 %s:%r", path.name, exc)
            raise ExcelReadError(_describe_failure(tmp_path)) from exc
        finally:
            tmp_path.unlink(missing_ok=True)


@dataclass
class _Entry:
    key: tuple[float, int]  # 讀取時檔案的修改時間與大小,用來偵測變動
    sheets: dict[str, Sheet]
    error: str | None = None


class Catalog:
    """資料夾中所有 Excel 的讀取結果。

    refresh() 由背景執行緒定期呼叫,讀取新增或變動的檔案;搜尋只使用已載入的資料,
    不會在搜尋時讀檔。檔案變動後重新載入完成前,搜尋仍使用舊資料。
    """

    def __init__(self, data_dir: Path, include_hidden: bool = False) -> None:
        self.data_dir = data_dir
        self.include_hidden = include_hidden
        self._entries: dict[str, _Entry] = {}
        self._refresh_lock = threading.Lock()

    def refresh(self) -> None:
        """載入新增或變動的檔案(小檔優先),移除已刪除的檔案。"""
        with self._refresh_lock:
            names = list_categories(self.data_dir)
            for gone in set(self._entries) - set(names):
                del self._entries[gone]
            pending = []
            for name in names:
                try:
                    stat = (self.data_dir / f"{name}.xlsx").stat()
                except OSError:  # 列出後被刪除或改名,下次再處理
                    continue
                key = (stat.st_mtime, stat.st_size)
                entry = self._entries.get(name)
                if entry is None or entry.key != key:
                    pending.append((stat.st_size, name, key))
            for _, name, key in sorted(pending):
                path = self.data_dir / f"{name}.xlsx"
                try:
                    self._entries[name] = _Entry(key, read_excel_safely(path, self.include_hidden))
                except ExcelReadError as exc:
                    self._entries[name] = _Entry(key, {}, str(exc))

    def categories(self) -> list[dict]:
        """分類清單;status 為 ok、error(附 error 原因)或 loading(尚未載入)。"""
        result = []
        for name in list_categories(self.data_dir):
            entry = self._entries.get(name)
            if entry is None:
                result.append({"name": name, "status": "loading", "error": None})
            elif entry.error:
                result.append({"name": name, "status": "error", "error": entry.error})
            else:
                result.append({"name": name, "status": "ok", "error": None})
        return result

    def search(self, query: str, category: str | None = None, max_results: int = 500) -> dict:
        """在指定分類(或全部分類)中搜尋包含關鍵字的列。

        不分大小寫、/ 與 - 視為相同、忽略數字間的千分位逗號。
        讀不了的檔案列在 errors,尚未載入完成的檔案列在 loading,都不影響其他檔案。
        """
        needle = normalize(query.strip())
        results: list[dict] = []
        errors: list[dict] = []
        loading: list[str] = []
        if not needle:
            return {"results": results, "errors": errors, "loading": loading}
        remaining = max_results
        for cat in [category] if category else list_categories(self.data_dir):
            entry = self._entries.get(cat)
            if entry is None:
                loading.append(cat)
                continue
            if entry.error:
                errors.append({"category": cat, "error": entry.error})
                continue
            for sheet_name, sheet in entry.sheets.items():
                if remaining <= 0:
                    break
                if not sheet.rows:
                    continue
                hits = sheet.haystack.str.contains(needle, regex=False)
                indexes = hits[hits].index[:remaining]
                if len(indexes) == 0:
                    continue
                remaining -= len(indexes)
                results.append(
                    {
                        "category": cat,
                        "sheet": sheet_name,
                        "columns": sheet.columns,
                        "rows": [sheet.rows[i] for i in indexes],
                    }
                )
        return {"results": results, "errors": errors, "loading": loading}
