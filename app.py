"""內網唯讀 Excel 查詢網頁。啟動:python app.py"""

from __future__ import annotations

import asyncio
import json
import logging
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal

import uvicorn
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse

import backup
import excel_reader

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
BACKUP_CHECK_SECONDS = 60 * 60  # 每小時檢查一次,當日已備份就略過

logger = logging.getLogger(__name__)


def load_config(config_path: Path | None = None) -> dict:
    """讀取設定檔。未指定時優先讀取 config.json(含真實路徑,不進版控),否則使用 config.example.json。

    相對路徑(data_dir、backup_dir)以專案資料夾為準。
    """
    candidates = [config_path] if config_path else [BASE_DIR / "config.json", BASE_DIR / "config.example.json"]
    for path in candidates:
        if path.is_file():
            # utf-8-sig:Windows 記事本可能存成含 BOM 的 UTF-8
            config = json.loads(path.read_text(encoding="utf-8-sig"))
            break
    else:
        config = {}
    for key, default in (("data_dir", "./data"), ("backup_dir", "./backups")):
        path = Path(config.get(key, default)).expanduser()
        config[key] = path if path.is_absolute() else BASE_DIR / path
    return config


async def backup_loop(data_dir: Path, backup_dir: Path, keep_days: int) -> None:
    while True:
        try:
            await asyncio.to_thread(backup.run_backup, data_dir, backup_dir, keep_days)
        except Exception:
            logger.exception("自動備份失敗")
        await asyncio.sleep(BACKUP_CHECK_SECONDS)


async def reload_loop(catalog: excel_reader.Catalog, interval: float) -> None:
    """啟動時先載入全部檔案,之後定期檢查並重新載入有變動的檔案。"""
    while True:
        try:
            await asyncio.to_thread(catalog.refresh)
        except Exception:
            logger.exception("載入 Excel 失敗")
        await asyncio.sleep(interval)


def create_app(config: dict | None = None) -> FastAPI:
    config = config or load_config()
    data_dir: Path = config["data_dir"]
    max_results: int = config.get("max_results", 500)
    include_hidden: bool = config.get("include_hidden_sheets", False)
    backup_dir: Path | None = config.get("backup_dir")
    keep_days: int = config.get("backup_keep_days", 30)
    reload_interval: float = config.get("reload_interval_seconds", 10)
    catalog = excel_reader.Catalog(data_dir, include_hidden)

    @asynccontextmanager
    async def lifespan(_: FastAPI):
        if not excel_reader.merge_support_ok():
            logger.warning(excel_reader.MSG_MERGE_UNSUPPORTED, excel_reader.SHEET_PATH_ATTR)
        tasks = [asyncio.create_task(reload_loop(catalog, reload_interval))]
        if backup_dir:
            tasks.append(asyncio.create_task(backup_loop(data_dir, backup_dir, keep_days)))
        yield
        for task in tasks:
            task.cancel()

    app = FastAPI(title="Excel 查詢", lifespan=lifespan)
    app.state.catalog = catalog

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(STATIC_DIR / "index.html")

    @app.get("/api/categories")
    def categories() -> dict:
        """資料夾連線狀態,以及檔案清單與各檔案的載入狀態(ok / error / loading)。"""
        return {"folder": catalog.folder_status(), "files": catalog.categories()}

    @app.get("/api/files/{name}")
    def file_info(name: str) -> dict:
        """單一檔案的摘要與工作表清單。"""
        try:
            return catalog.file_info(name)
        except KeyError:
            raise HTTPException(status_code=404, detail="找不到此檔案") from None

    @app.get("/api/files/{name}/rows")
    def file_rows(
        name: str,
        sheet: str | None = Query(None, description="工作表名稱,留空為第一個工作表"),
        q: str = Query("", description="篩選關鍵字,留空顯示全部"),
        sort: int | None = Query(None, ge=0, description="排序欄位(0 起算)"),
        order: Literal["asc", "desc"] = Query("asc", description="排序方向"),
        page: int = Query(1, ge=1),
        page_size: int = Query(100, ge=1, le=500),
    ) -> dict:
        """工作表的分頁資料,可篩選與排序;不帶關鍵字時回傳整份資料的第 page 頁。"""
        try:
            return catalog.rows(name, sheet, q, sort, order == "desc", page, page_size)
        except KeyError:
            raise HTTPException(status_code=404, detail="找不到此檔案") from None
        except excel_reader.SheetNotFound:
            raise HTTPException(status_code=404, detail="找不到此工作表") from None
        except excel_reader.SortColumnNotFound:
            raise HTTPException(status_code=400, detail="排序欄位不存在") from None

    @app.get("/api/search")
    def search(
        q: str = Query("", description="搜尋關鍵字"),
        category: str | None = Query(None, description="分類名稱,留空表示全部"),
    ) -> dict:
        if category and not catalog.has(category):
            raise HTTPException(status_code=404, detail="找不到此分類")
        return catalog.search(q, category or None, max_results)

    return app


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    cfg = load_config()
    uvicorn.run(create_app(cfg), host=cfg.get("host", "0.0.0.0"), port=cfg.get("port", 8000))
