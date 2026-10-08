"""內網唯讀 Excel 查詢網頁。啟動:python app.py"""

from __future__ import annotations

import json
from pathlib import Path

import uvicorn
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse

import excel_reader

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"


def load_config() -> dict:
    """優先讀取 config.json(含真實路徑,不進版控),否則使用 config.example.json。"""
    for name in ("config.json", "config.example.json"):
        path = BASE_DIR / name
        if path.is_file():
            config = json.loads(path.read_text(encoding="utf-8"))
            break
    else:
        config = {}
    data_dir = Path(config.get("data_dir", "./data")).expanduser()
    if not data_dir.is_absolute():
        data_dir = BASE_DIR / data_dir
    config["data_dir"] = data_dir
    return config


def create_app(config: dict | None = None) -> FastAPI:
    config = config or load_config()
    data_dir: Path = config["data_dir"]
    max_results: int = config.get("max_results", 500)

    app = FastAPI(title="Excel 查詢")

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(STATIC_DIR / "index.html")

    @app.get("/api/categories")
    def categories() -> list[str]:
        return excel_reader.list_categories(data_dir)

    @app.get("/api/search")
    def search(
        q: str = Query("", description="搜尋關鍵字"),
        category: str | None = Query(None, description="分類名稱,留空表示全部"),
    ) -> list[dict]:
        if category and category not in excel_reader.list_categories(data_dir):
            raise HTTPException(status_code=404, detail="找不到此分類")
        return excel_reader.search(data_dir, q, category or None, max_results)

    return app


if __name__ == "__main__":
    cfg = load_config()
    uvicorn.run(create_app(cfg), host=cfg.get("host", "0.0.0.0"), port=cfg.get("port", 8000))
