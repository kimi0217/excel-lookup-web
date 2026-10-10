from pathlib import Path

from fastapi.testclient import TestClient

import excel_reader
from app import create_app


def search(data_dir: Path, query: str, **kwargs) -> dict:
    """建立 Catalog、同步載入完成後搜尋。"""
    catalog = excel_reader.Catalog(data_dir, kwargs.pop("include_hidden", False))
    catalog.refresh()
    return catalog.search(query, **kwargs)


def loaded_client(config: dict, **kwargs) -> TestClient:
    """建立 app 並同步載入完成(不啟動背景工作),用於測試搜尋結果。"""
    app = create_app(config)
    app.state.catalog.refresh()
    return TestClient(app, **kwargs)
