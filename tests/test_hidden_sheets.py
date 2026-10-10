from pathlib import Path

import pytest
from openpyxl import Workbook

import excel_reader
from conftest import loaded_client


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    wb = Workbook()
    wb.active.title = "公開"
    wb.active.append(["項目", "內容"])
    wb.active.append(["公告", "共同關鍵字 公開"])
    for title, state in (("隱藏", "hidden"), ("極度隱藏", "veryHidden")):
        ws = wb.create_sheet(title)
        ws.append(["項目", "內容"])
        ws.append(["機密", f"共同關鍵字 {title}"])
        ws.sheet_state = state
    wb.save(d / "文件.xlsx")
    return d


def test_hidden_sheets_skipped_by_default(data_dir):
    sheets = excel_reader.read_excel_safely(data_dir / "文件.xlsx")
    assert list(sheets) == ["公開"]


def test_hidden_sheets_included_when_enabled(data_dir):
    sheets = excel_reader.read_excel_safely(data_dir / "文件.xlsx", include_hidden=True)
    assert list(sheets) == ["公開", "隱藏", "極度隱藏"]


def search_sheets(config: dict) -> list[str]:
    client = loaded_client(config)
    body = client.get("/api/search", params={"q": "共同關鍵字"}).json()
    return [g["sheet"] for g in body["results"]]


def test_api_default_hides_hidden_sheets(data_dir):
    assert search_sheets({"data_dir": data_dir}) == ["公開"]
    assert search_sheets({"data_dir": data_dir, "include_hidden_sheets": False}) == ["公開"]


def test_api_include_hidden_sheets(data_dir):
    assert search_sheets({"data_dir": data_dir, "include_hidden_sheets": True}) == ["公開", "隱藏", "極度隱藏"]


def test_example_config_defaults_to_false():
    import json

    config = json.loads((Path(__file__).parents[1] / "config.example.json").read_text(encoding="utf-8"))
    assert config["include_hidden_sheets"] is False
