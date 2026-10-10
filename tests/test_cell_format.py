from datetime import date, datetime, time
from pathlib import Path

import pytest
from openpyxl import Workbook

import excel_reader
from cell_format import display_value, format_number, normalize
from conftest import search


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (datetime(2026, 1, 15), "2026-01-15"),
        (datetime(2026, 12, 31, 17, 30), "2026-12-31 17:30"),
        (datetime(2026, 12, 31, 17, 30, 5), "2026-12-31 17:30:05"),
        (date(2026, 3, 1), "2026-03-01"),
        (time(8, 0), "08:00"),
    ],
)
def test_dates(value, expected):
    assert display_value(value) == expected


@pytest.mark.parametrize(
    ("value", "fmt", "expected"),
    [
        (123, "00000", "00123"),
        (7, "000", "007"),
        (123456, "00000", "123456"),
        (123, "General", "123"),
        (30000.0, "General", "30000"),
        (0.1 + 0.2, "General", "0.3"),
        (1234.5, "0.00", "1234.50"),
        (2.345, "0.00", "2.35"),  # 四捨五入與 Excel 相同
        (1234567, "#,##0", "1,234,567"),
        (-1234.5, "#,##0.00", "-1,234.50"),
        (0.153, "0%", "15%"),
        (0.1534, "0.0%", "15.3%"),
        (1.5, "0.##", "1.5"),
        (2.0, "0.##", "2"),
        (5, "[Red]0000", "0005"),
        (123, "@", "123"),
        (99, '"NT$"#,##0', "99"),  # 無法辨識的格式改用一般格式
    ],
)
def test_number_formats(value, fmt, expected):
    assert format_number(value, fmt) == expected


def test_bool_and_none():
    assert display_value(True) == "TRUE"
    assert display_value(None) == ""


def test_normalize():
    assert normalize("2026/01/15") == normalize("2026-01-15")
    assert normalize("1,234,567") == "1234567"
    assert normalize("王,小明") == "王,小明"  # 非數字間的逗號保留
    assert normalize("ABC") == "abc"


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    wb = Workbook()
    ws = wb.active
    ws.append(["料號", "品名", "簽約日", "到期時間", "金額"])
    ws.append([123, "螺絲", datetime(2026, 1, 15), datetime(2026, 12, 31, 17, 30), 1234567])
    ws["A2"].number_format = "00000"
    ws["C2"].number_format = "yyyy/mm/dd"
    ws["D2"].number_format = "yyyy/mm/dd hh:mm"
    ws["E2"].number_format = "#,##0"
    wb.save(d / "料號.xlsx")
    return d


def test_read_restores_display(data_dir):
    sheet = excel_reader.read_excel_safely(data_dir / "料號.xlsx")["Sheet"]
    assert sheet.rows == [["00123", "螺絲", "2026-01-15", "2026-12-31 17:30", "1,234,567"]]


@pytest.mark.parametrize("query", ["00123", "2026/01/15", "2026-01-15", "2026/12/31 17:30", "1234567", "1,234,567"])
def test_search_matches_displayed_text(data_dir, query):
    results = search(data_dir, query)["results"]
    assert results and results[0]["rows"][0][1] == "螺絲"


def test_search_does_not_match_across_cells(data_dir):
    # 「螺絲」與「2026」分屬相鄰儲存格,不應被當成連續文字
    assert search(data_dir, "螺絲2026")["results"] == []
