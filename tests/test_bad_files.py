"""壞檔(加密、損毀、舊版 .xls、讀取失敗)只影響該檔,不影響其他檔案。"""

import tempfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

import excel_reader
from app import create_app

OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")


def make_xlsx(path: Path, rows: list[list]) -> None:
    wb = Workbook()
    for row in rows:
        wb.active.append(row)
    wb.save(path)


def write_encrypted(path: Path) -> None:
    path.write_bytes(OLE_MAGIC + b"\x00" * 504 + "EncryptedPackage".encode("utf-16-le") + b"\x00" * 512)


def write_legacy_xls(path: Path) -> None:
    path.write_bytes(OLE_MAGIC + b"\x00" * 504 + "Workbook".encode("utf-16-le") + b"\x00" * 512)


def write_corrupted(path: Path) -> None:
    make_xlsx(path, [["項目"], ["王小明"]])
    data = path.read_bytes()
    path.write_bytes(data[: len(data) // 2])


@pytest.fixture
def data_dir(tmp_path: Path) -> Path:
    d = tmp_path / "data"
    d.mkdir()
    make_xlsx(d / "員工.xlsx", [["姓名", "部門"], ["王小明", "資訊部"]])
    make_xlsx(d / "設備.xlsx", [["編號", "保管人"], ["A001", "王小明"]])
    write_encrypted(d / "加密.xlsx")
    write_corrupted(d / "損毀.xlsx")
    write_legacy_xls(d / "舊版.xlsx")
    excel_reader._cache.clear()
    return d


@pytest.fixture
def client(data_dir: Path) -> TestClient:
    return TestClient(create_app({"data_dir": data_dir}), raise_server_exceptions=False)


@pytest.mark.parametrize(
    ("name", "message"),
    [
        ("加密", excel_reader.MSG_ENCRYPTED),
        ("損毀", excel_reader.MSG_CORRUPTED),
        ("舊版", excel_reader.MSG_LEGACY_XLS),
    ],
)
def test_read_error_message(data_dir, name, message):
    with pytest.raises(excel_reader.ExcelReadError, match=message):
        excel_reader.read_excel_safely(data_dir / f"{name}.xlsx")


def test_failed_read_leaves_no_temp_files(data_dir, tmp_path, monkeypatch):
    tmp_root = tmp_path / "tmp"
    tmp_root.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(tmp_root))
    for name in ("加密", "損毀", "舊版"):
        with pytest.raises(excel_reader.ExcelReadError):
            excel_reader.read_excel_safely(data_dir / f"{name}.xlsx")
    assert list(tmp_root.iterdir()) == []


def test_unexpected_error_only_affects_that_file(data_dir, monkeypatch):
    real_read = excel_reader.pd.read_excel

    def flaky_read(path, *args, **kwargs):
        if Path(path).stem == "設備":
            raise RuntimeError("意外錯誤")
        return real_read(path, *args, **kwargs)

    monkeypatch.setattr(excel_reader.pd, "read_excel", flaky_read)
    result = excel_reader.search(data_dir, "王小明")
    assert [g["category"] for g in result["results"]] == ["員工"]
    assert {"category": "設備", "error": excel_reader.MSG_CORRUPTED} in result["errors"]


def test_copy_failure_reports_unavailable(data_dir, monkeypatch):
    def deny(*args, **kwargs):
        raise PermissionError("檔案被鎖定")

    monkeypatch.setattr(excel_reader.shutil, "copyfile", deny)
    with pytest.raises(excel_reader.ExcelReadError, match=excel_reader.MSG_UNAVAILABLE):
        excel_reader.read_excel_safely(data_dir / "員工.xlsx")


def test_search_all_skips_bad_files(client):
    res = client.get("/api/search", params={"q": "王小明"})
    assert res.status_code == 200
    body = res.json()
    assert sorted(g["category"] for g in body["results"]) == ["員工", "設備"]
    assert sorted(body["errors"], key=lambda e: e["category"]) == [
        {"category": "加密", "error": excel_reader.MSG_ENCRYPTED},
        {"category": "損毀", "error": excel_reader.MSG_CORRUPTED},
        {"category": "舊版", "error": excel_reader.MSG_LEGACY_XLS},
    ]


def test_search_bad_category_returns_reason(client):
    res = client.get("/api/search", params={"q": "王小明", "category": "損毀"})
    assert res.status_code == 200
    assert res.json() == {
        "results": [],
        "errors": [{"category": "損毀", "error": excel_reader.MSG_CORRUPTED}],
    }


def test_categories_mark_failed_files_after_read(client):
    client.get("/api/search", params={"q": "任何字"})
    cats = {c["name"]: c["error"] for c in client.get("/api/categories").json()}
    assert cats == {
        "加密": excel_reader.MSG_ENCRYPTED,
        "員工": None,
        "損毀": excel_reader.MSG_CORRUPTED,
        "舊版": excel_reader.MSG_LEGACY_XLS,
        "設備": None,
    }


def test_fixed_file_is_read_again(data_dir, client):
    client.get("/api/search", params={"q": "王小明"})
    make_xlsx(data_dir / "損毀.xlsx", [["項目"], ["已修復 王小明"]])
    body = client.get("/api/search", params={"q": "已修復"}).json()
    assert [g["category"] for g in body["results"]] == ["損毀"]
    assert all(e["category"] != "損毀" for e in body["errors"])
