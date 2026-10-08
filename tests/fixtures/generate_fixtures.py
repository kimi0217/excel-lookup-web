"""產生各種格式的測試用 Excel,模擬客戶實際可能遇到的檔案。

用法:python tests/fixtures/generate_fixtures.py [輸出資料夾]
預設輸出到 tests/fixtures/generated/(不進版控)。每種狀況一個檔案,
避免單一檔案出錯時影響其他狀況的判斷。
"""

from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

DEFAULT_OUT = Path(__file__).resolve().parent / "generated"


def title_rows(path: Path) -> None:
    """表格上方有標題列與說明列,真正的表頭在第 4 列。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "員工名冊"
    ws.append(["2026 年度員工名冊"])
    ws.append(["製表日期:2026-10-01", None, "製表人:人資部"])
    ws.append([])
    ws.append(["工號", "姓名", "部門"])
    ws.append(["E001", "王小明", "資訊部"])
    ws.append(["E002", "李小華", "業務部"])
    wb.save(path)


def merged_cells(path: Path) -> None:
    """部門欄位用合併儲存格涵蓋多列,表頭也有跨欄合併。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "部門人員"
    ws.append(["部門", "人員", None])
    ws.merge_cells("B1:C1")
    ws.append([None, "姓名", "分機"])
    ws.merge_cells("A1:A2")
    ws.append(["資訊部", "王小明", "101"])
    ws.append([None, "陳大文", "102"])
    ws.append([None, "林美玲", "103"])
    ws.merge_cells("A3:A5")
    ws.append(["業務部", "李小華", "201"])
    wb.save(path)


def blank_rows_cols(path: Path) -> None:
    """表格前有空白列、左側與中間有空白欄、資料中有空白列。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "設備清單"
    ws.append([])
    ws.append([])
    ws.append([None, "編號", None, "名稱", "位置"])
    ws.append([None, "A001", None, "筆電", "3F"])
    ws.append([])
    ws.append([None, "A002", None, "螢幕", "5F"])
    ws.append([])
    ws.append([])
    ws.append([None, "A003", None, "印表機", "1F"])
    wb.save(path)


def dates(path: Path) -> None:
    """日期、日期時間、自訂日期格式(民國年風格的顯示格式)。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "合約"
    ws.append(["合約編號", "簽約日", "到期時間", "自訂格式日期"])
    rows = [
        ("C001", date(2026, 1, 15), datetime(2026, 12, 31, 17, 30), date(2026, 3, 1)),
        ("C002", date(2025, 7, 1), datetime(2027, 6, 30, 0, 0), date(2025, 11, 20)),
    ]
    for r, (code, signed, expires, custom) in enumerate(rows, start=2):
        ws.cell(r, 1, code)
        ws.cell(r, 2, signed).number_format = "yyyy/mm/dd"
        ws.cell(r, 3, expires).number_format = "yyyy/mm/dd hh:mm"
        ws.cell(r, 4, custom).number_format = 'yyyy"年"m"月"d"日"'
    wb.save(path)


def leading_zeros(path: Path) -> None:
    """代碼以數字儲存、用數字格式顯示開頭 0;另一欄是真正的文字代碼作對照。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "料號"
    ws.append(["料號(數字格式)", "郵遞區號(數字格式)", "料號(文字)", "品名"])
    for r, (num, zipcode, text, name) in enumerate(
        [(123, 100, "00123", "螺絲"), (45, 802, "00045", "螺帽"), (7, 51, "00007", "墊片")],
        start=2,
    ):
        ws.cell(r, 1, num).number_format = "00000"
        ws.cell(r, 2, zipcode).number_format = "000"
        ws.cell(r, 3, text)
        ws.cell(r, 4, name)
    wb.save(path)


def hidden_sheet(path: Path) -> None:
    """一般工作表 + 隱藏工作表 + 極度隱藏工作表。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "公開資料"
    ws.append(["項目", "內容"])
    ws.append(["公告", "公開內容"])
    hidden = wb.create_sheet("隱藏資料")
    hidden.append(["項目", "內容"])
    hidden.append(["機密", "隱藏內容"])
    hidden.sheet_state = "hidden"
    very_hidden = wb.create_sheet("極度隱藏")
    very_hidden.append(["項目", "內容"])
    very_hidden.append(["機密", "極度隱藏內容"])
    very_hidden.sheet_state = "veryHidden"
    wb.save(path)


def formulas(path: Path) -> None:
    """含公式的欄位。注意:openpyxl 存檔不會寫入公式的計算結果(快取值)。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "報價"
    ws.append(["品名", "單價", "數量", "小計", "說明"])
    ws.append(["筆電", 30000, 2, "=B2*C2", '=A2&"x"&C2'])
    ws.append(["螢幕", 5000, 3, "=B3*C3", '=A3&"x"&C3'])
    ws.append(["合計", None, None, "=SUM(D2:D3)", None])
    wb.save(path)


def duplicate_headers(path: Path) -> None:
    """重複的欄位名稱,以及空白的欄位名稱。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "聯絡人"
    ws.append(["姓名", "電話", "電話", None, "備註", "備註"])
    ws.append(["王小明", "02-1234", "0912-345678", "無表頭的值", "主要", "次要"])
    wb.save(path)


def protected(path: Path) -> None:
    """工作表與活頁簿結構以密碼保護(Excel 中可開啟檢視,但不能編輯)。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "受保護"
    ws.append(["項目", "內容"])
    ws.append(["密碼保護", "受保護內容"])
    ws.protection.sheet = True
    ws.protection.password = "1234"
    wb.security.workbookPassword = "1234"
    wb.security.lockStructure = True
    wb.save(path)


def encrypted_like(path: Path) -> None:
    """模擬「開啟需要密碼」的加密檔。

    真正的加密 .xlsx 是 OLE 複合文件(不是 zip),openpyxl 無法產生;
    這裡寫入 OLE 檔頭加填充資料,模擬程式讀到加密檔時看到的格式。
    """
    ole_magic = bytes.fromhex("D0CF11E0A1B11AE1")
    path.write_bytes(ole_magic + b"\x00" * 504 + b"EncryptedPackage" + b"\x00" * 1024)


def corrupted(path: Path) -> None:
    """損毀檔:正常產生後截斷一半。"""
    wb = Workbook()
    wb.active.append(["項目", "內容"])
    wb.active.append(["損毀", "損毀內容"])
    wb.save(path)
    data = path.read_bytes()
    path.write_bytes(data[: len(data) // 2])


def large(path: Path, rows: int = 50_000) -> None:
    """5 萬列 x 8 欄的大檔(使用 write_only 模式產生)。"""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("大量資料")
    ws.append(["編號", "姓名", "部門", "電話", "地址", "金額", "日期", "備註"])
    depts = ["資訊部", "業務部", "人資部", "財務部", "研發部"]
    for i in range(1, rows + 1):
        ws.append([
            f"N{i:06d}", f"員工{i}", depts[i % len(depts)], f"02-{i % 10000:04d}-{i % 1000:03d}",
            f"台北市信義路{i % 500}號", i * 10, date(2026, 1 + i % 12, 1 + i % 28), f"備註{i}",
        ])
    wb.save(path)


CASES = {
    "01_標題列": title_rows,
    "02_合併儲存格": merged_cells,
    "03_空白列與空白欄": blank_rows_cols,
    "04_日期欄位": dates,
    "05_開頭0代碼": leading_zeros,
    "06_隱藏工作表": hidden_sheet,
    "07_公式": formulas,
    "08_重複欄位名稱": duplicate_headers,
    "09_密碼保護": protected,
    "10_加密檔": encrypted_like,
    "11_損毀檔": corrupted,
    "12_五萬列大檔": large,
}


def main(out_dir: Path = DEFAULT_OUT) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, build in CASES.items():
        path = out_dir / f"{name}.xlsx"
        build(path)
        print(f"已產生 {path.name}({path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_OUT)
