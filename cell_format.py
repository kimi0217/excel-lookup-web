"""把儲存格的值轉成畫面上顯示的文字,以及搜尋用的正規化。"""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta
from decimal import ROUND_HALF_UP, Decimal

_BRACKETS = re.compile(r"\[[^\]]*\]")  # [Red]、[$-404] 等顏色或地區標記
# 支援的數字格式:0、00000、0.00、#,##0、#,##0.00、0%、0.0% 等
_NUMERIC_FMT = re.compile(r"^(?P<int>[#0,]*0)(?:\.(?P<dec>0*#*))?(?P<pct>%?)$")
_DIGIT_COMMA = re.compile(r"(?<=\d),(?=\d)")


def format_datetime(value: datetime | date | time) -> str:
    """日期顯示 YYYY-MM-DD;時間不是 00:00:00 才附上時間(秒數為 0 時省略)。"""
    if isinstance(value, datetime):
        if value.time() == time(0, 0):
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d ") + format_datetime(value.time())
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return value.strftime("%H:%M:%S" if value.second else "%H:%M")


def _general(value: int | float) -> str:
    if isinstance(value, int):
        return str(value)
    if value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    text = f"{value:.15g}"
    return format(Decimal(repr(value)), "f") if "e" in text else text


def format_number(value: int | float, number_format: str | None) -> str:
    """依數字格式還原顯示,例如格式 00000 的 123 顯示為 00123。

    只支援常見的數字格式;無法辨識的格式(貨幣符號、文字格式等)以一般格式顯示。
    """
    sections = (number_format or "General").split(";")
    section = _BRACKETS.sub("", sections[0]).strip()
    match = _NUMERIC_FMT.match(section)
    if not match:
        return _general(value)
    int_part, dec, pct = match["int"], match["dec"] or "", match["pct"]
    number = Decimal(repr(value)) * (100 if pct else 1)
    number = number.quantize(Decimal(1).scaleb(-len(dec)), rounding=ROUND_HALF_UP)
    comma = "," if "," in int_part else ""
    text = f"{abs(number):{comma}.{len(dec)}f}"
    optional = len(dec) - len(dec.rstrip("#"))  # 小數點後的 # 代表可省略的 0
    if optional:
        whole, _, frac = text.partition(".")
        frac = frac[: len(frac) - optional] + frac[len(frac) - optional :].rstrip("0")
        text = f"{whole}.{frac}" if frac else whole
    min_digits = int_part.count("0")
    if min_digits > 1 and not comma:
        whole, dot, frac = text.partition(".")
        text = whole.zfill(min_digits) + dot + frac
    sign = "-" if number < 0 else ""
    return f"{sign}{text}{pct}"


def display_value(value: object, number_format: str | None = None) -> str:
    """儲存格值 → 顯示文字。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, time)):
        return format_datetime(value)
    if isinstance(value, timedelta):
        return str(value)
    if isinstance(value, (int, float)):
        return format_number(value, number_format)
    return str(value)


def normalize(text: str) -> str:
    """搜尋用正規化:不分大小寫、/ 與 - 視為相同、忽略數字間的千分位逗號。"""
    return _DIGIT_COMMA.sub("", text.casefold().replace("/", "-"))
