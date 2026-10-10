"""把儲存格的值轉成畫面上顯示的文字,以及搜尋用的正規化。"""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta
from decimal import ROUND_HALF_UP, Decimal

# 數字部分支援:0、00000、0.00、#,##0、#,##0.00、0%、0.0%、0.## 等
_NUMERIC_FMT = re.compile(r"^(?P<int>[#0,]*0)(?:\.(?P<dec>0*#*))?(?P<pct>%?)$")
_CORE_CHARS = set("0#?.,%")
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


def _split_sections(number_format: str) -> list[str]:
    """以 ; 分段(正數;負數;零;文字),引號與方括號內的 ; 不算。"""
    sections, current, quoted, bracket = [], [], False, False
    for ch in number_format:
        if ch == '"' and not bracket:
            quoted = not quoted
        elif ch == "[" and not quoted:
            bracket = True
        elif ch == "]" and not quoted:
            bracket = False
        if ch == ";" and not quoted and not bracket:
            sections.append("".join(current))
            current = []
        else:
            current.append(ch)
    sections.append("".join(current))
    return sections


def _parse_section(section: str) -> tuple[str, str | None, str] | None:
    """拆成 (前綴文字, 數字部分, 後綴文字)。

    數字部分為 "General" 表示一般格式、None 表示只有文字(例如會計格式的零顯示為 -)。
    無法處理的格式(分數、科學記號、數字中間夾文字)回傳 None。
    """
    tokens: list[tuple[str, str]] = []  # (種類, 內容),種類為 lit / core / general
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            end = section.find('"', i + 1)
            end = len(section) if end < 0 else end
            tokens.append(("lit", section[i + 1 : end]))
            i = end + 1
            continue
        if ch == "\\":
            tokens.append(("lit", section[i + 1 : i + 2]))
            i += 2
            continue
        if ch in "_*":  # _x 補空白、*x 重複填滿,文字顯示時省略
            i += 2
            continue
        if ch == "[":
            end = section.find("]", i)
            end = len(section) if end < 0 else end
            inner = section[i + 1 : end]
            if inner.startswith("$"):  # [$NT$-404] 貨幣符號;[Red]、[>100] 等略過
                tokens.append(("lit", inner[1:].split("-")[0]))
            i = end + 1
            continue
        if section[i : i + 7].lower() == "general" or ch == "@":
            tokens.append(("general", ""))
            i += 7 if ch != "@" else 1
            continue
        if ch in "Ee" and section[i + 1 : i + 2] in ("+", "-"):
            return None  # 科學記號
        tokens.append(("core" if ch in _CORE_CHARS else "lit", ch))
        i += 1
    numeric = [n for n, (kind, _) in enumerate(tokens) if kind != "lit"]
    if not numeric:
        return "".join(text for _, text in tokens), None, ""
    first, last = numeric[0], numeric[-1]
    if any(kind == "lit" for kind, _ in tokens[first : last + 1]):
        return None
    kinds = {kind for kind, _ in tokens[first : last + 1]}
    core = "General" if "general" in kinds else "".join(text for _, text in tokens[first : last + 1])
    if "?" in core or (len(kinds) > 1):
        return None
    prefix = "".join(text for _, text in tokens[:first])
    suffix = "".join(text for _, text in tokens[last + 1 :])
    return prefix, core, suffix


def _format_core(number: Decimal, core: str) -> str | None:
    """依數字部分格式化絕對值;格式不支援時回傳 None。"""
    match = _NUMERIC_FMT.match(core)
    if not match:
        return None
    int_part, dec, pct = match["int"], match["dec"] or "", match["pct"]
    number = abs(number) * (100 if pct else 1)
    number = number.quantize(Decimal(1).scaleb(-len(dec)), rounding=ROUND_HALF_UP)
    comma = "," if "," in int_part else ""
    text = f"{number:{comma}.{len(dec)}f}"
    optional = len(dec) - len(dec.rstrip("#"))  # 小數點後的 # 代表可省略的 0
    if optional:
        whole, _, frac = text.partition(".")
        frac = frac[: len(frac) - optional] + frac[len(frac) - optional :].rstrip("0")
        text = f"{whole}.{frac}" if frac else whole
    min_digits = int_part.count("0")
    if min_digits > 1 and not comma:
        whole, dot, frac = text.partition(".")
        text = whole.zfill(min_digits) + dot + frac
    return text + pct


def format_number(value: int | float, number_format: str | None) -> str:
    """依數字格式還原顯示,例如格式 00000 的 123 顯示為 00123、"NT$"#,##0 的 1000 顯示為 NT$1,000。

    支援常見的數字、百分比、貨幣格式與正負零分段;無法處理的格式以一般格式顯示。
    """
    sections = _split_sections(number_format or "General")
    if value < 0 and len(sections) >= 2:
        section, auto_sign = sections[1], False  # 負數段自帶符號或括號,數值取絕對值
    elif value == 0 and len(sections) >= 3:
        section, auto_sign = sections[2], True
    else:
        section, auto_sign = sections[0], True
    parsed = _parse_section(section)
    if parsed is None:
        return _general(value)
    prefix, core, suffix = parsed
    if core is None:
        return prefix
    number = Decimal(repr(value))
    text = _general(abs(value)) if core == "General" else _format_core(number, core)
    if text is None:
        return _general(value)
    negative = number < 0 and auto_sign and any(ch in "123456789" for ch in text)
    return f"{'-' if negative else ''}{prefix}{text}{suffix}"


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
