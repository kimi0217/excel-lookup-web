"""從工作表的儲存格網格整理出表格:偵測表頭、填入合併儲存格、整理欄位名稱。

這裡只處理已轉成顯示文字的網格,不碰檔案,方便單獨測試。
"""

from __future__ import annotations

from dataclasses import dataclass

HEADER_SEARCH_ROWS = 10  # 在前 10 個非空白列中找表頭
HEADER_MIN_COVERAGE = 0.8  # 表頭的非空白格數至少要達到候選列最大值的 80%
HEADER_MIN_TEXT_RATIO = 0.5  # 表頭的非空白格中,文字要超過一半


@dataclass(frozen=True)
class Merge:
    """合併範圍,0 起算、包含兩端。"""

    min_row: int
    min_col: int
    max_row: int
    max_col: int


def drop_hidden(
    grid: list[list[str]],
    is_text: list[list[bool]],
    merges: list[Merge],
    hidden_rows: set[int],
    hidden_cols: set[int],
) -> tuple[list[list[str]], list[list[bool]], list[Merge]]:
    """移除隱藏的列與欄(0 起算),合併範圍對應到剩下的列與欄。

    合併範圍的左上角被隱藏時,值移到剩下範圍的新左上角,合併範圍內其他列仍會填入該值。
    """
    if not hidden_rows and not hidden_cols:
        return grid, is_text, merges
    width = len(grid[0]) if grid else 0
    rows = [r for r in range(len(grid)) if r not in hidden_rows]
    cols = [c for c in range(width) if c not in hidden_cols]
    row_pos = {r: i for i, r in enumerate(rows)}
    col_pos = {c: i for i, c in enumerate(cols)}
    new_grid = [[grid[r][c] for c in cols] for r in rows]
    new_text = [[is_text[r][c] for c in cols] for r in rows if r < len(is_text)]  # is_text 只有前幾列
    new_merges = []
    for m in merges:
        kept_rows = [r for r in range(m.min_row, m.max_row + 1) if r in row_pos]
        kept_cols = [c for c in range(m.min_col, m.max_col + 1) if c in col_pos]
        if not kept_rows or not kept_cols:
            continue
        moved = Merge(row_pos[kept_rows[0]], col_pos[kept_cols[0]], row_pos[kept_rows[-1]], col_pos[kept_cols[-1]])
        new_grid[moved.min_row][moved.min_col] = grid[m.min_row][m.min_col]
        if moved.min_row < len(new_text) and m.min_row < len(is_text):
            new_text[moved.min_row][moved.min_col] = is_text[m.min_row][m.min_col]
        new_merges.append(moved)
    return new_grid, new_text, new_merges


def _filled_row(grid: list[list[str]], is_text: list[list[bool]], r: int, merges: list[Merge]):
    """回傳第 r 列「合併範圍視為每格都有值」時的值與是否為文字(只用來偵測表頭)。"""
    values = list(grid[r])
    texts = list(is_text[r]) if r < len(is_text) else [False] * len(values)
    for m in merges:
        if m.min_row <= r <= m.max_row:
            top_value = grid[m.min_row][m.min_col]
            top_text = m.min_row < len(is_text) and is_text[m.min_row][m.min_col]
            for c in range(m.min_col, min(m.max_col + 1, len(values))):
                values[c], texts[c] = top_value, top_text
    return values, texts


def detect_header(grid: list[list[str]], is_text: list[list[bool]], merges: list[Merge]) -> int | None:
    """回傳表頭所在列(0 起算);整張表空白時回傳 None。

    在前 10 個非空白列中,選出最前面一個符合下列條件的列:
    - 非空白格數(合併範圍算每一格)達到候選列最大值的 80%
    - 非空白格大部分是文字
    - 至少有 2 個不同的值(跨整列合併的標題列只有 1 個值),整張表只有 1 欄時例外
    - 下面還有資料列
    找不到時使用第一個非空白列。
    """
    merged_rows: set[int] = set()  # 被有值的合併範圍涵蓋的列
    for m in merges:
        if grid[m.min_row][m.min_col]:
            merged_rows.update(range(m.min_row, min(m.max_row + 1, len(grid))))
    nonblank = [r for r in range(len(grid)) if r in merged_rows or any(grid[r])]
    if not nonblank:
        return None
    candidates = []
    for r in nonblank[:HEADER_SEARCH_ROWS]:
        values, texts = _filled_row(grid, is_text, r, merges)
        filled = [(v, t) for v, t in zip(values, texts) if v]
        candidates.append((r, len(filled), len({v for v, _ in filled}), sum(t for _, t in filled)))
    max_coverage = max(coverage for _, coverage, _, _ in candidates)
    last_row = nonblank[-1]
    for r, coverage, distinct, text_count in candidates:
        if (
            r < last_row
            and coverage >= max_coverage * HEADER_MIN_COVERAGE
            and text_count > coverage * HEADER_MIN_TEXT_RATIO
            and (distinct >= 2 or max_coverage == 1)
        ):
            return r
    return nonblank[0]


def tidy_columns(header: list[str]) -> list[str]:
    """空白欄名改為「欄位N」,重複欄名改為「電話」「電話(2)」。"""
    names = [name or f"欄位{i}" for i, name in enumerate(header, start=1)]
    result: list[str] = []
    used: set[str] = set()
    counts: dict[str, int] = {}
    for name in names:
        unique = name
        if unique in used:
            n = counts.get(name, 1)
            while unique in used:
                n += 1
                unique = f"{name}({n})"
            counts[name] = n
        used.add(unique)
        result.append(unique)
    return result


def build_table(
    grid: list[list[str]], is_text: list[list[bool]], merges: list[Merge], header_row: int | None = None
) -> tuple[list[str], list[list[str]]]:
    """回傳 (欄位名稱, 資料列)。

    header_row(0 起算)為手動指定的表頭列,有指定時不做自動偵測。
    - 表頭上方的列(標題、說明)不列入資料
    - 資料區的合併範圍每一格都填入合併格的值;起點在表頭或以上的合併不處理
    - 略過全空白的列,以及表頭與資料都空白的欄
    """
    if header_row is None:
        header_row = detect_header(grid, is_text, merges)
    if header_row is None:
        return [], []
    data_merges = [m for m in merges if m.min_row > header_row]
    if data_merges:  # 有需要填值時才複製網格,避免大檔多一份完整複製
        grid = [list(row) for row in grid]
    for m in data_merges:
        value = grid[m.min_row][m.min_col]
        for r in range(m.min_row, min(m.max_row + 1, len(grid))):
            for c in range(m.min_col, min(m.max_col + 1, len(grid[r]))):
                grid[r][c] = value
    header = grid[header_row]
    rows = [row for row in grid[header_row + 1 :] if any(row)]
    keep = [c for c in range(len(header)) if header[c] or any(row[c] for row in rows)]
    columns = tidy_columns([header[c] for c in keep])
    if len(keep) == len(header):  # 沒有空白欄要去掉
        return columns, rows
    return columns, [[row[c] for c in keep] for row in rows]
