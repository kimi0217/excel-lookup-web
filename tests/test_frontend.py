"""前端 JavaScript 的純函式(以 node 執行;沒有 node 時略過)。"""

import json
import shutil
import subprocess
from pathlib import Path

import pytest

HTML = (Path(__file__).resolve().parents[1] / "static" / "index.html").read_text(encoding="utf-8")
SCRIPT = HTML[HTML.index("<script>") + len("<script>") : HTML.index("</script>")]
HELPERS = SCRIPT[SCRIPT.index("function escapeHtml") : SCRIPT.index("function fmtNumber")]
SHORT_TIME = SCRIPT[SCRIPT.index("function shortTime") : SCRIPT.index("// 頂部顯示資料最後更新時間")]

pytestmark = pytest.mark.skipif(shutil.which("node") is None, reason="需要 node")


def highlight(query: str, text: str) -> str:
    code = HELPERS + f"\nprocess.stdout.write(makeHighlighter({json.dumps(query)})({json.dumps(text)}));"
    result = subprocess.run(["node", "-e", code], capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, result.stderr
    return result.stdout


def test_each_keyword_highlighted():
    assert highlight("資訊部 陳大文", "資訊部") == "<mark>資訊部</mark>"
    assert highlight("資訊部 陳大文", "陳大文") == "<mark>陳大文</mark>"
    assert highlight("資訊部　陳大文", "資訊部的陳大文") == "<mark>資訊部</mark>的<mark>陳大文</mark>"


def test_longer_keyword_wins_when_overlapping():
    assert highlight("王 王小明", "王小明") == "<mark>王小明</mark>"


def test_search_rules_still_apply():
    assert highlight("2026/01/15", "2026-01-15") == "<mark>2026-01-15</mark>"
    assert highlight("1000", "NT$1,000") == "NT$<mark>1,000</mark>"
    assert highlight("abc", "xABCx") == "x<mark>ABC</mark>x"


def test_html_escaped():
    assert highlight("b", "<b>") == "&lt;<mark>b</mark>&gt;"
    assert highlight("", "<i>") == "&lt;i&gt;"


def run_js(code: str) -> str:
    result = subprocess.run(["node", "-e", code], capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, result.stderr
    return result.stdout


def test_short_time_today_and_other_day():
    code = SHORT_TIME + """
    const d = new Date(); const p = (n) => String(n).padStart(2, "0");
    const today = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
    process.stdout.write(shortTime(`${today} 09:05`) + "|" + shortTime("2020-03-04 21:34"));"""
    assert run_js(code) == "09:05|03-04 21:34"


@pytest.mark.parametrize(
    ("query", "text", "expected"),
    [
        ("amp", "A&B amp", "A&amp;B <mark>amp</mark>"),  # 不可標到 &amp; 的內部
        ("lt", "a<b", "a&lt;b"),
        ("quot", '"x"', "&quot;x&quot;"),
        ("&", "A&B", "A<mark>&amp;</mark>B"),  # 搜尋特殊字元本身
        ("<b>", "x<b>y", "x<mark>&lt;b&gt;</mark>y"),
        ("a&b", "A&B", "<mark>A&amp;B</mark>"),
    ],
)
def test_highlight_matches_raw_text_before_escaping(query, text, expected):
    assert highlight(query, text) == expected


def group_by_folder(names: list[str], filter_text: str = "") -> list:
    start = SCRIPT.index("function groupByFolder")
    code = SCRIPT[start : SCRIPT.index("\n}\n", start) + 3] + (
        f"\nprocess.stdout.write(JSON.stringify(groupByFolder({json.dumps([{'name': n} for n in names])}, "
        f"{json.dumps(filter_text)}).map((g) => [g.folder, g.files.map((c) => c.name)])));"
    )
    return json.loads(run_js(code))


def test_group_by_folder_root_first():
    names = ["總表", "醫材/儀器/保養紀錄", "醫材/耗材清單", "A/b"]
    assert group_by_folder(names) == [
        ["", ["總表"]],
        ["A", ["A/b"]],
        ["醫材", ["醫材/耗材清單"]],
        ["醫材/儀器", ["醫材/儀器/保養紀錄"]],
    ]


def test_group_filter_matches_folder_name():
    names = ["總表", "醫材/耗材清單", "醫材/儀器/保養紀錄", "藥品/清單"]
    assert group_by_folder(names, "醫材") == [
        ["醫材", ["醫材/耗材清單"]],
        ["醫材/儀器", ["醫材/儀器/保養紀錄"]],
    ]
    assert group_by_folder(names, "清單") == [["藥品", ["藥品/清單"]], ["醫材", ["醫材/耗材清單"]]]  # 依字碼排序,與後端相同
