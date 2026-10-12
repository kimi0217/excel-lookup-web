"""部署腳本共用:路徑、中文訊息、內網 IP、連接埠檢查、第一次設定。

安裝前就會執行(install.py 用系統的 Python 執行),所以只能使用標準函式庫。
"""

from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.json"
EXAMPLE_CONFIG_PATH = ROOT / "config.example.json"
IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
MIN_PYTHON = (3, 11)
EDITOR = "記事本" if IS_WINDOWS else "「文字編輯」"  # 提示使用者修改 config.json 時用的程式


def script(base: str) -> str:
    """給使用者看的腳本檔名,例如「啟動.bat」或「啟動.command」。"""
    return f"「{base}.bat」" if IS_WINDOWS else f"「{base}.command」"


def venv_python(windowed: bool = False) -> Path:
    """venv 中的 Python;windowed 為 True 時在 Windows 使用不開視窗的 pythonw.exe。"""
    if IS_WINDOWS:
        return ROOT / "venv" / "Scripts" / ("pythonw.exe" if windowed else "python.exe")
    return ROOT / "venv" / "bin" / "python"


# ---------- 訊息 ----------

def print_banner(lines: list[str]) -> None:
    print()
    print("=" * 56)
    for line in lines:
        print(f"  {line}")
    print("=" * 56)
    print()


def print_error(problem: str, fixes: list[str], details: str | None = None) -> None:
    """顯示錯誤原因與處理方式。"""
    print()
    print(f"[錯誤] {problem}")
    if fixes:
        print()
        print("處理方式:")
        for i, fix in enumerate(fixes, start=1):
            print(f"  {i}. {fix}")
    if details:
        print()
        print("詳細訊息(請截圖給系統管理員):")
        print(details)
    print()


def ask_yes(question: str) -> bool:
    """在視窗中詢問,輸入 Y 為是;無法輸入(背景執行)時視為否。"""
    try:
        return input(f"{question}(輸入 Y 後按 Enter 表示是,直接按 Enter 表示否):").strip().lower() == "y"
    except (EOFError, OSError):
        return False


# ---------- 網路 ----------

def get_lan_ips() -> list[str]:
    """這台電腦的內網 IPv4 位址,最可能的排第一。"""
    candidates: list[str] = []
    try:
        # UDP connect 不會真的送出封包,只用來查出對外連線所用的網路卡位址
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            candidates.append(s.getsockname()[0])
    except OSError:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            candidates.append(info[4][0])
    except OSError:
        pass
    result: list[str] = []
    for ip in candidates:
        if ip.startswith(("127.", "169.254.", "0.")) or ip in result:
            continue
        result.append(ip)
    return result


def port_in_use(port: int, host: str = "0.0.0.0") -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if not IS_WINDOWS:
            # 與 uvicorn 相同;Windows 的 SO_REUSEADDR 語意不同(允許搶用),不可設定
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind((host, port))
        except OSError:
            return True
    return False


def is_our_service(port: int, timeout: float = 2) -> bool:
    """連接埠上執行的是否就是本服務。"""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/categories", timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            return isinstance(body, dict) and isinstance(body.get("files"), list)
    except (OSError, ValueError):
        return False


def service_urls(port: int) -> list[str]:
    return [f"http://{ip}:{port}" for ip in get_lan_ips()]


def url_lines(port: int) -> list[str]:
    """顯示給使用者的網址說明。"""
    urls = service_urls(port)
    lines = []
    if urls:
        lines.append(f"請同事用瀏覽器開啟:  {urls[0]}")
        for other in urls[1:]:
            lines.append(f"(若上面的網址打不開,試試:{other})")
    else:
        lines.append("找不到這台電腦的內網 IP,請確認已連上公司網路")
    lines.append(f"在這台電腦上開啟:    http://localhost:{port}")
    return lines


# ---------- 設定檔 ----------

class ConfigError(Exception):
    """config.json 有問題;訊息為中文說明。"""


def read_config(path: Path | None = None) -> dict:
    path = path or CONFIG_PATH
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ConfigError(
            f"設定檔 config.json 第 {exc.lineno} 行格式錯誤。"
            "常見原因:Windows 路徑的反斜線 \\ 只寫了一個(要寫成 \\\\ 或改用 /)、少了逗號或引號。"
        ) from exc
    except OSError as exc:
        raise ConfigError(f"無法讀取設定檔 config.json:{exc.strerror or exc}") from exc


def write_config(data_dir: str, existing: dict | None = None, path: Path | None = None) -> dict:
    """以 config.example.json 為預設值,寫入 data_dir 並保留既有的其他設定。"""
    path = path or CONFIG_PATH
    config: dict = {}
    if EXAMPLE_CONFIG_PATH.is_file():
        config.update(json.loads(EXAMPLE_CONFIG_PATH.read_text(encoding="utf-8-sig")))
    config.update(existing or {})
    config["data_dir"] = data_dir
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return config


def count_xlsx(folder: Path) -> int:
    try:
        return sum(1 for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    except OSError:
        return 0


def to_unc(path: str) -> str:
    """Windows 網路磁碟機代號(如 Z:\\資料)轉成 UNC 路徑(\\\\server\\share\\資料)。

    開機自動啟動時磁碟機代號可能還沒連上,UNC 路徑比較可靠。不是網路磁碟機或轉換失敗時原樣回傳。
    """
    if not IS_WINDOWS:
        return path
    drive, rest = os.path.splitdrive(path)
    if len(drive) != 2 or drive[1] != ":":
        return path
    try:
        import ctypes
        from ctypes import wintypes

        buffer = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(len(buffer))
        if ctypes.windll.mpr.WNetGetConnectionW(drive, buffer, ctypes.byref(size)) != 0:
            return path  # 不是網路磁碟機
        return buffer.value.rstrip("\\") + rest
    except Exception:
        return path


def pick_folder(prompt: str) -> str | None:
    """跳出資料夾選擇視窗;取消時回傳 None。沒有視窗環境時改為在畫面上輸入路徑。"""
    if IS_MAC:
        script_text = f'POSIX path of (choose folder with prompt "{prompt}")'
        try:
            result = subprocess.run(["osascript", "-e", script_text], capture_output=True, text=True)
        except OSError:
            return _input_folder()
        if result.returncode == 0:
            return result.stdout.strip().rstrip("/") or "/"
        if "-128" in result.stderr:  # 使用者按了取消
            return None
        return _input_folder()
    try:
        import tkinter
        from tkinter import filedialog

        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)  # 避免視窗被藏在黑色視窗後面
        folder = filedialog.askdirectory(title=prompt, mustexist=True, parent=root)
        root.destroy()
    except Exception:
        return _input_folder()
    return str(Path(folder)) if folder else None


def _input_folder() -> str | None:
    print("無法開啟資料夾選擇視窗,請改為輸入路徑。")
    try:
        text = input("請輸入(或把資料夾拖曳到這裡)Excel 資料夾的完整路徑,直接按 Enter 取消:")
    except (EOFError, OSError):
        return None
    text = text.strip().strip('"').strip("'")
    return str(Path(text)) if text else None


def choose_data_dir() -> str | None:
    """請使用者選擇 Excel 資料夾並確認;取消時回傳 None。"""
    while True:
        print("請在跳出的視窗中,選擇 Excel 檔案所在的資料夾……")
        folder = pick_folder("請選擇 Excel 檔案所在的資料夾")
        if not folder:
            return None
        path = Path(folder)
        if not path.is_dir():
            print_error(f"找不到資料夾:{folder}", ["請重新選擇一次"])
            continue
        count = count_xlsx(path)
        if count == 0:
            print(f"注意:資料夾「{folder}」中沒有任何 .xlsx 檔案。")
            if not ask_yes("仍要使用這個資料夾嗎?"):
                continue
        else:
            print(f"已選擇:{folder}(找到 {count} 個 Excel 檔案)")
        return to_unc(folder)
