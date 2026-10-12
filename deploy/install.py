"""安裝:建立 venv、安裝 requirements.txt、檢查結果。由「安裝.bat」/「安裝.command」以系統的 Python 執行。

重複執行不會出錯:venv 可用就沿用,已安裝的套件會略過。
"""

from __future__ import annotations

import os
import re
import shutil
import struct
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from deploy.common import MIN_PYTHON, ROOT, print_banner, print_error, script, venv_python  # noqa: E402

VENV_DIR = ROOT / "venv"
LOG_PATH = ROOT / "install.log"
REQUIRED_MODULES = "fastapi, uvicorn, pandas, openpyxl"

NETWORK_HINTS = (
    "ConnectionError", "Max retries exceeded", "NewConnectionError", "Temporary failure in name resolution",
    "getaddrinfo failed", "Could not find a version that satisfies", "No matching distribution found",
    "ReadTimeoutError", "ProxyError", "Network is unreachable",
)
SSL_HINTS = ("CERTIFICATE_VERIFY_FAILED", "SSLError", "SSL: ")
PERMISSION_HINTS = ("Permission denied", "Access is denied", "WinError 5", "WinError 32", "存取被拒")
DISK_HINTS = ("No space left", "WinError 112", "磁碟空間不足")


def classify_pip_error(output: str) -> tuple[str, list[str]]:
    """依 pip 的輸出判斷失敗原因,回傳 (問題, 處理方式)。"""
    if any(h in output for h in SSL_HINTS):
        return "公司網路的安全檢查擋住了套件下載(憑證驗證失敗)。", [
            "請 IT 人員協助允許連線到 pypi.org 與 files.pythonhosted.org",
            f"或改用可以直接上網的網路後,重新點兩下{script('安裝')}",
        ]
    if any(h in output for h in NETWORK_HINTS):
        return "無法從網路下載必要的套件。", [
            "確認這台電腦可以上網(試著用瀏覽器開啟 https://pypi.org)",
            "公司若使用 Proxy 或有網路管制,請 IT 人員允許連線到 pypi.org 與 files.pythonhosted.org",
            f"網路恢復後,重新點兩下{script('安裝')}",
        ]
    if any(h in output for h in PERMISSION_HINTS):
        return "檔案被占用或沒有寫入權限。", [
            f"若服務正在執行,請先關閉{script('啟動')}的視窗,並點兩下{script('取消開機自動啟動')}",
            "暫時關閉防毒軟體,或將此資料夾加入防毒軟體的例外",
            "不要把程式資料夾放在 C:\\Program Files 等系統資料夾,改放到例如 C:\\excel-lookup-web",
            f"完成後重新點兩下{script('安裝')}",
        ]
    if any(h in output for h in DISK_HINTS):
        return "磁碟空間不足。", ["清出至少 1 GB 的空間", f"重新點兩下{script('安裝')}"]
    return "安裝套件時發生錯誤。", [
        f"重新點兩下{script('安裝')}再試一次",
        f"仍然失敗時,請將此視窗截圖,連同 {LOG_PATH.name} 檔案一起交給系統管理員",
    ]


def check_python() -> bool:
    if sys.version_info < MIN_PYTHON:
        print_error(f"Python 版本太舊(目前 {sys.version.split()[0]},需要 3.11 以上)。", [
            "到 https://www.python.org/downloads/ 下載最新版並安裝",
            f"重新點兩下{script('安裝')}",
        ])
        return False
    if struct.calcsize("P") * 8 != 64:
        print_error("目前安裝的是 32 位元的 Python,本程式需要 64 位元版本。", [
            "到「設定 → 應用程式」解除安裝目前的 Python",
            "到 https://www.python.org/downloads/ 下載「Windows installer (64-bit)」並安裝",
            f"重新點兩下{script('安裝')}",
        ])
        return False
    return True


def venv_usable() -> bool:
    python = venv_python()
    if not python.is_file():
        return False
    check = f"import sys; sys.exit(0 if sys.version_info >= {MIN_PYTHON} else 1)"
    try:
        return subprocess.run([str(python), "-c", check], capture_output=True).returncode == 0
    except OSError:
        return False


def ensure_venv() -> bool:
    print("[1/3] 準備執行環境(venv 資料夾)……")
    if venv_usable():
        print("      執行環境已存在,沿用。")
        return True
    if VENV_DIR.exists():
        print("      舊的執行環境無法使用,重新建立。")
        try:
            shutil.rmtree(VENV_DIR)
        except OSError as exc:
            print_error("無法刪除舊的執行環境(venv 資料夾),檔案可能正在使用中。", [
                f"關閉{script('啟動')}的視窗,並點兩下{script('取消開機自動啟動')}",
                "仍無法刪除時,重新開機後再試",
                f"重新點兩下{script('安裝')}",
            ], details=str(exc))
            return False
    result = subprocess.run([sys.executable, "-m", "venv", str(VENV_DIR)], capture_output=True, text=True,
                            encoding="utf-8", errors="replace")
    if result.returncode != 0 or not venv_python().is_file():
        print_error("建立執行環境失敗。", [
            "確認此資料夾可以寫入(不要放在唯讀的位置或光碟)",
            "重新安裝 Python,安裝時使用預設選項",
            f"重新點兩下{script('安裝')}",
        ], details=(result.stdout + result.stderr).strip()[-2000:])
        return False
    print("      建立完成。")
    return True


def install_packages() -> bool:
    print("[2/3] 安裝必要套件(第一次約需 2~10 分鐘,視網路速度而定;請保持網路連線,不要關閉視窗)……")
    cmd = [str(venv_python()), "-m", "pip", "install", "--disable-pip-version-check",
           "-r", str(ROOT / "requirements.txt")]
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PIP_NO_INPUT="1")
    lines: list[str] = []
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                encoding="utf-8", errors="replace", env=env)
    except OSError as exc:
        print_error("無法執行套件安裝程式。", [f"重新點兩下{script('安裝')}"], details=str(exc))
        return False
    assert proc.stdout is not None
    for line in proc.stdout:
        lines.append(line)
        text = line.strip()
        if text.startswith("Collecting "):
            print(f"      下載:{re.split(r'[<>=!~;\[ ]', text.split()[1])[0]}")
        elif text.startswith("Installing collected packages"):
            print("      安裝中……")
    returncode = proc.wait()
    output = "".join(lines)
    LOG_PATH.write_text(output, encoding="utf-8")
    if returncode != 0:
        problem, fixes = classify_pip_error(output)
        print_error(problem, fixes, details="".join(lines[-15:]).rstrip())
        return False
    print("      套件安裝完成。")
    return True


def verify() -> bool:
    print("[3/3] 檢查安裝結果……")
    result = subprocess.run([str(venv_python()), "-c", f"import {REQUIRED_MODULES}"], capture_output=True,
                            text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        print_error("套件安裝不完整。", [
            f"重新點兩下{script('安裝')}",
            "仍然失敗時,請將此視窗截圖交給系統管理員",
        ], details=result.stderr.strip()[-2000:])
        return False
    print("      檢查通過。")
    return True


def main() -> int:
    print_banner(["Excel 查詢網頁 - 安裝"])
    print(f"使用的 Python:{sys.version.split()[0]}({sys.executable})")
    print()
    if not (check_python() and ensure_venv() and install_packages() and verify()):
        return 1
    print_banner([
        "安裝完成!",
        f"下一步:點兩下{script('啟動')}開始使用。",
    ])
    return 0


if __name__ == "__main__":
    sys.exit(main())
