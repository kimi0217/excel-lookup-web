"""Windows 防火牆規則。由「設定防火牆」/「移除防火牆設定」以 venv 的 Python 執行,需要以系統管理員身分執行。

規則允許私人與網域網路(不含公用網路)連線到 config.json 中 port 設定的 TCP 連接埠。
與自動啟動分開:自動啟動不可以系統管理員身分執行(排程工作可能建在別的帳號底下),防火牆則一定要。

用法:python deploy/firewall.py add | remove
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from deploy import common  # noqa: E402
from deploy.common import ConfigError, print_banner, print_error, run_command, script  # noqa: E402

RULE_NAME = "ExcelLookupWeb"  # 規則名稱(英文,避免編碼問題;「移除防火牆設定.bat」的備援流程也用這個名稱)
DEFAULT_PORT = 8000


def rule_command(action: str, port: int | None = None) -> list[str]:
    """netsh 防火牆規則指令:action 為 show / delete / add(add 需要 port)。"""
    cmd = ["netsh", "advfirewall", "firewall", action, "rule", f"name={RULE_NAME}"]
    if action == "add":
        cmd += ["dir=in", "action=allow", "protocol=TCP", f"localport={port}", "profile=private,domain"]
    return cmd


def explain_admin_needed(base: str) -> None:
    print_error("需要系統管理員權限才能設定防火牆。", [
        "關閉這個視窗",
        f"在{script(base)}上按滑鼠右鍵,選「以系統管理員身分執行」",
        "跳出「使用者帳戶控制」視窗時按「是」;需要輸入管理員密碼時,請找系統管理員協助",
    ])


def configured_port() -> int | None:
    """config.json 的 port;尚未設定時用預設值,config.json 有誤時顯示說明並回傳 None。"""
    if not common.CONFIG_PATH.is_file():
        print(f"尚未完成第一次設定(找不到 config.json),使用預設連接埠 {DEFAULT_PORT}。")
        return DEFAULT_PORT
    try:
        return int(common.read_config().get("port", DEFAULT_PORT))
    except ConfigError as exc:
        print_error(str(exc), [f"用{common.EDITOR}開啟 config.json 修正後,再執行一次"])
    except (TypeError, ValueError):
        print_error("config.json 中的 \"port\" 不是數字。", [f"用{common.EDITOR}開啟 config.json,把 \"port\" 改成數字(例如 8000)"])
    return None


def add_rule() -> int:
    """新增規則;已有同名規則時先刪除再新增(連接埠可能改過),所以重複執行會取代舊規則。"""
    if not common.is_windows_admin():
        explain_admin_needed("設定防火牆")
        return 1
    port = configured_port()
    if port is None:
        return 1
    run_command(rule_command("delete"))  # 沒有舊規則時會失敗,不影響
    code, output = run_command(rule_command("add", port))
    if code != 0:
        print_error(f"無法新增防火牆規則(連接埠 {port})。", [
            "確認是以系統管理員身分執行,再執行一次",
            "公司電腦的防火牆若由系統管理員統一管理,請將此視窗截圖交給系統管理員",
        ], details=output)
        return 1
    print(f"已新增防火牆規則「{RULE_NAME}」:允許私人與網域網路連線到 TCP 連接埠 {port}(不開放公用網路)。")
    print(f"之後若修改了 config.json 的 \"port\",請再執行一次{script('設定防火牆')}。")
    return 0


def remove_rule() -> int:
    code, _ = run_command(rule_command("show"))  # 查詢不需要系統管理員權限
    if code != 0:
        print("目前沒有設定防火牆規則,不需要移除。")
        return 0
    if not common.is_windows_admin():
        explain_admin_needed("移除防火牆設定")
        return 1
    code, output = run_command(rule_command("delete"))
    if code != 0:
        print_error(f"無法移除防火牆規則「{RULE_NAME}」。", [
            f"確認是以系統管理員身分執行{script('移除防火牆設定')},再執行一次",
        ], details=output)
        return 1
    print(f"已移除防火牆規則「{RULE_NAME}」。其他電腦將無法連線到這台電腦的網頁。")
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    action = argv[0] if argv else ""
    if action not in ("add", "remove"):
        print("用法:python deploy/firewall.py add | remove")
        return 2
    title = "設定防火牆" if action == "add" else "移除防火牆設定"
    print_banner([f"Excel 查詢網頁 - {title}"])
    if not common.IS_WINDOWS:
        print("這個功能只適用於 Windows。macOS 第一次啟動時若詢問是否允許傳入連線,請按「允許」。")
        return 0
    return add_rule() if action == "add" else remove_rule()


if __name__ == "__main__":
    sys.exit(main())
