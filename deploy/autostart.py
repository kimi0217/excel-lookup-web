"""設定/取消自動啟動。由「設定開機自動啟動」/「取消開機自動啟動」以 venv 的 Python 執行。

- Windows:工作排程器,以目前登入的使用者帳號在登入後於背景執行(讀得到該帳號可存取的網路共用資料夾,
  不需要輸入密碼)。以 XML 建立工作,才能取消預設的「執行超過 3 天自動停止」與「使用電池時不啟動」。
  不可以系統管理員身分執行:UAC 要求輸入另一個管理員帳號的密碼時,工作會建在那個帳號底下。
  防火牆規則由 firewall.py(「設定防火牆.bat」)另外設定。
- macOS:使用者的 LaunchAgent,登入後於背景執行。

用法:python deploy/autostart.py install | remove
"""

from __future__ import annotations

import os
import plistlib
import sys
import tempfile
import time
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from deploy import common  # noqa: E402
from deploy.common import ConfigError, print_banner, print_error, run_command, script  # noqa: E402

TASK_NAME = "ExcelLookupWeb"  # 工作排程器中的名稱(英文,避免編碼問題)
MAC_LABEL = "com.excel-lookup-web"
MAC_PLIST = Path.home() / "Library" / "LaunchAgents" / f"{MAC_LABEL}.plist"
LAUNCHER = common.ROOT / "deploy" / "launcher.py"
START_TIMEOUT = 40  # 啟動後最多等待幾秒確認服務可用


def build_task_xml(user: str, command: str, arguments: str, working_dir: str) -> str:
    """工作排程器的工作定義:登入後延遲 30 秒(等網路磁碟連線)於背景執行,不限執行時間。"""
    user, command, arguments, working_dir = (escape(v) for v in (user, command, arguments, working_dir))
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Excel 查詢網頁:登入 Windows 後自動在背景啟動</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{user}</UserId>
      <Delay>PT30S</Delay>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{user}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{command}</Command>
      <Arguments>{arguments}</Arguments>
      <WorkingDirectory>{working_dir}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def build_launch_agent() -> dict:
    """macOS LaunchAgent:登入後背景執行;非正常結束時自動重新啟動。"""
    return {
        "Label": MAC_LABEL,
        "ProgramArguments": [str(common.venv_python()), str(LAUNCHER), "--background"],
        "WorkingDirectory": str(common.ROOT),
        "RunAtLoad": True,
        "KeepAlive": {"SuccessfulExit": False},
        "StandardOutPath": os.devnull,
        "StandardErrorPath": str(common.ROOT / "app.log"),
    }


def current_windows_user() -> str:
    domain, user = os.environ.get("USERDOMAIN", ""), os.environ.get("USERNAME", "")
    return f"{domain}\\{user}" if domain else user


def mac_protected_paths(paths: list[Path]) -> list[Path]:
    """macOS 背景程式預設無權讀取的位置(桌面、文件、下載項目、iCloud)。"""
    home = Path.home()
    protected = [home / "Desktop", home / "Documents", home / "Downloads", home / "Library" / "Mobile Documents"]
    result = []
    for path in paths:
        resolved = path.expanduser().resolve()
        if any(resolved == p or p in resolved.parents for p in protected):
            result.append(path)
    return result


def wait_for_service(port: int) -> bool:
    deadline = time.monotonic() + START_TIMEOUT
    while time.monotonic() < deadline:
        if common.is_our_service(port):
            return True
        time.sleep(1)
    return False


def wait_until_stopped(port: int, timeout: float = 15) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not common.is_our_service(port, timeout=1):
            return True
        time.sleep(1)
    return False


def preflight() -> dict | None:
    """設定自動啟動前的檢查;有問題時顯示說明並回傳 None。"""
    if not common.venv_python(windowed=True).is_file():
        print_error("尚未安裝。", [f"先點兩下{script('安裝')}", f"再點兩下{script('啟動')}完成第一次設定", "最後再設定自動啟動"])
        return None
    if not common.CONFIG_PATH.is_file():
        print_error("尚未完成第一次設定(還沒選擇 Excel 資料夾)。", [
            f"先點兩下{script('啟動')},選擇 Excel 資料夾,確認網頁可以開啟",
            f"關閉{script('啟動')}的視窗後,再點兩下{script('設定開機自動啟動')}",
        ])
        return None
    try:
        return common.read_config()
    except ConfigError as exc:
        print_error(str(exc), [f"用{common.EDITOR}開啟 config.json 修正,或刪除後重新點兩下{script('啟動')}"])
        return None


def start_now_or_explain(port: int, start) -> None:
    """設定完成後:若服務沒在執行就立即在背景啟動,並確認可以連線。"""
    if common.port_in_use(port):
        if common.is_our_service(port):
            print(f"目前服務正由{script('啟動')}的視窗執行中。")
            print("關閉那個視窗後,下次登入時就會自動在背景啟動。")
        else:
            print_error(f"連接埠 {port} 已被其他程式使用,背景服務將無法啟動。", [
                f"用{common.EDITOR}開啟 config.json,把 \"port\" 改成其他數字(例如 8080)",
            ])
        return
    print("正在背景啟動服務……")
    start()
    if wait_for_service(port):
        print_banner(["服務已在背景啟動。", *common.url_lines(port), "", "不會出現任何視窗,關機或登出後,下次登入會自動啟動。"])
    else:
        print_error("背景服務沒有成功啟動。", [
            "開啟程式資料夾中的 app.log 查看原因",
            f"或點兩下{script('啟動')},直接在視窗中查看錯誤訊息",
        ])


# ---------- Windows ----------

def install_windows(config: dict) -> int:
    if common.is_windows_admin():
        # 以系統管理員身分執行時,USERNAME 可能是 UAC 輸入密碼的那個管理員帳號,工作會建錯帳號
        print_error("請不要以系統管理員身分執行這個檔案,直接點兩下即可。", [
            "關閉這個視窗",
            f"直接點兩下{script('設定開機自動啟動')}(不要按右鍵選「以系統管理員身分執行」)",
            f"防火牆需要系統管理員權限,請另外在{script('設定防火牆')}上按滑鼠右鍵,選「以系統管理員身分執行」",
        ])
        print("未建立自動啟動的排程工作。")
        return 1
    user = current_windows_user()
    xml = build_task_xml(
        user=user,
        command=str(common.venv_python(windowed=True)),
        arguments=f'"{LAUNCHER}" --background',
        working_dir=str(common.ROOT),
    )
    with tempfile.TemporaryDirectory() as tmp:
        xml_path = Path(tmp) / "task.xml"
        xml_path.write_text(xml, encoding="utf-16")  # 工作排程器要求 UTF-16
        code, output = run_command(["schtasks", "/Create", "/TN", TASK_NAME, "/XML", str(xml_path), "/F"])
    if code != 0:
        denied = any(k in output for k in ("存取被拒", "Access is denied", "拒絕存取"))
        fixes = (
            [  # 不可改用系統管理員身分執行:工作可能建在別的帳號底下
                "這台電腦的原則可能不允許一般使用者建立排程工作",
                "請將此視窗截圖交給系統管理員,由系統管理員開放權限後,再直接點兩下執行",
            ]
            if denied
            else [f"重新點兩下{script('設定開機自動啟動')}", "仍然失敗時,請將此視窗截圖交給系統管理員"]
        )
        print_error("無法建立自動啟動的排程工作。", fixes, details=output)
        return 1
    print(f"已設定:使用者「{user}」登入 Windows 後,約 30 秒會自動在背景啟動服務。")
    print("(需要有人登入這台電腦才會啟動;建議這台電腦保持登入狀態,不要登出。)")
    print(f"其他電腦若連不到網頁,請在{script('設定防火牆')}上按滑鼠右鍵,選「以系統管理員身分執行」。")
    print()
    port = int(config.get("port", 8000))
    start_now_or_explain(port, lambda: run_command(["schtasks", "/Run", "/TN", TASK_NAME]))
    return 0


def report_stopped(port: int) -> None:
    if wait_until_stopped(port):
        print("已取消自動啟動,背景執行中的服務也已停止。")
    else:
        print("已取消自動啟動。")
        print(f"連接埠 {port} 上仍有服務在執行,可能是{script('啟動')}的視窗還開著;關閉該視窗即可停止。")
    print(f"之後需要使用時,點兩下{script('啟動')}即可。")


def configured_port() -> int:
    try:
        return int(common.read_config().get("port", 8000))
    except (ConfigError, ValueError):
        return 8000


def remove_windows() -> int:
    result = 0
    code, _ = run_command(["schtasks", "/Query", "/TN", TASK_NAME])
    if code != 0:
        print("目前沒有設定自動啟動,不需要取消。")
    else:
        run_command(["schtasks", "/End", "/TN", TASK_NAME])  # 停止目前在背景執行的服務
        code, output = run_command(["schtasks", "/Delete", "/TN", TASK_NAME, "/F"])
        if code != 0:
            print_error("無法取消自動啟動。", [
                f"在{script('取消開機自動啟動')}上按滑鼠右鍵,選「以系統管理員身分執行」",
            ], details=output)
            result = 1
        else:
            report_stopped(configured_port())
    return result


# ---------- macOS ----------

def install_mac(config: dict) -> int:
    blocked = mac_protected_paths([common.ROOT, Path(config.get("data_dir", ""))])
    if blocked:
        print_error("程式或 Excel 資料夾位於「桌面」「文件」「下載項目」或 iCloud,背景服務預設無權讀取。", [
            "把程式資料夾與 Excel 資料夾搬到其他位置(例如「/Users/Shared」),搬移後重新安裝與設定",
            "或到「系統設定 → 隱私權與安全性 → 完整磁碟取用權限」加入 Python",
        ], details="\n".join(str(p) for p in blocked))
        if not common.ask_yes("仍要繼續設定嗎?"):
            return 1
    MAC_PLIST.parent.mkdir(parents=True, exist_ok=True)
    with MAC_PLIST.open("wb") as f:
        plistlib.dump(build_launch_agent(), f)
    domain = f"gui/{os.getuid()}"
    run_command(["launchctl", "bootout", f"{domain}/{MAC_LABEL}"])  # 先移除舊的設定(沒有也無妨)
    port = int(config.get("port", 8000))
    if common.port_in_use(port) and common.is_our_service(port):
        print("已設定:登入後會自動在背景啟動服務。")
        start_now_or_explain(port, lambda: None)
        return 0
    code, output = run_command(["launchctl", "bootstrap", domain, str(MAC_PLIST)])
    if code != 0:
        print_error("無法設定自動啟動。", ["重新執行一次", "仍然失敗時,請將此視窗截圖交給系統管理員"], details=output)
        return 1
    print("已設定:登入後會自動在背景啟動服務。")
    start_now_or_explain(port, lambda: None)  # bootstrap 時已依 RunAtLoad 啟動
    return 0


def remove_mac() -> int:
    if not MAC_PLIST.exists():
        print("目前沒有設定自動啟動,不需要取消。")
        return 0
    run_command(["launchctl", "bootout", f"gui/{os.getuid()}/{MAC_LABEL}"])
    MAC_PLIST.unlink(missing_ok=True)
    report_stopped(configured_port())
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    action = argv[0] if argv else ""
    if action == "install":
        print_banner(["Excel 查詢網頁 - 設定開機自動啟動"])
        config = preflight()
        if config is None:
            return 1
        return install_windows(config) if common.IS_WINDOWS else install_mac(config)
    if action == "remove":
        print_banner(["Excel 查詢網頁 - 取消開機自動啟動"])
        return remove_windows() if common.IS_WINDOWS else remove_mac()
    print("用法:python deploy/autostart.py install | remove")
    return 2


if __name__ == "__main__":
    sys.exit(main())
