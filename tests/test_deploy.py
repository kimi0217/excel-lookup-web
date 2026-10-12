"""部署腳本:第一次設定、啟動檢查、安裝錯誤判斷、自動啟動設定、批次檔格式。"""

import http.server
import json
import os
import plistlib
import re
import socket
import stat
import subprocess
import threading
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest
from openpyxl import Workbook

from deploy import autostart, common, install, launcher

ROOT = Path(__file__).resolve().parents[1]
TASK_NS = {"t": "http://schemas.microsoft.com/windows/2004/02/mit/task"}
REAL_GET_LAN_IPS = common.get_lan_ips


@pytest.fixture(autouse=True)
def fixed_lan_ip(monkeypatch):
    """測試不依賴實際網路(查詢主機名稱在部分網路環境可能卡住數十秒)。"""
    monkeypatch.setattr(common, "get_lan_ips", lambda: ["192.168.1.50"])


@pytest.fixture
def config_path(tmp_path, monkeypatch) -> Path:
    path = tmp_path / "config.json"
    monkeypatch.setattr(common, "CONFIG_PATH", path)
    return path


@pytest.fixture
def excel_dir(tmp_path) -> Path:
    d = tmp_path / "Excel 資料"
    d.mkdir()
    wb = Workbook()
    wb.active.append(["姓名"])
    wb.active.append(["王小明"])
    wb.save(d / "員工.xlsx")
    return d


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


# ---------- 網路 ----------

def test_lan_ips_filtered(monkeypatch):
    class FakeSocket:
        def __init__(self, *a):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *a):
            pass

        def connect(self, addr):
            raise OSError("沒有預設路由")

    monkeypatch.setattr(common.socket, "socket", FakeSocket)
    monkeypatch.setattr(common.socket, "getaddrinfo", lambda *a: [
        (None, None, None, None, (ip, 0)) for ip in ["127.0.0.1", "169.254.1.2", "192.168.1.50", "192.168.1.50", "10.0.0.8"]
    ])
    assert REAL_GET_LAN_IPS() == ["192.168.1.50", "10.0.0.8"]


def test_url_lines_show_lan_url_first(monkeypatch):
    monkeypatch.setattr(common, "get_lan_ips", lambda: ["192.168.1.50", "10.0.0.8"])
    lines = common.url_lines(8000)
    assert lines[0] == "請同事用瀏覽器開啟:  http://192.168.1.50:8000"
    assert "http://10.0.0.8:8000" in lines[1]
    assert lines[-1].endswith("http://localhost:8000")


def test_url_lines_without_network(monkeypatch):
    monkeypatch.setattr(common, "get_lan_ips", lambda: [])
    assert "找不到這台電腦的內網 IP" in common.url_lines(8000)[0]


def test_port_in_use():
    port = free_port()
    assert not common.port_in_use(port, "127.0.0.1")
    with socket.socket() as s:
        s.bind(("127.0.0.1", port))
        s.listen()
        assert common.port_in_use(port, "127.0.0.1")


@pytest.mark.parametrize(
    ("body", "expected"),
    [(b'{"folder": {"ok": true}, "files": []}', True), (b"[]", False), (b'{"other": 1}', False), (b"<html>", False)],
)
def test_is_our_service(body, expected):
    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *a):
            pass

    server = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        assert common.is_our_service(server.server_address[1]) is expected
    finally:
        server.shutdown()


def test_is_our_service_nothing_listening():
    assert common.is_our_service(free_port(), timeout=0.5) is False


# ---------- 設定檔 ----------

def test_write_config_uses_defaults_and_escapes_windows_path(config_path):
    common.write_config("D:\\共用資料\\Excel")
    text = config_path.read_text(encoding="utf-8")
    assert '"D:\\\\共用資料\\\\Excel"' in text  # JSON 中反斜線自動寫成兩個
    config = json.loads(text)
    assert config["data_dir"] == "D:\\共用資料\\Excel"
    assert config["port"] == 8000 and config["include_hidden_sheets"] is False


def test_write_config_keeps_existing_settings(config_path):
    common.write_config("/新資料夾", existing={"port": 8080, "data_dir": "/舊資料夾"})
    config = json.loads(config_path.read_text(encoding="utf-8"))
    assert config["port"] == 8080 and config["data_dir"] == "/新資料夾"


def test_read_config_explains_bad_backslash(config_path):
    config_path.write_text('{\n  "data_dir": "D:\\共用"\n}', encoding="utf-8")
    with pytest.raises(common.ConfigError, match="第 2 行.*反斜線"):
        common.read_config()


@pytest.mark.skipif(common.IS_WINDOWS, reason="Windows 上會實際查詢網路磁碟機")
def test_to_unc_unchanged_off_windows():
    assert common.to_unc("Z:\\資料") == "Z:\\資料"


def test_choose_data_dir_cancel(monkeypatch):
    monkeypatch.setattr(common, "pick_folder", lambda prompt: None)
    assert common.choose_data_dir() is None


def test_choose_data_dir_with_excel(monkeypatch, excel_dir, capsys):
    monkeypatch.setattr(common, "pick_folder", lambda prompt: str(excel_dir))
    assert common.choose_data_dir() == str(excel_dir)
    assert "找到 1 個 Excel 檔案" in capsys.readouterr().out


def test_choose_data_dir_empty_folder_asks_again(monkeypatch, tmp_path, excel_dir):
    empty = tmp_path / "空資料夾"
    empty.mkdir()
    picks = iter([str(empty), str(excel_dir)])
    monkeypatch.setattr(common, "pick_folder", lambda prompt: next(picks))
    monkeypatch.setattr(common, "ask_yes", lambda q: False)  # 不使用空資料夾
    assert common.choose_data_dir() == str(excel_dir)


def test_ask_yes_without_console(monkeypatch):
    def no_input(prompt):
        raise EOFError

    monkeypatch.setattr("builtins.input", no_input)
    assert common.ask_yes("?") is False


# ---------- 啟動 ----------

@pytest.fixture
def no_server(monkeypatch):
    """攔截 uvicorn.run,記錄參數而不真的啟動。"""
    import uvicorn

    calls = []
    monkeypatch.setattr(uvicorn, "run", lambda app, **kw: calls.append(kw))
    return calls


def test_first_run_creates_config_and_starts(config_path, excel_dir, monkeypatch, no_server, capsys):
    monkeypatch.setattr(common, "pick_folder", lambda prompt: str(excel_dir))
    monkeypatch.setattr(common, "port_in_use", lambda *a: False)
    assert launcher.run(background=False) == 0
    assert json.loads(config_path.read_text(encoding="utf-8"))["data_dir"] == str(excel_dir)
    out = capsys.readouterr().out
    assert "http://192.168.1.50:8000" in out and "Excel 查詢網頁已啟動" in out
    assert no_server == [{"host": "0.0.0.0", "port": 8000, "log_level": "warning", "access_log": False, "log_config": None}]


def test_first_run_cancel(config_path, monkeypatch, no_server, capsys):
    monkeypatch.setattr(common, "pick_folder", lambda prompt: None)
    assert launcher.run(background=False) == 0
    assert not config_path.exists() and no_server == []
    assert "已取消設定" in capsys.readouterr().out


def test_background_without_config_does_not_prompt(config_path, monkeypatch, no_server, caplog):
    monkeypatch.setattr(common, "pick_folder", lambda prompt: pytest.fail("背景模式不可跳出視窗"))
    assert launcher.run(background=True) == 0
    assert no_server == [] and "尚未完成第一次設定" in caplog.text


def test_bad_config_explained(config_path, no_server, capsys):
    config_path.write_text('{"data_dir": "D:\\共用"}', encoding="utf-8")
    assert launcher.run(background=False) == 1
    out = capsys.readouterr().out
    assert "格式錯誤" in out and "刪除 config.json" in out and no_server == []


def test_already_running(config_path, excel_dir, monkeypatch, no_server, capsys):
    common.write_config(str(excel_dir))
    monkeypatch.setattr(common, "port_in_use", lambda *a: True)
    monkeypatch.setattr(common, "is_our_service", lambda *a, **k: True)
    assert launcher.run(background=False) == 0
    assert "服務已經在執行中" in capsys.readouterr().out and no_server == []


def test_port_used_by_other_program(config_path, excel_dir, monkeypatch, no_server, capsys):
    common.write_config(str(excel_dir))
    monkeypatch.setattr(common, "port_in_use", lambda *a: True)
    monkeypatch.setattr(common, "is_our_service", lambda *a, **k: False)
    assert launcher.run(background=False) == 1
    out = capsys.readouterr().out
    assert "已被其他程式使用" in out and '"port"' in out and no_server == []


def test_missing_data_dir_can_reselect(config_path, excel_dir, tmp_path, monkeypatch, no_server):
    common.write_config(str(tmp_path / "已刪除的資料夾"), existing={"port": 8123})
    monkeypatch.setattr(common, "ask_yes", lambda q: True)
    monkeypatch.setattr(common, "pick_folder", lambda prompt: str(excel_dir))
    monkeypatch.setattr(common, "port_in_use", lambda *a: False)
    assert launcher.run(background=False) == 0
    config = json.loads(config_path.read_text(encoding="utf-8"))
    assert config["data_dir"] == str(excel_dir) and config["port"] == 8123
    assert no_server[0]["port"] == 8123


def test_missing_data_dir_background_still_starts(config_path, tmp_path, monkeypatch, no_server, caplog):
    common.write_config(str(tmp_path / "網路資料夾尚未連線"))
    monkeypatch.setattr(common, "port_in_use", lambda *a: False)
    assert launcher.run(background=True) == 0
    assert len(no_server) == 1 and "找不到 Excel 資料夾" in caplog.text


def test_unexpected_error_shown_in_chinese(monkeypatch, capsys):
    monkeypatch.setattr(launcher, "run", lambda background: 1 / 0)
    assert launcher.main([]) == 1
    assert "發生未預期的錯誤" in capsys.readouterr().out


# ---------- 安裝 ----------

@pytest.mark.parametrize(
    ("output", "keyword"),
    [
        ("ERROR: Could not find a version that satisfies the requirement fastapi==0.142.4", "無法從網路下載"),
        ("Retrying ... NewConnectionError: Failed to establish a new connection", "無法從網路下載"),
        ("SSLError(SSLCertVerificationError(1, '[SSL: CERTIFICATE_VERIFY_FAILED]", "憑證驗證失敗"),
        ("ERROR: Could not install packages due to an OSError: [WinError 5] Access is denied", "檔案被占用"),
        ("OSError: [Errno 28] No space left on device", "磁碟空間不足"),
        ("something unexpected", "安裝套件時發生錯誤"),
    ],
)
def test_classify_pip_error(output, keyword):
    problem, fixes = install.classify_pip_error(output)
    assert keyword in problem and fixes


def test_install_rerun_reuses_venv(monkeypatch, capsys):
    monkeypatch.setattr(install, "venv_usable", lambda: True)
    assert install.ensure_venv() is True
    assert "沿用" in capsys.readouterr().out


# ---------- 自動啟動 ----------

def test_task_xml():
    xml = autostart.build_task_xml(
        user="OFFICE\\王小明",
        command="C:\\R&D 工具\\venv\\Scripts\\pythonw.exe",
        arguments='"C:\\R&D 工具\\deploy\\launcher.py" --background',
        working_dir="C:\\R&D 工具",
    )
    root = ET.fromstring(xml.split("?>", 1)[1])  # 去掉宣告(宣告為 UTF-16,實際寫檔時才轉碼)
    find = lambda path: root.find(path, TASK_NS).text  # noqa: E731
    assert find("t:Triggers/t:LogonTrigger/t:UserId") == "OFFICE\\王小明"
    assert find("t:Principals/t:Principal/t:LogonType") == "InteractiveToken"
    assert find("t:Settings/t:ExecutionTimeLimit") == "PT0S"  # 不因執行超過 3 天被停止
    assert find("t:Settings/t:DisallowStartIfOnBatteries") == "false"
    assert find("t:Settings/t:StopIfGoingOnBatteries") == "false"
    assert find("t:Actions/t:Exec/t:Command") == "C:\\R&D 工具\\venv\\Scripts\\pythonw.exe"
    assert find("t:Actions/t:Exec/t:Arguments") == '"C:\\R&D 工具\\deploy\\launcher.py" --background'


def test_task_xml_written_as_utf16(tmp_path):
    path = tmp_path / "task.xml"
    path.write_text(autostart.build_task_xml("u", "c", "a", "w"), encoding="utf-16")
    assert path.read_bytes()[:2] in (b"\xff\xfe", b"\xfe\xff")  # 有 BOM
    ET.parse(path)


def test_launch_agent_plist(tmp_path):
    path = tmp_path / "agent.plist"
    with path.open("wb") as f:
        plistlib.dump(autostart.build_launch_agent(), f)
    agent = plistlib.loads(path.read_bytes())
    assert agent["ProgramArguments"][-1] == "--background"
    assert agent["RunAtLoad"] is True and agent["KeepAlive"] == {"SuccessfulExit": False}
    if Path("/usr/bin/plutil").exists():
        assert subprocess.run(["plutil", "-lint", str(path)], capture_output=True).returncode == 0


def test_mac_protected_paths():
    home = Path.home()
    paths = [home / "Desktop" / "excel-lookup-web", home / "Documents", Path("/Users/Shared/excel")]
    assert autostart.mac_protected_paths(paths) == paths[:2]


def test_preflight_requires_setup(config_path, monkeypatch, tmp_path, capsys):
    fake_python = tmp_path / "python"
    fake_python.touch()
    monkeypatch.setattr(common, "venv_python", lambda windowed=False: fake_python)
    assert autostart.preflight() is None
    assert "尚未完成第一次設定" in capsys.readouterr().out
    common.write_config("/資料")
    assert autostart.preflight()["data_dir"] == "/資料"


def test_preflight_requires_install(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(common, "venv_python", lambda windowed=False: tmp_path / "不存在")
    assert autostart.preflight() is None
    assert "尚未安裝" in capsys.readouterr().out


def test_windows_install_reports_access_denied(monkeypatch, capsys):
    monkeypatch.setattr(autostart, "run_command", lambda cmd: (1, "錯誤: 存取被拒。"))
    assert autostart.install_windows({"port": 8000}) == 1
    assert "以系統管理員身分執行" in capsys.readouterr().out


def test_windows_remove_when_not_configured(monkeypatch, capsys):
    monkeypatch.setattr(autostart, "run_command", lambda cmd: (1, ""))
    assert autostart.remove_windows() == 0
    assert "目前沒有設定自動啟動" in capsys.readouterr().out


# ---------- Windows 防火牆 ----------

class FakeWindows:
    """記錄執行的系統指令;schtasks 與 netsh 依設定回傳結果。"""

    def __init__(self, task_exists=True, rule_exists=False, netsh_ok=True):
        self.calls: list[list[str]] = []
        self.task_exists, self.rule_exists, self.netsh_ok = task_exists, rule_exists, netsh_ok

    def __call__(self, cmd):
        self.calls.append(cmd)
        if cmd[0] == "schtasks":
            return (0 if self.task_exists or cmd[1] == "/Create" else 1), ""
        if cmd[:5] == ["netsh", "advfirewall", "firewall", "show", "rule"]:
            return (0 if self.rule_exists else 1), ""
        return (0 if self.netsh_ok else 1), "" if self.netsh_ok else "要求的操作需要提高的權限。"

    def netsh(self, action):
        return [c for c in self.calls if c[:4] == ["netsh", "advfirewall", "firewall", action]]


@pytest.fixture
def windows(monkeypatch):
    fake = FakeWindows()
    monkeypatch.setattr(autostart, "run_command", fake)
    monkeypatch.setattr(autostart, "start_now_or_explain", lambda port, start: None)
    monkeypatch.setattr(autostart, "wait_until_stopped", lambda port: True)
    monkeypatch.setattr(autostart, "configured_port", lambda: 8080)
    return fake


def test_install_adds_firewall_rule_as_admin(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: True)
    assert autostart.install_windows({"port": 8080}) == 0
    added = windows.netsh("add")
    assert added == [[
        "netsh", "advfirewall", "firewall", "add", "rule", f"name={autostart.FIREWALL_RULE}",
        "dir=in", "action=allow", "protocol=TCP", "localport=8080", "profile=private,domain",
    ]]
    # 先刪除同名的舊規則(連接埠可能改過),再新增
    assert windows.calls.index(windows.netsh("delete")[0]) < windows.calls.index(added[0])
    out = capsys.readouterr().out
    assert "已新增防火牆規則" in out and "8080" in out


def test_install_without_admin_still_sets_autostart(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: False)
    assert autostart.install_windows({"port": 8000}) == 0
    assert any(c[:2] == ["schtasks", "/Create"] for c in windows.calls)
    assert windows.netsh("add") == []
    out = capsys.readouterr().out
    assert "已設定" in out
    assert "系統管理員" in out and "以系統管理員身分執行" in out and "防火牆" in out


def test_install_firewall_failure_explained(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: True)
    windows.netsh_ok = False
    assert autostart.install_windows({"port": 8000}) == 0  # 自動啟動已完成
    out = capsys.readouterr().out
    assert "無法新增防火牆規則" in out and "要求的操作需要提高的權限" in out


def test_remove_deletes_firewall_rule(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: True)
    windows.rule_exists = True
    assert autostart.remove_windows() == 0
    assert windows.netsh("delete") == [["netsh", "advfirewall", "firewall", "delete", "rule", f"name={autostart.FIREWALL_RULE}"]]
    assert "已移除防火牆規則" in capsys.readouterr().out


def test_remove_firewall_rule_even_without_task(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: True)
    windows.task_exists, windows.rule_exists = False, True
    assert autostart.remove_windows() == 0
    assert windows.netsh("delete")
    out = capsys.readouterr().out
    assert "目前沒有設定自動啟動" in out and "已移除防火牆規則" in out


def test_remove_without_admin_explains(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: False)
    windows.rule_exists = True
    assert autostart.remove_windows() == 0
    assert windows.netsh("delete") == []
    out = capsys.readouterr().out
    assert "已取消自動啟動" in out and "以系統管理員身分執行" in out


def test_remove_without_rule_is_quiet(windows, monkeypatch, capsys):
    monkeypatch.setattr(autostart, "is_windows_admin", lambda: False)
    assert autostart.remove_windows() == 0
    assert "防火牆" not in capsys.readouterr().out


def test_is_windows_admin_off_windows():
    assert autostart.is_windows_admin() is False


def test_remove_bat_fallback_deletes_firewall_rule():
    text = (ROOT / "取消開機自動啟動.bat").read_text(encoding="utf-8")
    fallback = text[text.index(":fallback"):]
    assert f"netsh advfirewall firewall delete rule name={autostart.FIREWALL_RULE}" in fallback


# ---------- 批次檔與 .command 格式 ----------

BAT_FILES = sorted(ROOT.glob("*.bat"))
COMMAND_FILES = sorted(ROOT.glob("*.command"))


def test_expected_scripts_exist():
    names = {"安裝", "啟動", "設定開機自動啟動", "取消開機自動啟動"}
    assert {p.stem for p in BAT_FILES} == names
    assert {p.stem for p in COMMAND_FILES} == names


@pytest.mark.parametrize("path", BAT_FILES, ids=lambda p: p.name)
def test_bat_format(path):
    data = path.read_bytes()
    assert not data.startswith(b"\xef\xbb\xbf"), "不可有 BOM,否則第一行指令會失效"
    assert b"\n" not in data.replace(b"\r\n", b""), "必須全部是 CRLF 換行"
    lines = data.decode("utf-8").split("\r\n")
    assert lines[0] == "@echo off" and lines[1] == "chcp 65001 >nul"
    labels = {line[1:].split()[0].lower() for line in lines if line.startswith(":")}
    for line in lines:
        stripped = line.strip()
        for target in re.findall(r"(?:goto|call)\s+:(\w+)", stripped, re.IGNORECASE):
            assert target.lower() in labels, f"找不到標籤 :{target}"
        if stripped.lower().startswith(("rem ", "title ")) or stripped.startswith(":"):
            continue
        echo = re.search(r"(?i)\becho\b", stripped)
        command = stripped[: echo.start()] if echo else stripped
        assert command.isascii(), f"指令中不可有中文(只能出現在 echo/rem/title):{stripped}"
        # 在 ( ) 區塊中,echo 內容的半形括號會提早結束區塊;訊息一律用全形括號
        if echo and ("(" in stripped[echo.end():] or ")" in stripped[echo.end():]):
            pytest.fail(f"echo 訊息請改用全形括號:{stripped}")


@pytest.mark.parametrize("path", COMMAND_FILES, ids=lambda p: p.name)
def test_command_format(path):
    data = path.read_bytes()
    assert b"\r" not in data, "必須是 LF 換行"
    assert data.startswith(b"#!/bin/bash\n")
    assert path.stat().st_mode & stat.S_IXUSR, "需要執行權限,否則無法點兩下執行"
    assert subprocess.run(["bash", "-n", str(path)]).returncode == 0
