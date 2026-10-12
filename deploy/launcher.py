"""啟動服務。由「啟動.bat」/「啟動.command」以 venv 的 Python 執行。

- 沒有 config.json 時跳出資料夾選擇視窗,自動產生設定檔
- 檢查 Excel 資料夾、連接埠,出問題時以中文說明原因與處理方式
- 顯示這台電腦的內網網址,方便告訴同事
- --background:開機自動啟動用,不開視窗、不詢問,記錄寫入 app.log
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from deploy import common  # noqa: E402
from deploy.common import ConfigError, print_banner, print_error, script  # noqa: E402

LOG_PATH = common.ROOT / "app.log"
FIREWALL_HINT = (
    "其他電腦連不到時:在「設定防火牆.bat」上按右鍵,選「以系統管理員身分執行」。"
    if common.IS_WINDOWS
    else "第一次啟動若詢問是否允許傳入連線,請按「允許」。"
)
logger = logging.getLogger("launcher")


def setup_logging(background: bool) -> None:
    if background:
        fmt = "%(asctime)s %(levelname)s %(name)s: %(message)s"
        # pythonw 沒有主控台,stdout/stderr 是 None;導向 devnull 避免寫入時出錯
        for name in ("stdout", "stderr"):
            if getattr(sys, name) is None:
                setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
        handler: logging.Handler = logging.handlers.RotatingFileHandler(
            LOG_PATH, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
        )
        handler.setFormatter(logging.Formatter(fmt))
        level = logging.INFO
    else:
        # 視窗中只顯示需要注意的訊息,格式簡化;輸出逐行送出,避免訊息延遲出現
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(line_buffering=True)
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(logging.Formatter("[%(asctime)s 注意] %(message)s", datefmt="%H:%M:%S"))
        level = logging.WARNING
    logging.basicConfig(level=level, handlers=[handler], force=True)


def first_time_setup() -> bool:
    print_banner(["Excel 查詢網頁 - 第一次設定"])
    data_dir = common.choose_data_dir()
    if not data_dir:
        print(f"已取消設定。下次點兩下{script('啟動')}時會再詢問。")
        return False
    common.write_config(data_dir)
    print(f"設定已儲存(config.json)。之後要更換資料夾,刪除 config.json 再點兩下{script('啟動')}即可。")
    return True


def check_data_dir(config: dict, background: bool) -> dict:
    """Excel 資料夾不存在時說明原因;互動模式下可重新選擇。回傳(可能更新後的)設定。"""
    data_dir = Path(config["data_dir"])
    if data_dir.is_dir():
        if common.count_xlsx(data_dir) == 0:
            if background:
                logger.warning("Excel 資料夾中目前沒有 Excel 檔案(.xlsx、.xlsm):%s", data_dir)
            else:
                print(f"注意:Excel 資料夾中目前沒有 Excel 檔案(.xlsx、.xlsm):{data_dir}")
                print("      放入檔案後幾秒內會自動出現在網頁上。\n")
        return config
    if background:
        logger.warning("找不到 Excel 資料夾:%s(網路資料夾可能尚未連線,連上後會自動載入)", data_dir)
        return config
    print_error(f"找不到 Excel 資料夾:{data_dir}", [
        "若是網路共用資料夾:確認網路已連線,並在檔案總管中可以開啟該資料夾",
        "若資料夾已搬移或改名:選擇重新設定",
    ])
    if not common.ask_yes("要重新選擇 Excel 資料夾嗎?"):
        print("繼續啟動。資料夾可以開啟後,幾秒內會自動載入。\n")
        return config
    new_dir = common.choose_data_dir()
    if not new_dir:
        return config
    raw = common.read_config()
    common.write_config(new_dir, existing=raw)
    print("設定已更新。\n")
    return load_app_config()


def load_app_config() -> dict:
    import app  # 延後匯入:需要 venv 中的套件

    common.read_config()  # 先用中文訊息檢查 JSON 格式
    return app.load_config(common.CONFIG_PATH)


def run(background: bool) -> int:
    try:
        import app
        import uvicorn
    except ImportError as exc:
        if background:
            logger.error("執行環境不完整:%s", exc)
        else:
            print_error("執行環境不完整,缺少必要的套件。", [f"重新點兩下{script('安裝')}", f"完成後再點兩下{script('啟動')}"])
        return 1

    if not common.CONFIG_PATH.is_file():
        if background:
            logger.error("尚未完成第一次設定,請先點兩下%s選擇 Excel 資料夾", script("啟動"))
            return 0
        if not first_time_setup():
            return 0

    try:
        config = load_app_config()
    except ConfigError as exc:
        if background:
            logger.error("%s", exc)
        else:
            print_error(str(exc), [
                f"用{common.EDITOR}開啟 config.json 修正",
                f"或刪除 config.json,再點兩下{script('啟動')}重新設定",
            ])
        return 1

    config = check_data_dir(config, background)
    host, port = config.get("host", "0.0.0.0"), int(config.get("port", 8000))

    if common.port_in_use(port, host):
        if common.is_our_service(port):
            if background:
                logger.info("服務已經在執行中(連接埠 %s),不需要再次啟動", port)
            else:
                print_banner(["服務已經在執行中,不需要再次啟動。", *common.url_lines(port)])
            return 0
        if background:
            logger.error("連接埠 %s 已被其他程式使用", port)
        else:
            print_error(f"連接埠 {port} 已被其他程式使用,無法啟動。", [
                f"用{common.EDITOR}開啟 config.json,把 \"port\": {port} 改成其他數字(例如 8080),存檔後再點兩下{script('啟動')}",
                "改了連接埠之後,同事要改用新的網址;防火牆也需要允許新的連接埠"
                + ("(以系統管理員身分重新執行「設定防火牆.bat」)" if common.IS_WINDOWS else ""),
            ])
        return 1

    if background:
        logger.info("服務啟動:%s", ", ".join(common.service_urls(port)) or f"http://localhost:{port}")
    else:
        print_banner([
            "Excel 查詢網頁已啟動",
            "",
            *common.url_lines(port),
            "",
            f"Excel 資料夾:{config['data_dir']}",
            "",
            "這個視窗要保持開啟;關閉視窗就會停止服務。",
            FIREWALL_HINT,
        ])
    uvicorn.run(app.create_app(config), host=host, port=port, log_level="warning", access_log=False,
                log_config=None)
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    background = "--background" in argv
    setup_logging(background)
    try:
        return run(background)
    except KeyboardInterrupt:
        return 0
    except Exception as exc:
        logger.exception("發生未預期的錯誤")
        if not background:
            print_error("發生未預期的錯誤,服務已停止。", [
                f"重新點兩下{script('啟動')}再試一次",
                "仍然失敗時,請將此視窗截圖交給系統管理員",
            ], details=f"{type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
