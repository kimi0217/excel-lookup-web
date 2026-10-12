#!/bin/bash
# Excel 查詢網頁 - 安裝(macOS)
cd "$(dirname "$0")" || exit 1
printf '\033]0;Excel 查詢網頁 - 安裝\007'

pause_and_exit() {
  echo
  read -n 1 -s -r -p "按任意鍵關閉視窗。"
  echo
  exit "$1"
}

# 尋找 Python 3.11 以上;略過 /usr/bin/python3(系統內建版本太舊,未安裝開發者工具時還會跳出安裝視窗)
PY=""
OLD_PY=""
for cand in python3.14 python3.13 python3.12 python3.11 python3 \
            /Library/Frameworks/Python.framework/Versions/Current/bin/python3; do
  path=$(command -v "$cand" 2>/dev/null) || continue
  [ "$path" = "/usr/bin/python3" ] && continue
  if "$path" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 10)' 2>/dev/null; then
    PY="$path"
    break
  elif [ $? -eq 10 ] && [ -z "$OLD_PY" ]; then
    OLD_PY="$path"
  fi
done

if [ -z "$PY" ]; then
  echo
  echo "========================================================"
  echo "  [錯誤] 找不到 Python 3.11 以上的版本"
  echo "========================================================"
  echo
  if [ -n "$OLD_PY" ]; then
    echo "這台電腦安裝的 Python 版本太舊:$("$OLD_PY" --version 2>&1)"
    echo
  fi
  echo "請依照下列步驟安裝 Python:"
  echo "  1. 開啟 https://www.python.org/downloads/(稍後會自動開啟)"
  echo "  2. 按黃色的「Download Python」按鈕下載安裝程式"
  echo "  3. 執行下載的 .pkg 檔,依照畫面指示完成安裝"
  echo "  4. 安裝完成後,重新點兩下「安裝.command」"
  open "https://www.python.org/downloads/"
  pause_and_exit 1
fi

if ! "$PY" deploy/install.py; then
  echo "安裝沒有完成,請依照上方的說明處理後,重新點兩下「安裝.command」。"
  pause_and_exit 1
fi
pause_and_exit 0
