#!/bin/bash
# Excel 查詢網頁 - 取消開機自動啟動(macOS)
cd "$(dirname "$0")" || exit 1
printf '\033]0;Excel 查詢網頁 - 取消開機自動啟動\007'

if [ -x venv/bin/python ]; then
  venv/bin/python deploy/autostart.py remove
else
  # 沒有 venv 時(例如已刪除),直接移除 LaunchAgent
  plist="$HOME/Library/LaunchAgents/com.excel-lookup-web.plist"
  if [ -f "$plist" ]; then
    launchctl bootout "gui/$(id -u)/com.excel-lookup-web" 2>/dev/null
    rm -f "$plist"
    echo "已取消自動啟動。"
  else
    echo "目前沒有設定自動啟動,不需要取消。"
  fi
fi
echo
read -n 1 -s -r -p "按任意鍵關閉視窗。"
echo
