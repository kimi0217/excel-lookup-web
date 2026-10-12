#!/bin/bash
# Excel 查詢網頁 - 設定開機自動啟動(macOS:登入後自動在背景啟動)
cd "$(dirname "$0")" || exit 1
printf '\033]0;Excel 查詢網頁 - 設定開機自動啟動\007'

if [ ! -x venv/bin/python ]; then
  echo
  echo "[錯誤] 尚未安裝。"
  echo
  echo "處理方式:"
  echo "  1. 先點兩下「安裝.command」"
  echo "  2. 再點兩下「啟動.command」完成第一次設定"
  echo "  3. 最後再點兩下「設定開機自動啟動.command」"
  echo
  read -n 1 -s -r -p "按任意鍵關閉視窗。"
  exit 1
fi

venv/bin/python deploy/autostart.py install
echo
read -n 1 -s -r -p "按任意鍵關閉視窗。"
echo
