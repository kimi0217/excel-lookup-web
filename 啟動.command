#!/bin/bash
# Excel 查詢網頁 - 啟動(macOS)
cd "$(dirname "$0")" || exit 1
printf '\033]0;Excel 查詢網頁\007'

if [ ! -x venv/bin/python ]; then
  echo
  echo "[錯誤] 尚未安裝。"
  echo
  echo "處理方式:"
  echo "  1. 先點兩下「安裝.command」,等待顯示「安裝完成」"
  echo "  2. 再點兩下「啟動.command」"
  echo
  read -n 1 -s -r -p "按任意鍵關閉視窗。"
  exit 1
fi

# 使用 venv 中的 Python 執行,等同啟用 venv
venv/bin/python deploy/launcher.py
echo
read -n 1 -s -r -p "服務已停止。按任意鍵關閉視窗。"
echo
