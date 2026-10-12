@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Excel 查詢網頁

if not exist "venv\Scripts\python.exe" goto :not_installed

rem 使用 venv 中的 Python 執行，等同啟用 venv
"venv\Scripts\python.exe" "%~dp0deploy\launcher.py"

echo.
echo 服務已停止。按任意鍵關閉視窗。
pause >nul
exit /b 0

:not_installed
echo.
echo [錯誤] 尚未安裝。
echo.
echo 處理方式：
echo   1. 先點兩下「安裝.bat」，等待顯示「安裝完成」
echo   2. 再點兩下「啟動.bat」
echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 1
