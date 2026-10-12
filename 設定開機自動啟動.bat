@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Excel 查詢網頁 - 設定開機自動啟動

if not exist "venv\Scripts\python.exe" goto :not_installed

"venv\Scripts\python.exe" "%~dp0deploy\autostart.py" install

echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 0

:not_installed
echo.
echo [錯誤] 尚未安裝。
echo.
echo 處理方式：
echo   1. 先點兩下「安裝.bat」
echo   2. 再點兩下「啟動.bat」完成第一次設定
echo   3. 最後再點兩下「設定開機自動啟動.bat」
echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 1
