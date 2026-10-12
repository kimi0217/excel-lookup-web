@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Excel 查詢網頁 - 設定防火牆
rem 需要以系統管理員身分執行（按右鍵選「以系統管理員身分執行」）；權限檢查與說明在 deploy\firewall.py

if not exist "venv\Scripts\python.exe" goto :not_installed

"venv\Scripts\python.exe" "%~dp0deploy\firewall.py" add

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
echo   3. 最後在「設定防火牆.bat」上按滑鼠右鍵，選「以系統管理員身分執行」
echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 1
