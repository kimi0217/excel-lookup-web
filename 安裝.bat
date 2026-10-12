@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Excel 查詢網頁 - 安裝

rem 尋找 Python 3.11 以上：先試 py 啟動器，再試 python
set "PY="
set "OLD_PY="
call :try_python py -3
if not defined PY call :try_python python
if not defined PY goto :no_python

%PY% "%~dp0deploy\install.py"
if errorlevel 1 goto :failed

echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 0

:try_python
%* -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 10)" >nul 2>&1
set "RC=%errorlevel%"
if "%RC%"=="0" set "PY=%*"
if "%RC%"=="10" if not defined OLD_PY set "OLD_PY=%*"
exit /b 0

:no_python
echo.
echo ========================================================
echo   [錯誤] 找不到 Python 3.11 以上的版本
echo ========================================================
echo.
if defined OLD_PY echo 這台電腦安裝的 Python 版本太舊：
if defined OLD_PY %OLD_PY% --version
if defined OLD_PY echo.
echo 請依照下列步驟安裝 Python：
echo   1. 開啟 https://www.python.org/downloads/ （稍後會自動開啟）
echo   2. 按黃色的「Download Python」按鈕下載安裝程式
echo   3. 執行安裝程式，第一個畫面最下方務必勾選
echo      「Add python.exe to PATH」，再按「Install Now」
echo   4. 安裝完成後，重新點兩下「安裝.bat」
echo.
start "" "https://www.python.org/downloads/"
echo 按任意鍵關閉視窗。
pause >nul
exit /b 1

:failed
echo 安裝沒有完成，請依照上方的說明處理後，重新點兩下「安裝.bat」。
echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 1
