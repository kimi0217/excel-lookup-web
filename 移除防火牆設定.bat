@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Excel 查詢網頁 - 移除防火牆設定
rem 需要以系統管理員身分執行（按右鍵選「以系統管理員身分執行」）

if not exist "venv\Scripts\python.exe" goto :fallback

"venv\Scripts\python.exe" "%~dp0deploy\firewall.py" remove
goto :done

:fallback
rem 沒有 venv 時（例如已刪除），直接用 netsh 移除
netsh advfirewall firewall show rule name=ExcelLookupWeb >nul 2>&1
if errorlevel 1 goto :no_rule
netsh advfirewall firewall delete rule name=ExcelLookupWeb >nul 2>&1
if errorlevel 1 goto :failed
echo 已移除防火牆規則 ExcelLookupWeb。
goto :done

:no_rule
echo 目前沒有設定防火牆規則，不需要移除。
goto :done

:failed
echo [錯誤] 無法移除防火牆規則。請在此檔案上按滑鼠右鍵，選「以系統管理員身分執行」。

:done
echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 0
