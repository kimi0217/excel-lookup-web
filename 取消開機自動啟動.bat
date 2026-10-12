@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Excel 查詢網頁 - 取消開機自動啟動

if not exist "venv\Scripts\python.exe" goto :fallback

"venv\Scripts\python.exe" "%~dp0deploy\autostart.py" remove
goto :done

:fallback
rem 沒有 venv 時（例如已刪除），直接用系統指令移除防火牆規則與排程工作
netsh advfirewall firewall show rule name=ExcelLookupWeb >nul 2>&1
if errorlevel 1 goto :fallback_task
netsh advfirewall firewall delete rule name=ExcelLookupWeb >nul 2>&1
if errorlevel 1 goto :firewall_failed
echo 已移除防火牆規則 ExcelLookupWeb。
goto :fallback_task

:firewall_failed
echo [注意] 無法移除防火牆規則 ExcelLookupWeb，請在此檔案上按滑鼠右鍵，選「以系統管理員身分執行」。

:fallback_task
schtasks /Query /TN ExcelLookupWeb >nul 2>&1
if errorlevel 1 (
  echo 目前沒有設定自動啟動，不需要取消。
  goto :done
)
schtasks /End /TN ExcelLookupWeb >nul 2>&1
schtasks /Delete /TN ExcelLookupWeb /F >nul 2>&1
if errorlevel 1 (
  echo [錯誤] 無法取消自動啟動。請在此檔案上按滑鼠右鍵，選「以系統管理員身分執行」。
  goto :done
)
echo 已取消自動啟動。

:done
echo.
echo 按任意鍵關閉視窗。
pause >nul
exit /b 0
