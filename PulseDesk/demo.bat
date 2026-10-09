@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo [錯誤] 找不到 Python，請先安裝 Python 3。
  start https://www.python.org/downloads/windows/
  pause
  exit /b 1
)
start "" pythonw pulsedesk.py --demo
