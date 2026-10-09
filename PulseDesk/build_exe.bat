@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo    PulseDesk - 一鍵建立 / 更新 EXE
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo [錯誤] 找不到 Python。
  echo 請先安裝 Python 3，安裝畫面記得勾選 "Add python.exe to PATH"。
  start https://www.python.org/downloads/windows/
  pause
  exit /b 1
)

echo [1/4] 關閉正在運行的 PulseDesk（包括背景中的舊版本）...
powershell -NoProfile -Command "try{Invoke-WebRequest -UseBasicParsing -Method Post -Uri http://127.0.0.1:8765/api/shutdown -Body '{}' -ContentType 'application/json' -TimeoutSec 3 | Out-Null}catch{}"
timeout /t 2 /nobreak >nul
taskkill /IM PulseDesk.exe /F >nul 2>nul
timeout /t 1 /nobreak >nul

echo.
echo [2/4] 安裝所需套件...
python -m pip install --upgrade pywin32 pyinstaller
if errorlevel 1 goto fail

echo.
echo [3/4] 建立 EXE，大約需要 1 至 2 分鐘...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PulseDesk --hidden-import win32timezone --add-data "ui.html;." pulsedesk.py
if errorlevel 1 goto fail

copy /y "dist\PulseDesk.exe" "PulseDesk.exe" >nul
if errorlevel 1 (
  echo.
  echo [錯誤] 無法取代 PulseDesk.exe，可能仍有 PulseDesk 在運行。
  echo 請在工作管理員結束所有 PulseDesk，然後再執行一次 build_exe.bat。
  pause
  exit /b 1
)
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q PulseDesk.spec 2>nul

echo.
echo [4/4] 更新桌面捷徑...
powershell -NoProfile -Command "$d=[Environment]::GetFolderPath('Desktop'); $exe=Join-Path '%~dp0' 'PulseDesk.exe'; $copy=Join-Path $d 'PulseDesk.exe'; if(Test-Path $copy){ Copy-Item $exe $copy -Force; Write-Host 'Desktop copy of PulseDesk.exe updated' } else { $s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'PulseDesk.lnk')); $s.TargetPath=$exe; $s.WorkingDirectory='%~dp0'; $s.Save(); Write-Host 'Desktop shortcut PulseDesk created' }"

echo.
echo 完成！以後雙擊桌面的 PulseDesk 就可以使用。
echo 左下角會顯示版本號碼（例如 v2.7），可用來確認已經更新。
start "" "%~dp0PulseDesk.exe"
pause
exit /b 0

:fail
echo.
echo [錯誤] 建立失敗，請把上面的訊息截圖發給我。
pause
exit /b 1
