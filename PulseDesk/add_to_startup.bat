@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "PulseDesk.exe" (
  echo [錯誤] 找不到 PulseDesk.exe，請先執行 build_exe.bat。
  pause
  exit /b 1
)
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Startup')+'\PulseDesk.lnk'); $s.TargetPath='%~dp0PulseDesk.exe'; $s.WorkingDirectory='%~dp0'; $s.Save()"
if errorlevel 1 (
  echo [錯誤] 未能建立捷徑。
  pause
  exit /b 1
)
echo 完成！以後開機登入 Windows 時會自動打開 PulseDesk。
echo 如要取消，刪除「啟動」資料夾內的 PulseDesk 捷徑即可（Win+R 輸入 shell:startup）。
pause
