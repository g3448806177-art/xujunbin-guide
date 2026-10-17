@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if errorlevel 1 (
  echo Installation failed. Keep this window open to read the error.
  pause
  exit /b 1
)
echo Installation complete.
pause
