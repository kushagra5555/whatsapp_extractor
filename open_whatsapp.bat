@echo off
set "PROJECT_DIR=%~dp0"
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" (
  echo Google Chrome was not found. Install Chrome, then run this file again.
  pause
  exit /b 1
)
start "" "%CHROME%" ^
  --user-data-dir="%PROJECT_DIR%wa_session" ^
  https://web.whatsapp.com/
