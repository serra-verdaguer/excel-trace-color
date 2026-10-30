@echo off
setlocal
rem One-click installer for the Trace ^& Color Excel add-in (Microsoft 365 for Windows).
rem It downloads manifest.xml and registers it with Excel. No admin rights needed.

set "DIR=%LOCALAPPDATA%\TraceColor"
set "MANIFEST=%DIR%\manifest.xml"
set "URL=https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/manifest.xml"

if not exist "%DIR%" mkdir "%DIR%"
echo Downloading Trace ^& Color...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -UseBasicParsing '%URL%' -OutFile '%MANIFEST%'"
if errorlevel 1 (
  echo.
  echo Download failed. Check your internet connection and try again.
  pause
  exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "226f2a76-31d8-4adc-95c6-053ff2729ee0" /t REG_SZ /d "%MANIFEST%" /f >nul
if errorlevel 1 (
  echo Could not register the add-in with Excel.
  pause
  exit /b 1
)

echo.
echo Installed!
echo 1. Close Excel completely (all windows) and open it again.
echo 2. Home ^> Add-ins ^> More Add-ins ^> MY ADD-INS ^> "Developer Add-ins"
echo    ^> click "Trace & Color" ^> Add.  (Only needed the first time.)
echo 3. The "Trace & Color" tab appears on the ribbon.
echo.
pause
