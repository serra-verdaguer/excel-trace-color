@echo off
setlocal
rem Permanent installer for the Trace ^& Color Excel add-in (Microsoft 365 for Windows).
rem Installs it through a "shared folder catalog", which Excel keeps between restarts
rem (unlike Developer add-ins). Windows asks once for permission to share the folder.
rem
rem What it does:
rem   1. Saves manifest.xml in %PUBLIC%\TraceColor
rem   2. Shares that folder read-only as \\%COMPUTERNAME%\TraceColor (needs admin, asked once)
rem   3. Registers the share as a trusted add-in catalog for Excel
rem   4. Removes the old Developer-mode registration, so you don't get two copies

set "DIR=%PUBLIC%\TraceColor"
set "SHARE=\\%COMPUTERNAME%\TraceColor"
set "URL=https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/manifest.xml"
set "CATALOG=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{C40A524E-4F95-44AD-B692-E1E752825EA9}"

echo Downloading Trace ^& Color...
if not exist "%DIR%" mkdir "%DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -UseBasicParsing '%URL%' -OutFile '%DIR%\manifest.xml'"
if errorlevel 1 goto :fail_download

echo Sharing the add-in folder (click Yes if Windows asks)...
powershell -NoProfile -Command "Start-Process powershell -Verb RunAs -Wait -WindowStyle Hidden -ArgumentList '-NoProfile','-EncodedCommand','JABwACAAPQAgAEoAbwBpAG4ALQBQAGEAdABoACAAJABlAG4AdgA6AFAAVQBCAEwASQBDACAAJwBUAHIAYQBjAGUAQwBvAGwAbwByACcACgBpAGYAIAAoAC0AbgBvAHQAIAAoAEcAZQB0AC0AUwBtAGIAUwBoAGEAcgBlACAALQBOAGEAbQBlACAAJwBUAHIAYQBjAGUAQwBvAGwAbwByACcAIAAtAEUAcgByAG8AcgBBAGMAdABpAG8AbgAgAFMAaQBsAGUAbgB0AGwAeQBDAG8AbgB0AGkAbgB1AGUAKQApACAAewAKACAAIAAkAGUAdgBlAHIAeQBvAG4AZQAgAD0AIAAoAE4AZQB3AC0ATwBiAGoAZQBjAHQAIABTAGUAYwB1AHIAaQB0AHkALgBQAHIAaQBuAGMAaQBwAGEAbAAuAFMAZQBjAHUAcgBpAHQAeQBJAGQAZQBuAHQAaQBmAGkAZQByACAAJwBTAC0AMQAtADEALQAwACcAKQAuAFQAcgBhAG4AcwBsAGEAdABlACgAWwBTAGUAYwB1AHIAaQB0AHkALgBQAHIAaQBuAGMAaQBwAGEAbAAuAE4AVABBAGMAYwBvAHUAbgB0AF0AKQAuAFYAYQBsAHUAZQAKACAAIABOAGUAdwAtAFMAbQBiAFMAaABhAHIAZQAgAC0ATgBhAG0AZQAgACcAVAByAGEAYwBlAEMAbwBsAG8AcgAnACAALQBQAGEAdABoACAAJABwACAALQBSAGUAYQBkAEEAYwBjAGUAcwBzACAAJABlAHYAZQByAHkAbwBuAGUAIAB8ACAATwB1AHQALQBOAHUAbABsAAoAfQA='"
if not exist "%SHARE%\manifest.xml" goto :fail_share

reg add "%CATALOG%" /v Id /t REG_SZ /d "{C40A524E-4F95-44AD-B692-E1E752825EA9}" /f >nul
reg add "%CATALOG%" /v Url /t REG_SZ /d "%SHARE%" /f >nul
reg add "%CATALOG%" /v Flags /t REG_DWORD /d 1 /f >nul
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "226f2a76-31d8-4adc-95c6-053ff2729ee0" /f >nul 2>&1

echo.
echo Installed!  Do this ONCE:
echo   1. Close Excel completely (all windows) and open it again.
echo   2. Home ^> Add-ins ^> More Add-ins ^> SHARED FOLDER tab ^> Trace ^& Color ^> Add.
echo From then on the Trace ^& Color tab is there every time you open Excel.
echo.
pause
exit /b 0

:fail_download
echo.
echo Download failed. Check your internet connection and try again.
pause
exit /b 1

:fail_share
echo.
echo Could not share the folder %DIR%.
echo You may not have administrator rights on this PC (common on company laptops).
echo In that case ask IT to deploy the add-in: Microsoft 365 admin center ^> Integrated apps.
pause
exit /b 1
