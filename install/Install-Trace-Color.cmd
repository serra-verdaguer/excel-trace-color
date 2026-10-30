@echo off
setlocal
rem One-file installer for the Trace ^& Color Excel add-in (Microsoft 365 for Windows).
rem
rem Permanent install (Windows asks once for permission):
rem   1. Saves manifest.xml in %PUBLIC%\TraceColor
rem   2. Shares that folder read-only as \\%COMPUTERNAME%\TraceColor
rem   3. Registers the share as a trusted add-in catalog, which Excel keeps between restarts
rem   4. Removes any older Developer-mode registration, so you don't get two copies
rem No admin rights (or permission declined)? It falls back to Developer mode, which
rem needs no permission but Excel may forget after a restart.

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
if not exist "%SHARE%\manifest.xml" goto :developer

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

:developer
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "226f2a76-31d8-4adc-95c6-053ff2729ee0" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
echo.
echo Installed in basic mode (this PC did not allow sharing the folder).
echo   1. Close Excel completely (all windows) and open it again.
echo   2. Home ^> Add-ins ^> More Add-ins ^> MY ADD-INS ^> Developer Add-ins ^> Trace ^& Color ^> Add.
echo Excel may forget it after a restart. To make it permanent, run this file again
echo and click Yes when Windows asks, or ask IT to deploy it (Microsoft 365 admin center ^> Integrated apps).
echo.
pause
exit /b 0
