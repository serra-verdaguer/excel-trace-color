@echo off
rem Removes the Trace ^& Color Excel add-in installed by Install-Windows.cmd.
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "226f2a76-31d8-4adc-95c6-053ff2729ee0" /f >nul 2>&1
if exist "%LOCALAPPDATA%\TraceColor" rmdir /s /q "%LOCALAPPDATA%\TraceColor"
echo Trace ^& Color removed. Restart Excel.
pause
