@echo off
rem Removes Trace ^& Color (both the permanent and the Developer-mode install).
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{C40A524E-4F95-44AD-B692-E1E752825EA9}" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "226f2a76-31d8-4adc-95c6-053ff2729ee0" /f >nul 2>&1
if exist "%LOCALAPPDATA%\TraceColor" rmdir /s /q "%LOCALAPPDATA%\TraceColor"
echo Removing the shared folder (click Yes if Windows asks)...
powershell -NoProfile -Command "Start-Process powershell -Verb RunAs -Wait -WindowStyle Hidden -ArgumentList '-NoProfile','-Command','Remove-SmbShare -Name TraceColor -Force -ErrorAction SilentlyContinue'"
if exist "%PUBLIC%\TraceColor" rmdir /s /q "%PUBLIC%\TraceColor"
echo Trace ^& Color removed. Restart Excel.
pause
