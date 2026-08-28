@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install ExcelChart Selection Add-in.ps1"
set "exitCode=%ERRORLEVEL%"
echo.
if not "%exitCode%"=="0" echo Installation did not complete.
pause
exit /b %exitCode%
