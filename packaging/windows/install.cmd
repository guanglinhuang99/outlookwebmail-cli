@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
set "WEBMAIL_INSTALL_EXIT=%ERRORLEVEL%"
echo.
pause
exit /b %WEBMAIL_INSTALL_EXIT%
