@echo off
setlocal
cd /d "%~dp0"

echo Building MailZen and cleaning previous release artifacts...
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-MailZen.ps1"
set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
    echo.
    echo Build failed with exit code %EXITCODE%.
    pause
    exit /b %EXITCODE%
)

echo.
echo Build, cleanup, verification, commit, and GitHub push completed.
pause
exit /b 0
