@echo off
setlocal

set "ROOT=%~dp0"
set "PROJECT=%ROOT%src\EmailManage.App\EmailManage.App.csproj"
set "PUBLISH_DIR=%ROOT%bin\Release\single-file"
set "PUBLISHED_EXE=%PUBLISH_DIR%\MailZen.exe"
set "ROOT_EXE=%ROOT%MailZen.exe"

echo Publishing MailZen (Release, single-file, win-x64)...
dotnet publish "%PROJECT%" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "%PUBLISH_DIR%"
if errorlevel 1 (
	echo Publish failed.
	pause
	exit /b 1
)

copy /y "%PUBLISHED_EXE%" "%ROOT_EXE%" >nul
if errorlevel 1 (
	echo Failed to copy MailZen.exe to repository root.
	pause
	exit /b 1
)

echo Launching %ROOT_EXE%
start "" "%ROOT_EXE%"
endlocal
