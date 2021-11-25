@echo off

REM -- Prepare the Command Processor --
SETLOCAL ENABLEEXTENSIONS
SETLOCAL DISABLEDELAYEDEXPANSION

if /I "%~1"=="/h" goto:help
if "%3"=="" goto:help

if "%~1"=="" findstr "^::" "%~f0"&goto:help
for /f "tokens=1,* delims=]" %%A in ('"type %3|find /n /v """') do (
    set "line=%%B"
    if defined line (
        call set "line=echo.%%line:%~1=%~2%%"
        for /f "delims=" %%X in ('"echo."%%line%%""')  do %%~X >> %3_temp
    ) ELSE echo.
)

move /Y %3_temp %3 >nul

goto:eof

:help
echo BatchSubstitude - parses a File line by line and replaces a substring
echo.
echo Usage: %0 [oldstr] [newstr] [filename]
echo.          oldstr  - string to be replaced
echo.          newstr  - string to replace with
echo.          filname - file to be parsed
echo.
goto:eof

:eof