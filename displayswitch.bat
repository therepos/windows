@echo off
setlocal

set STATE_DIR=%LOCALAPPDATA%\DisplayToggle
set STATE_FILE=%STATE_DIR%\state.txt

:: Create folder if it doesn't exist
if not exist "%STATE_DIR%" mkdir "%STATE_DIR%"

if exist "%STATE_FILE%" (
    set /p CURRENT=<"%STATE_FILE%"
) else (
    set CURRENT=internal
)

if "%CURRENT%"=="internal" (
    Start "" "%SystemRoot%\System32\DisplaySwitch.exe" /extend
    echo extend>"%STATE_FILE%"
) else (
    Start "" "%SystemRoot%\System32\DisplaySwitch.exe" /internal
    echo internal>"%STATE_FILE%"
)

endlocal