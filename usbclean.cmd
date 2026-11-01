@echo off
net session >nul 2>&1
if %errorlevel% neq 0 (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process '%~f0' -Verb RunAs"
  exit /b
)
echo === DISK LIST ===
(
  echo list disk
  echo exit
) > "%temp%\__listdisk.txt"
diskpart /s "%temp%\__listdisk.txt"
del "%temp%\__listdisk.txt" >nul 2>&1
echo.
set /p DISKNUM=Enter USB disk number to CLEAN: 
if "%DISKNUM%"=="" exit /b 1
echo.
echo *** WARNING: DISK %DISKNUM% WILL BE ERASED ***
set /p CONFIRM=Type YES to continue: 
if /I not "%CONFIRM%"=="YES" exit /b 2
set DP=%temp%\usb_clean.txt
(
  echo select disk %DISKNUM%
  echo clean
  echo convert mbr
  echo create partition primary
  echo format fs=fat32 quick
  echo assign
  echo exit
) > "%DP%"
diskpart /s "%DP%"
del "%DP%" >nul 2>&1
pause
