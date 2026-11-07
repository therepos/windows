@echo off
:: Purpose: Toggle Windows SMB client caching so mapped network drives refresh instantly
:: ============================================================================
:: Note:
::   Double-click to switch between "caching off" and Windows defaults. Asks for admin
::   Admin windows cannot see mapped drives, so it re-launches a copy from %TEMP%
:: ============================================================================

net session >nul 2>&1 || (
    copy /y "%~f0" "%TEMP%\smb-cache-toggle.cmd" >nul
    powershell -NoProfile -Command "Start-Process -FilePath '%TEMP%\smb-cache-toggle.cmd' -Verb RunAs"
    exit /b
)

set "K=HKLM\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters"
set "VALUES=DirectoryCacheLifetime FileNotFoundCacheLifetime FileInfoCacheLifetime"

reg query "%K%" /v DirectoryCacheLifetime 2>nul | find "0x0" >nul
if errorlevel 1 (
    for %%V in (%VALUES%) do reg add "%K%" /v %%V /t REG_DWORD /d 0 /f >nul
    echo.
    echo  Network drive caching is now OFF - folders refresh instantly.
) else (
    for %%V in (%VALUES%) do reg delete "%K%" /v %%V /f >nul 2>&1
    echo.
    echo  Network drive caching is now ON - Windows defaults restored.
)

echo  Restart Windows to apply. Run this file again to switch back.
echo.
choice /c YN /m " Restart now"
if not errorlevel 2 shutdown /r /t 5
