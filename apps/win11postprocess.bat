@echo off
:: Purpose: Apply post-install settings to a fresh Windows 11 install
:: ============================================================================
:: Note:
::   Double-click to run; asks for admin
:: ============================================================================

:: --- Self-elevate to Admin if needed ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process cmd.exe -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

:: --- Extract embedded PowerShell script to temp ---
set "PS_SCRIPT=%TEMP%\PostInstall_%RANDOM%.ps1"

setlocal enabledelayedexpansion
set "COPYING=0"
(
    for /f "usebackq delims=" %%L in ("%~f0") do (
        if "%%L"=="::END_PS1" set "COPYING=0"
        if !COPYING!==1 echo %%L
        if "%%L"=="::BEGIN_PS1" set "COPYING=1"
    )
) > "%PS_SCRIPT%"
endlocal

:: --- Run the script ---
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

:: --- Cleanup temp file ---
del "%PS_SCRIPT%" >nul 2>&1

echo.
echo =============================================
echo   Script finished. You can close this window.
echo =============================================
pause
exit /b

::BEGIN_PS1
# =============================================
# Post-Fresh Windows Install Automation Script
# =============================================

$ErrorActionPreference = "Continue"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = "$env:USERPROFILE\Desktop\PostInstall_Log_$Timestamp.txt"

# --- Logging ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry
    switch ($Level) {
        "INFO"    { Write-Host $Message -ForegroundColor Cyan }
        "WARN"    { Write-Host $Message -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
    }
}

# --- Counters ---
$script:Removed    = 0
$script:RmFailed   = 0
$script:Installed  = 0
$script:InstFailed = 0

# =============================================
# BLOATWARE REMOVAL
# =============================================

$BloatwareList = [ordered]@{
    "Snipping Tool (ScreenSketch)"     = @("*ScreenSketch*")
    "LinkedIn"                         = @("*LinkedIn*")
    "Microsoft Solitaire Collection"   = @("*MicrosoftSolitaireCollection*", "*Solitaire*")
    "Microsoft / Bing News"            = @("*BingNews*", "*News*")
    "Clipchamp"                        = @("*Clipchamp*")
}

function Remove-Bloatware {
    Write-Log "===== BLOATWARE REMOVAL =====" "INFO"
    foreach ($app in $BloatwareList.GetEnumerator()) {
        $name     = $app.Key
        $patterns = $app.Value
        Write-Log "Removing: $name ..." "INFO"
        $success = $true
        foreach ($pattern in $patterns) {
            try {
                $pkg = Get-AppxPackage $pattern -ErrorAction SilentlyContinue
                if ($pkg) { $pkg | Remove-AppxPackage -ErrorAction Stop }
                $prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                        Where-Object { $_.PackageName -like $pattern }
                if ($prov) { $prov | Remove-AppxProvisionedPackage -Online -ErrorAction Stop | Out-Null }
            }
            catch {
                Write-Log "  Failed on pattern '$pattern': $_" "ERROR"
                $success = $false
            }
        }
        if ($success) {
            Write-Log "  $name removed." "SUCCESS"
            $script:Removed++
        } else {
            Write-Log "  $name removal had errors (see above)." "WARN"
            $script:RmFailed++
        }
    }
}

# =============================================
# APP INSTALLATION
# =============================================

$InstallList = [ordered]@{
    "Python Manager"   = @{ Id = "9NQ7512CXL7T";    Source = "msstore" }
    "Snipaste"         = @{ Id = "liule.Snipaste";   Source = "winget"  }
    "Microsoft 365"    = @{ Id = "Microsoft.Office";  Source = "winget"  }
}

function Install-Apps {
    Write-Log "`n===== APP INSTALLATION =====" "INFO"
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "winget not found. Install App Installer from the Microsoft Store." "ERROR"
        $script:InstFailed += $InstallList.Count
        return
    }
    foreach ($app in $InstallList.GetEnumerator()) {
        $name   = $app.Key
        $id     = $app.Value.Id
        $source = $app.Value.Source
        Write-Log "Installing: $name ($id) ..." "INFO"
        try {
            $wArgs = @("install", "--id", $id, "-e",
                       "--accept-package-agreements",
                       "--accept-source-agreements")
            if ($source -eq "msstore") { $wArgs += @("-s", "msstore") }
            $result   = & winget @wArgs 2>&1
            $exitCode = $LASTEXITCODE
            if ($exitCode -eq 0 -or $result -match "already installed") {
                Write-Log "  $name installed successfully." "SUCCESS"
                $script:Installed++
            } else {
                Write-Log "  $name install returned exit code $exitCode." "WARN"
                Write-Log "  Output: $($result | Out-String)" "WARN"
                $script:InstFailed++
            }
        }
        catch {
            Write-Log "  $name install failed: $_" "ERROR"
            $script:InstFailed++
        }
    }
}

# =============================================
# SUMMARY REPORT
# =============================================

function Show-Summary {
    $divider = "=" * 50
    $summary = @"

$divider
  POST-INSTALL SUMMARY
$divider
  Bloatware removed :  $($script:Removed) / $($BloatwareList.Count)
  Removal failures  :  $($script:RmFailed)
  Apps installed    :  $($script:Installed) / $($InstallList.Count)
  Install failures  :  $($script:InstFailed)
$divider
  Log saved to: $LogFile
$divider
"@
    Write-Host $summary -ForegroundColor White
    Add-Content -Path $LogFile -Value $summary
    if (($script:RmFailed + $script:InstFailed) -gt 0) {
        Write-Host "  Some operations had issues. Review the log for details." -ForegroundColor Yellow
    } else {
        Write-Host "  All operations completed successfully!" -ForegroundColor Green
    }
    Write-Host "`n  A restart is recommended for all changes to take effect." -ForegroundColor Cyan
}

# =============================================
# MAIN
# =============================================

Write-Log "Post-install script started." "INFO"
Write-Log "Log file: $LogFile" "INFO"

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Remove-Bloatware
Install-Apps

$stopwatch.Stop()
Write-Log "Completed in $([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) seconds." "INFO"

Show-Summary
::END_PS1
