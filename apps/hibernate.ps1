# irm https://github.com/therepos/windows/raw/main/apps/hibernate.ps1 | iex
# Purpose: Enable hibernation on Windows (powercfg /hibernate on)
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/hibernate.ps1'

# --- Self-elevate: re-run the same one-liner in an elevated window -----------
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "  Requesting administrator privileges..." -ForegroundColor Yellow
    $cmd = "[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; irm '$ScriptUrl?$(Get-Random)' | iex"
    Start-Process -FilePath 'powershell.exe' -Verb RunAs `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-Command', $cmd
    Write-Host "  Continue in the new (elevated) window. You can close this one." -ForegroundColor Gray
    return
}

powercfg /hibernate on
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Hibernation enabled." -ForegroundColor Green
} else {
    Write-Host "  Failed (exit $LASTEXITCODE). Firmware may not support hibernate (S4)." -ForegroundColor Red
}

Write-Host ""
powercfg /availablesleepstates
