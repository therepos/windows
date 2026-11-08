# irm "https://github.com/therepos/windows/raw/main/apps/hibernate.ps1?$(Get-Random)" | iex
# Purpose: Enable hibernation on Windows (powercfg /hibernate on)
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/hibernate.ps1'

# >>> elevate
function Test-Admin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Start-Elevated([string]$ScriptUrl, [switch]$NoExit) {
    Write-Host "  Requesting administrator rights..." -ForegroundColor Yellow
    $run = if ($PSCommandPath) { "& '$($PSCommandPath -replace "'", "''")'" } else { "irm '${ScriptUrl}?$(Get-Random)' | iex" }
    $cmd = "[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; try { $run } catch { Write-Host `$_ -ForegroundColor Red; Read-Host 'Press Enter to exit' }"
    # Encoded, so Windows Terminal does not split the command at ';' or mangle quotes
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass')
    if ($NoExit) { $psArgs += '-NoExit' }
    $psArgs += @('-EncodedCommand', $enc)
    $opened = $false
    if (Get-Command wt.exe -ErrorAction SilentlyContinue) {
        try {
            Start-Process -FilePath 'wt.exe' -Verb RunAs -ArgumentList (@('powershell.exe') + $psArgs) -ErrorAction Stop
            $opened = $true
        } catch {
            # 'No' on the admin prompt: do not ask again for the classic window
            $e = $_.Exception
            while ($e.InnerException) { $e = $e.InnerException }
            if ($e -is [ComponentModel.Win32Exception] -and $e.NativeErrorCode -eq 1223) {
                Write-Host "  Cancelled." -ForegroundColor Yellow
                return
            }
        }
    }
    if (-not $opened) { Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $psArgs }
    Write-Host "  Continue in the new (elevated) window. You can close this one." -ForegroundColor Gray
}
# <<< elevate
$IsAdmin = Test-Admin
if (-not $IsAdmin) { Start-Elevated $ScriptUrl -NoExit; return }
# >>> layout
function Set-WindowRight {
    try {
        if (-not ('WinLayout.Native' -as [type])) {
            Add-Type -Namespace WinLayout -Name Native -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
[DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
[DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr h, uint f);
[DllImport("user32.dll")] public static extern bool GetMonitorInfo(IntPtr m, ref MONITORINFO mi);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
[DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int a, out RECT r, int s);
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
[StructLayout(LayoutKind.Sequential)] public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }
'@
        }
        $h = [WinLayout.Native]::GetConsoleWindow()
        if ($h -eq [IntPtr]::Zero -or -not [WinLayout.Native]::IsWindowVisible($h)) {
            # Windows Terminal hides the console window: use its own window, in front at start
            $h = [WinLayout.Native]::GetForegroundWindow()
            $owner = [uint32]0
            [void][WinLayout.Native]::GetWindowThreadProcessId($h, [ref]$owner)
            if ((Get-Process -Id $owner -ErrorAction SilentlyContinue).ProcessName -ne 'WindowsTerminal') { return }
        }
        $oldDpi = [WinLayout.Native]::SetThreadDpiAwarenessContext([IntPtr](-4))
        try {
            [void][WinLayout.Native]::ShowWindow($h, 9)
            $mi = New-Object WinLayout.Native+MONITORINFO
            $mi.cbSize = 40
            [void][WinLayout.Native]::GetMonitorInfo([WinLayout.Native]::MonitorFromWindow($h, 2), [ref]$mi)
            $w = $mi.rcWork
            $half = [int][Math]::Floor(($w.Right - $w.Left) / 2)
            # Grow by the invisible resize borders so the visible frame fills the half exactly
            $wr = New-Object WinLayout.Native+RECT
            $fr = New-Object WinLayout.Native+RECT
            [void][WinLayout.Native]::GetWindowRect($h, [ref]$wr)
            $l = 0; $r = 0; $b = 0
            if ([WinLayout.Native]::DwmGetWindowAttribute($h, 9, [ref]$fr, 16) -eq 0) {
                $l = $fr.Left - $wr.Left; $r = $wr.Right - $fr.Right; $b = $wr.Bottom - $fr.Bottom
            }
            [void][WinLayout.Native]::SetWindowPos($h, [IntPtr]::Zero, ($w.Left + $half - $l), $w.Top,
                ($w.Right - $w.Left - $half + $l + $r), ($w.Bottom - $w.Top + $b), 0x0014)
        } finally {
            if ($oldDpi -ne [IntPtr]::Zero) { [void][WinLayout.Native]::SetThreadDpiAwarenessContext($oldDpi) }
        }
    } catch {}
}
# <<< layout
Set-WindowRight

powercfg /hibernate on
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Hibernation enabled." -ForegroundColor Green
} else {
    Write-Host "  Failed (exit $LASTEXITCODE). Firmware may not support hibernate (S4)." -ForegroundColor Red
}

Write-Host ""
powercfg /availablesleepstates
