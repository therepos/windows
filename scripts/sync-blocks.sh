#!/usr/bin/env bash
# Purpose: Copy the shared blocks (say, menukey, downloads, elevate, layout) into every PowerShell script
# =============================================================================
# Usage:
#   scripts/sync-blocks.sh           copy each block into every script that has its markers
#   scripts/sync-blocks.sh --check   report scripts out of sync, change nothing
#
# Markers:
#   # >>> <name>  …  # <<< <name>
# =============================================================================
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

# --- Blocks ------------------------------------------------------------------
BLOCK_NAMES=(say menukey downloads elevate layout)

# say: coloured console output
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_say=$(cat <<'BLOCK'
function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }
BLOCK
)

# menukey: one keypress picks a menu option; Enter picks the default
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_menukey=$(cat <<'BLOCK'
function Read-MenuKey([string]$Prompt, [string]$Valid, [string]$Default = '') {
    Write-Host "  $Prompt$(if ($Default) { " [$Default]" }): " -NoNewline
    while ($true) {
        try {
            $k = [Console]::ReadKey($true)
            $c = if ($k.Key -eq 'Enter') { $Default } else { "$($k.KeyChar)" }
        } catch {
            # No keyboard (input redirected): fall back to a typed line
            $line = [Console]::ReadLine()
            if ($null -eq $line) { throw 'No keyboard input available' }
            $c = $line.Trim(); if (-not $c) { $c = $Default }
        }
        if ($c.Length -eq 1 -and $Valid.Contains($c)) {
            Write-Host $c
            # Swallow an Enter typed out of habit so it does not reach the next prompt
            Start-Sleep -Milliseconds 300
            try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch {}
            return $c
        }
    }
}
BLOCK
)

# downloads: the user's real Downloads folder (registry Shell Folders)
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_downloads=$(cat <<'BLOCK'
function Get-DownloadsFolder {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $raw = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'{374DE290-123F-4565-9164-39C4925E467B}'
    if ($raw) {
        $path = [Environment]::ExpandEnvironmentVariables($raw)
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return (Join-Path $env:USERPROFILE 'Downloads')
}
BLOCK
)

# elevate: re-run this script as admin, in Windows Terminal when installed (Windows ignores
# the default-terminal setting for admin windows). From GitHub it re-runs the bootstrap, from a
# file (shortcut) that file. -NoExit keeps the admin window open after a one-shot script.
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_elevate=$(cat <<'BLOCK'
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
BLOCK
)

# layout: put the script's window on the right half of its screen (Windows Terminal or the
# classic console), in physical pixels so it is exact at any display scaling
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_layout=$(cat <<'BLOCK'
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
BLOCK
)

# --- Sync --------------------------------------------------------------------
CHECK=0
[[ "${1:-}" == "--check" ]] && CHECK=1

TMP_BLOCK=$(mktemp)
trap 'rm -f "$TMP_BLOCK"' EXIT

drift=0
changed=0
for f in "$ROOT"/apps/*.ps1; do
    [[ -f "$f" ]] || continue
    cp "$f" "$f.new"
    for b in "${BLOCK_NAMES[@]}"; do
        grep -qE "^# >>> $b( |$)" "$f.new" || continue
        var="BLOCK_$b"
        printf '%s\n' "${!var}" > "$TMP_BLOCK"
        awk -v blockfile="$TMP_BLOCK" -v b="$b" '
            BEGIN { n=0; while ((getline l < blockfile) > 0) blk[n++]=l; close(blockfile) }
            $0 ~ "^# >>> " b "( |$)" { print; for (i=0;i<n;i++) print blk[i]; skip=1; next }
            $0 ~ "^# <<< " b "( |$)" { skip=0; print; next }
            !skip { print }
        ' "$f.new" > "$f.tmp" && mv "$f.tmp" "$f.new"
    done

    if cmp -s "$f" "$f.new"; then
        rm -f "$f.new"
    elif [[ $CHECK -eq 1 ]]; then
        rm -f "$f.new"
        echo "DRIFT: $(basename "$f")"
        drift=1
    else
        cat "$f.new" > "$f" && rm -f "$f.new"
        echo "updated: $(basename "$f")"
        changed=1
    fi
done

if [[ $CHECK -eq 1 ]]; then
    [[ $drift -eq 0 ]] && { echo "All in sync."; exit 0; } || { echo "Out of sync — run scripts/sync-blocks.sh"; exit 1; }
fi
[[ $changed -eq 0 ]] && echo "All in sync."
exit 0
