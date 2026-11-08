# irm "https://github.com/therepos/windows/raw/main/apps/appupdate.ps1?$(Get-Random)" | iex
# Purpose: Update all apps with winget every Sunday 23:45, closing open apps first and reopening them after
# =============================================================================
# Usage:
#   1 = Weekly update   on/off; scheduled task Sundays 23:45, or at next start if missed
#   2 = Update now      same update, in this window
#   3 = Open log        %LOCALAPPDATA%\appupdate\update.log
#   4 = Settings        'appupdate' shortcut (runs the local copy in %LOCALAPPDATA%\appupdate)
#   0 = Quit
#
# Note:
#   Open apps are closed first (forced after 20 s), updated, then reopened as they were (minimized or tray)
#   A failed update is retried twice (second with winget --force), closing apps that hold its files
#   Update now shows a bar per app; the weekly task runs hidden and ends with a notification
#   Skipped: apps pinned with 'winget pin add <id>', and Microsoft Store apps
#   Needs admin (asks); installs Microsoft.WinGet.Client if missing
#   The task runs a local copy; turning the weekly update on again refreshes it
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/appupdate.ps1'
$ToolDir   = Join-Path $env:LOCALAPPDATA 'appupdate'
$LocalCopy = Join-Path $ToolDir 'appupdate.ps1'
$LogFile   = Join-Path $ToolDir 'update.log'
$LastFile  = Join-Path $ToolDir 'last.txt'
$TaskName  = 'AppUpdate'
$RunDay    = 'Sunday'
$RunAt     = '23:45'
$RunMode   = ($args -contains '-Run')      # set by the scheduled task

$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

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
if (-not $IsAdmin -and -not $RunMode) { Start-Elevated $ScriptUrl; return }

function Say($msg, $color = 'Gray') { if (-not $RunMode) { Write-Host $msg -ForegroundColor $color } }

# Shown in the window (when there is one) and appended to update.log
function Log($msg, $color = 'Gray') {
    Say "  $msg" $color
    try {
        if (-not (Test-Path -LiteralPath $ToolDir)) { New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null }
        Add-Content -LiteralPath $LogFile -Value ('{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)
    } catch {}
}

# >>> menukey
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
# <<< menukey

# =====================================================================
#  WINGET
# =====================================================================
function Initialize-WinGet {
    if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client)) {
        Log 'Installing the Microsoft.WinGet.Client module...' 'DarkGray'
        $nuget = Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue |
            Where-Object { $_.Version -ge [version]'2.8.5.201' }
        if (-not $nuget) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -Force | Out-Null }
        Install-Module -Name Microsoft.WinGet.Client -Repository PSGallery -Scope AllUsers -Force
    }
    Import-Module Microsoft.WinGet.Client -ErrorAction Stop
    $ok = $false
    try { $null = Get-WinGetVersion -ErrorAction Stop; $ok = $true } catch {}
    if (-not $ok) {
        Log 'winget is missing or broken, repairing it...' 'Yellow'
        Repair-WinGetPackageManager -Latest -Force -AllUsers | Out-Null
    }
    try { & winget.exe source update --disable-interactivity *> $null } catch {}
}

# The task can start right after the PC wakes, before Wi-Fi is back: wait up to 10 min
function Wait-Network {
    for ($i = 0; $i -lt 60; $i++) {
        try { $null = [Net.Dns]::GetHostEntry('cdn.winget.microsoft.com'); return $true } catch {}
        Start-Sleep -Seconds 10
    }
    $false
}

# Raw lines of 'winget pin list'; a package is pinned when its id appears as a word
function Get-PinLines {
    try { return @(& winget.exe pin list --disable-interactivity 2>$null) } catch { return @() }
}
function Test-Pinned([string]$Id, $PinLines) {
    $rx = '(^|\s)' + [regex]::Escape($Id) + '(\s|$)'
    foreach ($l in $PinLines) { if ("$l" -match $rx) { return $true } }
    $false
}

# =====================================================================
#  FINDING AND CLOSING AN APP'S PROCESSES
# =====================================================================
$UninstallKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*')

# Folder of the exe in a registry value like '"C:\App\app.exe",0' or 'C:\App\uninst.exe /S'
function Get-ExeDir([string]$s) {
    if (-not $s) { return $null }
    $s = $s.Trim()
    if ($s -match '^"([^"]+)"') { $p = $Matches[1] }
    elseif ($s -match '^(.+?\.exe)') { $p = $Matches[1] }
    else { $p = ($s -split ',')[0] }
    if ($p -match 'msiexec') { return $null }
    try { Split-Path -Path $p -Parent } catch { $null }
}

# Never treat a shared folder (Program Files, Windows, AppData...) as an app's folder,
# or every program in it would be closed
function Test-AppDir([string]$d) {
    if (-not $d -or $d.Length -le 3) { return $false }
    $roots = @($env:windir, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432, $env:ProgramData,
        "$env:ProgramFiles\Common Files", "${env:ProgramFiles(x86)}\Common Files", "$env:ProgramFiles\WindowsApps",
        $env:USERPROFILE, $env:APPDATA, $env:LOCALAPPDATA, "$env:LOCALAPPDATA\Programs", "$env:LOCALAPPDATA\Microsoft",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links", "$env:LOCALAPPDATA\Temp", $env:TEMP, "$env:SystemDrive\Users") |
        Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') }
    foreach ($r in $roots) { if ($d -ieq $r) { return $false } }
    if ($d.StartsWith($env:windir.TrimEnd('\') + '\', 'OrdinalIgnoreCase')) { return $false }
    try { return (Test-Path -LiteralPath $d -PathType Container) } catch { return $false }
}

# Install folders of a package: from its Add/Remove Programs entry, plus winget's portable folder
function Get-AppDirs($pkg) {
    $dirs = New-Object System.Collections.Generic.List[string]
    $name = [WildcardPattern]::Escape("$($pkg.Name)")
    foreach ($e in @(Get-ItemProperty -Path $UninstallKeys -ErrorAction SilentlyContinue)) {
        $dn = "$($e.DisplayName)"
        if (-not $dn) { continue }
        if ($dn -ieq "$($pkg.Name)" -or $dn -like "$name *" -or $dn -like "$name (*") {
            foreach ($d in @("$($e.InstallLocation)", (Get-ExeDir "$($e.DisplayIcon)"), (Get-ExeDir "$($e.UninstallString)"))) {
                if (-not $d) { continue }
                $d = $d.Trim().Trim('"').TrimEnd('\')
                if ((Test-AppDir $d) -and -not $dirs.Contains($d)) { $dirs.Add($d) }
            }
        }
    }
    # Zip-based apps (e.g. adb.exe from Platform-Tools): per-user and machine-wide folders
    foreach ($portable in @((Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'),
                            "$env:ProgramFiles\WinGet\Packages", "${env:ProgramFiles(x86)}\WinGet\Packages")) {
        Get-ChildItem -LiteralPath $portable -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name.StartsWith("$($pkg.Id)_", 'OrdinalIgnoreCase') } |
            ForEach-Object { if (-not $dirs.Contains($_.FullName)) { $dirs.Add($_.FullName) } }
    }
    , $dirs.ToArray()
}

# Running processes of a package: exe inside one of its folders, or named like the last
# part of its id (qBittorrent.qBittorrent -> qbittorrent.exe). Nothing from C:\Windows, and
# no Windows services (session 0): their installers stop and restart those themselves.
function Get-AppProcesses($pkg, $dirs) {
    $short  = ("$($pkg.Id)" -split '\.')[-1]
    $winDir = $env:windir.TrimEnd('\') + '\'
    @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $path = $null
        try { $path = $_.Path } catch {}
        if (-not $path -or $_.Id -eq $PID -or $_.SessionId -eq 0) { return $false }
        if ($path.StartsWith($winDir, 'OrdinalIgnoreCase')) { return $false }
        foreach ($d in $dirs) { if ($path.StartsWith($d + '\', 'OrdinalIgnoreCase')) { return $true } }
        return ($_.ProcessName -ieq $short)
    })
}

# This window and the programs hosting it (e.g. Windows Terminal) must not be closed
function Get-ProtectedPids {
    $parents = @{}
    Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        ForEach-Object { $parents[[int]$_.ProcessId] = [int]$_.ParentProcessId }
    $ids = @($PID); $cur = $PID
    for ($i = 0; $i -lt 20 -and $parents.ContainsKey($cur); $i++) {
        $cur = $parents[$cur]
        if ($cur -le 4 -or $ids -contains $cur) { break }
        $ids += $cur
    }
    $ids
}

function Test-Running($p) {
    try { return (-not $p.HasExited) } catch { return ($null -ne (Get-Process -Id $p.Id -ErrorAction SilentlyContinue)) }
}

# Closes the processes (like clicking X, then forced after 20 s). Returns what to reopen
# afterwards (exe path and window state): only each app's top process, not its helpers.
function Close-AppProcesses($procs) {
    $ids = @($procs | ForEach-Object { $_.Id })
    $reopen = @()
    foreach ($p in $procs) {
        $parent = (Get-CimInstance Win32_Process -Filter "ProcessId=$($p.Id)" -ErrorAction SilentlyContinue).ParentProcessId
        $path = $null; try { $path = $p.Path } catch {}
        if (-not $path -or $ids -contains [int]$parent -or @($reopen | Where-Object { $_.Path -eq $path }).Count) { continue }
        # Reopen only what you started (from Start, desktop or startup: parent is explorer) or
        # what had a window. Background helpers are restarted by their own app or service.
        $state  = Get-WindowState $p
        $byUser = (Get-Process -Id $parent -ErrorAction SilentlyContinue).ProcessName -eq 'explorer'
        if ($byUser -or $state -ne 'hidden') { $reopen += [pscustomobject]@{ Path = $path; State = $state } }
    }
    foreach ($p in $procs) {
        Log "  Closing $($p.ProcessName) (PID $($p.Id))" 'DarkGray'
        & taskkill.exe /PID $p.Id *> $null
    }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline -and @($procs | Where-Object { Test-Running $_ }).Count) { Start-Sleep -Milliseconds 500 }
    foreach ($p in $procs) {
        if (Test-Running $p) {
            Log "  Force-closing $($p.ProcessName) (PID $($p.Id))" 'Yellow'
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 2
    , $reopen
}

# Other apps that have this app's files open, e.g. browsers and video-call apps using the
# OBS virtual camera (OBS installer code 6). Only done on a retry, since it can close a browser.
$ExtraDirs = @{ 'OBSProject.OBSStudio' = @("$env:ProgramData\obs-studio-hook") }
$StoppableServices = @('FrameServer', 'FrameServerMonitor')   # Windows camera service, restarts by itself

function Close-FileUsers($pkg, $dirs, $protect) {
    $dirs = @(@($dirs) + @($ExtraDirs["$($pkg.Id)"]) | Where-Object { $_ -and (Test-Path -LiteralPath $_) })
    if (-not $dirs.Count) { return , @() }
    $winDir = $env:windir.TrimEnd('\') + '\'
    $users = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
        if ($_.Id -le 4 -or $protect -contains $_.Id) { return $false }
        $mods = $null; try { $mods = $_.Modules } catch {}
        foreach ($m in $mods) {
            $f = "$($m.FileName)"
            foreach ($d in $dirs) { if ($f.StartsWith($d + '\', 'OrdinalIgnoreCase')) { return $true } }
        }
        return $false
    })
    if (-not $users.Count) { return , @() }
    $apps = @(); $services = @()
    foreach ($u in $users) {
        $path = $null; try { $path = $u.Path } catch {}
        if ($path -and $u.SessionId -ne 0 -and -not $path.StartsWith($winDir, 'OrdinalIgnoreCase')) {
            if ($apps -notcontains $path) { $apps += $path }
        } else {
            $svc  = @(Get-CimInstance Win32_Service -Filter "ProcessId=$($u.Id)" -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            $stop = @($svc | Where-Object { $StoppableServices -contains $_ })
            if ($stop.Count) { $services += $stop }
            else { Log "  $($u.ProcessName) (PID $($u.Id)) uses its files but is a Windows service, left open" 'Yellow' }
        }
    }
    foreach ($s in @($services | Select-Object -Unique)) {
        Log "  Stopping Windows service $s (uses its files)" 'DarkGray'
        Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
    }
    if (-not $apps.Count) { return , @() }
    # Close the whole app (every process of that exe), not only the one holding the file
    $procs = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $pp = $null; try { $pp = $_.Path } catch {}
        $pp -and ($apps -contains $pp) -and $_.SessionId -ne 0 -and ($protect -notcontains $_.Id) })
    if (-not $procs.Count) { return , @() }
    Log "  Other apps using its files: $((@($procs | ForEach-Object { $_.ProcessName }) | Select-Object -Unique) -join ', ')" 'Yellow'
    Close-AppProcesses $procs
}

# Window state before closing: 'normal', 'minimized', or 'hidden' (no visible window,
# e.g. in the system tray)
if (-not ('AppUpdate.Win' -as [type])) {
    Add-Type -Namespace AppUpdate -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
'@
}
function Get-WindowState($p) {
    $h = [IntPtr]::Zero
    try { $p.Refresh(); $h = $p.MainWindowHandle } catch {}
    if ($h -eq [IntPtr]::Zero) { return 'hidden' }
    if ([AppUpdate.Win]::IsIconic($h)) { return 'minimized' }
    'normal'
}

# Waits up to 15 s for the reopened app's window, then minimizes it like clicking its
# minimize button, so apps set to 'minimize to tray' go back to the tray
function Restore-WindowState([string]$path) {
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Milliseconds 500
        $w = Get-Process -ErrorAction SilentlyContinue | Where-Object {
            $pp = $null; try { $pp = $_.Path } catch {}
            $pp -eq $path -and $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
        if ($w) {
            [void][AppUpdate.Win]::PostMessage($w.MainWindowHandle, 0x0112, [IntPtr]0xF020, [IntPtr]::Zero)  # WM_SYSCOMMAND, SC_MINIMIZE
            return
        }
    }
}

# explorer.exe starts the app as the normal user, not as admin like this script
function Open-Apps($items) {
    $seen = @()
    foreach ($it in @($items)) {
        if (-not $it -or -not $it.Path -or $seen -contains $it.Path) { continue }
        $seen += $it.Path
        if (-not (Test-Path -LiteralPath $it.Path)) { continue }
        Log "  Reopening $(Split-Path -Path $it.Path -Leaf)$(if ($it.State -ne 'normal') { " ($($it.State))" })" 'DarkGray'
        Start-Process -FilePath (Join-Path $env:windir 'explorer.exe') -ArgumentList "`"$($it.Path)`""
        if ($it.State -ne 'normal') { Restore-WindowState $it.Path }
    }
}

# =====================================================================
#  UPDATE
# =====================================================================
# Live line under the app being updated, like winget's own bar (window only, not the task):
#   [solid blocks for apps done][shaded for apps left]  11 of 18   /  1:23
# The bar is apps done out of the total; spinner and time are for the current app.
$script:Bar = $null      # @{ Index; Total } of the app being updated
function Get-ConsoleWidth { try { [Math]::Max(40, [Console]::WindowWidth - 1) } catch { 79 } }
function Write-LiveLine([int]$Frame, [TimeSpan]$Elapsed) {
    $width  = 30
    $done   = [int][Math]::Floor(($script:Bar.Index - 1) * $width / [Math]::Max(1, $script:Bar.Total))
    $frames = '|', '/', '-', '\'
    $tail   = '  {0} of {1}   {2}  {3}:{4:00}' -f $script:Bar.Index, $script:Bar.Total, $frames[$Frame % 4],
                [int][Math]::Floor($Elapsed.TotalMinutes), $Elapsed.Seconds
    $pad    = [Math]::Max(0, (Get-ConsoleWidth) - 2 - $width - $tail.Length)
    Write-Host "`r  " -NoNewline
    Write-Host ([string][char]0x2588 * $done) -ForegroundColor Cyan -NoNewline
    Write-Host ([string][char]0x2591 * ($width - $done)) -ForegroundColor DarkGray -NoNewline
    Write-Host ($tail + (' ' * $pad)) -NoNewline
}
function Clear-LiveLine { Write-Host ("`r" + (' ' * (Get-ConsoleWidth)) + "`r") -NoNewline }

# Waits until $IsDone returns true, showing the live line meanwhile
function Wait-WithSpinner([scriptblock]$IsDone) {
    $live = (-not $RunMode) -and $script:Bar
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $n = 0
    try {
        while (-not (& $IsDone)) {
            if ($live) { Write-LiveLine $n $sw.Elapsed }
            $n++
            Start-Sleep -Milliseconds 250
        }
    } finally {
        if ($live) { Clear-LiveLine }
    }
}

# Runs Update-WinGetPackage in the background so the spinner can move meanwhile.
# Always returns @{ Status; Reboot; Code; Error } and never throws.
$UpdateWork = {
    param($p)
    $ProgressPreference = 'SilentlyContinue'   # winget's bar shows full during install: hide it
    try {
        Import-Module Microsoft.WinGet.Client -ErrorAction Stop
        $r = Update-WinGetPackage @p -ErrorAction Stop
        @{ Status = "$($r.Status)"; Reboot = [bool]$r.RebootRequired; Code = "$($r.InstallerErrorCode)"; Error = '' }
    } catch {
        @{ Status = 'Error'; Reboot = $false; Code = ''; Error = $_.Exception.Message }
    }
}
function Invoke-WinGetUpdate($params) {
    $ps = $null; $handle = $null
    try {
        $iss = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch {}
        $ps = [PowerShell]::Create($iss)
        [void]$ps.AddScript($UpdateWork.ToString()).AddArgument($params)
        $handle = $ps.BeginInvoke()
        Wait-WithSpinner ({ $handle.IsCompleted }.GetNewClosure())
        $out = $ps.EndInvoke($handle)
        if ($out.Count) {
            # Note: .BaseObject on a wrapped hashtable is read as a key (null), so use -is / cast
            $o = $out[$out.Count - 1]
            if ($o -is [hashtable]) { return [hashtable]$o }
        }
        $msg = if ($ps.Streams.Error.Count) { "$($ps.Streams.Error[0])" } else { 'no result from winget' }
        return @{ Status = 'Error'; Reboot = $false; Code = ''; Error = $msg }
    } catch {
        return @{ Status = 'Error'; Reboot = $false; Code = ''; Error = $_.Exception.Message }
    } finally {
        # If stopped with Ctrl+C mid-install, let the installer finish instead of cutting it off
        if ($ps -and (-not $handle -or $handle.IsCompleted)) { $ps.Dispose() }
    }
}

# Returns 'ok', 'failed' or 'skipped'
function Update-OnePackage($pkg, $protect) {
    $to = @($pkg.AvailableVersions)[0]
    Log "$($pkg.Name)  $($pkg.InstalledVersion) -> $to" 'White'
    $dirs   = Get-AppDirs $pkg
    $reopen = @()
    try {
        for ($try = 1; $try -le 3; $try++) {
            # Close it before every attempt: some apps get restarted by a helper or tray icon
            $procs = Get-AppProcesses $pkg $dirs
            if (@($procs | Where-Object { $protect -contains $_.Id }).Count) {
                Log '  Skipped: this window runs inside it. The weekly task will update it.' 'Yellow'
                return 'skipped'
            }
            if ($procs.Count) { $reopen += Close-AppProcesses $procs }
            if ($try -ge 2) { $reopen += Close-FileUsers $pkg $dirs $protect }

            $status = ''; $detail = ''; $reboot = $false
            if ($try -lt 3) {
                $params = @{ Id = $pkg.Id; Source = $pkg.Source; MatchOption = 'Equals'; Mode = 'Silent' }
                if ($try -eq 2) { $params.Force = $true }
                $r = Invoke-WinGetUpdate $params
                $status = "$($r.Status)"; $reboot = [bool]$r.Reboot
                $detail = if ($r.Error) { $r.Error } else { "installer code $($r.Code)" }
            } else {
                # Last try: winget.exe itself, in case the module is the problem
                $code = $null
                try {
                    $proc = Start-Process -FilePath 'winget.exe' -WindowStyle Hidden -PassThru -ArgumentList @(
                        'upgrade', '--id', $pkg.Id, '--exact', '--source', $pkg.Source, '--silent', '--force',
                        '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
                    $null = $proc.Handle      # keeps the exit code readable after it ends
                    Wait-WithSpinner ({ $proc.HasExited }.GetNewClosure())
                    $proc.WaitForExit()
                    $code = $proc.ExitCode
                } catch { $detail = $_.Exception.Message }
                $status = if ($code -eq 0) { 'Ok' } else { 'Error' }
                if ($null -ne $code) { $detail = 'winget exit code 0x{0:X8}' -f $code }
            }

            if ($status -eq 'Ok') {
                $now = Get-WinGetPackage -Id $pkg.Id -MatchOption Equals -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($now -and $now.IsUpdateAvailable -and "$($now.InstalledVersion)" -eq "$($pkg.InstalledVersion)") {
                    $status = 'NotChanged'; $detail = "still $($now.InstalledVersion) after install"
                } else {
                    Log "  Updated$(if ($reboot) { ' (restart needed)' })" 'Green'
                    if ($reboot) { $script:RebootNeeded = $true }
                    return 'ok'
                }
            }
            # winget lists it as installed but cannot find it to update (a winget bug, seen with
            # zip-based apps like FFmpeg): retrying does not help, a one-time reinstall does
            if ($detail -match 'No (installed )?packages? (matched|found)' -or $detail -match '8A150014') {
                Log "  Skipped: winget cannot find the installed copy to update. Fix once in PowerShell:" 'Yellow'
                Log "    winget uninstall --id $($pkg.Id) --exact; winget install --id $($pkg.Id) --exact" 'Yellow'
                return 'skipped'
            }
            if ($status -eq 'NoApplicableUpgrade') {
                Log '  Skipped: winget has no update that fits this install.' 'Yellow'
                return 'skipped'
            }
            # Trust what is installed over what was reported: if the update is in, it worked
            # (e.g. winget.exe 0x8A15002B 'no applicable update' after an earlier attempt installed it)
            $now = Get-WinGetPackage -Id $pkg.Id -MatchOption Equals -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($now -and -not $now.IsUpdateAvailable -and "$($now.InstalledVersion)" -ne "$($pkg.InstalledVersion)") {
                Log "  Updated (now $($now.InstalledVersion))" 'Green'
                return 'ok'
            }
            Log "  Attempt $try failed: $status, $detail" 'Yellow'
            # Often another installer (Windows Update, msiexec) was busy: wait before retrying
            if ($try -lt 3) { Start-Sleep -Seconds 60 }
        }
        Log '  FAILED after 3 attempts' 'Red'
        return 'failed'
    } finally {
        if ($reopen.Count) { Open-Apps $reopen }
    }
}

function Save-Last([string]$text) {
    try { Set-Content -LiteralPath $LastFile -Value ("{0}, {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $text) } catch {}
    if ($RunMode) { Show-Toast 'Apps updated' $text }
}

# Windows notification (shown as from Windows PowerShell); only used by the weekly task
function Show-Toast([string]$Title, [string]$Text) {
    try {
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
        $esc = [Security.SecurityElement]
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml("<toast><visual><binding template='ToastGeneric'><text>$($esc::Escape($Title))</text>" +
            "<text>$($esc::Escape($Text))</text><text>Details: App Update, option 3 (log)</text></binding></visual></toast>")
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show(
            [Windows.UI.Notifications.ToastNotification]::new($xml))
    } catch { Log "Notification failed: $($_.Exception.Message)" 'DarkGray' }
}

function Invoke-Update {
    # One run at a time (weekly task and 'Update now')
    $mutex = New-Object System.Threading.Mutex($false, 'Global\appupdate')
    $got = $false
    try { $got = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $got = $true }
    if (-not $got) { Log 'An update is already running.' 'Yellow'; return }
    try {
        # Keep the log small
        if (Test-Path -LiteralPath $LogFile) {
            $lines = @(Get-Content -LiteralPath $LogFile)
            if ($lines.Count -gt 5000) { $lines[-3000..-1] | Set-Content -LiteralPath $LogFile }
        }
        Log "==== Update started ($(if ($RunMode) { 'weekly task' } else { 'manual' })) ====" 'Cyan'
        if (-not (Wait-Network)) { Log 'No internet. Nothing updated.' 'Red'; Save-Last 'no internet'; return }
        Initialize-WinGet
        $script:RebootNeeded = $false
        $pins    = Get-PinLines
        $protect = Get-ProtectedPids
        $pkgs = @(Get-WinGetPackage -ErrorAction Stop | Where-Object {
            $_.IsUpdateAvailable -and $_.Source -eq 'winget' -and $_.Id -ne 'Microsoft.AppInstaller' })
        $n = @{ ok = 0; failed = 0; skipped = 0 }
        $i = 0
        foreach ($p in $pkgs) {
            $i++
            $script:Bar = @{ Index = $i; Total = $pkgs.Count }
            if (Test-Pinned $p.Id $pins) { Log "$($p.Name): pinned, skipped" 'DarkGray'; $n.skipped++; continue }
            $n[(Update-OnePackage $p $protect)]++
        }
        $script:Bar = $null
        $summary = if ($pkgs.Count) { "$($n.ok) updated, $($n.failed) failed, $($n.skipped) skipped" } else { 'all up to date' }
        if ($script:RebootNeeded) { $summary += ', restart needed' }
        Log "==== Done: $summary ====" $(if ($n.failed) { 'Red' } else { 'Green' })
        Save-Last $summary
    } catch {
        Log "Error: $($_.Exception.Message)" 'Red'
        Save-Last 'error, see log'
    } finally {
        $mutex.ReleaseMutex(); $mutex.Dispose()
    }
}

# =====================================================================
#  WEEKLY TASK
# =====================================================================
function Get-TaskState {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) { "on, $($RunDay.Substring(0,3)) $RunAt" } else { 'off' }
}

function Enable-WeeklyTask {
    New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null
    # Save the copy the task runs (from GitHub when started with irm | iex)
    if ($PSCommandPath -and $PSCommandPath -ine $LocalCopy) {
        Copy-Item -LiteralPath $PSCommandPath -Destination $LocalCopy -Force
    } elseif (-not $PSCommandPath) {
        $code = Invoke-RestMethod -Uri "${ScriptUrl}?$(Get-Random)" -UseBasicParsing
        Set-Content -LiteralPath $LocalCopy -Value $code -Encoding UTF8
    }
    Initialize-WinGet
    $user      = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' `
                    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$LocalCopy`" -Run"
    $trigger   = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek $RunDay -At $RunAt
    # Runs as you (so it sees your apps and per-user installs), as admin, without a prompt
    $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -AllowStartIfOnBatteries `
                    -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 3) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal `
        -Settings $settings -Description 'Updates all apps with winget (therepos/windows appupdate.ps1)' -Force | Out-Null
    $next = (Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue).NextRunTime
    Say "  Weekly update on. Next run: $(if ($next) { $next.ToString('ddd d MMM yyyy HH:mm') } else { "$RunDay $RunAt" })" 'Green'
}

function Disable-WeeklyTask {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Say '  Weekly update off.' 'Green'
}

# =====================================================================
#  SHORTCUT
# =====================================================================
# >>> shortcut
function Get-ShortcutName { [IO.Path]::GetFileNameWithoutExtension(([uri]$ScriptUrl).AbsolutePath) }
function Get-ShortcutCopy { $n = Get-ShortcutName; Join-Path (Join-Path $env:LOCALAPPDATA $n) "$n.ps1" }
function Get-ShortcutLine { 'function {0} {{ & "$env:LOCALAPPDATA\{0}\{0}.ps1" }}' -f (Get-ShortcutName) }
function Test-Shortcut {
    (Test-Path -LiteralPath $PROFILE) -and (Test-Path -LiteralPath (Get-ShortcutCopy)) -and
        (@(Get-Content -LiteralPath $PROFILE -ErrorAction SilentlyContinue) -contains (Get-ShortcutLine))
}
function Switch-Shortcut([switch]$KeepCopy) {
    $name = Get-ShortcutName
    $copy = Get-ShortcutCopy
    try {
        $on = Test-Shortcut
        # Remove any shortcut of this name, including older ones that downloaded the script each time
        if (Test-Path -LiteralPath $PROFILE) {
            @(Get-Content -LiteralPath $PROFILE) | Where-Object { $_ -notmatch "^function $name " } | Set-Content -LiteralPath $PROFILE
        }
        if ($on) {
            if (-not $KeepCopy) { Remove-Item -LiteralPath $copy -Force -ErrorAction SilentlyContinue }
            Say "  Shortcut removed." 'Green'
            return
        }
        New-Item -ItemType Directory -Path (Split-Path -Path $copy -Parent) -Force | Out-Null
        if ($PSCommandPath -and (Test-Path -LiteralPath $PSCommandPath)) {
            if ($PSCommandPath -ne $copy) { Copy-Item -LiteralPath $PSCommandPath -Destination $copy -Force }
        } else {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri "${ScriptUrl}?$(Get-Random)" -OutFile $copy -UseBasicParsing
        }
        Unblock-File -LiteralPath $copy -ErrorAction SilentlyContinue
        if ((Get-ExecutionPolicy -Scope CurrentUser) -in 'Undefined', 'Restricted', 'AllSigned') {
            Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
        }
        if (-not (Test-Path -LiteralPath $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }
        Add-Content -LiteralPath $PROFILE "`n$(Get-ShortcutLine)"
        Say "  Shortcut added. In a new PowerShell window, just type: $name" 'Green'
    } catch {
        Say "  Could not change the shortcut: $($_.Exception.Message)" 'Red'
    }
}
# Settings submenu for scripts whose only setting is the shortcut; returns 'back' or 'quit'
function Show-ShortcutSettings([switch]$KeepCopy) {
    while ($true) {
        $short = if (Test-Shortcut) { 'on' } else { 'off' }
        Say ""
        Say "  ---------------- SETTINGS -------------------" 'Cyan'
        Say ""
        Say ("    1  -  Shortcut '{0}'   [{1}]" -f (Get-ShortcutName), $short) 'White'
        Say "    2  -  Back" 'White'
        Say "    0  -  Quit" 'White'
        Say ""
        switch (Read-MenuKey 'Choose 0-2' '012') {
            '1' { Switch-Shortcut -KeepCopy:$KeepCopy }
            '2' { return 'back' }
            '0' { return 'quit' }
        }
    }
}
# <<< shortcut

# =====================================================================
#  MOVE FROM THE OLD NAME (autoupdate): weekly task, shortcut, log
# =====================================================================
function Move-FromOldName {
    try {
        $oldDir      = Join-Path $env:LOCALAPPDATA 'autoupdate'
        $oldTask     = Get-ScheduledTask -TaskName 'AutoUpdate Apps' -ErrorAction SilentlyContinue
        $oldShortcut = (Test-Path -LiteralPath $PROFILE) -and
            @(@(Get-Content -LiteralPath $PROFILE -ErrorAction SilentlyContinue) -match '^function autoupdate ').Count
        if (-not ($oldTask -or $oldShortcut -or (Test-Path -LiteralPath $oldDir))) { return }
        Say '  Moving from the old name (autoupdate)...' 'DarkGray'
        New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null
        foreach ($f in 'update.log', 'last.txt') {
            $src = Join-Path $oldDir $f
            if ((Test-Path -LiteralPath $src) -and -not (Test-Path -LiteralPath (Join-Path $ToolDir $f))) {
                Move-Item -LiteralPath $src -Destination $ToolDir
            }
        }
        if ($oldTask) {
            Unregister-ScheduledTask -TaskName 'AutoUpdate Apps' -Confirm:$false
            Enable-WeeklyTask
        }
        if ($oldShortcut) {
            @(Get-Content -LiteralPath $PROFILE) | Where-Object { $_ -notmatch '^function autoupdate ' } | Set-Content -LiteralPath $PROFILE
            if (-not (Test-Shortcut)) { Switch-Shortcut }
        }
        Remove-Item -LiteralPath $oldDir -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Say "  Could not move everything from the old name: $($_.Exception.Message)" 'Yellow'
    }
}

# =====================================================================
#  MAIN
# =====================================================================
if ($RunMode) {
    if (-not $IsAdmin) { Log 'Error: the weekly task is not running as admin. Turn the weekly update off and on again.' 'Red'; exit 1 }
    Invoke-Update
    exit 0
}

# >>> layout
function Set-WindowRight {
    try {
        if (-not ('WinLayout.Api' -as [type])) {
            Add-Type -Namespace WinLayout -Name Api -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
[DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
[DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr h, uint f);
[DllImport("user32.dll")] public static extern bool GetMonitorInfo(IntPtr m, ref MONITORINFO mi);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
[DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
[DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int a, out RECT r, int s);
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
[StructLayout(LayoutKind.Sequential)] public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }
'@
        }
        # Classic console: its own window. Windows Terminal: the hidden console window is owned
        # by the Terminal window hosting this tab (it may still be opening: wait up to 3 s)
        $h = [IntPtr]::Zero
        # Owner first: Terminal's hidden console window still reports itself as visible
        for ($i = 0; $i -lt 15 -and $h -eq [IntPtr]::Zero; $i++) {
            $c = [WinLayout.Api]::GetConsoleWindow()
            $o = if ($c -ne [IntPtr]::Zero) { [WinLayout.Api]::GetAncestor($c, 3) } else { [IntPtr]::Zero }
            if ($o -ne [IntPtr]::Zero -and $o -ne $c -and [WinLayout.Api]::IsWindowVisible($o)) { $h = $o; break }
            if ($c -ne [IntPtr]::Zero -and $o -eq $c -and [WinLayout.Api]::IsWindowVisible($c)) { $h = $c; break }
            Start-Sleep -Milliseconds 200
        }
        if ($h -eq [IntPtr]::Zero) { return }
        $owner = [uint32]0
        [void][WinLayout.Api]::GetWindowThreadProcessId($h, [ref]$owner)
        if ((Get-Process -Id $owner -ErrorAction SilentlyContinue).ProcessName -notin
            'WindowsTerminal', 'conhost', 'OpenConsole', 'powershell', 'pwsh') { return }
        $oldDpi = [WinLayout.Api]::SetThreadDpiAwarenessContext([IntPtr](-4))
        try {
            [void][WinLayout.Api]::ShowWindow($h, 9)
            $mi = New-Object WinLayout.Api+MONITORINFO
            $mi.cbSize = 40
            [void][WinLayout.Api]::GetMonitorInfo([WinLayout.Api]::MonitorFromWindow($h, 2), [ref]$mi)
            $w = $mi.rcWork
            $half = [int][Math]::Floor(($w.Right - $w.Left) / 2)
            # Grow by the invisible resize borders so the visible frame fills the half exactly
            $wr = New-Object WinLayout.Api+RECT
            $fr = New-Object WinLayout.Api+RECT
            [void][WinLayout.Api]::GetWindowRect($h, [ref]$wr)
            $l = 0; $r = 0; $b = 0
            if ([WinLayout.Api]::DwmGetWindowAttribute($h, 9, [ref]$fr, 16) -eq 0) {
                $l = $fr.Left - $wr.Left; $r = $wr.Right - $fr.Right; $b = $wr.Bottom - $fr.Bottom
            }
            [void][WinLayout.Api]::SetWindowPos($h, [IntPtr]::Zero, ($w.Left + $half - $l), $w.Top,
                ($w.Right - $w.Left - $half + $l + $r), ($w.Bottom - $w.Top + $b), 0x0014)
        } finally {
            if ($oldDpi -ne [IntPtr]::Zero) { [void][WinLayout.Api]::SetThreadDpiAwarenessContext($oldDpi) }
        }
    } catch {}
}
# <<< layout
Set-WindowRight
Clear-Host
Move-FromOldName
$last = ''
while ($true) {
    $lastRun = if (Test-Path -LiteralPath $LastFile) { (Get-Content -LiteralPath $LastFile -TotalCount 1) } else { 'never' }
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        APP UPDATE" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  Weekly update   [$(Get-TaskState)]" 'White'
    Say "    2  -  Update now" 'White'
    Say "    3  -  Open log        [last: $lastRun]" 'White'
    Say "    4  -  Settings" 'White'
    Say "    0  -  Quit" 'White'
    Say ""
    $choice = Read-MenuKey 'Choose 0-4' '01234' $last
    if ($choice -eq '0') { break }
    if ($choice -eq '4') {
        if (@(Show-ShortcutSettings -KeepCopy:((Get-TaskState) -ne 'off')) -contains 'quit') { break }
        Clear-Host
        continue
    }
    Say ""
    try {
        switch ($choice) {
            '1' { if ((Get-TaskState) -eq 'off') { Enable-WeeklyTask } else { Disable-WeeklyTask } }
            '2' { Invoke-Update }
            '3' { if (Test-Path -LiteralPath $LogFile) { Start-Process notepad.exe -ArgumentList "`"$LogFile`"" } else { Say '  No log yet.' 'Yellow' } }
        }
    } catch {
        Say "  Error: $($_.Exception.Message)" 'Red'
    }
    if ($choice -ne '3') { $last = $choice }
}
# Quit closes this window, not just the menu
[Environment]::Exit(0)
