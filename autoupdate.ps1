# irm https://github.com/therepos/windows/raw/main/autoupdate.ps1 | iex
# Purpose: Update all apps with winget every Sunday 23:45, closing open apps first and reopening them after
# =============================================================================
#  Auto Update:
#   1 = Weekly update  on/off. A scheduled task runs Sundays 23:45 (or as soon as the
#                      PC is on again if it was off or asleep then)
#   2 = Update now     same update, in this window
#   3 = Open log       %LOCALAPPDATA%\autoupdate\update.log
#   0 = Quit           closes this window
#  An app that is open is closed first (gracefully, forced after 20 s), updated, then
#  reopened minimized if it was minimized or in the tray. A failed update is retried twice (second retry with winget.exe --force);
#  retries also close other apps that have its files open (e.g. a browser using the OBS
#  virtual camera) and reopen them after.
#  'Update now' shows a progress bar (apps done of total, spinner and time for the current app); the weekly task runs hidden and shows a Windows
#  notification with the result when it is done.
#  Skipped: apps pinned with 'winget pin add <id>', Microsoft Store apps (the Store
#  updates those itself).
#  Needs admin (asks). Installs the Microsoft.WinGet.Client module if missing.
#  The task runs a local copy in %LOCALAPPDATA%\autoupdate (not fetched each week);
#  turning the weekly update on again refreshes that copy.
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/autoupdate.ps1'
$ToolDir   = Join-Path $env:LOCALAPPDATA 'autoupdate'
$LocalCopy = Join-Path $ToolDir 'autoupdate.ps1'
$LogFile   = Join-Path $ToolDir 'update.log'
$LastFile  = Join-Path $ToolDir 'last.txt'
$TaskName  = 'AutoUpdate Apps'
$RunDay    = 'Sunday'
$RunAt     = '23:45'
$RunMode   = ($args -contains '-Run')      # set by the scheduled task

$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# --- Self-elevate: re-run the same one-liner in an elevated window ------------
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin -and -not $RunMode) {
    Write-Host "  Requesting administrator rights..." -ForegroundColor Yellow
    $cmd = "[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; try { irm '${ScriptUrl}?$(Get-Random)' | iex } catch { Write-Host `$_ -ForegroundColor Red; Read-Host 'Press Enter to exit' }"
    Start-Process -FilePath 'powershell.exe' -Verb RunAs `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $cmd
    Write-Host "  Continue in the new (elevated) window. You can close this one." -ForegroundColor Gray
    return
}

function Say($msg, $color = 'Gray') { if (-not $RunMode) { Write-Host $msg -ForegroundColor $color } }

# Shown in the window (when there is one) and appended to update.log
function Log($msg, $color = 'Gray') {
    Say "  $msg" $color
    try {
        if (-not (Test-Path -LiteralPath $ToolDir)) { New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null }
        Add-Content -LiteralPath $LogFile -Value ('{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)
    } catch {}
}

# One keypress picks a menu option (no Enter needed); Enter picks the default, if any
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
    $portable = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    Get-ChildItem -LiteralPath $portable -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name.StartsWith("$($pkg.Id)_", 'OrdinalIgnoreCase') } |
        ForEach-Object { if (-not $dirs.Contains($_.FullName)) { $dirs.Add($_.FullName) } }
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
    Show-Bar
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
if (-not ('AutoUpdate.Win' -as [type])) {
    Add-Type -Namespace AutoUpdate -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
'@
}
function Get-WindowState($p) {
    $h = [IntPtr]::Zero
    try { $p.Refresh(); $h = $p.MainWindowHandle } catch {}
    if ($h -eq [IntPtr]::Zero) { return 'hidden' }
    if ([AutoUpdate.Win]::IsIconic($h)) { return 'minimized' }
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
            [void][AutoUpdate.Win]::PostMessage($w.MainWindowHandle, 0x0112, [IntPtr]0xF020, [IntPtr]::Zero)  # WM_SYSCOMMAND, SC_MINIMIZE
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
# Waits until $IsDone returns true. In the window it shows a spinner and the time taken
# on the progress bar, e.g. 'Downloading / installing  /  1:23'
function Wait-WithSpinner([scriptblock]$IsDone) {
    $frames = '|', '/', '-', '\'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $n = 0
    while (-not (& $IsDone)) {
        if (-not $RunMode -and $script:Bar) {
            $b = $script:Bar.Clone()
            $b.CurrentOperation = 'Downloading / installing  {0}  {1}:{2:00}' -f $frames[$n % 4],
                [int][Math]::Floor($sw.Elapsed.TotalMinutes), $sw.Elapsed.Seconds
            Write-Progress @b
        }
        $n++
        Start-Sleep -Milliseconds 250
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
            $o = $out[$out.Count - 1].BaseObject
            if ($o -is [hashtable]) { return $o }
        }
        $msg = if ($ps.Streams.Error.Count) { "$($ps.Streams.Error[0])" } else { 'no result from winget' }
        return @{ Status = 'Error'; Reboot = $false; Code = ''; Error = $msg }
    } catch {
        return @{ Status = 'Error'; Reboot = $false; Code = ''; Error = $_.Exception.Message }
    } finally {
        # If stopped with Ctrl+C mid-install, let the installer finish instead of cutting it off
        if ($ps -and (-not $handle -or $handle.IsCompleted)) { $ps.Dispose() }
        Show-Bar
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
            if ($status -eq 'NoApplicableUpgrade') {
                Log '  Skipped: winget has no update that fits this install.' 'Yellow'
                return 'skipped'
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

# The console hides the progress bar while a program like taskkill.exe or winget.exe runs,
# and only shows it again on the next Write-Progress: call this after running one
$script:Bar = $null
function Show-Bar { if ($script:Bar) { Write-Progress @script:Bar } }

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
            "<text>$($esc::Escape($Text))</text><text>Details: Auto Update, option 3 (log)</text></binding></visual></toast>")
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show(
            [Windows.UI.Notifications.ToastNotification]::new($xml))
    } catch { Log "Notification failed: $($_.Exception.Message)" 'DarkGray' }
}

function Invoke-Update {
    # One run at a time (weekly task and 'Update now')
    $mutex = New-Object System.Threading.Mutex($false, 'Global\autoupdate-apps')
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
        # Progress bar in the window ('Update now'): apps done out of the total
        if (-not $RunMode) { $ProgressPreference = 'Continue' }
        $i = 0
        foreach ($p in $pkgs) {
            $i++
            $script:Bar = @{ Id = 1; Activity = "Updating apps ($i of $($pkgs.Count))"; Status = "$($p.Name)"
                             PercentComplete = [int](($i - 1) * 100 / $pkgs.Count) }
            Show-Bar
            if (Test-Pinned $p.Id $pins) { Log "$($p.Name): pinned, skipped" 'DarkGray'; $n.skipped++; continue }
            $n[(Update-OnePackage $p $protect)]++
        }
        Write-Progress -Id 1 -Activity 'Updating apps' -Completed
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
        -Settings $settings -Description 'Updates all apps with winget (therepos/windows autoupdate.ps1)' -Force | Out-Null
    $next = (Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue).NextRunTime
    Say "  Weekly update on. Next run: $(if ($next) { $next.ToString('ddd d MMM yyyy HH:mm') } else { "$RunDay $RunAt" })" 'Green'
}

function Disable-WeeklyTask {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Say '  Weekly update off.' 'Green'
}

# =====================================================================
#  MAIN
# =====================================================================
if ($RunMode) {
    if (-not $IsAdmin) { Log 'Error: the weekly task is not running as admin. Turn the weekly update off and on again.' 'Red'; exit 1 }
    Invoke-Update
    exit 0
}

Clear-Host
$last = ''
while ($true) {
    $lastRun = if (Test-Path -LiteralPath $LastFile) { (Get-Content -LiteralPath $LastFile -TotalCount 1) } else { 'never' }
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        AUTO UPDATE" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  Weekly update   [$(Get-TaskState)]" 'White'
    Say "    2  -  Update now" 'White'
    Say "    3  -  Open log        [last: $lastRun]" 'White'
    Say "    0  -  Quit" 'White'
    Say ""
    $choice = Read-MenuKey 'Choose 0-3' '0123' $last
    if ($choice -eq '0') { break }
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
