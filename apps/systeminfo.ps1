# irm "https://github.com/therepos/windows/raw/main/apps/systeminfo.ps1?$(Get-Random)" | iex
# Purpose: Write a health report of a Windows laptop or Android phone, or find what makes a PC laggy, to a text file
# =============================================================================
# Usage:
#   1 = Windows laptop
#   2 = Android phone   USB cable, USB debugging on
#   3 = Both
#   4 = Lag check       3-5 minutes; samples CPU, memory, disk and processes, times real actions
#   5 = Settings        'systeminfo' shortcut (runs a local copy in %LOCALAPPDATA%\systeminfo)
#   0 = Quit
#
# Note:
#   Needs admin (asks once)
#   Output: Data System|Phone <model>, Diag Lag <model> yyyyMMdd-HHmm.txt (folder asked, default Downloads)
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/systeminfo.ps1'

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
if (-not $IsAdmin) { Start-Elevated $ScriptUrl; return }

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

$OutDir = $null
$Stamp = Get-Date -Format 'yyyyMMdd-HHmm'

# Model as one word for file names: spaces -> dashes, characters Windows forbids removed
function ConvertTo-OneWord([string]$s) {
    $s = ($s -replace '[\\/:*?"<>|]', '').Trim() -replace '\s+', '-'
    if ($s) { $s } else { 'Unknown' }
}
# Readable model name; Lenovo reports a part code (e.g. 21CB0020US) as Model and the name in Version
function Get-LaptopModelName {
    $cs = Get-CimInstance Win32_ComputerSystem
    $model = $cs.Model
    if ($cs.Manufacturer -match 'LENOVO') {
        $friendly = (Get-CimInstance Win32_ComputerSystemProduct).Version
        if ($friendly) { $model = $friendly }
    }
    "$model".Trim()
}

# >>> say
function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }
# <<< say

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

# >>> downloads
function Get-DownloadsFolder {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $raw = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'{374DE290-123F-4565-9164-39C4925E467B}'
    if ($raw) {
        $path = [Environment]::ExpandEnvironmentVariables($raw)
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return (Join-Path $env:USERPROFILE 'Downloads')
}
# <<< downloads

function Select-OutputFolder {
    $default = Get-DownloadsFolder
    while ($true) {
        Say ""
        Say "  Where should the report be saved?" 'Cyan'
        $answer = Read-Host "  Folder [$default]"
        $path = if ([string]::IsNullOrWhiteSpace($answer)) { $default } else { [Environment]::ExpandEnvironmentVariables($answer.Trim().Trim('"')) }
        New-Item -ItemType Directory -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
        if (Test-Path -LiteralPath $path -PathType Container) {
            return (Resolve-Path -LiteralPath $path).ProviderPath
        }
        Say "  Cannot use '$path'. Try another folder." 'Red'
    }
}

# =====================================================================
#  REPORT ENGINE
# =====================================================================
$script:Out     = New-Object System.Collections.Generic.List[string]
$script:OutPath = $null
$script:Rs      = $null

# Loaded once into the runspace that runs the sections
$RsInit = @'
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'
$T = "$([char]0x251C)$([char]0x2500)$([char]0x2500) "
$L = "$([char]0x2514)$([char]0x2500)$([char]0x2500) "
function Sub([string]$t) { ''; "---- $t ----" }
function Row([string]$k, $v) { '{0,-24} {1}' -f $k, $v }
function Tree([string[]]$rows) { for ($i = 0; $i -lt $rows.Count; $i++) { $(if ($i -eq $rows.Count - 1) { $L } else { $T }) + $rows[$i] } }
function OrNone($v, [string]$text = '(none)') { if (@($v | Where-Object { $null -ne $_ -and "$_" -ne '' }).Count) { $v } else { $text } }
function Short([string]$s, [int]$n = 160) {
    if (-not $s) { return '' }
    $s = ($s -replace '\s+', ' ').Trim()
    if ($s.Length -gt $n) { $s.Substring(0, $n) + '...' } else { $s }
}
function Norm([string]$s) { $s -replace '\{[0-9A-Fa-f-]{36}\}', '{GUID}' -replace '0x[0-9A-Fa-f]+', '0xN' -replace '\b\d+\b', 'N' }
function Chars($a) { -join ($a | Where-Object { $_ } | ForEach-Object { [char]$_ }) }

# --- Windows ---
$DevCodes = @{ 1 = 'not configured'; 3 = 'driver corrupt'; 10 = 'cannot start'; 12 = 'not enough resources'; 14 = 'needs restart'
    18 = 'reinstall driver'; 19 = 'registry problem'; 22 = 'disabled'; 24 = 'not present'; 28 = 'no driver'; 31 = 'driver failed to load'
    32 = 'driver disabled'; 37 = 'driver failed to start'; 39 = 'driver missing'; 43 = 'stopped after a problem'; 48 = 'driver blocked'; 52 = 'unsigned driver' }
function Get-ProblemDevices { @(Get-CimInstance Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0') }
function Get-FixedVolumes { @(Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq 'Fixed' -and $_.Size } | Sort-Object DriveLetter) }
function Get-PendingReboot {
    $r = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $r += 'component servicing' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $r += 'Windows Update' }
    $r -join ', '
}
function Get-LastUpdate { @(Get-HotFix | Where-Object InstalledOn | Sort-Object InstalledOn -Descending) | Select-Object -First 1 }
function Get-SecureBoot { try { if (Confirm-SecureBootUEFI -ErrorAction Stop) { 'on' } else { 'off' } } catch { 'not supported (legacy BIOS)' } }
# powercfg's battery report, read once per run
function Get-Battery {
    if (-not $global:BatDone) {
        $global:BatDone = $true
        $global:Bat = @()
        $x = Join-Path $env:TEMP "systeminfo-battery-$PID.xml"
        powercfg /batteryreport /xml /output $x | Out-Null
        if (Test-Path -LiteralPath $x) {
            $global:BatXml = [xml](Get-Content -LiteralPath $x -Raw)
            Remove-Item -LiteralPath $x -Force
            $global:Bat = @(foreach ($b in @($global:BatXml.BatteryReport.Batteries.Battery)) {
                if (-not $b) { continue }
                $design = [double]$b.DesignCapacity
                $full   = [double]$b.FullChargeCapacity
                [pscustomobject]@{
                    Id           = $b.Id
                    Manufacturer = $b.Manufacturer
                    Chemistry    = $b.Chemistry
                    Design_mWh   = $design
                    Full_mWh     = $full
                    HealthPct    = if ($design -gt 0) { [math]::Round($full / $design * 100, 1) } else { $null }
                    Cycles       = $b.CycleCount
                }
            })
        }
    }
    $global:Bat
}

# --- Android (adb) ---
$BatHealth = @{ '1' = 'unknown'; '2' = 'good'; '3' = 'overheat'; '4' = 'dead'; '5' = 'over voltage'; '6' = 'failure'; '7' = 'cold' }
$BatStatus = @{ '1' = 'unknown'; '2' = 'charging'; '3' = 'discharging'; '4' = 'not charging'; '5' = 'full' }
function A { & $Adb -s $Serial @args 2>$null | ForEach-Object { "$_".TrimEnd() } }
function Get-Props {
    if (-not $global:Props) {
        $h = @{}
        foreach ($l in (A shell getprop)) { if ($l -match '^\[(.+?)\]: \[(.*)\]$') { $h[$matches[1]] = $matches[2] } }
        $global:Props = $h
    }
    $global:Props
}
function Get-PhoneBattery {
    $h = @{}
    foreach ($l in (A shell dumpsys battery)) { if ($l -match '^\s+([^:]+):\s*(.*)$') { $h[$matches[1].Trim()] = $matches[2].Trim() } }
    $h
}
# Cycle count and capacity, where the phone exposes them without root
function Get-PhoneCapacity {
    $r = @{}
    foreach ($n in 'cycle_count', 'charge_full', 'charge_full_design') {
        $x = "$(A shell cat "/sys/class/power_supply/battery/$n")".Trim()
        if ($x -match '^\d+$') { $r[$n] = [double]$x }
    }
    if ($r['charge_full'] -and $r['charge_full_design']) { $r['pct'] = [math]::Round($r['charge_full'] / $r['charge_full_design'] * 100) }
    $r
}
function Get-DataUse {
    $f = -split (@(A shell df -k /data) | Select-Object -Last 1)
    if ($f.Count -ge 4 -and $f[1] -match '^\d+$' -and [double]$f[1] -gt 0) {
        [pscustomobject]@{ TotalGB = [math]::Round([double]$f[1] / 1MB, 1); UsedGB = [math]::Round([double]$f[2] / 1MB, 1); Pct = [math]::Round([double]$f[2] / [double]$f[1] * 100) }
    }
}
function Get-Crashes { @(A logcat -d -b crash) }
function Get-CrashApps($lines) { foreach ($l in $lines) { if ($l -match 'Process: ([\w.:]+)' -or $l -match '>>> (\S+) <<<') { $matches[1] } } }

# --- Lag check ---
function Add-Flag([string]$text) { if (-not $Flags.Contains($text)) { [void]$Flags.Add($text) } }
function Get-RegValue([string]$Path, [string]$Name) {
    try { (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name } catch { $null }
}
function Show-RegValues([string]$Path, [string[]]$Names) {
    $rows = foreach ($n in $Names) {
        $v = Get-RegValue $Path $n
        [pscustomobject]@{ Key = $Path; Value = $n; Data = $(if ($null -eq $v) { '(not set)' } else { "$v" }) }
    }
    $rows | Format-Table -AutoSize
}
function Get-Events($Filter, [int]$Max = 50) {
    try { @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $Max -ErrorAction Stop) } catch { @() }
}
function Get-DirSize([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]@{ Folder = $Path; SizeMB = 'n/a'; Files = 0 } }
    $m = Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum
    [pscustomobject]@{ Folder = $Path; SizeMB = [math]::Round($m.Sum / 1MB); Files = $m.Count }
}
function Get-CompanyOf([string]$file) {
    try { if ($file -and (Test-Path -LiteralPath $file)) { [Diagnostics.FileVersionInfo]::GetVersionInfo($file).CompanyName } } catch { $null }
}
# Performance counter names are localised; map the English names through the Perflib index
function Get-CounterMap {
    if ($null -ne $global:CounterMap) { return $global:CounterMap }
    $map = @{}
    try {
        $en  = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009').Counter
        $loc = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage').Counter
        $byIndex = @{}
        for ($i = 0; $i + 1 -lt $loc.Count; $i += 2) { $byIndex[$loc[$i]] = $loc[$i + 1] }
        for ($i = 0; $i + 1 -lt $en.Count; $i += 2) {
            if ($byIndex.ContainsKey($en[$i]) -and -not $map.ContainsKey($en[$i + 1])) { $map[$en[$i + 1]] = $byIndex[$en[$i]] }
        }
    } catch {}
    $global:CounterMap = $map
    return $map
}
function ConvertTo-LocalCounter([string]$path) {
    if ($path -match '^\\([^\\(]+)(\([^)]*\))?\\(.+)$') {
        $map = Get-CounterMap
        $obj = $matches[1]; $inst = $matches[2]; $ctr = $matches[3]
        if ($map[$obj]) { $obj = $map[$obj] }
        if ($map[$ctr]) { $ctr = $map[$ctr] }
        return "\$obj$inst\$ctr"
    }
    return $path
}
function Measure-Ms([scriptblock]$sb) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try { [void](& $sb) } catch {}
    return [math]::Round($sw.Elapsed.TotalMilliseconds, 1)
}
'@

function Save-Report { if ($script:OutPath) { [IO.File]::WriteAllLines($script:OutPath, [string[]]$script:Out) } }

function Get-SectionRunspace([hashtable]$Vars) {
    if (-not $script:Rs -or $script:Rs.RunspaceStateInfo.State -ne 'Opened') {
        $script:Rs = [runspacefactory]::CreateRunspace()
        # STA for the Shell.Application and Search COM objects the lag check times
        $script:Rs.ApartmentState = 'STA'
        $script:Rs.ThreadOptions = 'ReuseThread'
        $script:Rs.Open()
        $ps = [PowerShell]::Create()
        $ps.Runspace = $script:Rs
        [void]$ps.AddScript($RsInit).Invoke()
        $ps.Dispose()
    }
    foreach ($k in $Vars.Keys) { $script:Rs.SessionStateProxy.SetVariable($k, $Vars[$k]) }
    $script:Rs
}

# Runs one section in the background runspace: a failing command is noted, not fatal,
# and a section that hangs (a stuck driver query, an adb call to a phone that went away) is abandoned
function Invoke-Collect([string]$Label, [scriptblock]$Body, [hashtable]$Vars, [int]$Timeout) {
    $script:Out.Add('')
    $script:Out.Add("========== $Label ==========")
    Write-Host "  [....] $Label" -NoNewline -ForegroundColor DarkGray
    $ps = [PowerShell]::Create()
    $ps.Runspace = Get-SectionRunspace $Vars
    [void]$ps.AddScript("& {`n$Body`n} | Out-String -Width 220")
    $state = 'ok'
    $text = ''
    $h = $ps.BeginInvoke()
    if ($h.AsyncWaitHandle.WaitOne($Timeout * 1000)) {
        try { $text = (@($ps.EndInvoke($h)) | ForEach-Object { "$_" }) -join "`n" } catch { $state = 'failed' }
        $ps.Dispose()
    } else {
        [void]$ps.BeginStop($null, $null)
        $script:Rs = $null
        $state = 'timeout'
    }
    # Trailing spaces off, runs of blank lines collapsed to one
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($text -split '\r?\n')) {
        $l = $l.TrimEnd()
        if ($l -or ($lines.Count -and $lines[$lines.Count - 1])) { $lines.Add($l) }
    }
    while ($lines.Count -and -not $lines[$lines.Count - 1]) { $lines.RemoveAt($lines.Count - 1) }
    if (-not $lines.Count -and $state -eq 'ok') { $lines.Add('(none)') }
    $script:Out.AddRange($lines)
    Write-Host "`r" -NoNewline
    switch ($state) {
        'ok'      { Say "  [ OK ] $Label" 'Green' }
        'timeout' {
            $script:Out.Add("(timed out after ${Timeout}s - skipped)")
            Say "  [WARN] $Label - timed out after ${Timeout}s" 'Yellow'
        }
        default   {
            $script:Out.Add('(unavailable or command failed)')
            Say "  [WARN] $Label - unavailable, skipped" 'Yellow'
        }
    }
    Save-Report
}

# $Sections: label, scriptblock, label, scriptblock, ...
# A '[[CHECKS]]' line in the output is replaced by the issues sections add with Add-Flag
function Write-Report([string]$Path, [string]$Title, [string]$Subject, [string]$Hint, [string]$Privacy,
                      [hashtable]$Vars, [object[]]$Sections, [int]$Timeout = 90) {
    $script:Out.Clear()
    $script:OutPath = $Path
    $flags = New-Object System.Collections.Generic.List[string]
    $Vars['Flags'] = $flags
    $script:Out.Add("==================== $Title REPORT: $(Get-Date -Format 'yyyy-MM-dd HH:mm') ====================")
    $script:Out.Add($Subject)
    Save-Report
    Say ""
    Say "  [INFO] Writing $($Title.ToLower()) report to: $Path" 'Cyan'
    Say "  [INFO] $Hint" 'Cyan'
    for ($i = 0; $i -lt $Sections.Count; $i += 2) { Invoke-Collect $Sections[$i] $Sections[$i + 1] $Vars $Timeout }
    $at = $script:Out.IndexOf('[[CHECKS]]')
    if ($at -ge 0) {
        $script:Out.RemoveAt($at)
        if ($flags.Count) { $script:Out.InsertRange($at, [string[]]@($flags | ForEach-Object { "!! $_" })) }
        else { $script:Out.Insert($at, 'No issues detected by the automated checks.') }
    }
    $script:Out.Add('')
    $script:Out.Add('==================== END OF REPORT ====================')
    Save-Report
    Say "  [ OK ] Report complete: $Path" 'Green'
    if ($at -ge 0) {
        if ($flags.Count) { Say "  [WARN] $($flags.Count) issue(s) flagged - listed at the top of the report" 'Yellow' }
        else { Say "  [ OK ] No issues flagged" 'Green' }
    }
    Say "  [INFO] $Privacy" 'Yellow'
}

# =====================================================================
#  OPTION 1 - WINDOWS LAPTOP
# =====================================================================
$LaptopSections = @(
'SYSTEM MAP', {
    $cs = Get-CimInstance Win32_ComputerSystem
    $os = Get-CimInstance Win32_OperatingSystem
    $nt = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    "$env:COMPUTERNAME  $($cs.Manufacturer) $ModelName  $($os.Caption) $($nt.DisplayVersion) build $($os.BuildNumber).$($nt.UBR)"
    $rows = New-Object System.Collections.Generic.List[string]
    foreach ($c in @(Get-CimInstance Win32_Processor)) {
        $rows.Add(('cpu      {0}  {1}c/{2}t' -f (Short $c.Name 60), $c.NumberOfCores, $c.NumberOfLogicalProcessors))
    }
    $mem = @(Get-CimInstance Win32_PhysicalMemory)
    $slots = (@(Get-CimInstance Win32_PhysicalMemoryArray) | Measure-Object MemoryDevices -Sum).Sum
    $rows.Add(('memory   {0}G  {1} of {2} slots  {3} MT/s' -f [math]::Round(($mem | Measure-Object Capacity -Sum).Sum / 1GB), $mem.Count, $slots, $mem[0].ConfiguredClockSpeed))
    foreach ($g in @(Get-CimInstance Win32_VideoController)) {
        $res = if ($g.CurrentHorizontalResolution) { "$($g.CurrentHorizontalResolution)x$($g.CurrentVerticalResolution) @ $($g.CurrentRefreshRate)Hz" } else { 'no display' }
        $rows.Add(('gpu      {0}  {1}' -f $g.Name, $res))
    }
    foreach ($d in @(Get-PhysicalDisk | Sort-Object DeviceId)) {
        $rows.Add(('disk     {0}  {1} {2}  {3}G  {4}' -f $d.FriendlyName, $d.BusType, $d.MediaType, [math]::Round($d.Size / 1GB), $d.HealthStatus))
    }
    foreach ($v in Get-FixedVolumes) {
        $u = $v.Size - $v.SizeRemaining
        $rows.Add(('volume   {0}:  {1}  {2:n0}G/{3:n0}G ({4:n0}%)' -f $v.DriveLetter, $v.FileSystem, ($u / 1GB), ($v.Size / 1GB), ($u / $v.Size * 100)))
    }
    foreach ($n in @(Get-NetAdapter -Physical | Where-Object Status -eq 'Up')) {
        $ip = (@(Get-NetIPAddress -InterfaceIndex $n.ifIndex -AddressFamily IPv4) | Select-Object -First 1).IPAddress
        $rows.Add(('network  {0}  {1}  {2}' -f $n.Name, $n.LinkSpeed, $(if ($ip) { $ip } else { '-' })))
    }
    foreach ($b in @(Get-Battery)) { $rows.Add(('battery  {0}% of design capacity  {1} cycles' -f $b.HealthPct, $b.Cycles)) }
    $tpm = Get-Tpm
    $bl = (Get-BitLockerVolume -MountPoint $env:SystemDrive).ProtectionStatus
    $rows.Add(('security BitLocker {0} {1}  Secure Boot {2}  TPM {3}' -f $env:SystemDrive, $(if ($bl) { "$bl".ToLower() } else { 'n/a' }), (Get-SecureBoot),
        $(if ($tpm.TpmReady) { 'ready' } elseif ($tpm.TpmPresent) { 'not ready' } else { 'none' })))
    Tree $rows
},
'HEALTH SUMMARY', {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS
    $nt = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $up = (Get-Date) - $os.LastBootUpTime
    $last = Get-LastUpdate
    '---- At a glance ----'
    Row 'Windows:' "$($os.Caption) $($nt.DisplayVersion) build $($os.BuildNumber).$($nt.UBR)"
    Row 'Model:' "$($cs.Manufacturer) $ModelName"
    Row 'BIOS:' ('{0} ({1:yyyy-MM-dd})' -f $bios.SMBIOSBIOSVersion, $bios.ReleaseDate)
    Row 'Uptime:' ('{0}d {1}h (since {2:yyyy-MM-dd HH:mm})' -f $up.Days, $up.Hours, $os.LastBootUpTime)
    Row 'Memory free/total:' ('{0:n1}G/{1:n1}G' -f ($os.FreePhysicalMemory / 1MB), ($os.TotalVisibleMemorySize / 1MB))
    Row 'Last update:' $(if ($last) { '{0} on {1:yyyy-MM-dd}' -f $last.HotFixID, $last.InstalledOn } else { '(unknown)' })

    $issues = New-Object System.Collections.Generic.List[string]
    # Code 22 is a device disabled on purpose
    $bad = @(Get-ProblemDevices | Where-Object ConfigManagerErrorCode -ne 22)
    if ($bad.Count) { $issues.Add("$($bad.Count) device(s) with a driver problem - see DEVICES") }
    foreach ($v in Get-FixedVolumes) {
        $p = [math]::Round(($v.Size - $v.SizeRemaining) / $v.Size * 100)
        if ($p -ge 85) { $issues.Add("Drive $($v.DriveLetter): at $p% - see DISKS") }
    }
    foreach ($d in @(Get-PhysicalDisk)) {
        if ($d.HealthStatus -and "$($d.HealthStatus)" -ne 'Healthy') { $issues.Add("Disk $($d.FriendlyName): $($d.HealthStatus) - see DISKS") }
        $r = $d | Get-StorageReliabilityCounter
        if ($r.Wear -ge 80) { $issues.Add("Disk $($d.FriendlyName): $($r.Wear)% worn - see DISKS") }
        if ($r.ReadErrorsUncorrected -gt 0) { $issues.Add("Disk $($d.FriendlyName): $($r.ReadErrorsUncorrected) uncorrected read errors - see DISKS") }
    }
    foreach ($b in @(Get-Battery)) {
        if ($b.HealthPct -and $b.HealthPct -lt 80) { $issues.Add("Battery at $($b.HealthPct)% of design capacity ($($b.Cycles) cycles) - see BATTERY & POWER") }
    }
    $pend = Get-PendingReboot
    if ($pend) { $issues.Add("Restart pending ($pend)") }
    if ($last -and $last.InstalledOn -lt (Get-Date).AddDays(-60)) { $issues.Add("No update installed for $(((Get-Date) - $last.InstalledOn).Days) days - see UPDATES") }
    if ((Get-SecureBoot) -eq 'off') { $issues.Add('Secure Boot off - see SECURITY') }
    $tpm = Get-Tpm
    if ($tpm -and -not $tpm.TpmReady) { $issues.Add($(if ($tpm.TpmPresent) { 'TPM present but not ready - see SECURITY' } else { 'No TPM found - see SECURITY' })) }
    $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive
    if ($bl -and "$($bl.ProtectionStatus)" -ne 'On') { $issues.Add("BitLocker protection off on $env:SystemDrive - see SECURITY") }
    $mp = Get-MpComputerStatus
    if ($mp -and $mp.AMRunningMode -eq 'Normal') {
        if (-not $mp.RealTimeProtectionEnabled) { $issues.Add('Defender real-time protection off - see SECURITY') }
        $age = ((Get-Date) - $mp.AntivirusSignatureLastUpdated).Days
        if ($age -gt 7) { $issues.Add("Defender signatures $age days old - see SECURITY") }
    }
    $k41 = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 41; StartTime = (Get-Date).AddDays(-30) }).Count
    if ($k41) { $issues.Add("$k41 unexpected shutdown(s) in 30 days - see LOGS") }
    $bsod = @(Get-ChildItem -LiteralPath "$env:windir\Minidump" -File | Where-Object LastWriteTime -gt (Get-Date).AddDays(-30)).Count
    if ($bsod) { $issues.Add("$bsod blue screen(s) in 30 days - see LOGS") }

    ''
    '---- Checks ----'
    if ($issues.Count) { $issues | ForEach-Object { "!! $_" } } else { 'No issues detected by the automated checks.' }
},
'WINDOWS / BIOS', {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS
    $nt = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $fast = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power').HiberbootEnabled
    [pscustomobject]@{
        Edition      = $nt.EditionID
        Version      = "$($os.Caption) $($nt.DisplayVersion) build $($os.BuildNumber).$($nt.UBR)"
        Architecture = $os.OSArchitecture
        Installed    = '{0:yyyy-MM-dd}' -f $os.InstallDate
        LastBoot     = '{0:yyyy-MM-dd HH:mm}' -f $os.LastBootUpTime
        FastStartup  = if ($fast -eq 1) { 'on (Shut down does not reset uptime)' } else { 'off' }
        TimeZone     = (Get-TimeZone).Id
        Manufacturer = $cs.Manufacturer
        Model        = "$($cs.Model) ($ModelName)"
        SystemType   = $cs.SystemType
        Hypervisor   = $cs.HypervisorPresent
        BIOS         = '{0} {1} ({2:yyyy-MM-dd})' -f $bios.Manufacturer, $bios.SMBIOSBIOSVersion, $bios.ReleaseDate
        BootMode     = if (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State') { 'UEFI' } else { 'Legacy BIOS' }
        Serial       = $bios.SerialNumber
    } | Format-List
},
'CPU', {
    Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed,
        @{ n = 'L2CacheMB'; e = { [math]::Round($_.L2CacheSize / 1KB, 1) } }, @{ n = 'L3CacheMB'; e = { [math]::Round($_.L3CacheSize / 1KB, 1) } },
        VirtualizationFirmwareEnabled, LoadPercentage | Format-List
},
'MEMORY', {
    $mem = @(Get-CimInstance Win32_PhysicalMemory)
    $mem | Select-Object @{ n = 'Slot'; e = { $_.DeviceLocator } }, @{ n = 'GB'; e = { [math]::Round($_.Capacity / 1GB) } }, Speed, ConfiguredClockSpeed,
        Manufacturer, @{ n = 'PartNumber'; e = { "$($_.PartNumber)".Trim() } } | Format-Table -AutoSize
    $os = Get-CimInstance Win32_OperatingSystem
    Row 'Slots used:' ('{0} of {1}' -f $mem.Count, (@(Get-CimInstance Win32_PhysicalMemoryArray) | Measure-Object MemoryDevices -Sum).Sum)
    Row 'Free/total:' ('{0:n1}G/{1:n1}G' -f ($os.FreePhysicalMemory / 1MB), ($os.TotalVisibleMemorySize / 1MB))
    Row 'Page file:' (OrNone ((Get-CimInstance Win32_PageFileUsage | ForEach-Object { "$($_.Name) $($_.AllocatedBaseSize)M, peak $($_.PeakUsage)M" }) -join '; '))
},
'GRAPHICS', {
    Get-CimInstance Win32_VideoController | Select-Object Name, DriverVersion, @{ n = 'DriverDate'; e = { '{0:yyyy-MM-dd}' -f $_.DriverDate } },
        @{ n = 'Resolution'; e = { if ($_.CurrentHorizontalResolution) { "$($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution) @ $($_.CurrentRefreshRate)Hz" } } } | Format-List
    '---- Monitors ----'
    OrNone (Get-CimInstance -Namespace root\wmi WmiMonitorID | ForEach-Object { '{0} {1}' -f (Chars $_.ManufacturerName), (Chars $_.UserFriendlyName) })
},
'DISKS', {
    Get-PhysicalDisk | Sort-Object DeviceId | Select-Object DeviceId, FriendlyName, MediaType, BusType, @{ n = 'SizeGB'; e = { [math]::Round($_.Size / 1GB) } },
        HealthStatus, FirmwareVersion, SerialNumber | Format-Table -AutoSize
    '---- Reliability (blank = not reported by the disk) ----'
    OrNone (@(Get-PhysicalDisk | Sort-Object DeviceId) | ForEach-Object {
        $r = $_ | Get-StorageReliabilityCounter
        [pscustomobject]@{ Disk = $_.FriendlyName; TempC = $r.Temperature; WearPct = $r.Wear; PowerOnHours = $r.PowerOnHours
            ReadErrors = $r.ReadErrorsUncorrected; WriteErrors = $r.WriteErrorsUncorrected }
    } | Format-Table -AutoSize)
    Sub 'Volumes'
    Get-Volume | Where-Object DriveLetter | Sort-Object DriveLetter | Select-Object DriveLetter, FileSystemLabel, FileSystem, DriveType,
        @{ n = 'SizeGB'; e = { [math]::Round($_.Size / 1GB, 1) } }, @{ n = 'FreeGB'; e = { [math]::Round($_.SizeRemaining / 1GB, 1) } },
        @{ n = 'UsedPct'; e = { if ($_.Size) { [math]::Round(($_.Size - $_.SizeRemaining) / $_.Size * 100) } } }, HealthStatus | Format-Table -AutoSize
},
'SECURITY', {
    '---- BitLocker ----'
    OrNone (Get-BitLockerVolume | Select-Object MountPoint, VolumeStatus, ProtectionStatus, EncryptionMethod, EncryptionPercentage | Format-Table -AutoSize) '(not available on this edition)'
    Sub 'TPM'
    OrNone (Get-Tpm | Select-Object TpmPresent, TpmReady, TpmEnabled, ManufacturerIdTxt, ManufacturerVersion | Format-List) '(unavailable)'
    Sub 'Secure Boot'
    Get-SecureBoot
    Sub 'Antivirus registered'
    OrNone ((Get-CimInstance -Namespace root/SecurityCenter2 AntiVirusProduct).displayName -join ', ')
    Sub 'Microsoft Defender'
    OrNone (Get-MpComputerStatus | Select-Object AMRunningMode, RealTimeProtectionEnabled, AntivirusSignatureLastUpdated, QuickScanEndTime, FullScanEndTime | Format-List) '(unavailable)'
    Sub 'Firewall profiles'
    OrNone (Get-NetFirewallProfile | Select-Object Name, Enabled, DefaultInboundAction | Format-Table -AutoSize)
    Sub 'UAC'
    $uac = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    if ($uac.EnableLUA -eq 0) { 'off' } else { "on (ConsentPromptBehaviorAdmin=$($uac.ConsentPromptBehaviorAdmin))" }
},
'BATTERY & POWER', {
    $bats = @(Get-Battery)
    if ($bats.Count) {
        $bats | Format-List
        $est = $BatXml.BatteryReport.RuntimeEstimates
        if ($est) {
            Row 'Runtime at full charge:' $est.FullChargeCapacity.ActiveRuntime
            Row 'Runtime at design charge:' $est.DesignCapacity.ActiveRuntime
        }
        Sub 'Capacity history (last 10 periods)'
        @($BatXml.BatteryReport.History.HistoryEntry) | Select-Object -Last 10 | Select-Object StartDate, EndDate,
            @{ n = 'FullCharge_mWh'; e = { $_.FullChargeCapacity } }, @{ n = 'Design_mWh'; e = { $_.DesignCapacity } }, CycleCount | Format-Table -AutoSize
    } else {
        'No battery detected.'
    }
    Sub 'Power plan'
    powercfg /getactivescheme
    Sub 'Sleep states'
    powercfg /a
},
'NETWORK', {
    '---- Adapters ----'
    Get-NetAdapter | Sort-Object Status, Name | Select-Object Name, InterfaceDescription, Status, LinkSpeed, DriverVersion | Format-Table -AutoSize
    Sub 'MAC addresses'
    Get-NetAdapter | Sort-Object Name | ForEach-Object { '{0,-30} {1}' -f $_.Name, $_.MacAddress }
    Sub 'IP configuration'
    OrNone (Get-NetIPConfiguration | Where-Object IPv4Address | ForEach-Object {
        [pscustomobject]@{ Interface = $_.InterfaceAlias; IPv4 = ($_.IPv4Address.IPAddress -join ', '); Gateway = ($_.IPv4DefaultGateway.NextHop -join ', ')
            DNS = (@($_.DNSServer | Where-Object AddressFamily -eq 2).ServerAddresses -join ', ') }
    } | Format-Table -AutoSize)
    Sub 'Wi-Fi'
    OrNone (netsh wlan show interfaces | Where-Object { $_ -match '^\s+(Name|Description|State|SSID|Radio type|Band|Channel|Receive rate.*|Transmit rate.*|Signal)\s*:' } |
        ForEach-Object { $_.Trim() }) '(no Wi-Fi adapter)'
},
'DEVICES', {
    $bad = @(Get-ProblemDevices)
    if ($bad.Count) {
        $bad | Select-Object Name, PNPClass, @{ n = 'Problem'; e = { $c = [int]$_.ConfigManagerErrorCode; "$c $($DevCodes[$c])" } }, DeviceID | Format-Table -AutoSize
    } else {
        'All devices report OK.'
    }
},
'DRIVERS (THIRD-PARTY, NEWEST 40)', {
    Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceName -and $_.DriverProviderName -and $_.DriverProviderName -ne 'Microsoft' } |
        Sort-Object DeviceName, DriverVersion -Unique | Sort-Object DriverDate -Descending | Select-Object -First 40 DeviceName, DriverVersion,
        @{ n = 'Date'; e = { '{0:yyyy-MM-dd}' -f $_.DriverDate } }, DriverProviderName | Format-Table -AutoSize
},
'SOFTWARE', {
    $apps = @(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' |
        Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 -and -not $_.ParentKeyName -and $_.ReleaseType -notmatch 'Update|Hotfix' } |
        Sort-Object DisplayName -Unique)
    "---- Installed apps ($($apps.Count)) ----"
    $apps | Select-Object DisplayName, DisplayVersion, Publisher | Format-Table -AutoSize
    Sub 'Startup items'
    OrNone (Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location | Format-Table -AutoSize)
    Sub 'Optional features enabled beyond the defaults'
    # Enabled out of the box on Windows 10/11
    $defaults = 'MicrosoftWindowsPowerShellV2Root', 'MicrosoftWindowsPowerShellV2', 'Printing-PrintToPDFServices-Features', 'Printing-XPSServices-Features',
        'Printing-Foundation-Features', 'Printing-Foundation-InternetPrinting-Client', 'SearchEngine-Client-Package', 'MSRDC-Infrastructure',
        'WorkFolders-Client', 'MediaPlayback', 'WindowsMediaPlayer', 'Windows-Defender-Default-Definitions', 'SmbDirect', 'NetFx4-AdvSrvs',
        'WCF-Services45', 'WCF-TCP-PortSharing45', 'Internet-Explorer-Optional-amd64', 'Microsoft-RemoteDesktopConnection', 'Microsoft-SnippingTool', 'Recall'
    OrNone (Get-WindowsOptionalFeature -Online | Where-Object { $_.State -eq 'Enabled' -and $_.FeatureName -notin $defaults } | ForEach-Object FeatureName | Sort-Object)
},
'UPDATES', {
    '---- Recent updates (last 10) ----'
    OrNone (Get-HotFix | Where-Object InstalledOn | Sort-Object InstalledOn -Descending | Select-Object -First 10 HotFixID, Description,
        @{ n = 'InstalledOn'; e = { '{0:yyyy-MM-dd}' -f $_.InstalledOn } } | Format-Table -AutoSize)
    Sub 'Restart pending'
    $p = Get-PendingReboot
    if ($p) { "Yes ($p)" } else { 'No' }
},
'SERVICES', {
    '---- Automatic services not running (trigger-start and delayed ones left out) ----'
    OrNone (Get-CimInstance Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" |
        Where-Object { -not $_.DelayedAutoStart -and -not (Test-Path -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Services\$($_.Name)\TriggerInfo") } |
        Select-Object Name, DisplayName, State, ExitCode | Format-Table -AutoSize)
},
'LOGS', {
    # Errors repeat; collapse them by message, then show the newest verbatim
    $since = (Get-Date).AddDays(-7)
    $ev = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2; StartTime = $since } -MaxEvents 1000)
    '---- System errors, last 7 days (count x message, numbers normalised) ----'
    OrNone ($ev | Group-Object { Norm "$($_.ProviderName) $($_.Id): $(Short (($_.Message -split '\r?\n')[0]) 140)" } | Sort-Object Count -Descending |
        Select-Object -First 20 | ForEach-Object { '{0,5} x {1}' -f $_.Count, $_.Name }) '(no errors logged)'
    Sub 'Most recent errors (last 10, verbatim)'
    OrNone ($ev | Select-Object -First 10 | ForEach-Object { '{0:yyyy-MM-dd HH:mm}  {1} {2}: {3}' -f $_.TimeCreated, $_.ProviderName, $_.Id, (Short $_.Message 200) })
    Sub 'App crashes and hangs, last 7 days (count x app)'
    OrNone (Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; StartTime = $since } |
        Group-Object { '{0} {1}' -f $(if ($_.ProviderName -eq 'Application Hang') { 'hang ' } else { 'crash' }), $_.Properties[0].Value } |
        Sort-Object Count -Descending | Select-Object -First 15 | ForEach-Object { '{0,5} x {1}' -f $_.Count, $_.Name })
    Sub 'Unexpected shutdowns and blue screens, last 30 days'
    OrNone (Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 41, 6008, 1001; StartTime = (Get-Date).AddDays(-30) } |
        Where-Object { $_.Id -ne 1001 -or $_.ProviderName -match 'WER-SystemErrorReporting' } | Select-Object -First 15 |
        ForEach-Object { '{0:yyyy-MM-dd HH:mm}  {1} {2}: {3}' -f $_.TimeCreated, $_.ProviderName, $_.Id, (Short $_.Message 160) })
    $md = @(Get-ChildItem -LiteralPath "$env:windir\Minidump" -File | Sort-Object LastWriteTime -Descending)
    Row 'Minidumps:' $(if ($md.Count) { '{0}, newest {1:yyyy-MM-dd HH:mm}' -f $md.Count, $md[0].LastWriteTime } else { 'none' })
}
)

function Get-LaptopReport {
    $name = Get-LaptopModelName
    Write-Report (Join-Path $OutDir "Data System $(ConvertTo-OneWord $name) $Stamp.txt") 'WINDOWS' "Computer: $env:COMPUTERNAME" `
        'Takes 1-2 minutes.' 'Contains serials, MACs and IPs - review before sharing.' @{ ModelName = $name } $LaptopSections
}

# =====================================================================
#  OPTION 2 - ANDROID PHONE
# =====================================================================
function Sync-Path {
    $env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
}

function Ensure-Adb {
    Sync-Path
    if (Get-Command adb -ErrorAction SilentlyContinue) { return $true }

    Say ""
    Say "  ADB is not installed. Installing it now (about 8 MB)..." 'Yellow'
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Say "  Cannot auto-install: 'winget' is unavailable on this PC." 'Red'
        Say "  Install 'App Installer' from the Microsoft Store, then try again." 'Red'
        return $false
    }
    winget install --id Google.PlatformTools -e --accept-source-agreements --accept-package-agreements | Out-Null
    Sync-Path
    if (Get-Command adb -ErrorAction SilentlyContinue) { Say "  ADB installed." 'Green'; return $true }

    $guess = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links'
    if (Test-Path (Join-Path $guess 'adb.exe')) { $env:Path += ";$guess"; Say "  ADB installed." 'Green'; return $true }

    Say "  ADB installed but not on PATH. Close this window and run the tool again." 'Yellow'
    return $false
}

# Returns the serial of the first authorised phone, or nothing on timeout
function Wait-ForPhone {
    Say ""
    Say "  ON YOUR PHONE, do this first:" 'Cyan'
    Say "   1. Settings > About phone > Software information" 'White'
    Say "      Tap 'Build number' 7 times to unlock Developer options." 'White'
    Say "   2. Settings > Developer options > turn ON 'USB debugging'." 'White'
    Say "   3. Plug the phone into this laptop with a USB cable." 'White'
    Say "   4. Set the USB mode to 'File transfer' (swipe down to change it)." 'White'
    Say "   5. UNLOCK the screen and tap ALLOW on the debugging pop-up." 'White'
    Say ""
    Say "  Waiting for the phone..." 'Yellow'

    adb start-server | Out-Null
    $warned = $false
    for ($i = 0; $i -lt 120; $i++) {
        $list = @((adb devices) -split "`n" | Where-Object { $_ -match "`t" } | ForEach-Object { $_.Trim() })
        $ready = @($list | Where-Object { $_ -match "`tdevice$" })
        if ($ready.Count) { Say "  Phone connected and authorised." 'Green'; return ($ready[0] -split "`t")[0] }
        if (($list -match "`tunauthorized") -and -not $warned) {
            Say "  Phone seen, but not authorised yet." 'Yellow'
            Say "  Unlock the screen and tap ALLOW (tick 'Always allow')." 'Yellow'
            $warned = $true
        }
        Start-Sleep -Seconds 2
    }
    Say ""
    Say "  Timed out waiting for the phone." 'Red'
    Say "  Try a different USB cable or port. On the phone, Developer options >" 'Red'
    Say "  'Revoke USB debugging authorisations', then run this again." 'Red'
}

$PhoneSections = @(
'SYSTEM MAP', {
    $p = Get-Props
    $b = Get-PhoneBattery
    $cap = Get-PhoneCapacity
    $d = Get-DataUse
    $memKb = if ("$(A shell cat /proc/meminfo)" -match 'MemTotal:\s+(\d+)') { [double]$matches[1] } else { 0 }
    $soc = "$($p['ro.soc.manufacturer']) $($p['ro.soc.model'])".Trim()
    if (-not $soc) { $soc = $p['ro.board.platform'] }
    $ip = if ("$(A shell ip -4 -o addr show wlan0)" -match 'inet ([\d.]+)') { $matches[1] } else { '-' }
    '{0} {1} ({2})  Android {3} (SDK {4})  patch {5}' -f $p['ro.product.manufacturer'], $p['ro.product.model'], $p['ro.product.device'],
        $p['ro.build.version.release'], $p['ro.build.version.sdk'], $p['ro.build.version.security_patch']
    Tree @(
        ('cpu      {0}  {1} cores' -f $soc, "$(A shell nproc)".Trim())
        ('memory   {0:n1}G' -f ($memKb / 1MB))
        $(if ($d) { 'storage  /data  {0}G/{1}G ({2}%)' -f $d.UsedGB, $d.TotalGB, $d.Pct } else { 'storage  /data  -' })
        ('battery  {0}%  {1}  health {2}  {3:n1}C{4}' -f $b['level'], $BatStatus[$b['status']], $BatHealth[$b['health']], ([double]$b['temperature'] / 10),
            $(if ($cap['pct']) { "  $($cap['pct'])% of design capacity" }))
        ('display  {0}  {1}' -f (@(A shell wm size)[0] -replace '.*:\s*', ''), (@(A shell wm density)[0] -replace '.*:\s*', ''))
        ('network  wlan0  {0}  {1} {2}' -f $ip, $p['gsm.operator.alpha'], $p['gsm.network.type'])
        ('apps     {0} user-installed, {1} disabled' -f @(A shell pm list packages -3).Count, @(A shell pm list packages -d).Count)
    )
},
'HEALTH SUMMARY', {
    $p = Get-Props
    $b = Get-PhoneBattery
    $cap = Get-PhoneCapacity
    $d = Get-DataUse
    $upS = [double](("$(A shell cat /proc/uptime)" -split ' ')[0])
    '---- At a glance ----'
    Row 'Model:' "$($p['ro.product.manufacturer']) $($p['ro.product.model'])"
    Row 'Android:' "$($p['ro.build.version.release']) (SDK $($p['ro.build.version.sdk'])), build $($p['ro.build.display.id'])"
    Row 'Security patch:' $p['ro.build.version.security_patch']
    Row 'Uptime:' ('{0}d {1}h' -f [math]::Floor($upS / 86400), [math]::Floor(($upS % 86400) / 3600))
    Row 'Battery:' "$($b['level'])%, $($BatStatus[$b['status']])"

    $issues = New-Object System.Collections.Generic.List[string]
    $patch = [datetime]::MinValue
    if ([datetime]::TryParseExact("$($p['ro.build.version.security_patch'])", 'yyyy-MM-dd', $null, 'None', [ref]$patch)) {
        $age = ((Get-Date) - $patch).Days
        if ($age -gt 90) { $issues.Add("Security patch is $age days old - check for a system update") }
    }
    if ($b['health'] -and $b['health'] -ne '2') { $issues.Add("Battery health: $($BatHealth[$b['health']]) - see BATTERY") }
    if ($cap['pct'] -and $cap['pct'] -lt 80) { $issues.Add("Battery at $($cap['pct'])% of design capacity - see BATTERY") }
    if ([double]$b['temperature'] -ge 450) { $issues.Add(('Battery at {0:n1}C - see BATTERY' -f ([double]$b['temperature'] / 10))) }
    if ($d -and $d.Pct -ge 85) { $issues.Add("Storage at $($d.Pct)% - see STORAGE") }
    $cr = @(Get-CrashApps (Get-Crashes))
    if ($cr.Count) { $issues.Add("$($cr.Count) app crash(es) logged - see CRASHES") }

    ''
    '---- Checks ----'
    if ($issues.Count) { $issues | ForEach-Object { "!! $_" } } else { 'No issues detected by the automated checks.' }
},
'DEVICE', {
    $p = Get-Props
    foreach ($k in 'ro.product.manufacturer', 'ro.product.brand', 'ro.product.model', 'ro.product.device', 'ro.product.name',
            'ro.build.version.release', 'ro.build.version.sdk', 'ro.build.version.security_patch', 'ro.build.display.id', 'ro.build.fingerprint',
            'ro.build.date', 'ro.bootloader', 'gsm.version.baseband', 'ro.board.platform', 'ro.soc.manufacturer', 'ro.soc.model', 'ro.hardware',
            'ro.product.cpu.abi', 'ro.boot.verifiedbootstate', 'ro.boot.flash.locked', 'ro.crypto.state', 'ro.crypto.type', 'persist.sys.timezone') {
        if ($p[$k]) { '{0,-34} {1}' -f $k, $p[$k] }
    }
},
'CPU', {
    OrNone (@(A shell cat /proc/cpuinfo) | Where-Object { $_ -cmatch '^(Hardware|model name|Processor)\s*:' } | Sort-Object -Unique)
    Sub 'Cores by max clock'
    OrNone (@(A shell 'cat /sys/devices/system/cpu/cpu*/cpufreq/cpuinfo_max_freq') | Where-Object { $_ -match '^\d+$' } | Group-Object |
        Sort-Object { [int]$_.Name } | ForEach-Object { '{0} x {1:n2} GHz' -f $_.Count, ([double]$_.Name / 1e6) })
},
'MEMORY', {
    OrNone (@(A shell cat /proc/meminfo) | Where-Object { $_ -match '^(MemTotal|MemAvailable|SwapTotal|SwapFree):' })
    Sub 'Summary'
    OrNone (@(A shell dumpsys meminfo) | Where-Object { $_ -match '^\s*(Total RAM|Free RAM|Used RAM|Lost RAM|ZRAM):' } | ForEach-Object { $_.Trim() })
},
'STORAGE', {
    OrNone (@(A shell df -h) | Where-Object { $_ -match '^Filesystem' -or $_ -match '\s(/|/data|/system|/vendor|/product|/metadata|/storage/emulated)$' } | Select-Object -Unique)
},
'BATTERY', {
    OrNone (@(A shell dumpsys battery) | Where-Object { $_.Trim() })
    Sub 'Capacity (blank = not exposed by the phone)'
    $cap = Get-PhoneCapacity
    Row 'Cycle count:' $cap['cycle_count']
    Row 'Full charge (uAh):' $cap['charge_full']
    Row 'Design (uAh):' $cap['charge_full_design']
    Row 'Health:' $(if ($cap['pct']) { "$($cap['pct'])% of design capacity" })
},
'DISPLAY', {
    A shell wm size
    A shell wm density
    Row 'Screen timeout:' ('{0}s' -f ([double]"$(A shell settings get system screen_off_timeout)" / 1000))
    Row 'Peak refresh rate:' "$(A shell settings get system peak_refresh_rate)"
    Row 'Brightness:' $(if ("$(A shell settings get system screen_brightness_mode)" -eq '1') { 'auto' } else { 'manual' })
},
'NETWORK', {
    $p = Get-Props
    '---- Interfaces (up) ----'
    OrNone (@(A shell ip -br a) | Where-Object { $_ -notmatch '^(lo|dummy|sit|ip6tnl|ip_vti|ip6_vti|ip6gre|gre|tunl|erspan)' -and $_ -match '\s(UP|UNKNOWN)\s' })
    Sub 'Wi-Fi'
    OrNone (@(A shell dumpsys wifi) | Where-Object { $_ -match 'mWifiInfo' } | Select-Object -First 1 |
        ForEach-Object { Short ($_ -replace '.*mWifiInfo\s*', '' -replace '(BSSID|MAC): [0-9a-fA-F:]{17},\s*', '') 400 })
    Sub 'Mobile'
    Row 'Operator:' $p['gsm.operator.alpha']
    Row 'Network type:' $p['gsm.network.type']
},
'APPS', {
    $u = @(A shell pm list packages -3 | ForEach-Object { $_ -replace '^package:', '' } | Sort-Object)
    $d = @(A shell pm list packages -d | ForEach-Object { $_ -replace '^package:', '' } | Sort-Object)
    Row 'Total packages:' @(A shell pm list packages).Count
    Sub "User-installed ($($u.Count))"
    OrNone $u
    Sub "Disabled ($($d.Count))"
    OrNone $d
},
'POWER', {
    Row 'Battery saver:' $(if ("$(A shell settings get global low_power)" -eq '1') { 'on' } else { 'off' })
    OrNone (@(A shell dumpsys deviceidle) | Where-Object { $_ -match '^\s+(mDeepEnabled|mLightEnabled|mState|mLightState)=' } | ForEach-Object { $_.Trim() })
    Sub 'Apps exempt from battery optimisation'
    OrNone (@(A shell dumpsys deviceidle whitelist) | Where-Object { $_ -match '^user,' } | ForEach-Object { ($_ -split ',')[1] } | Sort-Object)
},
'CRASHES', {
    $c = Get-Crashes
    '---- Crashed apps (count x app) ----'
    OrNone (Get-CrashApps $c | Group-Object | Sort-Object Count -Descending | ForEach-Object { '{0,5} x {1}' -f $_.Count, $_.Name }) '(no crashes logged)'
    Sub 'Most recent crash lines (last 30, verbatim)'
    OrNone ($c | Where-Object { $_ -and $_ -notmatch '^-+ beginning of' } | Select-Object -Last 30)
}
)

function Get-PhoneReport {
    if (-not (Ensure-Adb)) { return }
    $serial = Wait-ForPhone
    if (-not $serial) { return }
    $adb = (Get-Command adb).Source
    $model = "$(& $adb -s $serial shell getprop ro.product.model)".Trim()
    Write-Report (Join-Path $OutDir "Data Phone $(ConvertTo-OneWord $model) $Stamp.txt") 'ANDROID' "Device: $model" `
        'Takes about a minute. Keep the phone plugged in and unlocked.' 'Contains serials and IPs - review before sharing.' `
        @{ Adb = $adb; Serial = $serial } $PhoneSections
}

# =====================================================================
#  OPTION 4 - LAG CHECK
# =====================================================================
# Samples, times real actions and reads settings; each finding goes through Add-Flag to the checks at the top
$LagSections = @(
'SYSTEM MAP', $LaptopSections[1],
'HEALTH SUMMARY', {
    $up = (Get-Date) - $os.LastBootUpTime
    '---- At a glance ----'
    Row 'Windows:' "$($os.Caption) build $($os.BuildNumber)"
    Row 'Model:' "$($cs.Manufacturer) $ModelName"
    Row 'CPU:' "$($cpu.Name) ($($cpu.NumberOfCores)c/$($cpu.NumberOfLogicalProcessors)t)"
    Row 'RAM:' ('{0:n1}G' -f ($cs.TotalPhysicalMemory / 1GB))
    Row 'Uptime:' ('{0}d {1}h' -f $up.Days, $up.Hours)
    Row 'User:' $env:USERNAME
    ''
    '---- Checks (auto-detected, check these first) ----'
    '[[CHECKS]]'
},
'OS AND BOOT', {
    $ntv = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $uptime = (Get-Date) - $os.LastBootUpTime
    OrNone ([pscustomobject]@{
        OS            = "$($os.Caption) $(Get-RegValue $ntv 'DisplayVersion') build $($os.BuildNumber).$(Get-RegValue $ntv 'UBR')"
        Architecture  = $os.OSArchitecture
        Installed     = $os.InstallDate
        LastBoot      = $os.LastBootUpTime
        UptimeDays    = [math]::Round($uptime.TotalDays, 1)
        Manufacturer  = $cs.Manufacturer
        Model         = "$($cs.Model) / $((Get-CimInstance Win32_ComputerSystemProduct).Version)"
        SystemType    = $cs.SystemType
        BIOS          = "$((Get-CimInstance Win32_BIOS).SMBIOSBIOSVersion) ($((Get-CimInstance Win32_BIOS).ReleaseDate))"
        CPU           = $cpu.Name
        Cores         = "$($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) threads"
        CPU_MaxMHz    = $cpu.MaxClockSpeed
        CPU_NowMHz    = $cpu.CurrentClockSpeed
        CPU_LoadNow   = "$($cpu.LoadPercentage) %"
        RAM_GB        = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
    } | Format-List)
    if ($uptime.TotalDays -gt 7) { Add-Flag "PC has not been restarted for $([math]::Round($uptime.TotalDays)) days (uptime). Try Restart, not Shut down." }
    if ($cs.TotalPhysicalMemory -lt 7.5GB) { Add-Flag "Only $([math]::Round($cs.TotalPhysicalMemory / 1GB, 1)) GB RAM installed." }

    $fast = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled'
    "Fast startup (HiberbootEnabled): $(if ($null -eq $fast) { '(not set)' } else { $fast })"
    Sub "Last boots (Kernel-Boot event 27: boot type 0x0 = full, 0x1 = fast startup, 0x2 = resume)"
    $boots = Get-Events @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Boot'; Id = 27 } 6
    OrNone ($boots | Select-Object TimeCreated, @{n = 'Message'; e = { Short $_.Message 120 } } | Format-Table -AutoSize)
    if ($fast -eq 1 -and $boots -and -not ($boots | Where-Object { $_.Message -match '0x0\b' })) {
        Add-Flag "No full restart in the last $($boots.Count) boots (fast startup only). A real Restart clears the kernel session."
    }
},

'MEMORY AND PAGE FILE', {
    $totalMB = $os.TotalVisibleMemorySize / 1KB
    $freeMB  = $os.FreePhysicalMemory / 1KB
    $usedPct = [math]::Round(100 - ($freeMB / $totalMB * 100))
    OrNone ([pscustomobject]@{
        InstalledGB      = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        UsableGB         = [math]::Round($totalMB / 1KB, 2)
        HardwareReservedMB = [math]::Round(($cs.TotalPhysicalMemory / 1MB) - $totalMB)
        FreeGB           = [math]::Round($freeMB / 1KB, 2)
        InUsePercent     = $usedPct
        CommitUsedGB     = [math]::Round(($os.TotalVirtualMemorySize - $os.FreeVirtualMemory) / 1MB, 2)
        CommitLimitGB    = [math]::Round($os.TotalVirtualMemorySize / 1MB, 2)
        AutoPageFile     = $cs.AutomaticManagedPagefile
    } | Format-List)
    Sub "Modules"
    OrNone (Get-CimInstance Win32_PhysicalMemory | Select-Object BankLabel, @{n = 'GB'; e = { [math]::Round($_.Capacity / 1GB, 1) }}, Speed, ConfiguredClockSpeed, Manufacturer, PartNumber | Format-Table -AutoSize)
    Sub "Page file"
    OrNone (Get-CimInstance Win32_PageFileUsage | Select-Object Name, AllocatedBaseSize, CurrentUsage, PeakUsage | Format-Table -AutoSize)
    $mc = Get-Process 'Memory Compression' -ErrorAction SilentlyContinue
    if ($mc) { "Memory Compression process: $([math]::Round($mc.WorkingSet64 / 1MB)) MB" }
    if ($usedPct -ge 85) { Add-Flag "RAM is $usedPct % in use. The PC is short of memory (see top processes)." }
    $commitPct = ($os.TotalVirtualMemorySize - $os.FreeVirtualMemory) / $os.TotalVirtualMemorySize * 100
    if ($commitPct -ge 90) { Add-Flag "Commit charge is $([math]::Round($commitPct)) % of the limit. Page file nearly exhausted." }
    if (($cs.TotalPhysicalMemory / 1MB - $totalMB) -gt 1024) { Add-Flag "$([math]::Round(($cs.TotalPhysicalMemory / 1MB) - $totalMB)) MB of RAM is hardware reserved (BIOS / iGPU)." }
},

'POWER AND THERMAL', {
    $bat = Get-CimInstance Win32_Battery | Select-Object -First 1
    $onBattery = ($bat -and $bat.BatteryStatus -eq 1)
    "Battery: $(if ($bat) { "present, status $($bat.BatteryStatus) (1 = discharging, 2 = on AC), charge $($bat.EstimatedChargeRemaining) %" } else { 'none (desktop)' })"
    OrNone (& powercfg /getactivescheme 2>&1)
    $ovKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes'
    $ovNames = @{ '961cc777-2547-4f9d-8174-7d86181b8a7a' = 'Best power efficiency'; '00000000-0000-0000-0000-000000000000' = 'Balanced'; 'ded574b5-45a0-4f42-8737-46345c09c238' = 'Best performance'; '3af9b8d9-7c97-431d-ad78-34a8bfea439f' = 'Better performance' }
    foreach ($n in 'ActiveOverlayAcPowerScheme', 'ActiveOverlayDcPowerScheme') {
        $g = Get-RegValue $ovKey $n
        if ($g) { ("Power mode ({0}): {1} {2}" -f $n.Substring(13, 2), $g, $(if ($ovNames["$g".ToLower()]) { "= $($ovNames["$g".ToLower()])" } else { '' })) }
    }
    Sub "Max processor state (PROCTHROTTLEMAX, hex: 0x64 = 100 %)"
    OrNone ((& powercfg /query SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 2>&1) | Where-Object { $_ -match 'Power Setting|AC Power|DC Power' })
    Sub "Thermal zones"
    $tz = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
    if ($tz) {
        OrNone ($tz | Select-Object InstanceName, @{n = 'TempC'; e = { [math]::Round($_.CurrentTemperature / 10 - 273.15, 1) }} | Format-Table -AutoSize)
        $hot = $tz | Where-Object { ($_.CurrentTemperature / 10 - 273.15) -ge 90 }
        if ($hot) { Add-Flag "A thermal zone reports $([math]::Round($hot[0].CurrentTemperature / 10 - 273.15)) C. Possible overheating / throttling." }
    } else { '(not exposed by this firmware)' }
    if ($cpu.CurrentClockSpeed -and $cpu.MaxClockSpeed -and $cpu.CurrentClockSpeed -lt ($cpu.MaxClockSpeed * 0.5)) {
        Add-Flag "CPU reports $($cpu.CurrentClockSpeed) MHz of $($cpu.MaxClockSpeed) MHz max right now (possible throttling or power saving)."
    }
    if ($onBattery) { 'NOTE: on battery; Windows limits performance and WinSAT will refuse to run.' }
},

'LIVE COUNTERS (10 S IDLE)', {
    $wanted = @(
        '\Processor(_Total)\% Processor Time', '\Processor(_Total)\% Interrupt Time', '\Processor(_Total)\% DPC Time',
        '\Processor Information(_Total)\% Processor Performance', '\System\Processor Queue Length', '\System\Context Switches/sec',
        '\Memory\Available MBytes', '\Memory\Pages Input/sec', '\Memory\Page Faults/sec', '\Memory\% Committed Bytes In Use',
        '\Memory\Pool Nonpaged Bytes', '\Memory\Pool Paged Bytes', '\Paging File(_Total)\% Usage',
        '\PhysicalDisk(_Total)\Avg. Disk sec/Read', '\PhysicalDisk(_Total)\Avg. Disk sec/Write', '\PhysicalDisk(_Total)\Avg. Disk Queue Length',
        '\PhysicalDisk(_Total)\% Idle Time', '\PhysicalDisk(_Total)\Disk Transfers/sec', '\PhysicalDisk(*)\Avg. Disk sec/Transfer',
        '\Process(_Total)\Handle Count', '\Process(_Total)\Thread Count'
    )
    $paths = @($wanted | ForEach-Object { ConvertTo-LocalCounter $_ })
    $data = Get-Counter -Counter $paths -SampleInterval 2 -MaxSamples 5 -ErrorAction SilentlyContinue
    $rows = @()
    if ($data) {
        $rows = $data.CounterSamples | Group-Object Path | ForEach-Object {
            $m = $_.Group | Measure-Object CookedValue -Average -Maximum
            [pscustomobject]@{ Counter = ($_.Name -replace '^\\\\[^\\]+', ''); Avg = [math]::Round($m.Average, 3); Max = [math]::Round($m.Maximum, 3) }
        }
    }
    OrNone ($rows | Format-Table -AutoSize)
    if (-not $rows) { 'Get-Counter returned nothing (counters disabled or names not resolvable).'; return }
    $avg = {
        param($eng)
        $loc = (ConvertTo-LocalCounter $eng).ToLower()
        $r = $rows | Where-Object { $_.Counter.ToLower() -eq $loc } | Select-Object -First 1
        if ($r) { $r.Avg } else { $null }
    }
    $v = & $avg '\Processor(_Total)\% Processor Time';           if ($v -ge 50) { Add-Flag "CPU averaged $([math]::Round($v)) % while idle. Something is burning CPU (see top processes)." }
    $i = & $avg '\Processor(_Total)\% Interrupt Time'; $d = & $avg '\Processor(_Total)\% DPC Time'
    if (($i + $d) -ge 8) { Add-Flag "Interrupt + DPC time is $([math]::Round($i + $d, 1)) %. A driver (GPU, network, storage, audio) is misbehaving." }
    $v = & $avg '\Processor Information(_Total)\% Processor Performance'; if ($null -ne $v -and $v -lt 60) { Add-Flag "CPU running at $([math]::Round($v)) % of its base speed (power saving or thermal throttling)." }
    $v = & $avg '\Memory\Available MBytes';                      if ($null -ne $v -and $v -lt 1024) { Add-Flag "Only $([math]::Round($v)) MB of RAM available." }
    $v = & $avg '\Memory\Pages Input/sec';                       if ($v -ge 50) { Add-Flag "Hard page faults $([math]::Round($v))/s while idle: Windows is swapping to disk (RAM shortage)." }
    $v = & $avg '\PhysicalDisk(_Total)\Avg. Disk sec/Read';      if ($v -ge 0.02) { Add-Flag "Disk read latency averages $([math]::Round($v * 1000)) ms (healthy SSD: under 1 ms, HDD: under 15 ms)." }
    $v = & $avg '\PhysicalDisk(_Total)\Avg. Disk sec/Write';     if ($v -ge 0.03) { Add-Flag "Disk write latency averages $([math]::Round($v * 1000)) ms." }
    $v = & $avg '\PhysicalDisk(_Total)\Avg. Disk Queue Length';  if ($v -ge 2) { Add-Flag "Disk queue length averages $([math]::Round($v, 1)) while idle: the disk is the bottleneck." }
    $v = & $avg '\PhysicalDisk(_Total)\% Idle Time';             if ($null -ne $v -and $v -lt 50) { Add-Flag "Disk is busy $([math]::Round(100 - $v)) % of the time while idle." }
    $v = & $avg '\Memory\Pool Nonpaged Bytes';                   if ($v -ge 1.5GB) { Add-Flag "Nonpaged pool is $([math]::Round($v / 1GB, 1)) GB: a driver is leaking kernel memory." }
    $v = & $avg '\Process(_Total)\Handle Count';                 if ($v -ge 150000) { Add-Flag "$([math]::Round($v)) handles open system-wide: a process is leaking handles." }
},

'PROCESSES (5 S SAMPLE)', {
    $s1 = @{}
    Get-Process | ForEach-Object { try { $s1[$_.Id] = $_.TotalProcessorTime.TotalMilliseconds } catch {} }
    $sw = [Diagnostics.Stopwatch]::StartNew(); Start-Sleep -Seconds 5
    $procs = Get-Process; $elapsed = $sw.Elapsed.TotalMilliseconds
    $rows = foreach ($p in $procs) {
        $pct = $null
        try { if ($s1.ContainsKey($p.Id)) { $pct = [math]::Round(($p.TotalProcessorTime.TotalMilliseconds - $s1[$p.Id]) / $elapsed / $cores * 100, 1) } } catch {}
        $path = $null; try { $path = $p.Path } catch {}
        $resp = ''; try { if ($p.MainWindowHandle -ne 0) { $resp = $(if ($p.Responding) { 'yes' } else { 'NO' }) } } catch {}
        $start = ''; try { $start = $p.StartTime.ToString('MM-dd HH:mm') } catch {}
        [pscustomobject]@{
            Name = $p.ProcessName; PID = $p.Id; 'CPU%' = $pct
            WS_MB = [math]::Round($p.WorkingSet64 / 1MB); Priv_MB = [math]::Round($p.PrivateMemorySize64 / 1MB)
            Handles = $p.HandleCount; Threads = $p.Threads.Count; Responding = $resp; Started = $start
            Company = (Get-CompanyOf $path)
        }
    }
    $global:ProcRows = $rows
    "Processes: $($procs.Count)   Threads: $(($rows | Measure-Object Threads -Sum).Sum)   Handles: $(($rows | Measure-Object Handles -Sum).Sum)"
    Sub "Top 15 by CPU (share of all $cores logical CPUs)"
    OrNone ($rows | Where-Object { $_.'CPU%' -gt 0 } | Sort-Object 'CPU%' -Descending | Select-Object -First 15 | Format-Table -AutoSize)
    Sub "Top 15 by working set"
    OrNone ($rows | Sort-Object WS_MB -Descending | Select-Object -First 15 | Format-Table -AutoSize)
    Sub "Top 10 by handle count"
    OrNone ($rows | Sort-Object Handles -Descending | Select-Object -First 10 Name, PID, Handles, Threads, WS_MB, Company | Format-Table -AutoSize)
    Sub "Totals per program name (top 15 by memory)"
    OrNone ($rows | Group-Object Name | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Count = $_.Count; 'CPU%' = [math]::Round(($_.Group | Measure-Object 'CPU%' -Sum).Sum, 1); WS_MB = ($_.Group | Measure-Object WS_MB -Sum).Sum; Priv_MB = ($_.Group | Measure-Object Priv_MB -Sum).Sum }
    } | Sort-Object WS_MB -Descending | Select-Object -First 15 | Format-Table -AutoSize)
    Sub "Shell, search and system processes"
    $shellNames = '^(explorer|dwm|SearchIndexer|SearchHost|SearchApp|SearchUI|SearchProtocolHost|SearchFilterHost|StartMenuExperienceHost|ShellExperienceHost|RuntimeBroker|TextInputHost|msedgewebview2|WidgetService|Widgets|MsMpEng|NisSrv|sihost|ctfmon|OneDrive|Teams|ms-teams|csrss|lsass|System|Memory Compression|audiodg|WmiPrvSE|MoUsoCoreWorker|TiWorker|TrustedInstaller|CompatTelRunner|backgroundTaskHost|PhoneExperienceHost|CrossDeviceService|Cortana|fontdrvhost|spoolsv|SgrmBroker|SecurityHealthService|smartscreen)$'
    OrNone ($rows | Where-Object { $_.Name -match $shellNames } | Sort-Object Name | Format-Table -AutoSize)
    if ($isWin11 -and -not ($rows | Where-Object Name -eq 'SearchHost')) { Add-Flag "SearchHost.exe (Windows 11 Start search UI) is not loaded right now; it normally stays resident. It may have crashed." }
    if (-not $isWin11 -and -not ($rows | Where-Object Name -eq 'SearchApp')) { Add-Flag "SearchApp.exe (Windows 10 Start search UI) is not loaded right now; it normally stays resident. It may have crashed." }
    if (-not ($rows | Where-Object Name -eq 'StartMenuExperienceHost')) { Add-Flag "StartMenuExperienceHost.exe (Start menu host) is not loaded right now." }
    $ex = $rows | Where-Object Name -eq 'explorer'
    if ($ex) {
        if (($ex | Measure-Object Handles -Sum).Sum -ge 6000) { Add-Flag "explorer.exe holds $(($ex | Measure-Object Handles -Sum).Sum) handles (leak, often a shell extension)." }
        if (($ex | Measure-Object WS_MB -Sum).Sum -ge 1000) { Add-Flag "explorer.exe uses $(($ex | Measure-Object WS_MB -Sum).Sum) MB." }
    } else { Add-Flag "explorer.exe is not running." }
    $hog = $rows | Where-Object { $_.'CPU%' -ge 25 } | Sort-Object 'CPU%' -Descending
    foreach ($h in $hog) { Add-Flag "$($h.Name) (PID $($h.PID), $($h.Company)) used $($h.'CPU%') % CPU during the idle sample." }
    $nr = $rows | Where-Object Responding -eq 'NO'
    foreach ($n in $nr) { Add-Flag "Window of $($n.Name) (PID $($n.PID)) is NOT RESPONDING right now." }
    Sub "Services inside the busiest svchost processes"
    $svc = Get-CimInstance Win32_Service
    $topSvc = $rows | Where-Object Name -eq 'svchost' | Sort-Object @{e = { $_.'CPU%' }; Descending = $true }, @{e = 'WS_MB'; Descending = $true } | Select-Object -First 6
    foreach ($t in $topSvc) {
        $names = ($svc | Where-Object ProcessId -eq $t.PID | ForEach-Object { $_.Name }) -join ', '
        ("PID {0,-6} CPU {1,5}%  WS {2,5} MB  -> {3}" -f $t.PID, $t.'CPU%', $t.WS_MB, $names)
    }
},

'TIMING TESTS', {
    # Open + read 4 KB of 300 random files. Healthy SSD: < 0.5 ms avg. HDD or AV-hooked: several ms.
    $rand = New-Object Random
    foreach ($set in @(@{ Name = 'System32'; Path = "$env:windir\System32"; Depth = 0 }, @{ Name = 'Program Files'; Path = $env:ProgramFiles; Depth = 3 })) {
        $files = @(Get-ChildItem -LiteralPath $set.Path -File -Recurse -Depth $set.Depth -ErrorAction SilentlyContinue | Where-Object Length -gt 8KB | Select-Object -First 4000)
        $pick = $files | Sort-Object { $rand.Next() } | Select-Object -First 300
        $buf = New-Object byte[] 4096; $n = 0; $max = 0; $sum = 0
        foreach ($f in $pick) {
            $t = [Diagnostics.Stopwatch]::StartNew()
            try {
                $fs = New-Object IO.FileStream($f.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete), 4096, [IO.FileOptions]::RandomAccess)
                [void]$fs.Read($buf, 0, 4096); $fs.Close(); $n++
            } catch {}
            $ms = $t.Elapsed.TotalMilliseconds; $sum += $ms; if ($ms -gt $max) { $max = $ms }
        }
        if ($n) {
            $avgMs = [math]::Round($sum / $n, 2)
            ("File open+read ({0}): {1} files, avg {2} ms, max {3} ms" -f $set.Name, $n, $avgMs, [math]::Round($max, 1))
            if ($avgMs -ge 2) { Add-Flag "Opening small files takes $avgMs ms on average ($($set.Name)). Slow disk or antivirus scanning every file." }
        } else { "File open test ($($set.Name)): no files read" }
    }
    $ms = Measure-Ms { Get-ChildItem -LiteralPath "$env:windir\System32" -Force }
    "List System32 folder (~4000 entries): $ms ms"
    $ms = Measure-Ms { Get-ChildItem -LiteralPath $env:USERPROFILE -Force }
    "List user profile folder: $ms ms"

    $times = foreach ($i in 1..3) { Measure-Ms { Start-Process cmd.exe -ArgumentList '/c exit' -WindowStyle Hidden -Wait } }
    "Start a process (cmd /c exit) x3: $($times -join ' / ') ms  (healthy: under 150 ms; high = security software hooking every start)"
    if (($times | Measure-Object -Minimum).Minimum -ge 500) { Add-Flag "Starting a trivial process takes $(($times | Measure-Object -Minimum).Minimum) ms. Security / monitoring software is slowing every launch." }

    $times = foreach ($i in 1..3) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $p = Start-Process charmap.exe -PassThru
        while ($p -and -not $p.HasExited -and $p.MainWindowHandle -eq 0 -and $sw.ElapsedMilliseconds -lt 15000) { Start-Sleep -Milliseconds 25; $p.Refresh() }
        $t = $sw.ElapsedMilliseconds
        if ($p) { try { Stop-Process -Id $p.Id -Force } catch {} }
        $t
    }
    "Open a small Win32 window (charmap) x3: $($times -join ' / ') ms  (healthy: under 300 ms after the first)"
    if (($times | Select-Object -Skip 1 | Measure-Object -Minimum).Minimum -ge 1500) { Add-Flag "Opening a simple window takes $(($times | Select-Object -Skip 1 | Measure-Object -Minimum).Minimum) ms even when warm." }

    try {
        $shell = New-Object -ComObject Shell.Application
        $times = foreach ($i in 1..2) {
            $before = @($shell.Windows() | ForEach-Object { $_.HWND })
            $sw = [Diagnostics.Stopwatch]::StartNew()
            Start-Process explorer.exe -ArgumentList "$env:SystemDrive\"
            $new = $null
            while ($sw.ElapsedMilliseconds -lt 20000) {
                $w = @($shell.Windows() | Where-Object { $before -notcontains $_.HWND })
                if ($w.Count -gt 0) { $new = $w[0]; break }
                Start-Sleep -Milliseconds 50
            }
            $t = $sw.ElapsedMilliseconds
            Start-Sleep -Milliseconds 400
            if ($new) { try { $new.Quit() } catch {} } else { $t = "timeout" }
            $t
        }
        "Open an Explorer window ($env:SystemDrive\) x2: $($times -join ' / ') ms  (healthy: under 800 ms)"
        $nums = @($times | Where-Object { $_ -is [long] -or $_ -is [int] })
        if ($times -contains 'timeout') { Add-Flag "An Explorer window did not open within 20 s." }
        elseif ($nums -and ($nums | Measure-Object -Minimum).Minimum -ge 3000) { Add-Flag "Opening an Explorer window takes $(($nums | Measure-Object -Minimum).Minimum) ms (shell extensions, mapped drives, OneDrive or disk)." }
    } catch { "Explorer window test failed: $($_.Exception.Message)" }

    Sub "Windows Search index query (what Start search uses)"
    try {
        $conn = New-Object System.Data.OleDb.OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
        $open = Measure-Ms { $conn.Open() }
        "Connect to index: $open ms"
        $queries = @(
            @{ Name = 'file name starts with notepad'; Sql = "SELECT TOP 50 System.ItemPathDisplay FROM SYSTEMINDEX WHERE System.FileName LIKE 'notepad%'" },
            @{ Name = 'word search "setup"';           Sql = "SELECT TOP 50 System.ItemPathDisplay FROM SYSTEMINDEX WHERE CONTAINS(System.FileName, 'setup')" },
            @{ Name = 'anything in user profile';      Sql = "SELECT TOP 50 System.ItemPathDisplay FROM SYSTEMINDEX WHERE SCOPE='file:$env:USERPROFILE'" },
            @{ Name = 'installed app (Notepad)';       Sql = "SELECT TOP 20 System.ItemNameDisplay FROM SYSTEMINDEX WHERE System.Kind = 'program' AND CONTAINS(System.ItemNameDisplay, 'notepad')" }
        )
        foreach ($q in $queries) {
            # The timed block runs in a child scope: results go in a hashtable the caller can read
            $r = @{ n = 0; err = '' }
            $ms = Measure-Ms {
                try {
                    $cmd = $conn.CreateCommand(); $cmd.CommandText = $q.Sql; $cmd.CommandTimeout = 30
                    $rd = $cmd.ExecuteReader(); while ($rd.Read()) { $r.n++ }; $rd.Close()
                } catch { $r.err = $_.Exception.Message }
            }
            $n = $r.n; $err = $r.err
            ("Query '{0}': {1} ms, {2} rows {3}" -f $q.Name, $ms, $n, $(if ($err) { "ERROR $(Short $err 150)" } else { '' }))
            if ($ms -ge 2000) { Add-Flag "Search index query '$($q.Name)' took $ms ms (index busy or broken)." }
            if ($err) { Add-Flag "Search index query '$($q.Name)' failed: $(Short $err 100)" }
            elseif ($n -eq 0 -and $q.Name -eq 'anything in user profile') { Add-Flag "The search index returns nothing for the user profile: user files are not indexed (index empty or rebuilding)." }
        }
        $conn.Close()
    } catch { "Could not open the Windows Search index: $($_.Exception.Message)"; Add-Flag "Windows Search index cannot be opened: $(Short $_.Exception.Message 120)" }
},

'WINDOWS SEARCH', {
    $ws = Get-CimInstance Win32_Service -Filter "Name='WSearch'"
    "WSearch service: $(if ($ws) { "$($ws.State), start mode $($ws.StartMode)" } else { 'NOT INSTALLED' })"
    if (-not $ws -or $ws.State -ne 'Running') { Add-Flag "Windows Search service (WSearch) is $(if ($ws) { $ws.State } else { 'missing' }). Start search cannot find files or apps without it." }
    $wsKey = 'HKLM:\SOFTWARE\Microsoft\Windows Search'
    Show-RegValues $wsKey @('SetupCompletedSuccessfully', 'DataDirectory')
    $enh = Get-RegValue "$wsKey\Gather\Windows\SystemIndex" 'EnableFindMyFiles'
    "Enhanced indexing (whole PC, EnableFindMyFiles): $(if ($null -eq $enh) { '(classic)' } else { $enh })"
    if ($enh -eq 1) { Add-Flag "Search is in Enhanced mode (indexes the whole PC). Heavy on CPU/disk; Classic is usually enough." }
    $dataDir = Get-RegValue $wsKey 'DataDirectory'
    if ($dataDir) {
        $dataDir = [Environment]::ExpandEnvironmentVariables($dataDir)
        Sub "Index files in $dataDir"
        OrNone (Get-ChildItem -LiteralPath (Join-Path $dataDir 'Applications\Windows') -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 10MB -or $_.Name -match '^Windows\.(edb|db)' } | Select-Object Name, @{n = 'MB'; e = { [math]::Round($_.Length / 1MB) }}, LastWriteTime | Format-Table -AutoSize)
        $db = Get-ChildItem -LiteralPath (Join-Path $dataDir 'Applications\Windows') -File -ErrorAction SilentlyContinue | Where-Object Name -match '^Windows\.(edb|db)$'
        if ($db -and $db.Length -gt 4GB) { Add-Flag "Search index database is $([math]::Round($db.Length / 1GB, 1)) GB. Very large; a rebuild with fewer locations helps." }
    }

    Sub "Catalog status (from the indexer itself)"
    $child = @'
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
namespace WSX {
 [ComImport, Guid("AB310581-AC80-11D1-8DF3-00C04FB6EF69"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface ISearchCatalogManager {
  [return: MarshalAs(UnmanagedType.LPWStr)] string get_Name();
  IntPtr GetParameter([MarshalAs(UnmanagedType.LPWStr)] string n);
  void SetParameter([MarshalAs(UnmanagedType.LPWStr)] string n, IntPtr v);
  void GetCatalogStatus(out int status, out int reason);
  void Reset(); void Reindex();
  void ReindexMatchingURLs([MarshalAs(UnmanagedType.LPWStr)] string p);
  void ReindexSearchRoot([MarshalAs(UnmanagedType.LPWStr)] string p);
  void put_ConnectTimeout(uint t); uint get_ConnectTimeout();
  void put_DataTimeout(uint t); uint get_DataTimeout();
  int NumberOfItems();
  void NumberOfItemsToIndex(out int inc, out int notif, out int high);
  [return: MarshalAs(UnmanagedType.LPWStr)] string URLBeingIndexed();
 }
 [ComImport, Guid("AB310581-AC80-11D1-8DF3-00C04FB6EF63"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface ISearchManager {
  [return: MarshalAs(UnmanagedType.LPWStr)] string GetIndexerVersionStr();
  void GetIndexerVersion(out uint major, out uint minor);
  IntPtr GetParameter([MarshalAs(UnmanagedType.LPWStr)] string n);
  void SetParameter([MarshalAs(UnmanagedType.LPWStr)] string n, IntPtr v);
  [return: MarshalAs(UnmanagedType.LPWStr)] string get_ProxyName();
  [return: MarshalAs(UnmanagedType.LPWStr)] string get_BypassList();
  void SetProxy(int a, int b, uint c, [MarshalAs(UnmanagedType.LPWStr)] string d, [MarshalAs(UnmanagedType.LPWStr)] string e);
  ISearchCatalogManager GetCatalog([MarshalAs(UnmanagedType.LPWStr)] string n);
 }
 [ComImport, Guid("7D096C5F-AC08-4F1F-BEB7-5C22C517CE39")] public class CSearchManager {}
}
"@
$m = [WSX.ISearchManager]([Activator]::CreateInstance([WSX.CSearchManager]))
"IndexerVersion=" + $m.GetIndexerVersionStr()
$c = $m.GetCatalog('SystemIndex')
$s = 0; $r = 0; $c.GetCatalogStatus([ref]$s, [ref]$r)
"Status=$s"; "PausedReason=$r"
"Items=" + $c.NumberOfItems()
$a = 0; $b = 0; $d = 0; $c.NumberOfItemsToIndex([ref]$a, [ref]$b, [ref]$d)
"ToIndex=$a"; "Notifications=$b"; "HighPriority=$d"
"Indexing=" + $c.URLBeingIndexed()
'@
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($child))
    $out = @(& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $enc 2>&1 | ForEach-Object { "$_" })
    $statusNames = @('Idle', 'Paused', 'Recovering', 'Full crawl', 'Incremental crawl', 'Processing notifications', 'Shutting down')
    $reasonNames = @('none', 'high I/O', 'high CPU', 'high notification rate', 'low battery', 'low memory', 'low disk', 'delayed recovery', 'user active', 'external', 'upgrading')
    $kv = @{}
    foreach ($l in $out) { if ($l -match '^(\w+)=(.*)$') { $kv[$matches[1]] = $matches[2] } }
    if ($kv.Count -eq 0) { "(could not query: $(Short ($out -join ' ') 300))" }
    else {
        $st = [int]$kv['Status']; $rs = [int]$kv['PausedReason']
        "Indexer version : $($kv['IndexerVersion'])"
        "Status          : $st = $(if ($st -lt $statusNames.Count) { $statusNames[$st] } else { '?' })   paused reason: $rs = $(if ($rs -lt $reasonNames.Count) { $reasonNames[$rs] } else { '?' })"
        "Items in index  : $($kv['Items'])   waiting to index: $($kv['ToIndex'])   notifications: $($kv['Notifications'])   high priority: $($kv['HighPriority'])"
        "Indexing now    : $($kv['Indexing'])"
        if ($st -in 2, 3) { Add-Flag "Search index is $($statusNames[$st]) ($($kv['ToIndex']) items queued). Search gives partial results until it finishes." }
        if ($st -eq 1) { Add-Flag "Search indexer is paused ($($reasonNames[$rs])). Results stay stale." }
        if ([int]$kv['ToIndex'] -gt 50000) { Add-Flag "$($kv['ToIndex']) items are waiting to be indexed." }
        if ([int]$kv['Items'] -lt 5000) { Add-Flag "Search index holds only $($kv['Items']) items: it was reset recently or never finished." }
    }

    Sub "Web / cloud search settings (affects the Start search box)"
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' @('BingSearchEnabled', 'CortanaConsent', 'SearchboxTaskbarMode')
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings' @('IsDynamicSearchBoxEnabled', 'IsAADCloudSearchEnabled', 'IsMSACloudSearchEnabled', 'IsDeviceSearchHistoryEnabled')
    Show-RegValues 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' @('DisableSearchBoxSuggestions')
    Show-RegValues 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' @('DisableWebSearch', 'ConnectedSearchUseWeb', 'AllowCortana', 'AllowCloudSearch')
    $webOff = (Get-RegValue 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' 'DisableSearchBoxSuggestions') -eq 1 -or (Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' 'DisableWebSearch') -eq 1
    "Web results in Start search: $(if ($webOff) { 'disabled' } else { 'enabled (search waits on Bing; slow internet or proxy = spinning with no results)' })"

    Sub "Search / Start app packages"
    $pk = foreach ($n in 'Microsoft.Windows.Search', 'MicrosoftWindows.Client.CBS', 'MicrosoftWindows.Client.Core', 'Microsoft.Windows.StartMenuExperienceHost', 'Microsoft.Windows.ShellExperienceHost', 'Microsoft.Windows.Cortana', 'MicrosoftWindows.Client.WebExperience') {
        Get-AppxPackage -Name $n -ErrorAction SilentlyContinue | Select-Object Name, Version, Status, @{n = 'Installed'; e = { $_.InstallLocation -and (Test-Path -LiteralPath $_.InstallLocation) } }
    }
    OrNone ($pk | Format-Table -AutoSize)
    foreach ($p in $pk) { if ("$($p.Status)" -ne 'Ok') { Add-Flag "App package $($p.Name) status is '$($p.Status)' (Start/search UI package damaged)." } }

    Sub "Search events, last 7 days (Application log)"
    $ev = Get-Events @{ LogName = 'Application'; ProviderName = 'Microsoft-Windows-Search', 'Microsoft-Windows-Search-ProfileNotify', 'Microsoft-Windows-Search-Core'; StartTime = (Get-Date).AddDays(-7) } 20
    OrNone ($ev | Select-Object TimeCreated, Id, LevelDisplayName, @{n = 'Message'; e = { Short $_.Message 220 } } | Format-Table -AutoSize -Wrap)
    if ($ev | Where-Object { $_.Id -in 1006, 3083, 3084, 3105, 7042, 7040 -or $_.Message -match 'corrupt|reset|rebuil' }) { Add-Flag "Search log reports index corruption / reset / rebuild in the last 7 days." }
},

'STORAGE', {
    $pds = @(Get-PhysicalDisk)
    OrNone ($pds | Select-Object DeviceId, FriendlyName, MediaType, BusType, @{n = 'SizeGB'; e = { [math]::Round($_.Size / 1GB) }}, HealthStatus, OperationalStatus, FirmwareVersion | Format-Table -AutoSize)
    $sysDiskNo = (Get-Partition -DriveLetter $sysDrive -ErrorAction SilentlyContinue).DiskNumber
    $sysPd = $pds | Where-Object { "$($_.DeviceId)" -eq "$sysDiskNo" } | Select-Object -First 1
    "System drive $sysDrive`: is on disk $sysDiskNo = $($sysPd.FriendlyName) [$($sysPd.MediaType), $($sysPd.BusType)]"
    if ($sysPd -and $sysPd.MediaType -eq 'HDD') { Add-Flag "Windows runs from a mechanical hard disk (HDD). This alone causes slow windows and search; an SSD is the fix." }
    if ($sysPd -and $sysPd.BusType -eq 'USB') { Add-Flag "Windows runs from a USB disk." }
    foreach ($p in $pds) { if ("$($p.HealthStatus)" -ne 'Healthy') { Add-Flag "Disk '$($p.FriendlyName)' health is $($p.HealthStatus)." } }
    Sub "Reliability counters (SMART)"
    $rel = $pds | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
    OrNone ($rel | Select-Object DeviceId, Temperature, TemperatureMax, Wear, PowerOnHours, ReadErrorsTotal, ReadErrorsUncorrected, WriteErrorsTotal, WriteErrorsUncorrected, ReadLatencyMax, WriteLatencyMax, FlushLatencyMax | Format-Table -AutoSize)
    foreach ($r in $rel) {
        if ($r.Wear -ge 80) { Add-Flag "SSD wear on disk $($r.DeviceId) is $($r.Wear) %." }
        if ($r.ReadErrorsUncorrected -gt 0 -or $r.WriteErrorsUncorrected -gt 0) { Add-Flag "Disk $($r.DeviceId) reports uncorrected read/write errors (failing drive)." }
        if ($r.Temperature -ge 65) { Add-Flag "Disk $($r.DeviceId) is at $($r.Temperature) C." }
    }
    Sub "Volumes"
    $vols = Get-Volume | Where-Object DriveLetter
    OrNone ($vols | Select-Object DriveLetter, FileSystemLabel, FileSystem, @{n = 'SizeGB'; e = { [math]::Round($_.Size / 1GB, 1) }}, @{n = 'FreeGB'; e = { [math]::Round($_.SizeRemaining / 1GB, 1) }}, @{n = 'FreePct'; e = { if ($_.Size) { [math]::Round($_.SizeRemaining / $_.Size * 100) } }}, HealthStatus, DriveType | Format-Table -AutoSize)
    $sv = $vols | Where-Object DriveLetter -eq $sysDrive
    if ($sv -and ($sv.SizeRemaining -lt 15GB -or ($sv.SizeRemaining / $sv.Size) -lt 0.10)) { Add-Flag "$sysDrive`: has only $([math]::Round($sv.SizeRemaining / 1GB, 1)) GB free. Windows needs 15-20 GB free to stay responsive." }
    "TRIM (DisableDeleteNotify, 0 = TRIM on): $((& fsutil behavior query DisableDeleteNotify 2>&1) -join ' | ')"
    Sub "Folder sizes"
    OrNone (@(
        Get-DirSize $env:TEMP
        Get-DirSize "$env:windir\Temp"
        Get-DirSize "$env:windir\SoftwareDistribution\Download"
        Get-DirSize "$env:ProgramData\Microsoft\Windows\WER"
        Get-DirSize "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
        Get-DirSize "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"
    ) | Format-Table -AutoSize)
    Sub "Last WinSAT assessment (scores 1-9.9)"
    OrNone (Get-CimInstance Win32_WinSAT | Select-Object WinSATAssessmentState, CPUScore, MemoryScore, DiskScore, GraphicsScore, D3DScore, TimeTaken | Format-List)
    Sub "WinSAT disk test on $sysDrive`: (random 16 KB read, then sequential read)"
    foreach ($mode in '-ran -read', '-seq -read') {
        $o = @(& winsat.exe disk $mode.Split(' ') -drive $sysDrive 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match 'Disk|Total Run Time|Error|battery|cannot|not supported|remote' })
        OrNone $o
        $m = $o | Where-Object { $_ -match 'Random 16\.0 Read\s+([\d.]+) MB/s' } | Select-Object -First 1
        if ($m -and $m -match '([\d.]+) MB/s') { $mbs = [double]$matches[1]; if ($mbs -lt 2) { Add-Flag "Random 16 KB read speed is $mbs MB/s (HDD-class). A healthy SSD does 20+ MB/s here." } }
        $m = $o | Where-Object { $_ -match 'Sequential 64\.0 Read\s+([\d.]+) MB/s' } | Select-Object -First 1
        if ($m -and $m -match '([\d.]+) MB/s') { $mbs = [double]$matches[1]; if ($mbs -lt 150) { Add-Flag "Sequential read speed is $mbs MB/s (slow disk, bad cable or SATA mode)." } }
    }
},

'GRAPHICS AND VISUAL SETTINGS', {
    $gpus = Get-CimInstance Win32_VideoController
    OrNone ($gpus | Select-Object Name, Status, DriverVersion, DriverDate, @{n = 'VRAM_MB'; e = { [math]::Round($_.AdapterRAM / 1MB) }}, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate, VideoProcessor | Format-List)
    foreach ($g in $gpus) {
        if ($g.Name -match 'Basic Display|Basic Render') { Add-Flag "'$($g.Name)' is active: no real graphics driver installed. Everything on screen is drawn by the CPU." }
        if ("$($g.Status)" -ne 'OK') { Add-Flag "Graphics adapter '$($g.Name)' status is $($g.Status)." }
        if ($g.DriverDate -and $g.DriverDate -lt (Get-Date).AddYears(-4)) { "NOTE: driver for $($g.Name) dates from $($g.DriverDate.ToString('yyyy-MM-dd'))." }
    }
    $dwm = $global:ProcRows | Where-Object Name -eq 'dwm'
    if ($dwm) { "Desktop Window Manager: CPU $($dwm.'CPU%') %, $($dwm.WS_MB) MB, $($dwm.Handles) handles" }
    if ($dwm -and $dwm.WS_MB -ge 800) { Add-Flag "dwm.exe (Desktop Window Manager) uses $($dwm.WS_MB) MB: graphics driver memory leak, restart or update the driver." }
    Sub "Settings"
    Show-RegValues 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers' @('HwSchMode')
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' @('VisualFXSetting')
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' @('EnableTransparency')
    Show-RegValues 'HKCU:\Control Panel\Desktop\WindowMetrics' @('MinAnimate')
    Show-RegValues 'HKCU:\Control Panel\Desktop' @('UserPreferencesMask', 'MenuShowDelay')
    "(VisualFXSetting: 0 = let Windows choose, 1 = best appearance, 2 = best performance, 3 = custom)"
    Sub "Explorer / Start preferences"
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' @('LaunchTo', 'Start_TrackDocs', 'Start_IrisRecommendations', 'ShowRecent', 'ShowFrequent', 'ShowCloudFilesInQuickAccess', 'TaskbarDa', 'TaskbarMn', 'ShowCopilotButton')
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer' @('ShowRecent', 'ShowFrequent', 'ShowCloudFilesInQuickAccess')
    "(LaunchTo: 1 = This PC, 2 = Home/Quick access. Home with cloud files on is a common slow-open cause.)"
},

'SECURITY SOFTWARE AND FILTERS', {
    $mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($mp) { OrNone ($mp | Select-Object AMRunningMode, RealTimeProtectionEnabled, BehaviorMonitorEnabled, AntivirusSignatureAge, QuickScanAge, FullScanAge, QuickScanStartTime, QuickScanEndTime, FullScanStartTime, FullScanEndTime, IsTamperProtected | Format-List) }
    else { 'Defender status not available (another product may own real-time protection).' }
    Sub "Registered security products"
    $av = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue)
    OrNone ($av | Select-Object displayName, @{n = 'state'; e = { '0x{0:X6}' -f $_.productState }}, timestamp, pathToSignedProductExe | Format-Table -AutoSize)
    $third = @($av | Where-Object { $_.displayName -notmatch 'Windows Defender|Microsoft Defender' })
    if ($third.Count -ge 1) { "Third-party antivirus: $(($third | ForEach-Object { $_.displayName }) -join ', ')" }
    if ($third.Count -ge 2 -or ($third.Count -ge 1 -and $mp -and $mp.RealTimeProtectionEnabled)) { Add-Flag "More than one real-time antivirus appears active ($(($av | ForEach-Object { $_.displayName }) -join ', ')). Two scanners double every file access." }
    OrNone (Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct -ErrorAction SilentlyContinue | Select-Object displayName | Format-Table -AutoSize)
    $msmp = $global:ProcRows | Where-Object Name -eq 'MsMpEng'
    if ($msmp -and $msmp.'CPU%' -ge 15) { Add-Flag "Defender (MsMpEng) was using $($msmp.'CPU%') % CPU: a scan is running or something keeps triggering scans." }
    Sub "Minifilter drivers (every file access passes through these; AV, backup, cloud sync)"
    OrNone (& fltmc.exe filters 2>&1)
    Sub "Virtualization-based security / Memory integrity"
    $dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction SilentlyContinue
    if ($dg) {
        "VBS status: $($dg.VirtualizationBasedSecurityStatus) (2 = running)   services running: $($dg.SecurityServicesRunning -join ',') (1 = Credential Guard, 2 = HVCI/Memory integrity)"
        $gen = $null
        if ($cpu.Name -match 'i[3579]-(\d{3,5})') { $d = $matches[1]; $gen = [int]$(if ($d.Length -eq 5 -or ($d.Length -eq 4 -and $d[0] -eq '1')) { $d.Substring(0, 2) } elseif ($d.Length -eq 4) { $d.Substring(0, 1) } else { 1 }) }
        if ($dg.SecurityServicesRunning -contains 2 -and $gen -and $gen -lt 8) { Add-Flag "Memory integrity (HVCI) is on with a $gen`th-gen Intel CPU (no MBEC). Costs noticeable performance on this CPU." }
    } else { '(not available)' }
},

'STARTUP, TASKS, SERVICES', {
    $approved = @{}
    foreach ($root in 'HKCU', 'HKLM') {
        foreach ($sub in 'Run', 'Run32', 'StartupFolder') {
            $k = "${root}:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\$sub"
            $p = Get-ItemProperty -LiteralPath $k -ErrorAction SilentlyContinue
            if ($p) { foreach ($n in ($p.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' })) { $b = $n.Value; if ($b -is [byte[]] -and $b.Length -gt 0) { $approved[$n.Name] = $(if ($b[0] % 2 -eq 1) { 'disabled' } else { 'enabled' }) } } }
        }
    }
    $su = Get-CimInstance Win32_StartupCommand | ForEach-Object {
        $st = $approved[$_.Name]; if (-not $st) { $st = 'enabled' }
        $exe = $null; if ($_.Command -match '^"([^"]+)"' -or $_.Command -match '^(\S+\.exe)') { $exe = [Environment]::ExpandEnvironmentVariables($matches[1]) }
        [pscustomobject]@{ Name = $_.Name; State = $st; Location = $_.Location; Company = (Get-CompanyOf $exe); Command = (Short $_.Command 120) }
    }
    OrNone ($su | Sort-Object State, Name | Format-Table -AutoSize)
    $enabledCount = @($su | Where-Object State -eq 'enabled').Count
    "Enabled startup entries: $enabledCount"
    if ($enabledCount -ge 15) { Add-Flag "$enabledCount programs start with Windows. Each one costs memory and CPU at every logon." }
    Sub "Non-Microsoft scheduled tasks that are enabled"
    $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskPath -notlike '\Microsoft\*' -and $_.State -ne 'Disabled' }
    OrNone ($tasks | Select-Object TaskName, TaskPath, State, @{n = 'Triggers'; e = { ($_.Triggers | ForEach-Object { $_.CimClass.CimClassName -replace 'MSFT_Task|Trigger', '' }) -join ',' }}, @{n = 'Action'; e = { Short (($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; ') 100 }} | Format-Table -AutoSize)
    Sub "Scheduled tasks running right now"
    OrNone (Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object State -eq 'Running' | Select-Object TaskName, TaskPath | Format-Table -AutoSize)
    Sub "Running services not from the Windows folder"
    $svcs = Get-CimInstance Win32_Service | Where-Object { $_.State -eq 'Running' -and $_.PathName -notmatch '\\Windows\\' }
    OrNone ($svcs | Select-Object Name, DisplayName, StartMode, ProcessId, @{n = 'Path'; e = { Short $_.PathName 90 }} | Sort-Object Name | Format-Table -AutoSize)
    "Running services total: $(@(Get-CimInstance Win32_Service | Where-Object State -eq 'Running').Count)   non-Windows: $(@($svcs).Count)"
    Sub "Key Windows services"
    OrNone (Get-CimInstance Win32_Service | Where-Object { $_.Name -in 'WSearch', 'SysMain', 'Themes', 'DiagTrack', 'wuauserv', 'DoSvc', 'BITS', 'WinDefend', 'Spooler', 'OneSyncSvc', 'ShellHWDetection', 'StorSvc', 'TrustedInstaller', 'cbdhsvc', 'UsoSvc', 'Schedule', 'FontCache', 'TabletInputService', 'TextInputManagementService', 'CDPSvc' -or $_.Name -like 'cbdhsvc_*' } | Select-Object Name, State, StartMode | Sort-Object Name | Format-Table -AutoSize)
},

'INSTALLED SOFTWARE', {
    $apps = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object DisplayName | Sort-Object DisplayName -Unique
    "Installed programs: $(@($apps).Count)   Store apps (this user): $(@(Get-AppxPackage -ErrorAction SilentlyContinue).Count)"
    Sub "Programs known to run agents, scanners, overlays or 'optimizers'"
    $heavy = 'McAfee|Norton|Avast|AVG|Bitdefender|Kaspersky|ESET|Malwarebytes|Webroot|Trend Micro|Sophos|CrowdStrike|SentinelOne|Carbon Black|Cylance|Cortex|CCleaner|Advanced SystemCare|Driver Booster|IObit|Wallpaper Engine|Rainmeter|SupportAssist|HP Support|HP Wolf|Omen|Lenovo Vantage|Armoury|MSI Center|Dragon Center|iCUE|Razer|Logitech|LGHUB|Nahimic|Dropbox|OneDrive|Google Drive|iCloud|Teams|Zoom|Discord|Steam|Epic Games|Adobe|Creative Cloud|VMware|VirtualBox|Docker|Citrix|GlobalProtect|Cisco|FortiClient|Zscaler|Netskope|Tanium|Splashtop|TeamViewer|AnyDesk|Webex|Slack|Spotify|Bing Wallpaper|Booster|Optimizer|Cleaner|Toolbar|Updater|McAfee|WebAdvisor|Killer|Intel Driver|Dell Optimizer|Dell Digital|Waves|DTS|Dolby|Alienware|Acer|ASUS|Samsung Magician|Backup|Acronis|Carbonite|Norton 360|Avira|Comodo|360 Total|Baidu|Tencent|Kingsoft|WPS'
    OrNone ($apps | Where-Object { $_.DisplayName -match $heavy -or $_.Publisher -match $heavy } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Format-Table -AutoSize)
    Sub "Installed or updated in the last 30 days"
    $recent = $apps | Where-Object {
        $d = [datetime]::MinValue
        "$($_.InstallDate)" -match '^\d{8}$' -and [datetime]::TryParseExact("$($_.InstallDate)", 'yyyyMMdd', $null, 'None', [ref]$d) -and $d -gt (Get-Date).AddDays(-30)
    }
    OrNone ($recent | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Format-Table -AutoSize)
},

'DRIVERS AND DEVICES', {
    Sub "Devices with problems"
    $bad = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'OK' -and $_.Present }
    if ($bad) { OrNone ($bad | Select-Object FriendlyName, Status, Class, @{n = 'Problem'; e = { $_.ProblemDescription }} | Format-Table -AutoSize); foreach ($b in $bad) { Add-Flag "Device problem: $($b.FriendlyName) ($($b.Class)) status $($b.Status)." } }
    else { 'None.' }
    Sub "Display, storage, network, chipset drivers"
    $drv = Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceName -and $_.DeviceClass -in 'DISPLAY', 'NET', 'HDC', 'SCSIADAPTER', 'DISKDRIVE', 'MEDIA', 'BLUETOOTH', 'SYSTEM' }
    OrNone ($drv | Where-Object { $_.DeviceClass -ne 'SYSTEM' -or $_.DriverProviderName -notmatch 'Microsoft' } | Sort-Object DeviceClass, DeviceName | Select-Object DeviceClass, DeviceName, DriverVersion, DriverDate, DriverProviderName | Format-Table -AutoSize)
    Sub "Non-Microsoft kernel drivers running"
    $sd = Get-CimInstance Win32_SystemDriver | Where-Object { $_.State -eq 'Running' } | ForEach-Object {
        $path = $_.PathName -replace '^\\\?\?\\', '' -replace '^\\SystemRoot', $env:windir -replace '^System32', "$env:windir\System32"
        $co = Get-CompanyOf $path
        if ($co -and $co -notmatch 'Microsoft') { [pscustomobject]@{ Name = $_.Name; Company = $co; Path = (Short $path 80) } }
    }
    OrNone ($sd | Sort-Object Company, Name | Format-Table -AutoSize)
},

'WINDOWS UPDATE', {
    OrNone (Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object HotFixID, Description, InstalledOn -First 10 | Format-Table -AutoSize)
    $pending = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pending += 'WindowsUpdate RebootRequired' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pending += 'CBS RebootPending' }
    if (Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations') { $pending += 'PendingFileRenameOperations' }
    "Pending reboot: $(if ($pending) { $pending -join ', ' } else { 'no' })"
    if ($pending) { Add-Flag "A reboot is pending ($($pending -join ', ')). Restart before judging anything else." }
    $serv = $global:ProcRows | Where-Object { $_.Name -in 'TiWorker', 'TrustedInstaller', 'MoUsoCoreWorker', 'CompatTelRunner', 'wuauclt', 'UsoClient', 'MicrosoftEdgeUpdate', 'OneDriveSetup', 'DismHost' }
    if ($serv) { OrNone ($serv | Select-Object Name, PID, 'CPU%', WS_MB | Format-Table -AutoSize); if ($serv | Where-Object { $_.'CPU%' -ge 5 }) { Add-Flag "Windows servicing/telemetry ($(($serv | ForEach-Object { $_.Name }) -join ', ')) is busy right now. Updates are installing in the background." } }
    else { 'No servicing processes running.' }
    Sub "Update services"
    OrNone (Get-CimInstance Win32_Service | Where-Object { $_.Name -in 'wuauserv', 'UsoSvc', 'DoSvc', 'BITS', 'TrustedInstaller', 'WaaSMedicSvc' } | Select-Object Name, State, StartMode | Format-Table -AutoSize)
    Sub "Component store (DISM CheckHealth)"
    OrNone (@(& dism.exe /Online /Cleanup-Image /CheckHealth 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match '\S' -and $_ -notmatch '^\[=' }))
},

'NETWORK', {
    OrNone (Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up' | Select-Object Name, InterfaceDescription, LinkSpeed, MediaType, DriverVersion | Format-Table -AutoSize)
    OrNone (@(& netsh.exe wlan show interfaces 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match '^\s*(State|Radio type|Band|Channel|Receive rate|Transmit rate|Signal)\s*:' }))
    Sub "Proxy"
    Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' @('ProxyEnable', 'ProxyServer', 'AutoConfigURL')
    OrNone (@(& netsh.exe winhttp show proxy 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match '\S' }))
    if ((Get-RegValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyEnable') -eq 1 -or (Get-RegValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'AutoConfigURL')) { Add-Flag "A proxy / PAC script is configured. If it is unreachable, Start search and Store apps hang waiting for it." }
    Sub "Latency"
    $ms = Measure-Ms { Resolve-DnsName www.bing.com -ErrorAction Stop -DnsOnly }
    "DNS lookup www.bing.com: $ms ms"
    if ($ms -ge 1000) { Add-Flag "DNS lookup took $ms ms. Slow DNS makes web-enabled search spin." }
    foreach ($h in 'www.bing.com', 'www.microsoft.com') {
        $sw = [Diagnostics.Stopwatch]::StartNew(); $ok = $false
        try { $c = New-Object Net.Sockets.TcpClient; $ar = $c.BeginConnect($h, 443, $null, $null); $ok = $ar.AsyncWaitHandle.WaitOne(5000) -and $c.Connected; $c.Close() } catch {}
        "TCP connect $h`:443: $(if ($ok) { "$([math]::Round($sw.Elapsed.TotalMilliseconds)) ms" } else { 'FAILED / timeout' })"
        if (-not $ok -and $h -eq 'www.bing.com') { Add-Flag "Cannot reach www.bing.com. With web results enabled, Start search waits on it." }
    }
    $gw = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Sort-Object RouteMetric | Select-Object -First 1).NextHop
    if ($gw) { $r = Test-Connection $gw -Count 3 -ErrorAction SilentlyContinue; "Ping gateway x3: $(if ($r) { ($r | ForEach-Object { $_.ResponseTime }) -join ' / ' } else { 'no reply' }) ms" }
    Sub "Mapped network drives (as seen by this elevated session)"
    $maps = @(Get-SmbMapping -ErrorAction SilentlyContinue)
    OrNone ($maps | Select-Object LocalPath, RemotePath, Status | Format-Table -AutoSize)
    OrNone (Get-CimInstance Win32_NetworkConnection -ErrorAction SilentlyContinue | Select-Object LocalName, RemoteName, Status, ConnectionState | Format-Table -AutoSize)
    foreach ($m in $maps) { if ("$($m.Status)" -ne 'OK') { Add-Flag "Mapped drive $($m.LocalPath) -> $($m.RemotePath) is $($m.Status). Explorer stalls on unreachable network drives." } }
    $od = $global:ProcRows | Where-Object Name -eq 'OneDrive'
    "OneDrive running: $(if ($od) { "yes (CPU $($od.'CPU%') %, $($od.WS_MB) MB)" } else { 'no' })"
},

'SHELL EXTENSIONS', {
    $roots = @(
        'HKLM:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers', 'HKLM:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers',
        'HKLM:\SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers', 'HKLM:\SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers',
        'HKLM:\SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers', 'HKLM:\SOFTWARE\Classes\Drive\shellex\ContextMenuHandlers',
        'HKLM:\SOFTWARE\Classes\*\shellex\PropertySheetHandlers', 'HKLM:\SOFTWARE\Classes\Folder\shellex\ColumnHandlers',
        'HKLM:\SOFTWARE\Classes\*\shellex\DragDropHandlers', 'HKLM:\SOFTWARE\Classes\Directory\shellex\DragDropHandlers',
        'HKLM:\SOFTWARE\Classes\Directory\shellex\CopyHookHandlers',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers',
        'HKCU:\SOFTWARE\Classes\*\shellex\ContextMenuHandlers', 'HKCU:\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers',
        'HKCU:\SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers'
    )
    $rows = foreach ($r in $roots) {
        foreach ($k in (Get-ChildItem -LiteralPath $r -ErrorAction SilentlyContinue)) {
            $clsid = (Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue).'(default)'
            if (-not $clsid -or $clsid -notmatch '^\{') { $clsid = $k.PSChildName }
            if ($clsid -notmatch '^\{[0-9A-Fa-f-]+\}$') { continue }
            $dll = $null
            foreach ($c in "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32", "HKCU:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32", "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\$clsid\InprocServer32") {
                $d = (Get-ItemProperty -LiteralPath $c -ErrorAction SilentlyContinue).'(default)'
                if ($d) { $dll = [Environment]::ExpandEnvironmentVariables($d); break }
            }
            $co = Get-CompanyOf $dll
            [pscustomobject]@{ Type = ($r -replace '^HK(LM|CU):\\SOFTWARE\\(Classes\\)?', '' -replace '\\shellex', ''); Name = $k.PSChildName; Company = $(if ($co) { $co } elseif ($dll) { '?' } else { '(not registered)' }); Dll = (Short $dll 90) }
        }
    }
    $nonMs = @($rows | Where-Object { $_.Company -notmatch 'Microsoft' })
    "Total handlers: $(@($rows).Count)   non-Microsoft: $($nonMs.Count)"
    Sub "Non-Microsoft handlers (each runs inside Explorer when a window, menu or icon is drawn)"
    OrNone ($nonMs | Sort-Object Company, Type | Format-Table -AutoSize)
    if ($nonMs.Count -ge 12) { Add-Flag "$($nonMs.Count) non-Microsoft shell extensions are loaded into Explorer. Known slow-window cause; disable with ShellExView." }
    $overlays = @($rows | Where-Object Type -like '*ShellIconOverlayIdentifiers*')
    if ($overlays.Count -gt 15) { Add-Flag "$($overlays.Count) icon overlay handlers registered (Windows uses only 15; cloud sync apps fight over them)." }
},

'EVENT LOGS', {
    Sub "Errors in System + Application, last 3 days: count by source"
    $errs = @(Get-Events @{ LogName = 'System', 'Application'; Level = 1, 2; StartTime = (Get-Date).AddDays(-3) } 600)
    OrNone ($errs | Group-Object ProviderName, Id | Sort-Object Count -Descending | Select-Object -First 20 Count, Name | Format-Table -AutoSize)
    Sub "Newest 25 errors"
    OrNone ($errs | Select-Object -First 25 TimeCreated, LogName, ProviderName, Id, @{n = 'Message'; e = { Short $_.Message 260 } } | Format-Table -AutoSize -Wrap)

    Sub "Application hangs (1002) and crashes (1000), last 14 days"
    $hangs = @(Get-Events @{ LogName = 'Application'; ProviderName = 'Application Hang', 'Application Error'; StartTime = (Get-Date).AddDays(-14) } 200)
    OrNone ($hangs | ForEach-Object { [pscustomobject]@{ Kind = $(if ($_.Id -eq 1002) { 'hang' } else { 'crash' }); App = "$($_.Properties[0].Value)"; Module = $(if ($_.Id -eq 1000) { "$($_.Properties[3].Value)" } else { '' }) } } | Group-Object Kind, App, Module | Sort-Object Count -Descending | Select-Object -First 20 Count, Name | Format-Table -AutoSize)
    OrNone ($hangs | Select-Object -First 10 TimeCreated, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    $shellHang = $hangs | Where-Object { "$($_.Properties[0].Value)" -match '^(explorer|SearchHost|SearchApp|StartMenuExperienceHost|ShellExperienceHost|dwm)\.exe' }
    if ($shellHang) { Add-Flag "$(@($shellHang).Count) hang/crash events for the shell ($(($shellHang | ForEach-Object { "$($_.Properties[0].Value)" } | Sort-Object -Unique) -join ', ')) in 14 days." }

    Sub "Disk / storage events, last 30 days (7, 11, 51, 55, 98, 129, 153, 157 = bad sectors, resets, NTFS damage)"
    $disk = @(Get-Events @{ LogName = 'System'; Id = 7, 11, 51, 55, 98, 129, 153, 157, 140; StartTime = (Get-Date).AddDays(-30) } 40)
    OrNone ($disk | Select-Object TimeCreated, Id, ProviderName, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    if ($disk | Where-Object { $_.ProviderName -match 'disk|ntfs|stor|volmgr|nvme|iaStor|partmgr' }) { Add-Flag "$(@($disk).Count) storage error events in 30 days (IDs $(($disk | ForEach-Object Id | Sort-Object -Unique) -join ',')). The disk or its cable/driver is faulty." }

    Sub "Hardware errors (WHEA), last 30 days"
    $whea = @(Get-Events @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WHEA-Logger'; StartTime = (Get-Date).AddDays(-30) } 20)
    OrNone ($whea | Select-Object TimeCreated, Id, LevelDisplayName, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    if ($whea) { Add-Flag "$($whea.Count) WHEA hardware error events in 30 days (CPU, memory, PCIe or firmware fault)." }

    Sub "Unexpected shutdowns / crashes (41, 6008, 1001), last 30 days"
    $crash = @(Get-Events @{ LogName = 'System'; Id = 41, 6008, 1001; StartTime = (Get-Date).AddDays(-30) } 20)
    OrNone ($crash | Select-Object TimeCreated, Id, ProviderName, @{n = 'Message'; e = { Short $_.Message 160 } } | Format-Table -AutoSize -Wrap)
    if ($crash | Where-Object { $_.Id -in 41, 6008 }) { Add-Flag "$(@($crash | Where-Object { $_.Id -in 41, 6008 }).Count) unexpected shutdowns / blue screens in 30 days." }

    Sub "Diagnostics-Performance, last 14 days (1xx boot, 2xx shutdown, 3xx standby, 4xx desktop responsiveness, 5xx Explorer)"
    $perf = @(Get-Events @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; StartTime = (Get-Date).AddDays(-14) } 300)
    OrNone ($perf | Group-Object Id | Sort-Object Count -Descending | Select-Object Count, Name | Format-Table -AutoSize)
    OrNone ($perf | Where-Object { $_.Id -ge 100 -and $_.Id -lt 200 } | Select-Object -First 5 TimeCreated, Id, @{n = 'Boot'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    OrNone ($perf | Where-Object { $_.Id -ge 400 } | Select-Object -First 20 TimeCreated, Id, LevelDisplayName, @{n = 'Message'; e = { Short $_.Message 220 } } | Format-Table -AutoSize -Wrap)
    $boot = $perf | Where-Object Id -eq 100 | Select-Object -First 1
    if ($boot -and $boot.Message -match 'Boot Duration\s*:\s*(\d+)ms') { $bs = [int]$matches[1] / 1000; "Last measured boot: $bs s"; if ($bs -ge 120) { Add-Flag "Last boot took $bs s (events 101-110 below name the slow drivers/services/apps)." } }
    $slow = $perf | Where-Object { $_.Id -in 101..110 } | Select-Object -First 15
    OrNone ($slow | Select-Object TimeCreated, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    if (@($perf | Where-Object { $_.Id -ge 500 -and $_.Id -lt 600 }).Count -ge 3) { Add-Flag "Windows logged $(@($perf | Where-Object { $_.Id -ge 500 -and $_.Id -lt 600 }).Count) Explorer responsiveness-degradation events in 14 days." }

    Sub "Store app deployment errors, last 7 days"
    OrNone (Get-Events @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 1, 2; StartTime = (Get-Date).AddDays(-7) } 10 | Select-Object TimeCreated, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    Sub "Shell-Core / TWinUI errors, last 7 days"
    OrNone (Get-Events @{ LogName = 'Microsoft-Windows-Shell-Core/Operational', 'Microsoft-Windows-TWinUI/Operational'; Level = 1, 2; StartTime = (Get-Date).AddDays(-7) } 15 | Select-Object TimeCreated, LogName, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
},

'RELIABILITY AND CRASH DUMPS', {
    $stab = Get-CimInstance Win32_ReliabilityStabilityMetrics -ErrorAction SilentlyContinue | Sort-Object TimeGenerated -Descending | Select-Object -First 1
    if ($stab) { "Reliability index (1-10, Reliability Monitor): $($stab.SystemStabilityIndex) at $($stab.TimeGenerated)" }
    Sub "Reliability records, last 14 days"
    $rr = Get-CimInstance Win32_ReliabilityRecords -ErrorAction SilentlyContinue | Where-Object { $_.TimeGenerated -gt (Get-Date).AddDays(-14) }
    OrNone ($rr | Group-Object SourceName, ProductName | Sort-Object Count -Descending | Select-Object -First 15 Count, Name | Format-Table -AutoSize)
    OrNone ($rr | Sort-Object TimeGenerated -Descending | Select-Object -First 25 TimeGenerated, SourceName, EventIdentifier, ProductName, @{n = 'Message'; e = { Short $_.Message 150 } } | Format-Table -AutoSize -Wrap)
    Sub "Crash dumps"
    $md = @(Get-ChildItem -LiteralPath "$env:windir\Minidump" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    "Minidumps (blue screens): $($md.Count)$(if ($md) { ", newest $($md[0].LastWriteTime)" })"
    $lk = @(Get-ChildItem -LiteralPath "$env:windir\LiveKernelReports" -Recurse -File -ErrorAction SilentlyContinue | Where-Object LastWriteTime -gt (Get-Date).AddDays(-30))
    "Live kernel reports (driver hangs, GPU timeouts), last 30 days: $($lk.Count)"
    if ($lk) { OrNone ($lk | Select-Object -First 10 Name, LastWriteTime, @{n = 'MB'; e = { [math]::Round($_.Length / 1MB) }} | Format-Table -AutoSize); Add-Flag "$($lk.Count) live kernel reports in 30 days (a driver stopped responding; folder names say which: WATCHDOG, USB, PCI, NDIS...)." }
    $wer = @(Get-ChildItem -LiteralPath "$env:ProgramData\Microsoft\Windows\WER\ReportQueue", "$env:ProgramData\Microsoft\Windows\WER\ReportArchive", "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue", "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive" -Directory -ErrorAction SilentlyContinue | Where-Object LastWriteTime -gt (Get-Date).AddDays(-14))
    "Error reports (WER), last 14 days: $($wer.Count)"
    OrNone ($wer | Group-Object { ($_.Name -split '_')[0, 1] -join '_' } | Sort-Object Count -Descending | Select-Object -First 15 Count, Name | Format-Table -AutoSize)
    Sub "Logged-on sessions"
    OrNone (@(& quser.exe 2>&1 | ForEach-Object { "$_" }))
}
)

function Get-LagReport {
    $name = Get-LaptopModelName
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $vars = @{
        ModelName = $name; os = $os; cs = $cs; cpu = $cpu; cores = [int]$env:NUMBER_OF_PROCESSORS
        sysDrive = $env:SystemDrive.Substring(0, 1); isWin11 = [int]$os.BuildNumber -ge 22000
    }
    Write-Report (Join-Path $OutDir "Diag Lag $(ConvertTo-OneWord $name) $Stamp.txt") 'LAG' "Computer: $env:COMPUTERNAME" `
        'Takes 3-5 minutes and opens a few windows. Leave the PC alone until it finishes.' `
        'Contains the computer and user name, no serials or addresses.' $vars $LagSections 240
}

# =====================================================================
#  MENU
# =====================================================================
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

# lagcheck.ps1 became option 4: move its shortcut to this script and remove its local copy
function Move-Lagcheck {
    try {
        if ((Test-Path -LiteralPath $PROFILE) -and @(Get-Content -LiteralPath $PROFILE) -match '^function lagcheck ') {
            @(Get-Content -LiteralPath $PROFILE) | Where-Object { $_ -notmatch '^function lagcheck ' } | Set-Content -LiteralPath $PROFILE
            if (-not (Test-Shortcut)) { Switch-Shortcut }
            Say "  Lag check is now option 4 here; type 'systeminfo' instead of 'lagcheck'." 'Yellow'
        }
        $old = Join-Path $env:LOCALAPPDATA 'lagcheck'
        if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Recurse -Force }
    } catch {}
}

Set-WindowRight
Clear-Host
Move-Lagcheck
while ($true) {
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        SYSTEM INFO COLLECTOR" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  This Windows laptop" 'White'
    Say "    2  -  Android phone (USB cable needed)" 'White'
    Say "    3  -  Both" 'White'
    Say "    4  -  Lag check (3-5 min)" 'White'
    Say "    5  -  Settings" 'White'
    Say "    0  -  Quit" 'White'
    Say ""
    $choice = Read-MenuKey 'Choose 0-5' '012345'
    if ($choice -eq '0') { [Environment]::Exit(0) }   # Quit closes this window
    if ($choice -ne '5') { break }
    if (@(Show-ShortcutSettings) -contains 'quit') { [Environment]::Exit(0) }
    Clear-Host
}

$OutDir = Select-OutputFolder
switch ($choice) {
    '1' { Get-LaptopReport }
    '2' { Get-PhoneReport }
    '3' { Get-LaptopReport; Get-PhoneReport }
    '4' { Get-LagReport }
}

Say ""
Say "  Reports are in: $OutDir" 'Green'
Say ""
# This window was opened just for the report, so pause before it closes
Say "  Press any key to close this window" 'DarkGray'
try { [void][Console]::ReadKey($true) } catch { [void][Console]::ReadLine() }
