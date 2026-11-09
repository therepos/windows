# irm "https://github.com/therepos/windows/raw/main/apps/systeminfo.ps1?$(Get-Random)" | iex
# Purpose: Write a configuration/health report of a Windows laptop and/or Android phone to a text file
# =============================================================================
# Usage:
#   1 = Windows laptop
#   2 = Android phone   USB cable, USB debugging on
#   3 = Both
#   4 = Settings        'systeminfo' shortcut (runs a local copy in %LOCALAPPDATA%\systeminfo)
#   0 = Quit
#
# Note:
#   Needs admin (asks once)
#   Output: Data System <model> yyyyMMdd-HHmm.txt, Data Phone <model> yyyyMMdd-HHmm.txt (folder asked, default Downloads)
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
$SectionTimeout = 90
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
'@

function Save-Report { if ($script:OutPath) { [IO.File]::WriteAllLines($script:OutPath, [string[]]$script:Out) } }

function Get-SectionRunspace([hashtable]$Vars) {
    if (-not $script:Rs -or $script:Rs.RunspaceStateInfo.State -ne 'Opened') {
        $script:Rs = [runspacefactory]::CreateRunspace()
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
function Invoke-Collect([string]$Label, [scriptblock]$Body, [hashtable]$Vars) {
    $script:Out.Add('')
    $script:Out.Add("========== $Label ==========")
    $ps = [PowerShell]::Create()
    $ps.Runspace = Get-SectionRunspace $Vars
    [void]$ps.AddScript("& {`n$Body`n} | Out-String -Width 220")
    $state = 'ok'
    $text = ''
    $h = $ps.BeginInvoke()
    if ($h.AsyncWaitHandle.WaitOne($SectionTimeout * 1000)) {
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
    switch ($state) {
        'ok'      { Say "  [ OK ] $Label" 'Green' }
        'timeout' {
            $script:Out.Add("(timed out after ${SectionTimeout}s - skipped)")
            Say "  [WARN] $Label - timed out after ${SectionTimeout}s" 'Yellow'
        }
        default   {
            $script:Out.Add('(unavailable or command failed)')
            Say "  [WARN] $Label - unavailable, skipped" 'Yellow'
        }
    }
    Save-Report
}

# $Sections: label, scriptblock, label, scriptblock, ...
function Write-Report([string]$Path, [string]$Title, [string]$Subject, [string]$Hint, [hashtable]$Vars, [object[]]$Sections) {
    $script:Out.Clear()
    $script:OutPath = $Path
    $script:Out.Add("==================== $Title REPORT: $(Get-Date -Format 'yyyy-MM-dd HH:mm') ====================")
    $script:Out.Add($Subject)
    Save-Report
    Say ""
    Say "  [INFO] Writing $Title report to: $Path" 'Cyan'
    Say "  [INFO] $Hint" 'Cyan'
    for ($i = 0; $i -lt $Sections.Count; $i += 2) { Invoke-Collect $Sections[$i] $Sections[$i + 1] $Vars }
    $script:Out.Add('')
    $script:Out.Add('==================== END OF REPORT ====================')
    Save-Report
    Say "  [ OK ] Report complete: $Path" 'Green'
    Say "  [INFO] Contains serials, MACs and IPs - review before sharing." 'Yellow'
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
        'Takes 1-2 minutes.' @{ ModelName = $name } $LaptopSections
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
        'Takes about a minute. Keep the phone plugged in and unlocked.' @{ Adb = $adb; Serial = $serial } $PhoneSections
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
Set-WindowRight
Clear-Host
while ($true) {
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        SYSTEM INFO COLLECTOR" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  This Windows laptop" 'White'
    Say "    2  -  Android phone (USB cable needed)" 'White'
    Say "    3  -  Both" 'White'
    Say "    4  -  Settings" 'White'
    Say "    0  -  Quit" 'White'
    Say ""
    $choice = Read-MenuKey 'Choose 0-4' '01234'
    if ($choice -eq '0') { [Environment]::Exit(0) }   # Quit closes this window
    if ($choice -ne '4') { break }
    if (@(Show-ShortcutSettings) -contains 'quit') { [Environment]::Exit(0) }
    Clear-Host
}

$OutDir = Select-OutputFolder
switch ($choice) {
    '1' { Get-LaptopReport }
    '2' { Get-PhoneReport }
    '3' { Get-LaptopReport; Get-PhoneReport }
}

Say ""
Say "  Reports are in: $OutDir" 'Green'
Say ""
# This window was opened just for the report, so pause before it closes
Say "  Press any key to close this window" 'DarkGray'
try { [void][Console]::ReadKey($true) } catch { [void][Console]::ReadLine() }
