# irm "https://github.com/therepos/windows/raw/main/apps/lagcheck.ps1?$(Get-Random)" | iex
# Purpose: Find what makes a Windows PC laggy (slow windows, Explorer, Start search) and save one report
# =============================================================================
# Usage:
#   1 = Run the check   3-5 minutes, changes nothing
#   0 = Quit
#
# Note:
#   Needs admin (asks once)
#   Samples CPU, memory, disk and processes, times real actions, then reads Search, storage, drivers, power, startup and event logs
#   Output: Diag Lag <model> yyyyMMdd-HHmm.txt (folder asked, default Downloads); FLAGS at the top are the detected problems
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/lagcheck.ps1'

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
$Stamp = Get-Date -Format 'yyyyMMdd-HHmm'

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

function ConvertTo-OneWord([string]$s) {
    $s = ($s -replace '[\\/:*?"<>|]', '').Trim() -replace '\s+', '-'
    if ($s) { $s } else { 'Unknown' }
}

function Get-LaptopModel {
    $cs = Get-CimInstance Win32_ComputerSystem
    $model = $cs.Model
    if ($cs.Manufacturer -match 'LENOVO') {
        $friendly = (Get-CimInstance Win32_ComputerSystemProduct).Version
        if ($friendly) { $model = $friendly }
    }
    ConvertTo-OneWord $model
}

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
#  REPORT HELPERS
# =====================================================================
$script:Report     = New-Object System.Collections.Generic.List[string]
$script:Flags      = New-Object System.Collections.Generic.List[string]
$script:ReportPath = $null

function Add-Line([string]$text = '') { [void]$script:Report.Add($text) }

# Any object or table -> plain text lines in the report
function Add-Text($obj, [int]$width = 200) {
    if ($null -eq $obj) { Add-Line '(none)'; return }
    $s = ($obj | Out-String -Width $width)
    $s = ($s -replace '(?m)[ \t]+$', '').Trim("`r", "`n")
    if (-not $s.Trim()) { Add-Line '(none)'; return }
    foreach ($l in ($s -split "`r?`n")) { [void]$script:Report.Add($l) }
}

function Add-Flag([string]$text) { if (-not $script:Flags.Contains($text)) { [void]$script:Flags.Add($text) } }

function Save-Report {
    if ($script:ReportPath) { [IO.File]::WriteAllLines($script:ReportPath, [string[]]$script:Report) }
}

# Runs one section; a failure inside is written to the report, not fatal
function Invoke-Section([string]$Title, [string]$Progress, [scriptblock]$Body) {
    Say "  $Progress" 'DarkGray'
    Add-Line ''; Add-Line "=== $Title ==="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try { [void](& $Body) } catch { Add-Line "ERROR in this section: $($_.Exception.Message)" }
    Add-Line ("(section took {0:n1} s)" -f $sw.Elapsed.TotalSeconds)
    Save-Report
}

function Get-RegValue([string]$Path, [string]$Name) {
    try { (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name } catch { $null }
}

function Show-RegValues([string]$Path, [string[]]$Names) {
    $rows = foreach ($n in $Names) {
        $v = Get-RegValue $Path $n
        [pscustomobject]@{ Key = $Path; Value = $n; Data = $(if ($null -eq $v) { '(not set)' } else { "$v" }) }
    }
    Add-Text ($rows | Format-Table -AutoSize)
}

function Short([string]$s, [int]$n = 300) {
    if (-not $s) { return '' }
    $s = ($s -replace '\s+', ' ').Trim()
    if ($s.Length -gt $n) { $s.Substring(0, $n) + '...' } else { $s }
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
$script:CounterMap = $null
function Get-CounterMap {
    if ($null -ne $script:CounterMap) { return $script:CounterMap }
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
    $script:CounterMap = $map
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

# =====================================================================
#  THE CHECK
# =====================================================================
function Start-LagCheck([string]$OutDir) {
    $model = Get-LaptopModel
    $script:ReportPath = Join-Path $OutDir "Diag Lag $model $Stamp.txt"
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $cores = [int]$env:NUMBER_OF_PROCESSORS
    $sysDrive = $env:SystemDrive.Substring(0, 1)
    $isWin11 = [int]$os.BuildNumber -ge 22000

    Say ""
    Say "  Running the lag check. Leave the PC alone until it finishes (3-5 minutes)." 'Cyan'
    Say ""

    Add-Line "LAG CHECK REPORT"
    Add-Line "Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')   Computer: $env:COMPUTERNAME   User: $env:USERNAME"
    Add-Line "PowerShell: $($PSVersionTable.PSVersion)   Admin: $IsAdmin   64-bit process: $([Environment]::Is64BitProcess)"
    Add-Line "Contains the computer and user name. No serial numbers, MAC or IP addresses."
    Add-Line ""
    Add-Line "=== FLAGS (auto-detected, check these first) ==="
    Add-Line "[[FLAGS]]"
    Save-Report

    # ------------------------------------------------------------------
    Invoke-Section 'OS, HARDWARE, BOOT' '[1/18] OS and boot...' {
        $ntv = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $uptime = (Get-Date) - $os.LastBootUpTime
        Add-Text ([pscustomobject]@{
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
        Add-Line "Fast startup (HiberbootEnabled): $(if ($null -eq $fast) { '(not set)' } else { $fast })"
        Add-Line "--- Last boots (Kernel-Boot event 27: boot type 0x0 = full, 0x1 = fast startup, 0x2 = resume) ---"
        $boots = Get-Events @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Boot'; Id = 27 } 6
        Add-Text ($boots | Select-Object TimeCreated, @{n = 'Message'; e = { Short $_.Message 120 } } | Format-Table -AutoSize)
        if ($fast -eq 1 -and $boots -and -not ($boots | Where-Object { $_.Message -match '0x0\b' })) {
            Add-Flag "No full restart in the last $($boots.Count) boots (fast startup only). A real Restart clears the kernel session."
        }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'MEMORY AND PAGE FILE' '[2/18] Memory...' {
        $totalMB = $os.TotalVisibleMemorySize / 1KB
        $freeMB  = $os.FreePhysicalMemory / 1KB
        $usedPct = [math]::Round(100 - ($freeMB / $totalMB * 100))
        Add-Text ([pscustomobject]@{
            InstalledGB      = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
            UsableGB         = [math]::Round($totalMB / 1KB, 2)
            HardwareReservedMB = [math]::Round(($cs.TotalPhysicalMemory / 1MB) - $totalMB)
            FreeGB           = [math]::Round($freeMB / 1KB, 2)
            InUsePercent     = $usedPct
            CommitUsedGB     = [math]::Round(($os.TotalVirtualMemorySize - $os.FreeVirtualMemory) / 1MB, 2)
            CommitLimitGB    = [math]::Round($os.TotalVirtualMemorySize / 1MB, 2)
            AutoPageFile     = $cs.AutomaticManagedPagefile
        } | Format-List)
        Add-Line "--- Modules ---"
        Add-Text (Get-CimInstance Win32_PhysicalMemory | Select-Object BankLabel, @{n = 'GB'; e = { [math]::Round($_.Capacity / 1GB, 1) }}, Speed, ConfiguredClockSpeed, Manufacturer, PartNumber | Format-Table -AutoSize)
        Add-Line "--- Page file ---"
        Add-Text (Get-CimInstance Win32_PageFileUsage | Select-Object Name, AllocatedBaseSize, CurrentUsage, PeakUsage | Format-Table -AutoSize)
        $mc = Get-Process 'Memory Compression' -ErrorAction SilentlyContinue
        if ($mc) { Add-Line "Memory Compression process: $([math]::Round($mc.WorkingSet64 / 1MB)) MB" }
        if ($usedPct -ge 85) { Add-Flag "RAM is $usedPct % in use. The PC is short of memory (see top processes)." }
        $commitPct = ($os.TotalVirtualMemorySize - $os.FreeVirtualMemory) / $os.TotalVirtualMemorySize * 100
        if ($commitPct -ge 90) { Add-Flag "Commit charge is $([math]::Round($commitPct)) % of the limit. Page file nearly exhausted." }
        if (($cs.TotalPhysicalMemory / 1MB - $totalMB) -gt 1024) { Add-Flag "$([math]::Round(($cs.TotalPhysicalMemory / 1MB) - $totalMB)) MB of RAM is hardware reserved (BIOS / iGPU)." }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'POWER AND THERMAL' '[3/18] Power and thermal...' {
        $bat = Get-CimInstance Win32_Battery | Select-Object -First 1
        $onBattery = ($bat -and $bat.BatteryStatus -eq 1)
        Add-Line "Battery: $(if ($bat) { "present, status $($bat.BatteryStatus) (1 = discharging, 2 = on AC), charge $($bat.EstimatedChargeRemaining) %" } else { 'none (desktop)' })"
        Add-Text (& powercfg /getactivescheme 2>&1)
        $ovKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes'
        $ovNames = @{ '961cc777-2547-4f9d-8174-7d86181b8a7a' = 'Best power efficiency'; '00000000-0000-0000-0000-000000000000' = 'Balanced'; 'ded574b5-45a0-4f42-8737-46345c09c238' = 'Best performance'; '3af9b8d9-7c97-431d-ad78-34a8bfea439f' = 'Better performance' }
        foreach ($n in 'ActiveOverlayAcPowerScheme', 'ActiveOverlayDcPowerScheme') {
            $g = Get-RegValue $ovKey $n
            if ($g) { Add-Line ("Power mode ({0}): {1} {2}" -f $n.Substring(13, 2), $g, $(if ($ovNames["$g".ToLower()]) { "= $($ovNames["$g".ToLower()])" } else { '' })) }
        }
        Add-Line "--- Max processor state (PROCTHROTTLEMAX, hex: 0x64 = 100 %) ---"
        Add-Text ((& powercfg /query SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 2>&1) | Where-Object { $_ -match 'Power Setting|AC Power|DC Power' })
        Add-Line "--- Thermal zones ---"
        $tz = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
        if ($tz) {
            Add-Text ($tz | Select-Object InstanceName, @{n = 'TempC'; e = { [math]::Round($_.CurrentTemperature / 10 - 273.15, 1) }} | Format-Table -AutoSize)
            $hot = $tz | Where-Object { ($_.CurrentTemperature / 10 - 273.15) -ge 90 }
            if ($hot) { Add-Flag "A thermal zone reports $([math]::Round($hot[0].CurrentTemperature / 10 - 273.15)) C. Possible overheating / throttling." }
        } else { Add-Line '(not exposed by this firmware)' }
        if ($cpu.CurrentClockSpeed -and $cpu.MaxClockSpeed -and $cpu.CurrentClockSpeed -lt ($cpu.MaxClockSpeed * 0.5)) {
            Add-Flag "CPU reports $($cpu.CurrentClockSpeed) MHz of $($cpu.MaxClockSpeed) MHz max right now (possible throttling or power saving)."
        }
        if ($onBattery) { Add-Line 'NOTE: on battery; Windows limits performance and WinSAT will refuse to run.' }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'LIVE COUNTERS (10 s, PC should be idle)' '[4/18] Sampling CPU, memory and disk for 10 seconds...' {
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
        Add-Text ($rows | Format-Table -AutoSize)
        if (-not $rows) { Add-Line 'Get-Counter returned nothing (counters disabled or names not resolvable).'; return }
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
    }

    # ------------------------------------------------------------------
    Invoke-Section 'PROCESSES (5 s CPU sample)' '[5/18] Processes...' {
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
        $script:ProcRows = $rows
        Add-Line "Processes: $($procs.Count)   Threads: $(($rows | Measure-Object Threads -Sum).Sum)   Handles: $(($rows | Measure-Object Handles -Sum).Sum)"
        Add-Line "--- Top 15 by CPU (share of all $cores logical CPUs) ---"
        Add-Text ($rows | Where-Object { $_.'CPU%' -gt 0 } | Sort-Object 'CPU%' -Descending | Select-Object -First 15 | Format-Table -AutoSize)
        Add-Line "--- Top 15 by working set ---"
        Add-Text ($rows | Sort-Object WS_MB -Descending | Select-Object -First 15 | Format-Table -AutoSize)
        Add-Line "--- Top 10 by handle count ---"
        Add-Text ($rows | Sort-Object Handles -Descending | Select-Object -First 10 Name, PID, Handles, Threads, WS_MB, Company | Format-Table -AutoSize)
        Add-Line "--- Totals per program name (top 15 by memory) ---"
        Add-Text ($rows | Group-Object Name | ForEach-Object {
            [pscustomobject]@{ Name = $_.Name; Count = $_.Count; 'CPU%' = [math]::Round(($_.Group | Measure-Object 'CPU%' -Sum).Sum, 1); WS_MB = ($_.Group | Measure-Object WS_MB -Sum).Sum; Priv_MB = ($_.Group | Measure-Object Priv_MB -Sum).Sum }
        } | Sort-Object WS_MB -Descending | Select-Object -First 15 | Format-Table -AutoSize)
        Add-Line "--- Shell, search and system processes ---"
        $shellNames = '^(explorer|dwm|SearchIndexer|SearchHost|SearchApp|SearchUI|SearchProtocolHost|SearchFilterHost|StartMenuExperienceHost|ShellExperienceHost|RuntimeBroker|TextInputHost|msedgewebview2|WidgetService|Widgets|MsMpEng|NisSrv|sihost|ctfmon|OneDrive|Teams|ms-teams|csrss|lsass|System|Memory Compression|audiodg|WmiPrvSE|MoUsoCoreWorker|TiWorker|TrustedInstaller|CompatTelRunner|backgroundTaskHost|PhoneExperienceHost|CrossDeviceService|Cortana|fontdrvhost|spoolsv|SgrmBroker|SecurityHealthService|smartscreen)$'
        Add-Text ($rows | Where-Object { $_.Name -match $shellNames } | Sort-Object Name | Format-Table -AutoSize)
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
        Add-Line "--- Services inside the busiest svchost processes ---"
        $svc = Get-CimInstance Win32_Service
        $topSvc = $rows | Where-Object Name -eq 'svchost' | Sort-Object @{e = { $_.'CPU%' }; Descending = $true }, @{e = 'WS_MB'; Descending = $true } | Select-Object -First 6
        foreach ($t in $topSvc) {
            $names = ($svc | Where-Object ProcessId -eq $t.PID | ForEach-Object { $_.Name }) -join ', '
            Add-Line ("PID {0,-6} CPU {1,5}%  WS {2,5} MB  -> {3}" -f $t.PID, $t.'CPU%', $t.WS_MB, $names)
        }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'TIMING TESTS (real actions)' '[6/18] Timing real actions: files, process start, windows, search...' {
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
                Add-Line ("File open+read ({0}): {1} files, avg {2} ms, max {3} ms" -f $set.Name, $n, $avgMs, [math]::Round($max, 1))
                if ($avgMs -ge 2) { Add-Flag "Opening small files takes $avgMs ms on average ($($set.Name)). Slow disk or antivirus scanning every file." }
            } else { Add-Line "File open test ($($set.Name)): no files read" }
        }
        $ms = Measure-Ms { Get-ChildItem -LiteralPath "$env:windir\System32" -Force }
        Add-Line "List System32 folder (~4000 entries): $ms ms"
        $ms = Measure-Ms { Get-ChildItem -LiteralPath $env:USERPROFILE -Force }
        Add-Line "List user profile folder: $ms ms"

        $times = foreach ($i in 1..3) { Measure-Ms { Start-Process cmd.exe -ArgumentList '/c exit' -WindowStyle Hidden -Wait } }
        Add-Line "Start a process (cmd /c exit) x3: $($times -join ' / ') ms  (healthy: under 150 ms; high = security software hooking every start)"
        if (($times | Measure-Object -Minimum).Minimum -ge 500) { Add-Flag "Starting a trivial process takes $(($times | Measure-Object -Minimum).Minimum) ms. Security / monitoring software is slowing every launch." }

        $times = foreach ($i in 1..3) {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $p = Start-Process charmap.exe -PassThru
            while ($p -and -not $p.HasExited -and $p.MainWindowHandle -eq 0 -and $sw.ElapsedMilliseconds -lt 15000) { Start-Sleep -Milliseconds 25; $p.Refresh() }
            $t = $sw.ElapsedMilliseconds
            if ($p) { try { Stop-Process -Id $p.Id -Force } catch {} }
            $t
        }
        Add-Line "Open a small Win32 window (charmap) x3: $($times -join ' / ') ms  (healthy: under 300 ms after the first)"
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
            Add-Line "Open an Explorer window ($env:SystemDrive\) x2: $($times -join ' / ') ms  (healthy: under 800 ms)"
            $nums = @($times | Where-Object { $_ -is [long] -or $_ -is [int] })
            if ($times -contains 'timeout') { Add-Flag "An Explorer window did not open within 20 s." }
            elseif ($nums -and ($nums | Measure-Object -Minimum).Minimum -ge 3000) { Add-Flag "Opening an Explorer window takes $(($nums | Measure-Object -Minimum).Minimum) ms (shell extensions, mapped drives, OneDrive or disk)." }
        } catch { Add-Line "Explorer window test failed: $($_.Exception.Message)" }

        Add-Line "--- Windows Search index query (what Start search uses) ---"
        try {
            $conn = New-Object System.Data.OleDb.OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
            $open = Measure-Ms { $conn.Open() }
            Add-Line "Connect to index: $open ms"
            $queries = @(
                @{ Name = 'file name starts with notepad'; Sql = "SELECT TOP 50 System.ItemPathDisplay FROM SYSTEMINDEX WHERE System.FileName LIKE 'notepad%'" },
                @{ Name = 'word search "setup"';           Sql = "SELECT TOP 50 System.ItemPathDisplay FROM SYSTEMINDEX WHERE CONTAINS(System.FileName, 'setup')" },
                @{ Name = 'anything in user profile';      Sql = "SELECT TOP 50 System.ItemPathDisplay FROM SYSTEMINDEX WHERE SCOPE='file:$env:USERPROFILE'" },
                @{ Name = 'installed app (Notepad)';       Sql = "SELECT TOP 20 System.ItemNameDisplay FROM SYSTEMINDEX WHERE System.Kind = 'program' AND CONTAINS(System.ItemNameDisplay, 'notepad')" }
            )
            foreach ($q in $queries) {
                $n = 0; $err = ''
                $ms = Measure-Ms {
                    try {
                        $cmd = $conn.CreateCommand(); $cmd.CommandText = $q.Sql; $cmd.CommandTimeout = 30
                        $rd = $cmd.ExecuteReader(); while ($rd.Read()) { $n++ }; $rd.Close()
                    } catch { $err = $_.Exception.Message }
                }
                Add-Line ("Query '{0}': {1} ms, {2} rows {3}" -f $q.Name, $ms, $n, $(if ($err) { "ERROR $(Short $err 150)" } else { '' }))
                if ($ms -ge 2000) { Add-Flag "Search index query '$($q.Name)' took $ms ms (index busy or broken)." }
                if ($err) { Add-Flag "Search index query '$($q.Name)' failed: $(Short $err 100)" }
                elseif ($n -eq 0 -and $q.Name -eq 'anything in user profile') { Add-Flag "The search index returns nothing for the user profile: user files are not indexed (index empty or rebuilding)." }
            }
            $conn.Close()
        } catch { Add-Line "Could not open the Windows Search index: $($_.Exception.Message)"; Add-Flag "Windows Search index cannot be opened: $(Short $_.Exception.Message 120)" }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'WINDOWS SEARCH STATE' '[7/18] Windows Search...' {
        $ws = Get-CimInstance Win32_Service -Filter "Name='WSearch'"
        Add-Line "WSearch service: $(if ($ws) { "$($ws.State), start mode $($ws.StartMode)" } else { 'NOT INSTALLED' })"
        if (-not $ws -or $ws.State -ne 'Running') { Add-Flag "Windows Search service (WSearch) is $(if ($ws) { $ws.State } else { 'missing' }). Start search cannot find files or apps without it." }
        $wsKey = 'HKLM:\SOFTWARE\Microsoft\Windows Search'
        Show-RegValues $wsKey @('SetupCompletedSuccessfully', 'DataDirectory')
        $enh = Get-RegValue "$wsKey\Gather\Windows\SystemIndex" 'EnableFindMyFiles'
        Add-Line "Enhanced indexing (whole PC, EnableFindMyFiles): $(if ($null -eq $enh) { '(classic)' } else { $enh })"
        if ($enh -eq 1) { Add-Flag "Search is in Enhanced mode (indexes the whole PC). Heavy on CPU/disk; Classic is usually enough." }
        $dataDir = Get-RegValue $wsKey 'DataDirectory'
        if ($dataDir) {
            $dataDir = [Environment]::ExpandEnvironmentVariables($dataDir)
            Add-Line "--- Index files in $dataDir ---"
            Add-Text (Get-ChildItem -LiteralPath (Join-Path $dataDir 'Applications\Windows') -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 10MB -or $_.Name -match '^Windows\.(edb|db)' } | Select-Object Name, @{n = 'MB'; e = { [math]::Round($_.Length / 1MB) }}, LastWriteTime | Format-Table -AutoSize)
            $db = Get-ChildItem -LiteralPath (Join-Path $dataDir 'Applications\Windows') -File -ErrorAction SilentlyContinue | Where-Object Name -match '^Windows\.(edb|db)$'
            if ($db -and $db.Length -gt 4GB) { Add-Flag "Search index database is $([math]::Round($db.Length / 1GB, 1)) GB. Very large; a rebuild with fewer locations helps." }
        }

        Add-Line "--- Catalog status (from the indexer itself) ---"
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
        if ($kv.Count -eq 0) { Add-Line "(could not query: $(Short ($out -join ' ') 300))" }
        else {
            $st = [int]$kv['Status']; $rs = [int]$kv['PausedReason']
            Add-Line "Indexer version : $($kv['IndexerVersion'])"
            Add-Line "Status          : $st = $(if ($st -lt $statusNames.Count) { $statusNames[$st] } else { '?' })   paused reason: $rs = $(if ($rs -lt $reasonNames.Count) { $reasonNames[$rs] } else { '?' })"
            Add-Line "Items in index  : $($kv['Items'])   waiting to index: $($kv['ToIndex'])   notifications: $($kv['Notifications'])   high priority: $($kv['HighPriority'])"
            Add-Line "Indexing now    : $($kv['Indexing'])"
            if ($st -in 2, 3) { Add-Flag "Search index is $($statusNames[$st]) ($($kv['ToIndex']) items queued). Search gives partial results until it finishes." }
            if ($st -eq 1) { Add-Flag "Search indexer is paused ($($reasonNames[$rs])). Results stay stale." }
            if ([int]$kv['ToIndex'] -gt 50000) { Add-Flag "$($kv['ToIndex']) items are waiting to be indexed." }
            if ([int]$kv['Items'] -lt 5000) { Add-Flag "Search index holds only $($kv['Items']) items: it was reset recently or never finished." }
        }

        Add-Line "--- Web / cloud search settings (affects the Start search box) ---"
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' @('BingSearchEnabled', 'CortanaConsent', 'SearchboxTaskbarMode')
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings' @('IsDynamicSearchBoxEnabled', 'IsAADCloudSearchEnabled', 'IsMSACloudSearchEnabled', 'IsDeviceSearchHistoryEnabled')
        Show-RegValues 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' @('DisableSearchBoxSuggestions')
        Show-RegValues 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' @('DisableWebSearch', 'ConnectedSearchUseWeb', 'AllowCortana', 'AllowCloudSearch')
        $webOff = (Get-RegValue 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' 'DisableSearchBoxSuggestions') -eq 1 -or (Get-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' 'DisableWebSearch') -eq 1
        Add-Line "Web results in Start search: $(if ($webOff) { 'disabled' } else { 'enabled (search waits on Bing; slow internet or proxy = spinning with no results)' })"

        Add-Line "--- Search / Start app packages ---"
        $pk = foreach ($n in 'Microsoft.Windows.Search', 'MicrosoftWindows.Client.CBS', 'MicrosoftWindows.Client.Core', 'Microsoft.Windows.StartMenuExperienceHost', 'Microsoft.Windows.ShellExperienceHost', 'Microsoft.Windows.Cortana', 'MicrosoftWindows.Client.WebExperience') {
            Get-AppxPackage -Name $n -ErrorAction SilentlyContinue | Select-Object Name, Version, Status, @{n = 'Installed'; e = { $_.InstallLocation -and (Test-Path -LiteralPath $_.InstallLocation) } }
        }
        Add-Text ($pk | Format-Table -AutoSize)
        foreach ($p in $pk) { if ("$($p.Status)" -ne 'Ok') { Add-Flag "App package $($p.Name) status is '$($p.Status)' (Start/search UI package damaged)." } }

        Add-Line "--- Search events, last 7 days (Application log) ---"
        $ev = Get-Events @{ LogName = 'Application'; ProviderName = 'Microsoft-Windows-Search', 'Microsoft-Windows-Search-ProfileNotify', 'Microsoft-Windows-Search-Core'; StartTime = (Get-Date).AddDays(-7) } 20
        Add-Text ($ev | Select-Object TimeCreated, Id, LevelDisplayName, @{n = 'Message'; e = { Short $_.Message 220 } } | Format-Table -AutoSize -Wrap)
        if ($ev | Where-Object { $_.Id -in 1006, 3083, 3084, 3105, 7042, 7040 -or $_.Message -match 'corrupt|reset|rebuil' }) { Add-Flag "Search log reports index corruption / reset / rebuild in the last 7 days." }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'STORAGE' '[8/18] Storage health and speed (WinSAT may take a minute)...' {
        $pds = @(Get-PhysicalDisk)
        Add-Text ($pds | Select-Object DeviceId, FriendlyName, MediaType, BusType, @{n = 'SizeGB'; e = { [math]::Round($_.Size / 1GB) }}, HealthStatus, OperationalStatus, FirmwareVersion | Format-Table -AutoSize)
        $sysDiskNo = (Get-Partition -DriveLetter $sysDrive -ErrorAction SilentlyContinue).DiskNumber
        $sysPd = $pds | Where-Object { "$($_.DeviceId)" -eq "$sysDiskNo" } | Select-Object -First 1
        Add-Line "System drive $sysDrive`: is on disk $sysDiskNo = $($sysPd.FriendlyName) [$($sysPd.MediaType), $($sysPd.BusType)]"
        if ($sysPd -and $sysPd.MediaType -eq 'HDD') { Add-Flag "Windows runs from a mechanical hard disk (HDD). This alone causes slow windows and search; an SSD is the fix." }
        if ($sysPd -and $sysPd.BusType -eq 'USB') { Add-Flag "Windows runs from a USB disk." }
        foreach ($p in $pds) { if ("$($p.HealthStatus)" -ne 'Healthy') { Add-Flag "Disk '$($p.FriendlyName)' health is $($p.HealthStatus)." } }
        Add-Line "--- Reliability counters (SMART) ---"
        $rel = $pds | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
        Add-Text ($rel | Select-Object DeviceId, Temperature, TemperatureMax, Wear, PowerOnHours, ReadErrorsTotal, ReadErrorsUncorrected, WriteErrorsTotal, WriteErrorsUncorrected, ReadLatencyMax, WriteLatencyMax, FlushLatencyMax | Format-Table -AutoSize)
        foreach ($r in $rel) {
            if ($r.Wear -ge 80) { Add-Flag "SSD wear on disk $($r.DeviceId) is $($r.Wear) %." }
            if ($r.ReadErrorsUncorrected -gt 0 -or $r.WriteErrorsUncorrected -gt 0) { Add-Flag "Disk $($r.DeviceId) reports uncorrected read/write errors (failing drive)." }
            if ($r.Temperature -ge 65) { Add-Flag "Disk $($r.DeviceId) is at $($r.Temperature) C." }
        }
        Add-Line "--- Volumes ---"
        $vols = Get-Volume | Where-Object DriveLetter
        Add-Text ($vols | Select-Object DriveLetter, FileSystemLabel, FileSystem, @{n = 'SizeGB'; e = { [math]::Round($_.Size / 1GB, 1) }}, @{n = 'FreeGB'; e = { [math]::Round($_.SizeRemaining / 1GB, 1) }}, @{n = 'FreePct'; e = { if ($_.Size) { [math]::Round($_.SizeRemaining / $_.Size * 100) } }}, HealthStatus, DriveType | Format-Table -AutoSize)
        $sv = $vols | Where-Object DriveLetter -eq $sysDrive
        if ($sv -and ($sv.SizeRemaining -lt 15GB -or ($sv.SizeRemaining / $sv.Size) -lt 0.10)) { Add-Flag "$sysDrive`: has only $([math]::Round($sv.SizeRemaining / 1GB, 1)) GB free. Windows needs 15-20 GB free to stay responsive." }
        Add-Line "TRIM (DisableDeleteNotify, 0 = TRIM on): $((& fsutil behavior query DisableDeleteNotify 2>&1) -join ' | ')"
        Add-Line "--- Folder sizes ---"
        Add-Text (@(
            Get-DirSize $env:TEMP
            Get-DirSize "$env:windir\Temp"
            Get-DirSize "$env:windir\SoftwareDistribution\Download"
            Get-DirSize "$env:ProgramData\Microsoft\Windows\WER"
            Get-DirSize "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
            Get-DirSize "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"
        ) | Format-Table -AutoSize)
        Add-Line "--- Last WinSAT assessment (scores 1-9.9) ---"
        Add-Text (Get-CimInstance Win32_WinSAT | Select-Object WinSATAssessmentState, CPUScore, MemoryScore, DiskScore, GraphicsScore, D3DScore, TimeTaken | Format-List)
        Add-Line "--- WinSAT disk test on $sysDrive`: (random 16 KB read, then sequential read) ---"
        foreach ($mode in '-ran -read', '-seq -read') {
            $o = @(& winsat.exe disk $mode.Split(' ') -drive $sysDrive 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match 'Disk|Total Run Time|Error|battery|cannot|not supported|remote' })
            Add-Text $o
            $m = $o | Where-Object { $_ -match 'Random 16\.0 Read\s+([\d.]+) MB/s' } | Select-Object -First 1
            if ($m -and $m -match '([\d.]+) MB/s') { $mbs = [double]$matches[1]; if ($mbs -lt 2) { Add-Flag "Random 16 KB read speed is $mbs MB/s (HDD-class). A healthy SSD does 20+ MB/s here." } }
            $m = $o | Where-Object { $_ -match 'Sequential 64\.0 Read\s+([\d.]+) MB/s' } | Select-Object -First 1
            if ($m -and $m -match '([\d.]+) MB/s') { $mbs = [double]$matches[1]; if ($mbs -lt 150) { Add-Flag "Sequential read speed is $mbs MB/s (slow disk, bad cable or SATA mode)." } }
        }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'GRAPHICS AND VISUAL SETTINGS' '[9/18] Graphics...' {
        $gpus = Get-CimInstance Win32_VideoController
        Add-Text ($gpus | Select-Object Name, Status, DriverVersion, DriverDate, @{n = 'VRAM_MB'; e = { [math]::Round($_.AdapterRAM / 1MB) }}, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate, VideoProcessor | Format-List)
        foreach ($g in $gpus) {
            if ($g.Name -match 'Basic Display|Basic Render') { Add-Flag "'$($g.Name)' is active: no real graphics driver installed. Everything on screen is drawn by the CPU." }
            if ("$($g.Status)" -ne 'OK') { Add-Flag "Graphics adapter '$($g.Name)' status is $($g.Status)." }
            if ($g.DriverDate -and $g.DriverDate -lt (Get-Date).AddYears(-4)) { Add-Line "NOTE: driver for $($g.Name) dates from $($g.DriverDate.ToString('yyyy-MM-dd'))." }
        }
        $dwm = $script:ProcRows | Where-Object Name -eq 'dwm'
        if ($dwm) { Add-Line "Desktop Window Manager: CPU $($dwm.'CPU%') %, $($dwm.WS_MB) MB, $($dwm.Handles) handles" }
        if ($dwm -and $dwm.WS_MB -ge 800) { Add-Flag "dwm.exe (Desktop Window Manager) uses $($dwm.WS_MB) MB: graphics driver memory leak, restart or update the driver." }
        Add-Line "--- Settings ---"
        Show-RegValues 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers' @('HwSchMode')
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' @('VisualFXSetting')
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' @('EnableTransparency')
        Show-RegValues 'HKCU:\Control Panel\Desktop\WindowMetrics' @('MinAnimate')
        Show-RegValues 'HKCU:\Control Panel\Desktop' @('UserPreferencesMask', 'MenuShowDelay')
        Add-Line "(VisualFXSetting: 0 = let Windows choose, 1 = best appearance, 2 = best performance, 3 = custom)"
        Add-Line "--- Explorer / Start preferences ---"
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' @('LaunchTo', 'Start_TrackDocs', 'Start_IrisRecommendations', 'ShowRecent', 'ShowFrequent', 'ShowCloudFilesInQuickAccess', 'TaskbarDa', 'TaskbarMn', 'ShowCopilotButton')
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer' @('ShowRecent', 'ShowFrequent', 'ShowCloudFilesInQuickAccess')
        Add-Line "(LaunchTo: 1 = This PC, 2 = Home/Quick access. Home with cloud files on is a common slow-open cause.)"
    }

    # ------------------------------------------------------------------
    Invoke-Section 'SECURITY SOFTWARE AND KERNEL FILTERS' '[10/18] Security software...' {
        $mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($mp) { Add-Text ($mp | Select-Object AMRunningMode, RealTimeProtectionEnabled, BehaviorMonitorEnabled, AntivirusSignatureAge, QuickScanAge, FullScanAge, QuickScanStartTime, QuickScanEndTime, FullScanStartTime, FullScanEndTime, IsTamperProtected | Format-List) }
        else { Add-Line 'Defender status not available (another product may own real-time protection).' }
        Add-Line "--- Registered security products ---"
        $av = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue)
        Add-Text ($av | Select-Object displayName, @{n = 'state'; e = { '0x{0:X6}' -f $_.productState }}, timestamp, pathToSignedProductExe | Format-Table -AutoSize)
        $third = @($av | Where-Object { $_.displayName -notmatch 'Windows Defender|Microsoft Defender' })
        if ($third.Count -ge 1) { Add-Line "Third-party antivirus: $(($third | ForEach-Object { $_.displayName }) -join ', ')" }
        if ($third.Count -ge 2 -or ($third.Count -ge 1 -and $mp -and $mp.RealTimeProtectionEnabled)) { Add-Flag "More than one real-time antivirus appears active ($(($av | ForEach-Object { $_.displayName }) -join ', ')). Two scanners double every file access." }
        Add-Text (Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct -ErrorAction SilentlyContinue | Select-Object displayName | Format-Table -AutoSize)
        $msmp = $script:ProcRows | Where-Object Name -eq 'MsMpEng'
        if ($msmp -and $msmp.'CPU%' -ge 15) { Add-Flag "Defender (MsMpEng) was using $($msmp.'CPU%') % CPU: a scan is running or something keeps triggering scans." }
        Add-Line "--- Minifilter drivers (every file access passes through these; AV, backup, cloud sync) ---"
        Add-Text (& fltmc.exe filters 2>&1)
        Add-Line "--- Virtualization-based security / Memory integrity ---"
        $dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction SilentlyContinue
        if ($dg) {
            Add-Line "VBS status: $($dg.VirtualizationBasedSecurityStatus) (2 = running)   services running: $($dg.SecurityServicesRunning -join ',') (1 = Credential Guard, 2 = HVCI/Memory integrity)"
            $gen = $null
            if ($cpu.Name -match 'i[3579]-(\d{3,5})') { $d = $matches[1]; $gen = [int]$(if ($d.Length -eq 5 -or ($d.Length -eq 4 -and $d[0] -eq '1')) { $d.Substring(0, 2) } elseif ($d.Length -eq 4) { $d.Substring(0, 1) } else { 1 }) }
            if ($dg.SecurityServicesRunning -contains 2 -and $gen -and $gen -lt 8) { Add-Flag "Memory integrity (HVCI) is on with a $gen`th-gen Intel CPU (no MBEC). Costs noticeable performance on this CPU." }
        } else { Add-Line '(not available)' }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'STARTUP ITEMS, TASKS, SERVICES' '[11/18] Startup items and services...' {
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
        Add-Text ($su | Sort-Object State, Name | Format-Table -AutoSize)
        $enabledCount = @($su | Where-Object State -eq 'enabled').Count
        Add-Line "Enabled startup entries: $enabledCount"
        if ($enabledCount -ge 15) { Add-Flag "$enabledCount programs start with Windows. Each one costs memory and CPU at every logon." }
        Add-Line "--- Non-Microsoft scheduled tasks that are enabled ---"
        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskPath -notlike '\Microsoft\*' -and $_.State -ne 'Disabled' }
        Add-Text ($tasks | Select-Object TaskName, TaskPath, State, @{n = 'Triggers'; e = { ($_.Triggers | ForEach-Object { $_.CimClass.CimClassName -replace 'MSFT_Task|Trigger', '' }) -join ',' }}, @{n = 'Action'; e = { Short (($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; ') 100 }} | Format-Table -AutoSize)
        Add-Line "--- Scheduled tasks running right now ---"
        Add-Text (Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object State -eq 'Running' | Select-Object TaskName, TaskPath | Format-Table -AutoSize)
        Add-Line "--- Running services not from the Windows folder ---"
        $svcs = Get-CimInstance Win32_Service | Where-Object { $_.State -eq 'Running' -and $_.PathName -notmatch '\\Windows\\' }
        Add-Text ($svcs | Select-Object Name, DisplayName, StartMode, ProcessId, @{n = 'Path'; e = { Short $_.PathName 90 }} | Sort-Object Name | Format-Table -AutoSize)
        Add-Line "Running services total: $(@(Get-CimInstance Win32_Service | Where-Object State -eq 'Running').Count)   non-Windows: $(@($svcs).Count)"
        Add-Line "--- Key Windows services ---"
        Add-Text (Get-CimInstance Win32_Service | Where-Object { $_.Name -in 'WSearch', 'SysMain', 'Themes', 'DiagTrack', 'wuauserv', 'DoSvc', 'BITS', 'WinDefend', 'Spooler', 'OneSyncSvc', 'ShellHWDetection', 'StorSvc', 'TrustedInstaller', 'cbdhsvc', 'UsoSvc', 'Schedule', 'FontCache', 'TabletInputService', 'TextInputManagementService', 'CDPSvc' -or $_.Name -like 'cbdhsvc_*' } | Select-Object Name, State, StartMode | Sort-Object Name | Format-Table -AutoSize)
    }

    # ------------------------------------------------------------------
    Invoke-Section 'INSTALLED SOFTWARE' '[12/18] Installed software...' {
        $apps = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object DisplayName | Sort-Object DisplayName -Unique
        Add-Line "Installed programs: $(@($apps).Count)   Store apps (this user): $(@(Get-AppxPackage -ErrorAction SilentlyContinue).Count)"
        Add-Line "--- Programs known to run agents, scanners, overlays or 'optimizers' ---"
        $heavy = 'McAfee|Norton|Avast|AVG|Bitdefender|Kaspersky|ESET|Malwarebytes|Webroot|Trend Micro|Sophos|CrowdStrike|SentinelOne|Carbon Black|Cylance|Cortex|CCleaner|Advanced SystemCare|Driver Booster|IObit|Wallpaper Engine|Rainmeter|SupportAssist|HP Support|HP Wolf|Omen|Lenovo Vantage|Armoury|MSI Center|Dragon Center|iCUE|Razer|Logitech|LGHUB|Nahimic|Dropbox|OneDrive|Google Drive|iCloud|Teams|Zoom|Discord|Steam|Epic Games|Adobe|Creative Cloud|VMware|VirtualBox|Docker|Citrix|GlobalProtect|Cisco|FortiClient|Zscaler|Netskope|Tanium|Splashtop|TeamViewer|AnyDesk|Webex|Slack|Spotify|Bing Wallpaper|Booster|Optimizer|Cleaner|Toolbar|Updater|McAfee|WebAdvisor|Killer|Intel Driver|Dell Optimizer|Dell Digital|Waves|DTS|Dolby|Alienware|Acer|ASUS|Samsung Magician|Backup|Acronis|Carbonite|Norton 360|Avira|Comodo|360 Total|Baidu|Tencent|Kingsoft|WPS'
        Add-Text ($apps | Where-Object { $_.DisplayName -match $heavy -or $_.Publisher -match $heavy } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Format-Table -AutoSize)
        Add-Line "--- Installed or updated in the last 30 days ---"
        $recent = $apps | Where-Object {
            $d = [datetime]::MinValue
            "$($_.InstallDate)" -match '^\d{8}$' -and [datetime]::TryParseExact("$($_.InstallDate)", 'yyyyMMdd', $null, 'None', [ref]$d) -and $d -gt (Get-Date).AddDays(-30)
        }
        Add-Text ($recent | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Format-Table -AutoSize)
        Add-Line "--- Full list ---"
        Add-Text ($apps | Select-Object DisplayName, DisplayVersion, Publisher | Format-Table -AutoSize)
    }

    # ------------------------------------------------------------------
    Invoke-Section 'DRIVERS AND DEVICES' '[13/18] Drivers and devices...' {
        Add-Line "--- Devices with problems ---"
        $bad = Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'OK' -and $_.Present }
        if ($bad) { Add-Text ($bad | Select-Object FriendlyName, Status, Class, @{n = 'Problem'; e = { $_.ProblemDescription }} | Format-Table -AutoSize); foreach ($b in $bad) { Add-Flag "Device problem: $($b.FriendlyName) ($($b.Class)) status $($b.Status)." } }
        else { Add-Line 'None.' }
        Add-Line "--- Display, storage, network, chipset drivers ---"
        $drv = Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceName -and $_.DeviceClass -in 'DISPLAY', 'NET', 'HDC', 'SCSIADAPTER', 'DISKDRIVE', 'MEDIA', 'BLUETOOTH', 'SYSTEM' }
        Add-Text ($drv | Where-Object { $_.DeviceClass -ne 'SYSTEM' -or $_.DriverProviderName -notmatch 'Microsoft' } | Sort-Object DeviceClass, DeviceName | Select-Object DeviceClass, DeviceName, DriverVersion, DriverDate, DriverProviderName | Format-Table -AutoSize)
        Add-Line "--- Newest 15 drivers ---"
        Add-Text (Get-CimInstance Win32_PnPSignedDriver | Where-Object DeviceName | Sort-Object DriverDate -Descending | Select-Object DeviceName, DriverVersion, DriverDate, DriverProviderName -First 15 | Format-Table -AutoSize)
        Add-Line "--- Non-Microsoft kernel drivers running ---"
        $sd = Get-CimInstance Win32_SystemDriver | Where-Object { $_.State -eq 'Running' } | ForEach-Object {
            $path = $_.PathName -replace '^\\\?\?\\', '' -replace '^\\SystemRoot', $env:windir -replace '^System32', "$env:windir\System32"
            $co = Get-CompanyOf $path
            if ($co -and $co -notmatch 'Microsoft') { [pscustomobject]@{ Name = $_.Name; Company = $co; Path = (Short $path 80) } }
        }
        Add-Text ($sd | Sort-Object Company, Name | Format-Table -AutoSize)
    }

    # ------------------------------------------------------------------
    Invoke-Section 'WINDOWS UPDATE AND SERVICING' '[14/18] Windows Update state...' {
        Add-Text (Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object HotFixID, Description, InstalledOn -First 10 | Format-Table -AutoSize)
        $pending = @()
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pending += 'WindowsUpdate RebootRequired' }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pending += 'CBS RebootPending' }
        if (Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations') { $pending += 'PendingFileRenameOperations' }
        Add-Line "Pending reboot: $(if ($pending) { $pending -join ', ' } else { 'no' })"
        if ($pending) { Add-Flag "A reboot is pending ($($pending -join ', ')). Restart before judging anything else." }
        $serv = $script:ProcRows | Where-Object { $_.Name -in 'TiWorker', 'TrustedInstaller', 'MoUsoCoreWorker', 'CompatTelRunner', 'wuauclt', 'UsoClient', 'MicrosoftEdgeUpdate', 'OneDriveSetup', 'DismHost' }
        if ($serv) { Add-Text ($serv | Select-Object Name, PID, 'CPU%', WS_MB | Format-Table -AutoSize); if ($serv | Where-Object { $_.'CPU%' -ge 5 }) { Add-Flag "Windows servicing/telemetry ($(($serv | ForEach-Object { $_.Name }) -join ', ')) is busy right now. Updates are installing in the background." } }
        else { Add-Line 'No servicing processes running.' }
        Add-Line "--- Update services ---"
        Add-Text (Get-CimInstance Win32_Service | Where-Object { $_.Name -in 'wuauserv', 'UsoSvc', 'DoSvc', 'BITS', 'TrustedInstaller', 'WaaSMedicSvc' } | Select-Object Name, State, StartMode | Format-Table -AutoSize)
        Add-Line "--- Component store (DISM CheckHealth) ---"
        Add-Text (@(& dism.exe /Online /Cleanup-Image /CheckHealth 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match '\S' -and $_ -notmatch '^\[=' }))
    }

    # ------------------------------------------------------------------
    Invoke-Section 'NETWORK (web search, cloud, mapped drives)' '[15/18] Network...' {
        Add-Text (Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up' | Select-Object Name, InterfaceDescription, LinkSpeed, MediaType, DriverVersion | Format-Table -AutoSize)
        Add-Text (@(& netsh.exe wlan show interfaces 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match '^\s*(State|Radio type|Band|Channel|Receive rate|Transmit rate|Signal)\s*:' }))
        Add-Line "--- Proxy ---"
        Show-RegValues 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' @('ProxyEnable', 'ProxyServer', 'AutoConfigURL')
        Add-Text (@(& netsh.exe winhttp show proxy 2>&1 | ForEach-Object { "$_" } | Where-Object { $_ -match '\S' }))
        if ((Get-RegValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'ProxyEnable') -eq 1 -or (Get-RegValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'AutoConfigURL')) { Add-Flag "A proxy / PAC script is configured. If it is unreachable, Start search and Store apps hang waiting for it." }
        Add-Line "--- Latency ---"
        $ms = Measure-Ms { Resolve-DnsName www.bing.com -ErrorAction Stop -DnsOnly }
        Add-Line "DNS lookup www.bing.com: $ms ms"
        if ($ms -ge 1000) { Add-Flag "DNS lookup took $ms ms. Slow DNS makes web-enabled search spin." }
        foreach ($h in 'www.bing.com', 'www.microsoft.com') {
            $sw = [Diagnostics.Stopwatch]::StartNew(); $ok = $false
            try { $c = New-Object Net.Sockets.TcpClient; $ar = $c.BeginConnect($h, 443, $null, $null); $ok = $ar.AsyncWaitHandle.WaitOne(5000) -and $c.Connected; $c.Close() } catch {}
            Add-Line "TCP connect $h`:443: $(if ($ok) { "$([math]::Round($sw.Elapsed.TotalMilliseconds)) ms" } else { 'FAILED / timeout' })"
            if (-not $ok -and $h -eq 'www.bing.com') { Add-Flag "Cannot reach www.bing.com. With web results enabled, Start search waits on it." }
        }
        $gw = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Sort-Object RouteMetric | Select-Object -First 1).NextHop
        if ($gw) { $r = Test-Connection $gw -Count 3 -ErrorAction SilentlyContinue; Add-Line "Ping gateway x3: $(if ($r) { ($r | ForEach-Object { $_.ResponseTime }) -join ' / ' } else { 'no reply' }) ms" }
        Add-Line "--- Mapped network drives (as seen by this elevated session) ---"
        $maps = @(Get-SmbMapping -ErrorAction SilentlyContinue)
        Add-Text ($maps | Select-Object LocalPath, RemotePath, Status | Format-Table -AutoSize)
        Add-Text (Get-CimInstance Win32_NetworkConnection -ErrorAction SilentlyContinue | Select-Object LocalName, RemoteName, Status, ConnectionState | Format-Table -AutoSize)
        foreach ($m in $maps) { if ("$($m.Status)" -ne 'OK') { Add-Flag "Mapped drive $($m.LocalPath) -> $($m.RemotePath) is $($m.Status). Explorer stalls on unreachable network drives." } }
        $od = $script:ProcRows | Where-Object Name -eq 'OneDrive'
        Add-Line "OneDrive running: $(if ($od) { "yes (CPU $($od.'CPU%') %, $($od.WS_MB) MB)" } else { 'no' })"
    }

    # ------------------------------------------------------------------
    Invoke-Section 'EXPLORER SHELL EXTENSIONS' '[16/18] Explorer shell extensions...' {
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
        Add-Line "Total handlers: $(@($rows).Count)   non-Microsoft: $($nonMs.Count)"
        Add-Line "--- Non-Microsoft handlers (each runs inside Explorer when a window, menu or icon is drawn) ---"
        Add-Text ($nonMs | Sort-Object Company, Type | Format-Table -AutoSize)
        if ($nonMs.Count -ge 12) { Add-Flag "$($nonMs.Count) non-Microsoft shell extensions are loaded into Explorer. Known slow-window cause; disable with ShellExView." }
        $overlays = @($rows | Where-Object Type -like '*ShellIconOverlayIdentifiers*')
        if ($overlays.Count -gt 15) { Add-Flag "$($overlays.Count) icon overlay handlers registered (Windows uses only 15; cloud sync apps fight over them)." }
    }

    # ------------------------------------------------------------------
    Invoke-Section 'EVENT LOGS' '[17/18] Event logs (this takes a while)...' {
        Add-Line "--- Errors in System + Application, last 3 days: count by source ---"
        $errs = @(Get-Events @{ LogName = 'System', 'Application'; Level = 1, 2; StartTime = (Get-Date).AddDays(-3) } 600)
        Add-Text ($errs | Group-Object ProviderName, Id | Sort-Object Count -Descending | Select-Object -First 20 Count, Name | Format-Table -AutoSize)
        Add-Line "--- Newest 25 errors ---"
        Add-Text ($errs | Select-Object -First 25 TimeCreated, LogName, ProviderName, Id, @{n = 'Message'; e = { Short $_.Message 260 } } | Format-Table -AutoSize -Wrap)

        Add-Line "--- Application hangs (1002) and crashes (1000), last 14 days ---"
        $hangs = @(Get-Events @{ LogName = 'Application'; ProviderName = 'Application Hang', 'Application Error'; StartTime = (Get-Date).AddDays(-14) } 200)
        Add-Text ($hangs | ForEach-Object { [pscustomobject]@{ Kind = $(if ($_.Id -eq 1002) { 'hang' } else { 'crash' }); App = "$($_.Properties[0].Value)"; Module = $(if ($_.Id -eq 1000) { "$($_.Properties[3].Value)" } else { '' }) } } | Group-Object Kind, App, Module | Sort-Object Count -Descending | Select-Object -First 20 Count, Name | Format-Table -AutoSize)
        Add-Text ($hangs | Select-Object -First 10 TimeCreated, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
        $shellHang = $hangs | Where-Object { "$($_.Properties[0].Value)" -match '^(explorer|SearchHost|SearchApp|StartMenuExperienceHost|ShellExperienceHost|dwm)\.exe' }
        if ($shellHang) { Add-Flag "$(@($shellHang).Count) hang/crash events for the shell ($(($shellHang | ForEach-Object { "$($_.Properties[0].Value)" } | Sort-Object -Unique) -join ', ')) in 14 days." }

        Add-Line "--- Disk / storage events, last 30 days (7, 11, 51, 55, 98, 129, 153, 157 = bad sectors, resets, NTFS damage) ---"
        $disk = @(Get-Events @{ LogName = 'System'; Id = 7, 11, 51, 55, 98, 129, 153, 157, 140; StartTime = (Get-Date).AddDays(-30) } 40)
        Add-Text ($disk | Select-Object TimeCreated, Id, ProviderName, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
        if ($disk | Where-Object { $_.ProviderName -match 'disk|ntfs|stor|volmgr|nvme|iaStor|partmgr' }) { Add-Flag "$(@($disk).Count) storage error events in 30 days (IDs $(($disk | ForEach-Object Id | Sort-Object -Unique) -join ',')). The disk or its cable/driver is faulty." }

        Add-Line "--- Hardware errors (WHEA), last 30 days ---"
        $whea = @(Get-Events @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WHEA-Logger'; StartTime = (Get-Date).AddDays(-30) } 20)
        Add-Text ($whea | Select-Object TimeCreated, Id, LevelDisplayName, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
        if ($whea) { Add-Flag "$($whea.Count) WHEA hardware error events in 30 days (CPU, memory, PCIe or firmware fault)." }

        Add-Line "--- Unexpected shutdowns / crashes (41, 6008, 1001), last 30 days ---"
        $crash = @(Get-Events @{ LogName = 'System'; Id = 41, 6008, 1001; StartTime = (Get-Date).AddDays(-30) } 20)
        Add-Text ($crash | Select-Object TimeCreated, Id, ProviderName, @{n = 'Message'; e = { Short $_.Message 160 } } | Format-Table -AutoSize -Wrap)
        if ($crash | Where-Object { $_.Id -in 41, 6008 }) { Add-Flag "$(@($crash | Where-Object { $_.Id -in 41, 6008 }).Count) unexpected shutdowns / blue screens in 30 days." }

        Add-Line "--- Diagnostics-Performance, last 14 days (1xx boot, 2xx shutdown, 3xx standby, 4xx desktop responsiveness, 5xx Explorer) ---"
        $perf = @(Get-Events @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; StartTime = (Get-Date).AddDays(-14) } 300)
        Add-Text ($perf | Group-Object Id | Sort-Object Count -Descending | Select-Object Count, Name | Format-Table -AutoSize)
        Add-Text ($perf | Where-Object { $_.Id -ge 100 -and $_.Id -lt 200 } | Select-Object -First 5 TimeCreated, Id, @{n = 'Boot'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
        Add-Text ($perf | Where-Object { $_.Id -ge 400 } | Select-Object -First 20 TimeCreated, Id, LevelDisplayName, @{n = 'Message'; e = { Short $_.Message 220 } } | Format-Table -AutoSize -Wrap)
        $boot = $perf | Where-Object Id -eq 100 | Select-Object -First 1
        if ($boot -and $boot.Message -match 'Boot Duration\s*:\s*(\d+)ms') { $bs = [int]$matches[1] / 1000; Add-Line "Last measured boot: $bs s"; if ($bs -ge 120) { Add-Flag "Last boot took $bs s (events 101-110 below name the slow drivers/services/apps)." } }
        $slow = $perf | Where-Object { $_.Id -in 101..110 } | Select-Object -First 15
        Add-Text ($slow | Select-Object TimeCreated, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
        if (@($perf | Where-Object { $_.Id -ge 500 -and $_.Id -lt 600 }).Count -ge 3) { Add-Flag "Windows logged $(@($perf | Where-Object { $_.Id -ge 500 -and $_.Id -lt 600 }).Count) Explorer responsiveness-degradation events in 14 days." }

        Add-Line "--- Store app deployment errors, last 7 days ---"
        Add-Text (Get-Events @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 1, 2; StartTime = (Get-Date).AddDays(-7) } 10 | Select-Object TimeCreated, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
        Add-Line "--- Shell-Core / TWinUI errors, last 7 days ---"
        Add-Text (Get-Events @{ LogName = 'Microsoft-Windows-Shell-Core/Operational', 'Microsoft-Windows-TWinUI/Operational'; Level = 1, 2; StartTime = (Get-Date).AddDays(-7) } 15 | Select-Object TimeCreated, LogName, Id, @{n = 'Message'; e = { Short $_.Message 200 } } | Format-Table -AutoSize -Wrap)
    }

    # ------------------------------------------------------------------
    Invoke-Section 'RELIABILITY, CRASH DUMPS, SESSIONS' '[18/18] Reliability history...' {
        $stab = Get-CimInstance Win32_ReliabilityStabilityMetrics -ErrorAction SilentlyContinue | Sort-Object TimeGenerated -Descending | Select-Object -First 1
        if ($stab) { Add-Line "Reliability index (1-10, Reliability Monitor): $($stab.SystemStabilityIndex) at $($stab.TimeGenerated)" }
        Add-Line "--- Reliability records, last 14 days ---"
        $rr = Get-CimInstance Win32_ReliabilityRecords -ErrorAction SilentlyContinue | Where-Object { $_.TimeGenerated -gt (Get-Date).AddDays(-14) }
        Add-Text ($rr | Group-Object SourceName, ProductName | Sort-Object Count -Descending | Select-Object -First 15 Count, Name | Format-Table -AutoSize)
        Add-Text ($rr | Sort-Object TimeGenerated -Descending | Select-Object -First 25 TimeGenerated, SourceName, EventIdentifier, ProductName, @{n = 'Message'; e = { Short $_.Message 150 } } | Format-Table -AutoSize -Wrap)
        Add-Line "--- Crash dumps ---"
        $md = @(Get-ChildItem -LiteralPath "$env:windir\Minidump" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        Add-Line "Minidumps (blue screens): $($md.Count)$(if ($md) { ", newest $($md[0].LastWriteTime)" })"
        $lk = @(Get-ChildItem -LiteralPath "$env:windir\LiveKernelReports" -Recurse -File -ErrorAction SilentlyContinue | Where-Object LastWriteTime -gt (Get-Date).AddDays(-30))
        Add-Line "Live kernel reports (driver hangs, GPU timeouts), last 30 days: $($lk.Count)"
        if ($lk) { Add-Text ($lk | Select-Object -First 10 Name, LastWriteTime, @{n = 'MB'; e = { [math]::Round($_.Length / 1MB) }} | Format-Table -AutoSize); Add-Flag "$($lk.Count) live kernel reports in 30 days (a driver stopped responding; folder names say which: WATCHDOG, USB, PCI, NDIS...)." }
        $wer = @(Get-ChildItem -LiteralPath "$env:ProgramData\Microsoft\Windows\WER\ReportQueue", "$env:ProgramData\Microsoft\Windows\WER\ReportArchive", "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue", "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive" -Directory -ErrorAction SilentlyContinue | Where-Object LastWriteTime -gt (Get-Date).AddDays(-14))
        Add-Line "Error reports (WER), last 14 days: $($wer.Count)"
        Add-Text ($wer | Group-Object { ($_.Name -split '_')[0, 1] -join '_' } | Sort-Object Count -Descending | Select-Object -First 15 Count, Name | Format-Table -AutoSize)
        Add-Line "--- Logged-on sessions ---"
        Add-Text (@(& quser.exe 2>&1 | ForEach-Object { "$_" }))
    }

    # ------------------------------------------------------------------
    Add-Line ''
    Add-Line 'END OF REPORT'
    $idx = $script:Report.IndexOf('[[FLAGS]]')
    if ($idx -ge 0) {
        $script:Report.RemoveAt($idx)
        if ($script:Flags.Count -eq 0) { $script:Report.Insert($idx, 'No obvious problem was auto-detected. The details below still show where the time goes.') }
        else { $i = 0; foreach ($f in $script:Flags) { $i++; $script:Report.Insert($idx + $i - 1, "$i. $f") } }
    }
    Save-Report

    Say ""
    Say "  DONE - Lag check" 'Green'
    Say "  Saved to: $script:ReportPath" 'Green'
    Say ""
    if ($script:Flags.Count -eq 0) { Say "  Nothing obvious was flagged. Send the file for a closer look." 'Yellow' }
    else {
        Say "  Flagged ($($script:Flags.Count)):" 'Yellow'
        $i = 0; foreach ($f in $script:Flags) { $i++; Say "   $i. $f" 'Yellow' }
    }
    Say ""
    Say "  Send the file back for analysis. It names the PC and user but no serials or addresses." 'DarkYellow'
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
Set-WindowRight
Clear-Host
Say ""
Say "  =============================================" 'Cyan'
Say "        LAG CHECK" 'Cyan'
Say "  =============================================" 'Cyan'
Say ""
Say "    1  -  Run the check (3-5 min)" 'White'
Say "    0  -  Quit" 'White'
Say ""

$choice = Read-MenuKey 'Choose 0-1' '01' '1'
if ($choice -eq '0') { [Environment]::Exit(0) }   # Quit closes this window

$OutDir = Select-OutputFolder
Start-LagCheck -OutDir $OutDir

Say ""
# This window was opened just for the report, so pause before it closes
Say "  Press any key to close this window" 'DarkGray'
try { [void][Console]::ReadKey($true) } catch { [void][Console]::ReadLine() }
