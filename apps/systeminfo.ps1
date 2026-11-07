# irm "https://github.com/therepos/windows/raw/main/apps/systeminfo.ps1?$(Get-Random)" | iex
# Purpose: Collect raw hardware/software info from a Windows laptop and/or Android phone
# =============================================================================
# Usage:
#   1 = Windows laptop
#   2 = Android phone
#   3 = Both
#   0 = Quit
#
# Note:
#   Output: folder asked, default Downloads
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/systeminfo.ps1'

# --- Self-elevate: re-run the same one-liner in an elevated window ------------
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()

           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "  Requesting administrator rights..." -ForegroundColor Yellow
    $cmd = "[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; try { irm '${ScriptUrl}?$(Get-Random)' | iex } catch { Write-Host `$_ -ForegroundColor Red; Read-Host 'Press Enter to exit' }"
    Start-Process -FilePath 'powershell.exe' -Verb RunAs `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $cmd
    Write-Host "  Continue in the new (elevated) window. You can close this one." -ForegroundColor Gray
    return
}

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

$OutDir = $null
$Stamp = Get-Date -Format 'yyyyMMdd-HHmm'

# Model as one word for file names: spaces -> dashes, characters Windows forbids removed
function ConvertTo-OneWord([string]$s) {
    $s = ($s -replace '[\\/:*?"<>|]', '').Trim() -replace '\s+', '-'
    if ($s) { $s } else { 'Unknown' }
}

function Get-LaptopModel {
    $cs = Get-CimInstance Win32_ComputerSystem
    $model = $cs.Model
    # Lenovo reports a part code (e.g. 21CB0020US) as Model; the readable name is in Version
    if ($cs.Manufacturer -match 'LENOVO') {
        $friendly = (Get-CimInstance Win32_ComputerSystemProduct).Version
        if ($friendly) { $model = $friendly }
    }
    ConvertTo-OneWord $model
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

function Finish-Report {
    param([string]$Path, [string]$Label)
    Say ""
    Say "  DONE - $Label" 'Green'
    Say "  Saved to: $Path" 'Green'
    Say "  NOTE: raw output. Contains serials, MACs and IPs - check before sharing." 'DarkYellow'
    Say ""
}

# =====================================================================
#  OPTION 1 - WINDOWS LAPTOP
# =====================================================================
function Get-LaptopReport {
    $out = Join-Path $OutDir "Data System $(Get-LaptopModel) $Stamp.txt"
    Say ""
    Say "  Collecting laptop info. This takes about 1-2 minutes..." 'Cyan'
    Say ""

    Start-Transcript -Path $out -Force | Out-Null

    Say "  [1/12] Operating system..." 'DarkGray'
    "=== OS ==="
    Get-ComputerInfo | Select-Object OsName,OsVersion,OsBuildNumber,WindowsVersion,OsArchitecture,CsName,CsManufacturer,CsModel,CsSystemType,BiosSMBIOSBIOSVersion,BiosReleaseDate,BiosSeralNumber,OsUptime,OsInstallDate,TimeZone | Format-List

    Say "  [2/12] Processor..." 'DarkGray'
    "=== CPU ==="
    Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed,L3CacheSize,VirtualizationFirmwareEnabled | Format-List

    Say "  [3/12] Memory..." 'DarkGray'
    "=== RAM ==="
    Get-CimInstance Win32_PhysicalMemory | Select-Object BankLabel,@{n='CapacityGB';e={[math]::Round($_.Capacity/1GB,1)}},Speed,ConfiguredClockSpeed,Manufacturer,PartNumber,SerialNumber | Format-Table -AutoSize
    Get-CimInstance Win32_OperatingSystem | Select-Object @{n='TotalRAM_GB';e={[math]::Round($_.TotalVisibleMemorySize/1MB,1)}},@{n='FreeRAM_GB';e={[math]::Round($_.FreePhysicalMemory/1MB,1)}} | Format-List

    Say "  [4/12] Graphics..." 'DarkGray'
    "=== GPU / DISPLAY ==="
    Get-CimInstance Win32_VideoController | Select-Object Name,DriverVersion,DriverDate,VideoModeDescription,CurrentRefreshRate,VideoProcessor | Format-List

    Say "  [5/12] Storage..." 'DarkGray'
    "=== PHYSICAL DISKS ==="
    Get-PhysicalDisk | Select-Object FriendlyName,SerialNumber,MediaType,BusType,@{n='SizeGB';e={[math]::Round($_.Size/1GB,0)}},HealthStatus,FirmwareVersion | Format-Table -AutoSize
    "=== VOLUMES ==="
    Get-Volume | Where-Object DriveLetter | Select-Object DriveLetter,FileSystemLabel,FileSystem,@{n='SizeGB';e={[math]::Round($_.Size/1GB,1)}},@{n='FreeGB';e={[math]::Round($_.SizeRemaining/1GB,1)}},HealthStatus | Format-Table -AutoSize

    Say "  [6/12] Security (TPM / BitLocker / Secure Boot)..." 'DarkGray'
    "=== SECURITY ==="
    Get-BitLockerVolume | Select-Object MountPoint,VolumeStatus,ProtectionStatus,EncryptionMethod | Format-Table -AutoSize
    Get-Tpm | Select-Object TpmPresent,TpmReady,TpmEnabled,ManufacturerVersion | Format-List
    "SecureBoot enabled: " + (Confirm-SecureBootUEFI)
    Get-MpComputerStatus | Select-Object AMServiceEnabled,RealTimeProtectionEnabled,AntivirusSignatureLastUpdated | Format-List

    Say "  [7/12] Battery and power..." 'DarkGray'
    "=== BATTERY / POWER ==="
    Get-CimInstance Win32_Battery | Select-Object Name,EstimatedChargeRemaining,BatteryStatus,DesignVoltage | Format-List
    powercfg /getactivescheme
    "=== BATTERY HEALTH ==="
    $xmlPath = Join-Path $env:TEMP "battery_$Stamp.xml"
    powercfg /batteryreport /xml /output $xmlPath | Out-Null
    if (Test-Path $xmlPath) {
        [xml]$br = Get-Content $xmlPath -Raw
        Remove-Item $xmlPath -Force
        foreach ($b in @($br.BatteryReport.Batteries.Battery)) {
            $design = [double]$b.DesignCapacity
            $full   = [double]$b.FullChargeCapacity
            [pscustomobject]@{
                Battery            = $b.Id
                Manufacturer       = $b.Manufacturer
                Chemistry          = $b.Chemistry
                DesignCapacity_mWh = $design
                FullCharge_mWh     = $full
                HealthPercent      = if ($design -gt 0) { [math]::Round($full / $design * 100, 1) } else { 'n/a' }
                CycleCount         = $b.CycleCount
            } | Format-List
        }
        $est = $br.BatteryReport.RuntimeEstimates
        if ($est) {
            "Estimated runtime at full charge:   " + $est.FullChargeCapacity.ActiveRuntime
            "Estimated runtime at design charge: " + $est.DesignCapacity.ActiveRuntime
        }
        "--- Capacity history (last 10 periods) ---"
        @($br.BatteryReport.History.HistoryEntry) | Select-Object -Last 10 |
            Select-Object StartDate,EndDate,@{n='FullCharge_mWh';e={$_.FullChargeCapacity}},@{n='Design_mWh';e={$_.DesignCapacity}},CycleCount |
            Format-Table -AutoSize
    } else {
        "No battery detected."
    }

    Say "  [8/12] Network..." 'DarkGray'
    "=== NETWORK ADAPTERS ==="
    Get-NetAdapter | Select-Object Name,InterfaceDescription,Status,LinkSpeed,MacAddress,DriverVersion | Format-Table -AutoSize
    "=== IP CONFIGURATION ==="
    Get-NetIPConfiguration | Format-List

    Say "  [9/12] Devices with problems..." 'DarkGray'
    "=== PROBLEM DEVICES ==="
    $bad = Get-PnpDevice | Where-Object { $_.Status -ne 'OK' -and $_.Present }
    if ($bad) { $bad | Select-Object FriendlyName,Status,Class,InstanceId | Format-Table -AutoSize }
    else { "None - all devices reporting OK." }

    Say "  [10/12] Drivers..." 'DarkGray'
    "=== NEWEST 40 DRIVERS ==="
    Get-CimInstance Win32_PnPSignedDriver | Where-Object DeviceName | Sort-Object DriverDate -Descending | Select-Object DeviceName,DriverVersion,DriverDate,DriverProviderName -First 40 | Format-Table -AutoSize

    Say "  [11/12] Installed software and startup..." 'DarkGray'
    "=== INSTALLED APPS ==="
    Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object DisplayName | Select-Object DisplayName,DisplayVersion,Publisher | Sort-Object DisplayName -Unique | Format-Table -AutoSize
    "=== STARTUP ITEMS ==="
    Get-CimInstance Win32_StartupCommand | Select-Object Name,Command,Location | Format-Table -AutoSize
    "=== WINDOWS FEATURES ENABLED ==="
    Get-WindowsOptionalFeature -Online | Where-Object State -eq 'Enabled' | Select-Object FeatureName | Format-Table -AutoSize

    Say "  [12/12] Updates and recent errors..." 'DarkGray'
    "=== RECENT UPDATES ==="
    Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object HotFixID,Description,InstalledOn -First 20 | Format-Table -AutoSize
    "=== SYSTEM ERRORS (LAST 7 DAYS) ==="
    Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 25 | Select-Object TimeCreated,Id,ProviderName,Message | Format-List
    "=== UNEXPECTED SHUTDOWNS ==="
    Get-WinEvent -FilterHashtable @{LogName='System';Id=41,6008} -MaxEvents 10 | Select-Object TimeCreated,Id,Message | Format-List

    Stop-Transcript | Out-Null
    Finish-Report -Path $out -Label 'Laptop report'
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
        $list = (adb devices) -split "`n" | Where-Object { $_ -match "`t" }
        if ($list -match "`tdevice$") { Say "  Phone connected and authorised." 'Green'; return $true }
        if ($list -match "`tunauthorized" -and -not $warned) {
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
    return $false
}

function Get-PhoneReport {
    if (-not (Ensure-Adb))    { return }
    if (-not (Wait-ForPhone)) { return }

    $out = Join-Path $OutDir "Data Phone $(ConvertTo-OneWord "$(adb shell getprop ro.product.model)") $Stamp.txt"
    Say ""
    Say "  Collecting phone info. Keep the phone plugged in and unlocked..." 'Cyan'
    Say ""

    Start-Transcript -Path $out -Force | Out-Null

    Say "  [1/9] Device identity..." 'DarkGray'
    "=== DEVICE ==="
    adb shell getprop ro.product.manufacturer
    adb shell getprop ro.product.model
    adb shell getprop ro.product.name
    adb shell getprop ro.build.version.release
    adb shell getprop ro.build.version.sdk
    adb shell getprop ro.build.version.security_patch
    adb shell getprop ro.build.display.id
    adb shell getprop ro.board.platform

    Say "  [2/9] All system properties..." 'DarkGray'
    "=== FULL PROPERTIES ==="
    adb shell getprop

    Say "  [3/9] Hardware..." 'DarkGray'
    "=== CPU ==="
    adb shell cat /proc/cpuinfo
    "=== MEMORY ==="
    adb shell cat /proc/meminfo
    "=== MEMORY SUMMARY ==="
    adb shell dumpsys meminfo | Select-Object -First 30

    Say "  [4/9] Storage..." 'DarkGray'
    "=== STORAGE ==="
    adb shell df -h

    Say "  [5/9] Battery..." 'DarkGray'
    "=== BATTERY ==="
    adb shell dumpsys battery
    "=== BATTERY USAGE (TOP) ==="
    adb shell dumpsys batterystats --charged | Select-Object -First 150

    Say "  [6/9] Display..." 'DarkGray'
    "=== DISPLAY ==="
    adb shell wm size
    adb shell wm density
    adb shell settings get system screen_off_timeout

    Say "  [7/9] Network..." 'DarkGray'
    "=== NETWORK ==="
    adb shell ip -br a
    adb shell dumpsys wifi | Select-Object -First 25

    Say "  [8/9] Apps..." 'DarkGray'
    "=== USER-INSTALLED APPS ==="
    adb shell pm list packages -3
    "=== DISABLED APPS ==="
    adb shell pm list packages -d
    "=== TOTAL PACKAGE COUNT ==="
    (adb shell pm list packages).Count

    Say "  [9/9] Power management..." 'DarkGray'
    "=== DOZE / STANDBY ==="
    adb shell dumpsys deviceidle | Select-Object -First 50
    "=== RECENT CRASHES ==="
    adb shell logcat -d -b crash -t 100

    Stop-Transcript | Out-Null
    Finish-Report -Path $out -Label 'Phone report'
}

# =====================================================================
#  MENU
# =====================================================================
Clear-Host
Say ""
Say "  =============================================" 'Cyan'
Say "        SYSTEM INFO COLLECTOR" 'Cyan'
Say "  =============================================" 'Cyan'
Say ""
Say "    1  -  This Windows laptop" 'White'
Say "    2  -  Android phone (USB cable needed)" 'White'
Say "    3  -  Both" 'White'
Say "    0  -  Quit" 'White'
Say ""

$choice = Read-MenuKey 'Choose 0-3' '0123'
if ($choice -eq '0') { [Environment]::Exit(0) }   # Quit closes this window

$OutDir = Select-OutputFolder
switch ($choice) {
    '1' { Get-LaptopReport }
    '2' { Get-PhoneReport }
    '3' { Get-LaptopReport; Get-PhoneReport }
}

Say "  Reports are in: $OutDir" 'Green'
Say ""
# This window was opened just for the report, so pause before it closes
Say "  Press any key to close this window" 'DarkGray'
try { [void][Console]::ReadKey($true) } catch { [void][Console]::ReadLine() }
