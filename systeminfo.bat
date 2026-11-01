@echo off
title System Info Collector
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Requesting administrator rights...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$f=[IO.File]::ReadAllText('%~f0'); $i=$f.LastIndexOf('#PSSTART'); Invoke-Expression $f.Substring($i)"

echo.
pause
exit /b

#PSSTART
# =====================================================================
#  System Info Collector  -  embedded PowerShell section
#  1 = Windows laptop   2 = Android phone   3 = Both
#  Raw output, no redaction. Reports go to Desktop\SystemReports.
# =====================================================================

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

$Desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($Desktop) -or -not (Test-Path $Desktop)) { $Desktop = $env:USERPROFILE }
$OutDir = Join-Path $Desktop 'SystemReports'
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$Stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }

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
    $out = Join-Path $OutDir "laptop_$Stamp.txt"
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
    powercfg /batteryreport /output (Join-Path $OutDir "battery_$Stamp.html") | Out-Null
    "Detailed battery health report saved as battery_$Stamp.html"

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
    Say "  A battery health file (battery_$Stamp.html) is in the same folder." 'DarkGray'
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

    $out = Join-Path $OutDir "phone_$Stamp.txt"
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
while ($true) {
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        SYSTEM INFO COLLECTOR" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  This Windows laptop" 'White'
    Say "    2  -  Android phone (USB cable needed)" 'White'
    Say "    3  -  Both" 'White'
    Say "    O  -  Open the reports folder" 'White'
    Say "    Q  -  Quit" 'White'
    Say ""
    $choice = Read-Host "  Type a number and press Enter"

    switch ($choice.Trim().ToUpper()) {
        '1' { Get-LaptopReport }
        '2' { Get-PhoneReport }
        '3' { Get-LaptopReport; Get-PhoneReport }
        'O' { Start-Process explorer.exe $OutDir }
        'Q' { Say ""; Say "  Reports are in: $OutDir" 'Green'; Say ""; return }
        default { Say "  Please type 1, 2, 3, O or Q." 'Red' }
    }
}
