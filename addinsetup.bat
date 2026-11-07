@echo off
:: Purpose: Install or remove Office add-ins (.xlam, .ppam, .dotm) placed next to this file
:: ============================================================================
:: Note:
::   Double-click to run. No Python needed
:: ============================================================================
setlocal EnableDelayedExpansion

set "PS_TEMP=%TEMP%\addin-mgr-%RANDOM%.ps1"

:: Use PowerShell to extract everything after ::__PS_BEGIN__ from this file
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$lines = [IO.File]::ReadAllLines('%~f0');" ^
  "$start = -1;" ^
  "for ($i=0; $i -lt $lines.Count; $i++) { if ($lines[$i] -eq '::__PS_BEGIN__') { $start = $i + 1; break } };" ^
  "if ($start -ge 0) { [IO.File]::WriteAllLines('%PS_TEMP%', $lines[$start..($lines.Count-1)]) }"

:: Run it, passing the folder where this .bat lives
:: Note: trailing \ in %~dp0 can escape the closing quote, so we append a dot
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_TEMP%" "%~dp0."
set "RC=%ERRORLEVEL%"

:: Clean up
del /f /q "%PS_TEMP%" >nul 2>&1
exit /b %RC%

::__PS_BEGIN__
# =============================================================
# Office Add-in Manager - PowerShell
# =============================================================

param([string]$ScriptDir)
$ScriptDir = $ScriptDir.TrimEnd('.').TrimEnd('\').TrimEnd('"').TrimEnd("'")

$ErrorActionPreference = 'SilentlyContinue'
$Host.UI.RawUI.WindowTitle = 'Office Add-in Manager'

# =============================================================
# CONFIG
# =============================================================
$appdata = $env:APPDATA

# Where each add-in type gets installed to
$installTargets = @{
    '.xlam' = "$appdata\Microsoft\AddIns"
    '.xla'  = "$appdata\Microsoft\AddIns"
    '.ppam' = "$appdata\Microsoft\PowerPoint\AddIns"
    '.ppa'  = "$appdata\Microsoft\PowerPoint\AddIns"
    '.dotm' = "$appdata\Microsoft\Word\STARTUP"
    '.dot'  = "$appdata\Microsoft\Word\STARTUP"
}

# Folders to scan for installed add-in files
$scanPaths = @(
    @{ Folder = "$appdata\Microsoft\AddIns";              Exts = @('.xlam','.xla','.ppam','.ppa') }
    @{ Folder = "$appdata\Microsoft\Word\STARTUP";        Exts = @('.dotm','.dot') }
    @{ Folder = "$appdata\Microsoft\PowerPoint\AddIns";   Exts = @('.ppam','.ppa') }
    @{ Folder = "$appdata\Microsoft\Excel\XLSTART";       Exts = @('.xlam','.xla','.xls','.xlsx') }
)

$appLabels = @{
    '.xlam'='Excel'; '.xla'='Excel'
    '.dotm'='Word';  '.dot'='Word'
    '.ppam'='PowerPoint'; '.ppa'='PowerPoint'
    '.xls'='Excel'; '.xlsx'='Excel'
}

# Extensions we can install
$installExts = @('.xlam','.xla','.ppam','.ppa','.dotm','.dot')

# Registry keys for auto-load (OPEN/OPEN1) and COM subkeys
$regKeys = @(
    @{ Path='HKCU:\Software\Microsoft\Office\Excel\Addins';            App='Excel';       Type='subkey' }
    @{ Path='HKCU:\Software\Microsoft\Office\Word\Addins';             App='Word';        Type='subkey' }
    @{ Path='HKCU:\Software\Microsoft\Office\PowerPoint\Addins';       App='PowerPoint';  Type='subkey' }
    @{ Path='HKCU:\Software\Microsoft\Office\16.0\Excel\Options';      App='Excel';       Type='open' }
    @{ Path='HKCU:\Software\Microsoft\Office\15.0\Excel\Options';      App='Excel';       Type='open' }
    @{ Path='HKCU:\Software\Microsoft\Office\14.0\Excel\Options';      App='Excel';       Type='open' }
    @{ Path='HKCU:\Software\Microsoft\Office\16.0\Word\Options';       App='Word';        Type='open' }
    @{ Path='HKCU:\Software\Microsoft\Office\15.0\Word\Options';       App='Word';        Type='open' }
    @{ Path='HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options'; App='PowerPoint';  Type='open' }
    @{ Path='HKCU:\Software\Microsoft\Office\15.0\PowerPoint\Options'; App='PowerPoint';  Type='open' }
)

# Which Office version registry path to use for OPEN values when installing
# We detect the installed version, falling back to 16.0
$regOptionsKeys = @{
    'Excel'      = @('HKCU:\Software\Microsoft\Office\16.0\Excel\Options',
                     'HKCU:\Software\Microsoft\Office\15.0\Excel\Options',
                     'HKCU:\Software\Microsoft\Office\14.0\Excel\Options')
    'PowerPoint' = @('HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options',
                     'HKCU:\Software\Microsoft\Office\15.0\PowerPoint\Options',
                     'HKCU:\Software\Microsoft\Office\14.0\PowerPoint\Options')
}

# =============================================================
# SHARED UTILITIES
# =============================================================

function Parse-Selection($text, $max) {
    $text = $text.Trim().ToLower()
    if ($text -in 'all','a','*') { return @(0..($max - 1)) }
    $indices = [System.Collections.Generic.SortedSet[int]]::new()
    foreach ($part in ($text -split ',')) {
        $p = $part.Trim()
        if ($p -match '^\d+\s*-\s*\d+$') {
            $bounds = $p -split '-'
            $lo = [int]$bounds[0]; $hi = [int]$bounds[1]
            if ($lo -gt $hi) { $lo, $hi = $hi, $lo }
            for ($v = $lo; $v -le $hi; $v++) { [void]$indices.Add($v) }
        } elseif ($p -match '^\d+$') {
            [void]$indices.Add([int]$p)
        } else { return $null }
    }
    $result = @()
    foreach ($v in $indices) {
        $idx = $v - 1
        if ($idx -lt 0 -or $idx -ge $max) { return $null }
        $result += $idx
    }
    return $result
}

function Detect-OfficeRegPath($app) {
    # Find the first existing Options key for this app
    $candidates = $regOptionsKeys[$app]
    if (-not $candidates) { return $null }
    foreach ($c in $candidates) {
        # Check if parent exists (e.g. Office\16.0\Excel)
        $parent = Split-Path $c -Parent
        if (Test-Path $parent) { return $c }
    }
    # Default to 16.0
    return $candidates[0]
}

function Get-NextOpenName($regPath) {
    # Find the next available OPEN, OPEN1, OPEN2, ... value name
    $props = Get-ItemProperty -Path $regPath 2>$null
    if (-not $props) { return 'OPEN' }

    $existing = @()
    foreach ($pn in $props.PSObject.Properties.Name) {
        if ($pn -match '^OPEN(\d*)$') {
            if ($Matches[1] -eq '') { $existing += 0 }
            else { $existing += [int]$Matches[1] }
        }
    }

    if ($existing.Count -eq 0) { return 'OPEN' }

    # OPEN=0, OPEN1=1, OPEN2=2 ...
    $existing = $existing | Sort-Object
    $next = 0
    foreach ($n in $existing) {
        if ($n -eq $next) { $next++ }
    }
    if ($next -eq 0) { return 'OPEN' }
    return "OPEN$next"
}

function Is-AlreadyRegistered($fileName, $app) {
    # Check if this file is already referenced in an OPEN value
    foreach ($rk in $regKeys) {
        if ($rk.App -ne $app) { continue }
        if ($rk.Type -ne 'open') { continue }
        $props = Get-ItemProperty -Path $rk.Path 2>$null
        if (-not $props) { continue }
        foreach ($pn in $props.PSObject.Properties.Name) {
            if ($pn -match '^OPEN\d*$') {
                $val = $props.$pn
                if ($val -and $val.ToLower().Contains($fileName.ToLower())) {
                    return $true
                }
            }
        }
    }
    return $false
}


# =============================================================
# INSTALL LOGIC
# =============================================================

function Invoke-Install {
    Write-Host ''
    Write-Host '----------------------------------------------------' -ForegroundColor Green
    Write-Host '  INSTALL Add-ins' -ForegroundColor Green
    Write-Host '----------------------------------------------------' -ForegroundColor Green

    # Scan the script's own folder for add-in files
    $localAddins = @()
    foreach ($f in Get-ChildItem -Path $ScriptDir -File) {
        $ext = $f.Extension.ToLower()
        if ($installExts -contains $ext) {
            $app = if ($appLabels[$ext]) { $appLabels[$ext] } else { 'Office' }
            $localAddins += @{
                Name   = $f.Name
                Path   = $f.FullName
                Ext    = $ext
                App    = $app
                Target = $installTargets[$ext]
            }
        }
    }

    $localAddins = @($localAddins | Sort-Object { $_.App }, { $_.Name.ToLower() })

    if ($localAddins.Count -eq 0) {
        Write-Host ''
        Write-Host '  No add-in files found next to this script.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host "  Place .xlam, .ppam, or .dotm files in:" -ForegroundColor Yellow
        Write-Host "    $ScriptDir" -ForegroundColor Yellow
        Write-Host '  Then run this again.' -ForegroundColor Yellow
        return
    }

    Write-Host "`n  Add-in files found in script folder:`n"
    for ($i = 0; $i -lt $localAddins.Count; $i++) {
        $a = $localAddins[$i]
        $num = ($i + 1).ToString().PadLeft(2)

        # Check if already installed
        $targetFile = Join-Path $a.Target $a.Name
        $installed = (Test-Path $targetFile)
        $registered = Is-AlreadyRegistered $a.Name $a.App

        $status = ''
        if ($installed -and $registered) {
            $status = ' [already installed]'
        } elseif ($installed) {
            $status = ' [file exists, not registered]'
        }

        $color = if ($status) { 'DarkGray' } else { 'White' }
        Write-Host "  [$num]  $($a.Name)  ($($a.App))$status" -ForegroundColor $color
    }

    Write-Host ''
    Write-Host '  Select items to install:'
    Write-Host '    Single:   3'
    Write-Host '    Multiple: 1,3,5'
    Write-Host '    Range:    1-4'
    Write-Host '    All:      all'
    Write-Host ''

    $choice = Read-Host '  Your selection (or q to cancel)'
    if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq 'q') {
        Write-Host '  Cancelled.'
        return
    }

    $selIndices = Parse-Selection $choice $localAddins.Count
    if ($null -eq $selIndices) {
        Write-Host '  Invalid selection.' -ForegroundColor Red
        return
    }

    $selected = @(); foreach ($i in $selIndices) { $selected += $localAddins[$i] }

    Write-Host "`n  Installing $($selected.Count) add-in(s)...`n"

    $okCount = 0; $errCount = 0

    foreach ($s in $selected) {
        $targetDir  = $s.Target
        $targetFile = Join-Path $targetDir $s.Name

        # Step 1: Create target folder if needed
        if (-not (Test-Path $targetDir)) {
            try {
                New-Item -Path $targetDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Host "  [OK]  Created folder: $targetDir" -ForegroundColor Green
            } catch {
                Write-Host "  [ERR] Cannot create folder: $targetDir - $($_.Exception.Message)" -ForegroundColor Red
                $errCount++
                continue
            }
        }

        # Step 2: Copy file
        try {
            Copy-Item -Path $s.Path -Destination $targetFile -Force -ErrorAction Stop
            Write-Host "  [OK]  Copied: $($s.Name) -> $targetDir" -ForegroundColor Green
            $okCount++
        } catch {
            Write-Host "  [ERR] Copy failed: $($_.Exception.Message)" -ForegroundColor Red
            $errCount++
            continue
        }

        # Step 3: Register in registry (Excel and PowerPoint need OPEN values)
        #          Word auto-loads from STARTUP, no registry needed
        $ext = $s.Ext
        $app = $s.App

        if ($app -eq 'Word') {
            Write-Host "  [OK]  $($s.Name) will auto-load from Word STARTUP folder" -ForegroundColor Green
            $okCount++
            continue
        }

        # Skip if already registered
        if (Is-AlreadyRegistered $s.Name $app) {
            Write-Host "  [OK]  $($s.Name) already registered in $app" -ForegroundColor Green
            $okCount++
            continue
        }

        # Find the right registry key
        $optionsKey = Detect-OfficeRegPath $app
        if (-not $optionsKey) {
            Write-Host "  [WARN] Cannot detect $app registry path, skipping registration" -ForegroundColor Yellow
            Write-Host "         Add-in file is in place and may need manual activation in $app" -ForegroundColor Yellow
            continue
        }

        # Ensure registry key exists
        if (-not (Test-Path $optionsKey)) {
            try {
                New-Item -Path $optionsKey -Force -ErrorAction Stop | Out-Null
            } catch {
                Write-Host "  [ERR] Cannot create registry key: $optionsKey" -ForegroundColor Red
                $errCount++
                continue
            }
        }

        # Get next available OPEN name
        $openName = Get-NextOpenName $optionsKey

        # Value is the filename (for AddIns folder) — Excel resolves relative to the AddIns path
        try {
            Set-ItemProperty -Path $optionsKey -Name $openName -Value $s.Name -ErrorAction Stop
            Write-Host "  [OK]  Registered: $optionsKey -> $openName = $($s.Name)" -ForegroundColor Green
            $okCount++
        } catch {
            Write-Host "  [ERR] Registry write failed: $($_.Exception.Message)" -ForegroundColor Red
            $errCount++
        }
    }

    Write-Host ''
    Write-Host "  Install done: $okCount succeeded, $errCount failed." -ForegroundColor Cyan
    if ($okCount -gt 0) {
        Write-Host ''
        Write-Host '  Tip: Restart the Office app(s) for add-ins to load.' -ForegroundColor Yellow
    }
}


# =============================================================
# REMOVE LOGIC
# =============================================================

function Invoke-Remove {
    Write-Host ''
    Write-Host '----------------------------------------------------' -ForegroundColor Red
    Write-Host '  REMOVE Add-ins  (Clean Uninstall)' -ForegroundColor Red
    Write-Host '----------------------------------------------------' -ForegroundColor Red

    # Scan installed files
    $allItems  = [System.Collections.ArrayList]::new()
    $fileNames = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($sp in $scanPaths) {
        if (-not (Test-Path $sp.Folder)) { continue }
        foreach ($f in Get-ChildItem -Path $sp.Folder -File) {
            $ext = $f.Extension.ToLower()
            if ($sp.Exts -contains $ext) {
                $app = if ($appLabels[$ext]) { $appLabels[$ext] } else { 'Office' }
                [void]$allItems.Add(@{
                    Name   = $f.Name
                    Path   = $f.FullName
                    App    = $app
                    Source = 'file'
                    Tag    = $app
                })
                [void]$fileNames.Add($f.Name.ToLower())
            }
        }
    }

    # Scan registry
    $regItems = [System.Collections.ArrayList]::new()

    foreach ($rk in $regKeys) {
        if (-not (Test-Path $rk.Path)) { continue }

        if ($rk.Type -eq 'subkey') {
            $subs = Get-ChildItem -Path $rk.Path 2>$null
            foreach ($sub in $subs) {
                $manifest = ''
                try { $manifest = (Get-ItemProperty -Path $sub.PSPath -Name 'Manifest' -ErrorAction Stop).Manifest } catch {}
                if (-not $manifest) {
                    try { $manifest = (Get-ItemProperty -Path $sub.PSPath -Name 'Path' -ErrorAction Stop).Path } catch {}
                }
                [void]$regItems.Add(@{
                    Name     = $sub.PSChildName
                    RegPath  = $rk.Path
                    RegSub   = $sub.PSChildName
                    Manifest = $manifest
                    App      = $rk.App
                    Source   = 'reg_subkey'
                    Tag      = "$($rk.App) / COM add-in"
                })
            }
        }

        if ($rk.Type -eq 'open') {
            $props = Get-ItemProperty -Path $rk.Path 2>$null
            if (-not $props) { continue }
            foreach ($pn in $props.PSObject.Properties.Name) {
                if ($pn -match '^OPEN\d*$') {
                    $val = $props.$pn
                    $display = $val.Trim() -replace '^\s*/R\s*', '' -replace '^"', '' -replace '"$', ''
                    $leaf = Split-Path $display -Leaf -ErrorAction SilentlyContinue
                    if (-not $leaf) { $leaf = $val }
                    [void]$regItems.Add(@{
                        Name    = $leaf
                        RegPath = $rk.Path
                        RegVal  = $pn
                        RegData = $val
                        App     = $rk.App
                        Source  = 'reg_value'
                        Tag     = "$($rk.App) / registry-only"
                    })
                }
            }
        }
    }

    # Merge registry-only items (skip those matching a file on disk)
    # Also deduplicate OPEN values: if multiple Office versions reference
    # the same filename, show it once but track all for cleanup
    $openSeen = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($ri in $regItems) {
        $dominated = $false
        foreach ($fn in $fileNames) {
            if ($ri.Source -eq 'reg_value' -and $ri.RegData -and $ri.RegData.ToLower().Contains($fn)) { $dominated = $true; break }
            if ($ri.Source -eq 'reg_subkey' -and $ri.Manifest -and $ri.Manifest.ToLower().Contains($fn)) { $dominated = $true; break }
            if ($ri.Name.ToLower() -eq $fn) { $dominated = $true; break }
        }
        if ($dominated) { continue }

        # Deduplicate OPEN values by displayed filename
        if ($ri.Source -eq 'reg_value') {
            $key = "$($ri.App)|$($ri.Name.ToLower())"
            if ($openSeen.Contains($key)) {
                # Already shown — skip display but it's still in $regItems for cleanup
                continue
            }
            [void]$openSeen.Add($key)
        }

        [void]$allItems.Add($ri)
    }

    if ($allItems.Count -eq 0) {
        Write-Host "`n  No installed add-ins found (files or registry).`n" -ForegroundColor Yellow
        return
    }

    Write-Host "`n  Installed add-ins:`n"
    for ($i = 0; $i -lt $allItems.Count; $i++) {
        $a = $allItems[$i]
        $num = ($i + 1).ToString().PadLeft(2)
        Write-Host "  [$num]  $($a.Name)  ($($a.Tag))"
    }

    Write-Host "`n  Total: $($allItems.Count)"
    Write-Host ''
    Write-Host '  Select items to remove:'
    Write-Host '    Single:   3'
    Write-Host '    Multiple: 1,3,5'
    Write-Host '    Range:    1-4'
    Write-Host '    Mixed:    1-3,5,7'
    Write-Host '    All:      all'
    Write-Host ''

    $choice = Read-Host '  Your selection (or q to cancel)'
    if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq 'q') {
        Write-Host '  Cancelled.'
        return
    }

    $selIndices = Parse-Selection $choice $allItems.Count
    if ($null -eq $selIndices) {
        Write-Host '  Invalid selection.' -ForegroundColor Red
        return
    }

    $selected = @(); foreach ($i in $selIndices) { $selected += $allItems[$i] }

    Write-Host "`n  You selected $($selected.Count) add-in(s):`n"
    foreach ($s in $selected) {
        $loc = if ($s.Path) { $s.Path } else { "$($s.RegPath)\$($s.RegSub)$($s.RegVal)" }
        Write-Host "    - $($s.Name)  =  $loc"
    }

    $confirm = Read-Host "`n  Delete these $($selected.Count) add-in(s)? (y/n)"
    if ($confirm -ne 'y') {
        Write-Host '  Cancelled.'
        return
    }

    # Removal helpers
    function Remove-RegSubkey($item) {
        try {
            Remove-Item -Path "$($item.RegPath)\$($item.RegSub)" -Recurse -Force -ErrorAction Stop
            Write-Host "  [OK]  Removed registry key: $($item.RegPath)\$($item.RegSub)" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "  [ERR] $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    function Remove-RegValue($item) {
        try {
            Remove-ItemProperty -Path $item.RegPath -Name $item.RegVal -Force -ErrorAction Stop
            Write-Host "  [OK]  Removed registry value: $($item.RegPath) -> $($item.RegVal)" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "  [ERR] $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    $okCount = 0; $errCount = 0
    Write-Host ''

    foreach ($s in $selected) {

        if ($s.Source -eq 'file') {
            try {
                Remove-Item -Path $s.Path -Force -ErrorAction Stop
                Write-Host "  [OK]  Deleted file: $($s.Path)" -ForegroundColor Green
                $okCount++
            } catch {
                Write-Host "  [ERR] $($_.Exception.Message)" -ForegroundColor Red
                $errCount++
            }
            # Auto-clean matching registry entries
            $fn = $s.Name.ToLower()
            foreach ($ri in $regItems) {
                $match = $false
                if ($ri.Source -eq 'reg_value' -and $ri.RegData -and $ri.RegData.ToLower().Contains($fn)) { $match = $true }
                if ($ri.Source -eq 'reg_subkey' -and $ri.Manifest -and $ri.Manifest.ToLower().Contains($fn)) { $match = $true }
                if ($ri.Name.ToLower() -eq $fn) { $match = $true }
                if ($match) {
                    if ($ri.Source -eq 'reg_subkey') {
                        if (Remove-RegSubkey $ri) { $okCount++ } else { $errCount++ }
                    } elseif ($ri.Source -eq 'reg_value') {
                        if (Remove-RegValue $ri) { $okCount++ } else { $errCount++ }
                    }
                }
            }
        }

        if ($s.Source -eq 'reg_subkey') {
            if (Remove-RegSubkey $s) { $okCount++ } else { $errCount++ }
        }

        if ($s.Source -eq 'reg_value') {
            # Remove this entry AND all other OPEN entries referencing the same file
            # (covers duplicates across Office 16.0, 15.0, 14.0 etc.)
            $targetName = $s.Name.ToLower()
            $cleaned = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($ri in $regItems) {
                if ($ri.Source -ne 'reg_value') { continue }
                if ($ri.Name.ToLower() -ne $targetName) { continue }
                $cleanKey = "$($ri.RegPath)|$($ri.RegVal)"
                if ($cleaned.Contains($cleanKey)) { continue }
                [void]$cleaned.Add($cleanKey)
                if (Remove-RegValue $ri) { $okCount++ } else { $errCount++ }
            }
        }
    }

    Write-Host ''
    Write-Host "  Remove done: $okCount succeeded, $errCount failed." -ForegroundColor Cyan
    if ($errCount -gt 0) {
        Write-Host "`n  Tip: Close all Office apps and retry if you got permission errors." -ForegroundColor Yellow
    }
}


# =============================================================
# MAIN MENU
# =============================================================

# Count local add-in files to show in menu
$localCount = @(Get-ChildItem -Path $ScriptDir -File | Where-Object { $installExts -contains $_.Extension.ToLower() }).Count

:menuLoop
while ($true) {
    Write-Host ''
    Write-Host '====================================================' -ForegroundColor Cyan
    Write-Host '  Office Add-in Manager' -ForegroundColor Cyan
    Write-Host '====================================================' -ForegroundColor Cyan
    Write-Host ''

    if ($localCount -gt 0) {
        Write-Host "  [I]  Install add-ins   ($localCount file(s) found next to script)" -ForegroundColor Green
    } else {
        Write-Host '  [I]  Install add-ins   (no files found next to script)' -ForegroundColor DarkGray
    }

    Write-Host '  [R]  Remove add-ins    (clean uninstall + registry)' -ForegroundColor Red
    Write-Host '  [Q]  Quit' -ForegroundColor Gray
    Write-Host ''

    $menuChoice = Read-Host '  Choose action'

    switch ($menuChoice.Trim().ToUpper()) {
        'I' {
            if ($localCount -eq 0) {
                Write-Host ''
                Write-Host '  No add-in files found next to this script.' -ForegroundColor Yellow
                Write-Host "  Place .xlam, .ppam, or .dotm files in:" -ForegroundColor Yellow
                Write-Host "    $ScriptDir" -ForegroundColor Yellow
            } else {
                Invoke-Install
            }
        }
        'R' { Invoke-Remove }
        'Q' {
            Write-Host '  Bye.'
            exit
        }
        default {
            Write-Host '  Invalid choice.' -ForegroundColor Yellow
        }
    }

    Write-Host ''
    Read-Host '  Press Enter to return to menu'

    # Recount local files (user might have added some)
    $localCount = @(Get-ChildItem -Path $ScriptDir -File | Where-Object { $installExts -contains $_.Extension.ToLower() }).Count
}
