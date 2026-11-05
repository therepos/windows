# irm https://github.com/therepos/windows/raw/main/filesync.ps1 | iex
# Purpose: Compare two folders and copy the files missing from either side, one way or both
# =============================================================================
#  File Sync:
#   1 = Compare   count the files in A and B and list what is missing from each
#   2 = A -> B    copy the files missing from B
#   3 = B -> A    copy the files missing from A
#   4 = Both      copy both ways so A and B end up with the same files
#   5 = Update A -> B   files that differ and are newer in A: replace B's copy
#   6 = Update B -> A   files that differ and are newer in B: replace A's copy
#   7 = Update both     replace every older copy with the newer one
#                 (same date but different size: always left alone)
#   8 = Settings  change Folder A or Folder B, and the 'filesync' shortcut,
#                 which runs a local copy in %LOCALAPPDATA%\filesync
#   0 = Quit      closes this window
#  Files match by relative path and name. Nothing is ever deleted. Options 2-4
#  never replace a file; only Update (5-7) does, and only an older copy with a newer
#  one, after showing the list and asking.
#  OneDrive / SharePoint folders work too; online-only files are downloaded
#  when copied. Local, network and mapped drives are fine.
#  One keypress picks an option (no Enter). After each action the menu comes
#  back with the last action as the default, so Enter repeats it.
#  Folders are asked each run and not saved (fine on a shared PC); they are
#  kept only while this window is open. No admin, no downloads.
# =============================================================================

$ToolDir      = Join-Path $env:LOCALAPPDATA 'filesync'
$ScriptUrl    = 'https://github.com/therepos/windows/raw/main/filesync.ps1'
$MaxListed    = 40                      # lines shown per list before "... and N more"

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }

# =====================================================================
#  INPUT HELPERS
# =====================================================================

# Test-Path that never prints an error, e.g. for text that is not one path ("a" "b")
function Test-PathQuiet([string]$path, [string]$type = 'Any') {
    if (-not $path) { return $false }
    try { Test-Path -LiteralPath $path -PathType $type -ErrorAction Stop } catch { $false }
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

# Typed or dragged folder -> full path without a trailing slash, or $null (reason shown)
function Resolve-Folder([string]$answer) {
    $path = [Environment]::ExpandEnvironmentVariables($answer.Trim().Trim('"').Trim("'"))
    if ($path -match '^[A-Za-z]:$') { $path += '\' }
    if (-not (Test-PathQuiet $path 'Container')) {
        Say "  Cannot open '$path'. Check the path, or that OneDrive is signed in." 'Red'
        return $null
    }
    try { Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop | Select-Object -First 1 | Out-Null }
    catch { Say "  Cannot read '$path': $($_.Exception.Message)" 'Red'; return $null }
    $p = (Resolve-Path -LiteralPath $path).ProviderPath
    if ($p.Length -gt 3) { $p = $p.TrimEnd('\') }
    return $p
}

# True when the two folders are the same or one is inside the other
function Test-Overlap([string]$a, [string]$b) {
    if (-not $a -or -not $b) { return $false }
    $a = $a.TrimEnd('\') + '\'
    $b = $b.TrimEnd('\') + '\'
    $a.StartsWith($b, [StringComparison]::OrdinalIgnoreCase) -or $b.StartsWith($a, [StringComparison]::OrdinalIgnoreCase)
}

# =====================================================================
#  FOLDERS  (asked on first use in each run; nothing is saved to disk)
# =====================================================================
$Settings = [pscustomobject]@{ FolderA = $null; FolderB = $null }

function Set-FolderSetting([string]$Which) {
    $current = $Settings."Folder$Which"
    $other   = if ($Which -eq 'A') { $Settings.FolderB } else { $Settings.FolderA }
    Say ""
    Say "  Folder ${Which}: type or drag it here (a drive like Z:\ works too)." 'Cyan'
    while ($true) {
        $answer = Read-Host "  Folder $Which$(if ($current) { " [$current]" })"
        if ([string]::IsNullOrWhiteSpace($answer)) {
            if ($current) { Say "  Folder ${Which}: $current" 'Gray'; return }
            continue
        }
        $p = Resolve-Folder $answer
        if (-not $p) { continue }
        if (Test-Overlap $p $other) {
            Say "  Folder A and Folder B must be separate folders, not one inside the other." 'Red'
            continue
        }
        $Settings."Folder$Which" = $p
        Say "  Folder ${Which}: $p" 'Green'
        return
    }
}

# Both folders must be set (asked on first use) and reachable right now
function Confirm-Folders {
    foreach ($w in 'A', 'B') {
        $cur = $Settings."Folder$w"
        if ($cur -and (Test-PathQuiet $cur 'Container')) { continue }
        if ($cur) { Say "  Folder $w is not available right now: $cur" 'Yellow' }
        Set-FolderSetting $w
    }
}

# Optional one-word shortcut: typing `filesync` in any new PowerShell window runs a local copy
# of this script saved in %LOCALAPPDATA%\filesync (nothing is fetched from the internet)
$ScriptFile   = $PSCommandPath
$LocalScript  = Join-Path $ToolDir 'filesync.ps1'
$ShortcutLine = 'function filesync { & "$env:LOCALAPPDATA\filesync\filesync.ps1" }'

function Test-Shortcut {
    (Test-Path -LiteralPath $PROFILE) -and (Test-Path -LiteralPath $LocalScript) -and
        (@(Get-Content -LiteralPath $PROFILE -ErrorAction SilentlyContinue) -contains $ShortcutLine)
}

function Switch-Shortcut {
    try {
        $on = Test-Shortcut
        if (Test-Path -LiteralPath $PROFILE) {
            @(Get-Content -LiteralPath $PROFILE) | Where-Object { $_ -notmatch '^function filesync ' } | Set-Content -LiteralPath $PROFILE
        }
        if ($on) {
            Remove-Item -LiteralPath $LocalScript -Force -ErrorAction SilentlyContinue
            Say "  Shortcut removed." 'Green'
            return
        }
        New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null
        if ($ScriptFile -and (Test-Path -LiteralPath $ScriptFile)) {
            if ($ScriptFile -ne $LocalScript) { Copy-Item -LiteralPath $ScriptFile -Destination $LocalScript -Force }
        } else {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $ScriptUrl -OutFile $LocalScript -UseBasicParsing
        }
        Unblock-File -LiteralPath $LocalScript -ErrorAction SilentlyContinue
        if ((Get-ExecutionPolicy -Scope CurrentUser) -in 'Undefined', 'Restricted', 'AllSigned') {
            Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
        }
        if (-not (Test-Path -LiteralPath $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }
        Add-Content -LiteralPath $PROFILE "`n$ShortcutLine"
        Say "  Shortcut added. In a new PowerShell window, just type: filesync" 'Green'
    } catch {
        Say "  Could not change the shortcut: $($_.Exception.Message)" 'Red'
    }
}

function Show-Settings {
    while ($true) {
        $a     = if ($Settings.FolderA) { $Settings.FolderA } else { 'not set' }
        $b     = if ($Settings.FolderB) { $Settings.FolderB } else { 'not set' }
        $short = if (Test-Shortcut) { 'on' } else { 'off' }
        Say ""
        Say "  ---------------- SETTINGS -------------------" 'Cyan'
        Say ""
        Say "    1  -  Set Folder A           [$a]" 'White'
        Say "    2  -  Set Folder B           [$b]" 'White'
        Say "    3  -  Shortcut 'filesync'    [$short]" 'White'
        Say "    4  -  Back" 'White'
        Say "    0  -  Quit" 'White'
        Say ""
        switch (Read-MenuKey 'Choose 0-4' '01234') {
            '1' { Set-FolderSetting 'A' }
            '2' { Set-FolderSetting 'B' }
            '3' { Switch-Shortcut }
            '4' { return 'back' }
            '0' { return 'quit' }
        }
    }
}

# =====================================================================
#  COMPARE
# =====================================================================

# Every file under $Root keyed by relative path (case-insensitive); unreadable items go to Errors
function Get-FileMap([string]$Root) {
    $map    = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    $errors = New-Object System.Collections.Generic.List[string]
    $scanErrors = @()
    $items = @(Get-ChildItem -LiteralPath $Root -File -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors)
    foreach ($f in $items) {
        $rel = $f.FullName.Substring($Root.Length).TrimStart('\')
        if (-not $map.ContainsKey($rel)) { $map.Add($rel, $f) }
    }
    foreach ($e in $scanErrors) { $errors.Add("$($e.TargetObject): $($e.Exception.Message)") }
    [pscustomobject]@{ Map = $map; Errors = $errors }
}

function Get-Comparison {
    Say ""
    Say "  Reading A: $($Settings.FolderA)" 'DarkGray'
    $a = Get-FileMap $Settings.FolderA
    Say "  Reading B: $($Settings.FolderB)" 'DarkGray'
    $b = Get-FileMap $Settings.FolderB
    # In both, but not the same file: Newer = 'A' or 'B', or '' when the dates match (size differs)
    $different = foreach ($k in $a.Map.Keys) {
        if (-not $b.Map.ContainsKey($k)) { continue }
        $fa = $a.Map[$k]; $fb = $b.Map[$k]
        $gap = ($fa.LastWriteTimeUtc - $fb.LastWriteTimeUtc).TotalSeconds
        if ($fa.Length -eq $fb.Length -and [Math]::Abs($gap) -le 2) { continue }
        $newer = if ([Math]::Abs($gap) -le 2) { '' } elseif ($gap -gt 0) { 'A' } else { 'B' }
        [pscustomobject]@{ Path = $k; Newer = $newer }
    }
    [pscustomobject]@{
        A         = $a
        B         = $b
        MissingB  = @($a.Map.Keys | Where-Object { -not $b.Map.ContainsKey($_) } | Sort-Object)
        MissingA  = @($b.Map.Keys | Where-Object { -not $a.Map.ContainsKey($_) } | Sort-Object)
        Different = @($different | Sort-Object Path)
        Errors    = @($a.Errors) + @($b.Errors)
    }
}

function Show-List([string]$title, [string[]]$items, [string]$color, [string]$indent = '  ') {
    Say "$indent$title ($($items.Count))" $color
    if (-not $items.Count) { Say "$indent  none" 'Green'; return }
    $items | Select-Object -First $MaxListed | ForEach-Object { Say "$indent  $_" 'White' }
    if ($items.Count -gt $MaxListed) { Say "$indent  ... and $($items.Count - $MaxListed) more" 'DarkGray' }
}

# Files in both but different, grouped: A Newer, B Newer, Same Date, Different Size
function Show-Different([string]$title, $items) {
    Say "  $title ($(@($items).Count))" 'Cyan'
    if (-not @($items).Count) { Say "    none" 'Green'; return }
    $groups = @(
        @{ Name = 'A Newer';                   Items = @($items | Where-Object { $_.Newer -eq 'A' }) },
        @{ Name = 'B Newer';                   Items = @($items | Where-Object { $_.Newer -eq 'B' }) },
        @{ Name = 'Same Date, Different Size'; Items = @($items | Where-Object { -not $_.Newer }) }
    )
    foreach ($g in $groups) {
        if (-not $g.Items.Count) { continue }
        Say ""
        Show-List $g.Name @($g.Items | ForEach-Object { $_.Path }) 'White' '    '
    }
}

function Show-Errors([string[]]$errors) {
    if (-not $errors.Count) { return }
    Say ""
    Say "  $($errors.Count) item(s) could not be read:" 'Yellow'
    $errors | Select-Object -First 10 | ForEach-Object { Say "    $_" 'Yellow' }
    if ($errors.Count -gt 10) { Say "    ... and $($errors.Count - 10) more" 'DarkGray' }
    Say "  Usual fixes: close the file if it is open, check the folder opens in File Explorer," 'DarkGray'
    Say "  or for OneDrive files choose 'Always keep on this device'." 'DarkGray'
}

function Invoke-Compare {
    $c = Get-Comparison
    Say ""
    Say "  A: $($c.A.Map.Count) file(s)    B: $($c.B.Map.Count) file(s)" 'White'
    Say ""
    Show-List 'MISSING FROM B (ONLY IN A)' $c.MissingB 'Cyan'
    Say ""
    Show-List 'MISSING FROM A (ONLY IN B)' $c.MissingA 'Cyan'
    Say ""
    Show-Different 'IN BOTH BUT DIFFERENT' $c.Different
    Show-Errors $c.Errors
    Say ""
    if (-not $c.MissingA.Count -and -not $c.MissingB.Count -and -not $c.Different.Count) { Say "  DONE - A and B have the same files." 'Green' }
    else { Say "  DONE - 2, 3 or 4 copies the missing files; 5, 6 or 7 replaces older copies with newer ones." 'Green' }
}

# =====================================================================
#  COPY
# =====================================================================

# Copy the files in $Missing from $Source into Folder $ToName; nothing there is ever replaced
function Copy-Missing([string]$FromName, [string]$ToName, $Source, [string[]]$Missing) {
    $to = $Settings."Folder$ToName"
    Say ""
    if (-not $Missing.Count) { Say "  $FromName -> ${ToName}: nothing to copy, $ToName already has every file in $FromName." 'Green'; return }
    Show-List "$FromName -> ${ToName}: FILES TO COPY" $Missing 'Cyan'
    Say ""
    if ((Read-MenuKey "Copy $($Missing.Count) file(s) into Folder ${ToName}? y/n" 'ynYN') -notin 'y', 'Y') { Say "  Skipped." 'Yellow'; return }

    $copied = 0; $n = 0
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($rel in $Missing) {
        $n++
        try {
            $dst = [IO.Path]::Combine($to, $rel)
            [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dst)) | Out-Null
            if ([IO.File]::Exists($dst)) { continue }        # appeared meanwhile: never replace
            Say "    [$n/$($Missing.Count)] $rel" 'DarkGray'
            [IO.File]::Copy($Source[$rel].FullName, $dst, $false)
            $copied++
        } catch {
            $failed.Add("${rel}: $($_.Exception.Message)")
        }
    }
    Say "  DONE - $copied of $($Missing.Count) copied into Folder $ToName" $(if ($failed.Count) { 'Yellow' } else { 'Green' })
    Show-Errors $failed
}

# Scan errors mean the comparison may be incomplete: say so and let the user decide
function Confirm-Scan($c) {
    if (-not $c.Errors.Count) { return $true }
    Show-Errors $c.Errors
    Say "  Some items could not be read, so the comparison may be incomplete." 'Yellow'
    if ((Read-MenuKey 'Continue anyway? y/n' 'ynYN') -in 'y', 'Y') { return $true }
    Say "  Cancelled." 'Yellow'
    return $false
}

function Invoke-Sync([bool]$AtoB, [bool]$BtoA) {
    $c = Get-Comparison
    if (-not (Confirm-Scan $c)) { return }
    if ($AtoB) { Copy-Missing 'A' 'B' $c.A.Map $c.MissingB }
    if ($BtoA) { Copy-Missing 'B' 'A' $c.B.Map $c.MissingA }
}

# =====================================================================
#  UPDATE: replace the older of two differing copies with the newer one
# =====================================================================
# $AtoB: replace B's copy where A is newer; $BtoA: replace A's copy where B is newer
function Invoke-Update([bool]$AtoB, [bool]$BtoA) {
    $c = Get-Comparison
    if (-not (Confirm-Scan $c)) { return }
    $todo = @($c.Different | Where-Object { ($AtoB -and $_.Newer -eq 'A') -or ($BtoA -and $_.Newer -eq 'B') })
    $tied = @($c.Different | Where-Object { -not $_.Newer })
    $other = @($c.Different | Where-Object { (-not $AtoB -and $_.Newer -eq 'A') -or (-not $BtoA -and $_.Newer -eq 'B') })
    Say ""
    if ($tied.Count) {
        Show-List 'SAME DATE, DIFFERENT SIZE - LEFT ALONE, CHECK THESE YOURSELF' @($tied | ForEach-Object { $_.Path }) 'Yellow'
        Say ""
    }
    if ($other.Count) {
        $side = if ($AtoB) { 'B' } else { 'A' }
        Say "  $($other.Count) file(s) newer in $side left alone (use $(if ($AtoB) { '6' } else { '5' }) or 7 for those)." 'Yellow'
        Say ""
    }
    if (-not $todo.Count) { Say "  Nothing to update." 'Green'; return }
    Show-Different 'OLDER COPIES TO REPLACE WITH THE NEWER VERSION' $todo
    Say ""
    if ((Read-MenuKey "Replace $($todo.Count) older file(s)? y/n" 'ynYN') -notin 'y', 'Y') { Say "  Skipped." 'Yellow'; return }

    $done = 0; $n = 0
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($d in $todo) {
        $n++
        try {
            $src = $(if ($d.Newer -eq 'A') { $c.A.Map } else { $c.B.Map })[$d.Path].FullName
            $dst = $(if ($d.Newer -eq 'A') { $c.B.Map } else { $c.A.Map })[$d.Path].FullName
            Say "    [$n/$($todo.Count)] $($d.Path)  ($($d.Newer) -> $(if ($d.Newer -eq 'A') { 'B' } else { 'A' }))" 'DarkGray'
            [IO.File]::Copy($src, $dst, $true)
            $done++
        } catch {
            $failed.Add("$($d.Path): $($_.Exception.Message)")
        }
    }
    Say "  DONE - $done of $($todo.Count) updated" $(if ($failed.Count) { 'Yellow' } else { 'Green' })
    Show-Errors $failed
}

# =====================================================================
#  MENU
# =====================================================================
Clear-Host
$last = ''
while ($true) {
    $a = if ($Settings.FolderA) { $Settings.FolderA } else { 'not set' }
    $b = if ($Settings.FolderB) { $Settings.FolderB } else { 'not set' }
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        FILE SYNC" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say "    A  $a" 'DarkGray'
    Say "    B  $b" 'DarkGray'
    Say ""
    Say "    1  -  Compare" 'White'
    Say "    2  -  Copy A -> B" 'White'
    Say "    3  -  Copy B -> A" 'White'
    Say "    4  -  Copy both ways" 'White'
    Say "    5  -  Update A -> B" 'White'
    Say "    6  -  Update B -> A" 'White'
    Say "    7  -  Update both ways" 'White'
    Say "    8  -  Settings" 'White'
    Say "    0  -  Quit" 'White'
    Say ""
    $choice = Read-MenuKey 'Choose 0-8' '012345678' $last
    if ($choice -eq '0') { break }
    if ($choice -eq '8') {
        if (@(Show-Settings) -contains 'quit') { break }
        Clear-Host
        continue
    }
    try {
        Confirm-Folders
        switch ($choice) {
            '1' { Invoke-Compare }
            '2' { Invoke-Sync $true  $false }
            '3' { Invoke-Sync $false $true }
            '4' { Invoke-Sync $true  $true }
            '5' { Invoke-Update $true  $false }
            '6' { Invoke-Update $false $true }
            '7' { Invoke-Update $true  $true }
        }
    } catch {
        Say "  FAILED: $($_.Exception.Message)" 'Red'
    }
    $last = $choice
    Say ""
}
# Quit closes this window, not just the menu (errors above skip this and leave it open)
[Environment]::Exit(0)
