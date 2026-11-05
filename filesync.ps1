# irm https://github.com/therepos/windows/raw/main/filesync.ps1 | iex
# Purpose: Compare two folders by file content, copy what is missing, replace older versions, remove duplicates
# =============================================================================
#  File Sync:
#   1 = Compare A and B   count the files in A and B; each list can be viewed a page at a
#                         time or saved as Report FileSync yyyyMMdd-HHmm.txt in Downloads
#   2 = Add: A -> B       copy the files B does not have
#   3 = Add: B -> A       copy the files A does not have
#   4 = Add: A <> B       2 and 3 together; never replaces a file
#   5 = Replace: A -> B   same file name in the same place but newer in A: replace B's copy
#   6 = Replace: B -> A   same file name in the same place but newer in B: replace A's copy
#   7 = Sync              add missing files both ways, then replace every older copy
#                         with the newer one, so A and B end up the same
#                         (same date but different content: always left alone)
#   8 = Duplicates  find identical files inside A or inside B, keep the oldest
#                   copy of each and send the rest to the Recycle Bin
#   9 = Settings    change Folder A or Folder B, and the 'filesync' shortcut,
#                   which runs a local copy in %LOCALAPPDATA%\filesync
#   0 = Quit        closes this window
#  Two files are the same file when their name, size and contents (SHA-256) match,
#  whatever folder they sit in. So a file that B keeps in a different folder is not
#  "missing" and is never copied again. Only files whose name and size appear on the
#  other side have their contents read, so most runs stay quick.
#  Options 2-4 never replace a file; only Replace (5-6) and Sync (7) do,
#  and only an older copy with a newer one, after showing the list and asking.
#  Only Duplicates (8) deletes, and only to the Recycle Bin, after showing the list
#  and asking. Network and mapped drives have no Recycle Bin, so there it warns
#  that the delete is permanent and asks you to type yes.
#  OneDrive / SharePoint folders work too; online-only files are downloaded when
#  read or copied. Local, network and mapped drives are fine.
#  One keypress picks an option (no Enter). After each action the menu comes
#  back with the last action as the default, so Enter repeats it. Before a copy,
#  replace or delete the first 10 files are shown; v on the y/n/v prompt shows all.
#  Folders are asked each run and not saved (fine on a shared PC); they are
#  kept only while this window is open. No admin, no downloads.
# =============================================================================

$ToolDir      = Join-Path $env:LOCALAPPDATA 'filesync'
$ScriptUrl    = 'https://github.com/therepos/windows/raw/main/filesync.ps1'
$Preview      = 10                      # lines shown before "... and N more"
$PageSize     = 40                      # lines per page when viewing a full list

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

function Get-DownloadsFolder {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $raw = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'{374DE290-123F-4565-9164-39C4925E467B}'
    if ($raw) {
        $path = [Environment]::ExpandEnvironmentVariables($raw)
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return (Join-Path $env:USERPROFILE 'Downloads')
}

function Format-Size([long]$bytes) {
    if ($bytes -ge 1GB) { return '{0:N1} GB' -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return '{0:N0} MB' -f ($bytes / 1MB) }
    return '{0:N0} KB' -f [Math]::Max(1, $bytes / 1KB)
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

# The named folders must be set (asked on first use) and reachable right now
function Confirm-Folders([string[]]$Which = @('A', 'B')) {
    foreach ($w in $Which) {
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
#  SCAN
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

# "name|size" -> list of relative paths with that name and size. Only files that share
# a name and size with another file can be the same file, so only those get read.
function Get-NameSizeIndex($map) {
    $ix = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($rel in $map.Keys) {
        $f = $map[$rel]
        $k = "$($f.Name)|$($f.Length)"
        if (-not $ix.ContainsKey($k)) { $ix[$k] = New-Object System.Collections.Generic.List[string] }
        $ix[$k].Add($rel)
    }
    $ix
}

$Sha256 = [System.Security.Cryptography.SHA256]::Create()

# Content key "name|size|sha256" (lower case), or $null when the file cannot be read
function Get-ContentKey($f, $errors) {
    try {
        $s = [IO.File]::OpenRead($f.FullName)
        try { $h = [BitConverter]::ToString($Sha256.ComputeHash($s)).Replace('-', '') } finally { $s.Dispose() }
        return "$($f.Name)|$($f.Length)|$h".ToLowerInvariant()
    } catch {
        $errors.Add("$($f.FullName): $($_.Exception.Message)")
        return $null
    }
}

# Read the contents of the given files, with a one-line progress counter
function Read-Contents($files, $errors) {
    $keys = @{}
    $n = 0; $total = @($files).Count
    foreach ($f in $files) {
        $n++
        Write-Host "`r  Checking contents... $n of $total" -NoNewline -ForegroundColor DarkGray
        $keys[$f.FullName] = Get-ContentKey $f $errors
    }
    if ($total) { Write-Host "`r  Checking contents... $total file(s) read          " -ForegroundColor DarkGray }
    $keys
}

# Key per file: the content key when read, "name|size" when nothing on the other side
# could match it, $null when it could not be read (then it is left out of every action)
function Get-Keys($map, $hashes) {
    $keys = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($rel in $map.Keys) {
        $f = $map[$rel]
        if ($hashes.ContainsKey($f.FullName)) { $keys[$rel] = $hashes[$f.FullName] }
        else { $keys[$rel] = "$($f.Name)|$($f.Length)".ToLowerInvariant() }
    }
    $keys
}

# Content key -> relative paths holding it
function Get-ByKey($keys) {
    $by = @{}
    foreach ($rel in $keys.Keys) {
        $k = $keys[$rel]
        if (-not $k) { continue }
        if (-not $by.ContainsKey($k)) { $by[$k] = New-Object System.Collections.Generic.List[string] }
        $by[$k].Add($rel)
    }
    $by
}

# =====================================================================
#  COMPARE
# =====================================================================

function Get-Comparison {
    Say ""
    Say "  Reading A: $($Settings.FolderA)" 'DarkGray'
    $a = Get-FileMap $Settings.FolderA
    Say "  Reading B: $($Settings.FolderB)" 'DarkGray'
    $b = Get-FileMap $Settings.FolderB
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($e in $a.Errors) { $errors.Add($e) }
    foreach ($e in $b.Errors) { $errors.Add($e) }

    # Files whose name and size appear on the other side are read; the rest cannot match
    $ixA = Get-NameSizeIndex $a.Map
    $ixB = Get-NameSizeIndex $b.Map
    $toRead = New-Object System.Collections.Generic.List[object]
    foreach ($k in $ixA.Keys) {
        if (-not $ixB.ContainsKey($k)) { continue }
        foreach ($rel in $ixA[$k]) { $toRead.Add($a.Map[$rel]) }
        foreach ($rel in $ixB[$k]) { $toRead.Add($b.Map[$rel]) }
    }
    $hashes = Read-Contents $toRead $errors
    $keysA = Get-Keys $a.Map $hashes
    $keysB = Get-Keys $b.Map $hashes
    $byA = Get-ByKey $keysA
    $byB = Get-ByKey $keysB

    $missingB  = New-Object System.Collections.Generic.List[string]
    $missingA  = New-Object System.Collections.Generic.List[string]
    $different = New-Object System.Collections.Generic.List[object]
    $elsewhere = New-Object System.Collections.Generic.List[string]
    $shown     = @{}

    foreach ($rel in $a.Map.Keys) {
        $k = $keysA[$rel]
        if (-not $k) { continue }
        if ($b.Map.ContainsKey($rel)) {
            $kb = $keysB[$rel]
            if (-not $kb) { continue }
            if ($kb -eq $k) { continue }                                      # same file, same place
            # Same place, different content: an edited version. Newer = 'A' or 'B', '' when dates match
            $gap = ($a.Map[$rel].LastWriteTimeUtc - $b.Map[$rel].LastWriteTimeUtc).TotalSeconds
            $newer = if ([Math]::Abs($gap) -le 2) { '' } elseif ($gap -gt 0) { 'A' } else { 'B' }
            $different.Add([pscustomobject]@{ Path = $rel; Newer = $newer })
        } elseif ($byB.ContainsKey($k)) {
            $elsewhere.Add("A: $rel`n        = B: $($byB[$k][0])")            # same file, kept elsewhere in B
            $shown[$k] = $true
        } else {
            $missingB.Add($rel)
        }
    }
    foreach ($rel in $b.Map.Keys) {
        $k = $keysB[$rel]
        if (-not $k -or $a.Map.ContainsKey($rel)) { continue }
        if ($byA.ContainsKey($k)) {
            if (-not $shown.ContainsKey($k)) { $elsewhere.Add("B: $rel`n        = A: $($byA[$k][0])") }
        } else {
            $missingA.Add($rel)
        }
    }

    [pscustomobject]@{
        A         = $a
        B         = $b
        MissingB  = @($missingB | Sort-Object)
        MissingA  = @($missingA | Sort-Object)
        Different = @($different | Sort-Object Path)
        Elsewhere = @($elsewhere | Sort-Object)
        Errors    = @($errors)
    }
}

# Title, count and the first few lines; the rest is a count ("v" on the next prompt views all)
function Show-List([string]$title, [string[]]$items, [string]$color, [string]$indent = '  ') {
    Say "$indent$title ($($items.Count))" $color
    if (-not $items.Count) { Say "$indent  none" 'Green'; return }
    $items | Select-Object -First $Preview | ForEach-Object { Say "$indent  $_" 'White' }
    if ($items.Count -gt $Preview) { Say "$indent  ... and $($items.Count - $Preview) more" 'DarkGray' }
}

# The whole list, a page at a time
function Show-Page([string]$title, [string[]]$items) {
    if (-not $items.Count) { Say ""; Say "  $title (0)" 'Cyan'; Say "    none" 'Green'; return }
    for ($i = 0; $i -lt $items.Count; $i += $PageSize) {
        $end = [Math]::Min($i + $PageSize, $items.Count)
        Say ""
        Say "  $title  $($i + 1)-$end of $($items.Count)" 'Cyan'
        for ($j = $i; $j -lt $end; $j++) { Say "    $($items[$j])" 'White' }
        if ($end -ge $items.Count) { break }
        if ((Read-MenuKey "Next $([Math]::Min($PageSize, $items.Count - $end))? y/n" 'ynYN') -notin 'y', 'Y') { break }
    }
}

# y/n question that also accepts v to view the full list first
function Confirm-Action([string]$Prompt, [string]$Title, [string[]]$Items) {
    while ($true) {
        $k = Read-MenuKey "$Prompt y/n/v" 'ynvYNV'
        if ($k -in 'v', 'V') { Show-Page $Title $Items; Say ""; continue }
        return ($k -in 'y', 'Y')
    }
}

# Files in both but different, as "path  (A newer)" lines
function Format-Different($items) {
    @($items | ForEach-Object {
        $tag = if ($_.Newer) { "$($_.Newer) newer" } else { 'same date, different content' }
        "$($_.Path)  ($tag)"
    })
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

# Everything Compare found, as one text file in Downloads
function Save-Report($c, $sections) {
    $name = "Report FileSync $(Get-Date -Format 'yyyyMMdd-HHmm').txt"
    $path = Join-Path (Get-DownloadsFolder) $name
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("File Sync report  $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    $lines.Add("A: $($Settings.FolderA)  ($($c.A.Map.Count) files)")
    $lines.Add("B: $($Settings.FolderB)  ($($c.B.Map.Count) files)")
    foreach ($sec in $sections) {
        $lines.Add(''); $lines.Add("$($sec.Title) ($($sec.Items.Count))")
        if (-not $sec.Items.Count) { $lines.Add('  none') }
        foreach ($i in $sec.Items) { $lines.Add("  $($i.Replace("`n", "`n  "))") }
    }
    if ($c.Errors.Count) {
        $lines.Add(''); $lines.Add("COULD NOT BE READ ($($c.Errors.Count))")
        foreach ($e in $c.Errors) { $lines.Add("  $e") }
    }
    $lines | Out-File -LiteralPath $path -Encoding UTF8
    Say "  Saved: $path" 'Green'
}

# Counts first; each list can be viewed a page at a time or saved as a report
function Invoke-Compare {
    $c = Get-Comparison
    $sections = @(
        @{ Title = 'MISSING FROM B (ONLY IN A)';                              Items = $c.MissingB },
        @{ Title = 'MISSING FROM A (ONLY IN B)';                              Items = $c.MissingA },
        @{ Title = 'SAME FILE, DIFFERENT PLACE - NOT COPIED, THE OTHER SIDE HAS IT'; Items = $c.Elsewhere },
        @{ Title = 'IN BOTH BUT DIFFERENT (SAME PLACE, DIFFERENT CONTENT)';  Items = (Format-Different $c.Different) }
    )
    Show-Errors $c.Errors
    while ($true) {
        Say ""
        Say "  A: $($c.A.Map.Count) file(s)    B: $($c.B.Map.Count) file(s)" 'White'
        if (-not $c.MissingA.Count -and -not $c.MissingB.Count -and -not $c.Different.Count) { Say "  A and B have the same files." 'Green' }
        Say ""
        Say "    1  -  Missing from B (only in A)      $($c.MissingB.Count.ToString().PadLeft(6))" 'White'
        Say "    2  -  Missing from A (only in B)      $($c.MissingA.Count.ToString().PadLeft(6))" 'White'
        Say "    3  -  Same file, different place      $($c.Elsewhere.Count.ToString().PadLeft(6))" 'White'
        Say "    4  -  In both but different           $($c.Different.Count.ToString().PadLeft(6))" 'White'
        Say "    5  -  Save report to Downloads" 'White'
        Say "    6  -  Back" 'White'
        Say "    0  -  Quit" 'White'
        $k = Read-MenuKey 'Choose 0-6' '0123456'
        switch ($k) {
            '1' { Show-Page $sections[0].Title $sections[0].Items }
            '2' { Show-Page $sections[1].Title $sections[1].Items }
            '3' { Show-Page $sections[2].Title $sections[2].Items }
            '4' { Show-Page $sections[3].Title $sections[3].Items }
            '5' { Save-Report $c $sections }
            '6' { return 'back' }
            '0' { return 'quit' }
        }
    }
}

# =====================================================================
#  COPY
# =====================================================================

# Copy the files in $Missing from $Source into Folder $ToName; nothing there is ever replaced
function Copy-Missing([string]$FromName, [string]$ToName, $Source, [string[]]$Missing) {
    $to = $Settings."Folder$ToName"
    Say ""
    if (-not $Missing.Count) { Say "  $FromName -> ${ToName}: nothing to copy, $ToName already has every file in $FromName." 'Green'; return }
    $title = "$FromName -> ${ToName}: FILES TO COPY"
    Show-List $title $Missing 'Cyan'
    Say ""
    if (-not (Confirm-Action "Copy $($Missing.Count) file(s) into Folder ${ToName}?" $title $Missing)) { Say "  Skipped." 'Yellow'; return }

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
    Say "  Some items could not be read, so the comparison may be incomplete. Those files are left alone." 'Yellow'
    if ((Read-MenuKey 'Continue anyway? y/n' 'ynYN') -in 'y', 'Y') { return $true }
    Say "  Cancelled." 'Yellow'
    return $false
}

function Invoke-Sync([bool]$AtoB, [bool]$BtoA) {
    $c = Get-Comparison
    if (-not (Confirm-Scan $c)) { return }
    if ($c.Elsewhere.Count) { Say "  $($c.Elsewhere.Count) file(s) already on the other side in a different folder: not copied (see 1 - Compare)." 'Yellow' }
    if ($AtoB) { Copy-Missing 'A' 'B' $c.A.Map $c.MissingB }
    if ($BtoA) { Copy-Missing 'B' 'A' $c.B.Map $c.MissingA }
}

# =====================================================================
#  UPDATE: replace the older of two differing copies with the newer one
# =====================================================================
# $AtoB: replace B's copy where A is newer; $BtoA: replace A's copy where B is newer
function Invoke-Update([bool]$AtoB, [bool]$BtoA, $c = $null) {
    if (-not $c) {
        $c = Get-Comparison
        if (-not (Confirm-Scan $c)) { return }
    }
    $todo = @($c.Different | Where-Object { ($AtoB -and $_.Newer -eq 'A') -or ($BtoA -and $_.Newer -eq 'B') })
    $tied = @($c.Different | Where-Object { -not $_.Newer })
    $other = @($c.Different | Where-Object { (-not $AtoB -and $_.Newer -eq 'A') -or (-not $BtoA -and $_.Newer -eq 'B') })
    Say ""
    if ($tied.Count) {
        Show-List 'SAME DATE, DIFFERENT CONTENT - LEFT ALONE, CHECK THESE YOURSELF' @($tied | ForEach-Object { $_.Path }) 'Yellow'
        Say ""
    }
    if ($other.Count) {
        $side = if ($AtoB) { 'B' } else { 'A' }
        Say "  $($other.Count) file(s) newer in $side left alone (use $(if ($AtoB) { '6' } else { '5' }) or 7 for those)." 'Yellow'
        Say ""
    }
    if (-not $todo.Count) {
        Say "  Nothing to replace." 'Green'
        $miss = $null
        if ($AtoB -and -not $BtoA) { $miss = @{ To = 'B'; N = $c.MissingB.Count; Key = '2' } }
        if ($BtoA -and -not $AtoB) { $miss = @{ To = 'A'; N = $c.MissingA.Count; Key = '3' } }
        if ($miss -and $miss.N) { Say "  $($miss.N) file(s) are missing from $($miss.To); use $($miss.Key) - Add to copy them." 'Yellow' }
        return
    }
    $title = 'OLDER COPIES TO REPLACE WITH THE NEWER VERSION'
    $lines = Format-Different $todo
    Show-List $title $lines 'Cyan'
    Say ""
    if (-not (Confirm-Action "Replace $($todo.Count) older file(s)?" $title $lines)) { Say "  Skipped." 'Yellow'; return }

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

# Sync: copy the missing files both ways, then replace older copies with newer ones.
# One scan is enough: copying only adds files that were missing, so the list of
# files in both but different is still right.
function Invoke-FullSync {
    $c = Get-Comparison
    if (-not (Confirm-Scan $c)) { return }
    if ($c.Elsewhere.Count) { Say "  $($c.Elsewhere.Count) file(s) already on the other side in a different folder: not copied (see 1 - Compare)." 'Yellow' }
    Copy-Missing 'A' 'B' $c.A.Map $c.MissingB
    Copy-Missing 'B' 'A' $c.B.Map $c.MissingA
    Invoke-Update $true $true $c
}

# =====================================================================
#  DUPLICATES: identical files inside one folder; keep the oldest copy
# =====================================================================

# Local drives (OneDrive included) have a Recycle Bin; network and mapped drives do not
function Test-RecycleBin([string]$folder) {
    $root = [IO.Path]::GetPathRoot($folder)
    if ($root.StartsWith('\\')) { return $false }
    try { return ((New-Object IO.DriveInfo $root).DriveType -ne 'Network') } catch { return $false }
}

# Delete through the Windows shell: the only route that behaves on OneDrive folders
function Remove-ToBin([string]$path, [bool]$bin, [bool]$folder = $false) {
    $how = if ($bin) { 'SendToRecycleBin' } else { 'DeletePermanently' }
    if ($folder) { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($path, 'OnlyErrorDialogs', $how) }
    else         { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($path, 'OnlyErrorDialogs', $how) }
}

function Invoke-Duplicates([string]$Which) {
    Confirm-Folders @($Which)
    $root = $Settings."Folder$Which"
    Say ""
    Say "  Reading ${Which}: $root" 'DarkGray'
    $m = Get-FileMap $root
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($e in $m.Errors) { $errors.Add($e) }

    # Only files that share a name and size with another file in the folder are read
    $ix = Get-NameSizeIndex $m.Map
    $toRead = New-Object System.Collections.Generic.List[object]
    foreach ($k in $ix.Keys) {
        if ($ix[$k].Count -lt 2) { continue }
        foreach ($rel in $ix[$k]) { $toRead.Add($m.Map[$rel]) }
    }
    $hashes = Read-Contents $toRead $errors

    $groups = @{}
    foreach ($f in $toRead) {
        $k = $hashes[$f.FullName]
        if (-not $k) { continue }
        if (-not $groups.ContainsKey($k)) { $groups[$k] = New-Object System.Collections.Generic.List[object] }
        $groups[$k].Add($f)
    }
    # In each group keep the copy that was there first (creation date), then the shortest path
    $todo = New-Object System.Collections.Generic.List[object]
    $bytes = [long]0
    foreach ($k in $groups.Keys) {
        if ($groups[$k].Count -lt 2) { continue }
        $sorted = @($groups[$k] | Sort-Object CreationTimeUtc, { $_.FullName.Length }, FullName)
        $keep = $sorted[0]
        foreach ($f in $sorted[1..($sorted.Count - 1)]) {
            $todo.Add([pscustomobject]@{ File = $f; Keep = $keep; Rel = $f.FullName.Substring($root.Length).TrimStart('\'); KeepRel = $keep.FullName.Substring($root.Length).TrimStart('\') })
            $bytes += $f.Length
        }
    }
    $todo = @($todo | Sort-Object Rel)
    $groupCount = @($groups.Keys | Where-Object { $groups[$_].Count -gt 1 }).Count
    Show-Errors $errors
    Say ""
    Say "  $Which has $($m.Map.Count) file(s)." 'White'
    if (-not $todo.Count) { Say "  DONE - no duplicates in $Which." 'Green'; return }

    Say ""
    Say "  DUPLICATES IN $Which ($groupCount file(s) exist more than once, $($todo.Count) extra cop$(if ($todo.Count -eq 1) { 'y' } else { 'ies' }), $(Format-Size $bytes))" 'Cyan'
    $title = 'TO DELETE   (the copy being kept is shown after =)'
    $lines = @($todo | ForEach-Object { "$($_.Rel)`n      = $($_.KeepRel)" })
    Show-List $title $lines 'Cyan'
    Say ""

    $bin = Test-RecycleBin $root
    if ($bin) {
        if (-not (Confirm-Action "Send $($todo.Count) file(s) to the Recycle Bin?" $title $lines)) { Say "  Skipped." 'Yellow'; return }
    } else {
        if ((Read-MenuKey 'View the full list first? y/n' 'ynYN') -in 'y', 'Y') { Show-Page $title $lines; Say "" }
        Say "  NO RECYCLE BIN ON THIS DRIVE - DELETING HERE IS PERMANENT AND CANNOT BE UNDONE." 'Red'
        if ((Read-Host "  Type yes to delete $($todo.Count) file(s) permanently, anything else to skip").Trim() -ne 'yes') { Say "  Skipped." 'Yellow'; return }
    }

    Add-Type -AssemblyName Microsoft.VisualBasic
    $done = 0; $n = 0
    $failed = New-Object System.Collections.Generic.List[string]
    $dirs = @{}
    foreach ($d in $todo) {
        $n++
        try {
            Say "    [$n/$($todo.Count)] $($d.Rel)" 'DarkGray'
            Remove-ToBin $d.File.FullName $bin
            $dirs[$d.File.DirectoryName] = $true
            $done++
        } catch {
            $failed.Add("$($d.Rel): $($_.Exception.Message)")
        }
    }
    # Folders left empty by the deletes go too, walking up but never past the folder itself
    $removed = 0
    foreach ($dir in @($dirs.Keys | Sort-Object { $_.Length } -Descending)) {
        $cur = $dir
        while ($cur.Length -gt $root.Length -and $cur.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
            try {
                if (-not [IO.Directory]::Exists($cur)) { break }
                if (@([IO.Directory]::EnumerateFileSystemEntries($cur) | Select-Object -First 1).Count) { break }
                Remove-ToBin $cur $bin $true
                $removed++
            } catch { break }
            $cur = [IO.Path]::GetDirectoryName($cur)
        }
    }
    $where = if ($bin) { 'sent to the Recycle Bin' } else { 'deleted' }
    Say "  DONE - $done of $($todo.Count) $where$(if ($removed) { ", $removed empty folder(s) removed" })" $(if ($failed.Count) { 'Yellow' } else { 'Green' })
    Show-Errors $failed
}

function Show-Duplicates {
    $a = if ($Settings.FolderA) { $Settings.FolderA } else { 'not set' }
    $b = if ($Settings.FolderB) { $Settings.FolderB } else { 'not set' }
    Say ""
    Say "  ---------------- DUPLICATES -----------------" 'Cyan'
    Say "    A  $a" 'DarkGray'
    Say "    B  $b" 'DarkGray'
    Say ""
    Say "    1  -  Find in A" 'White'
    Say "    2  -  Find in B" 'White'
    Say "    3  -  Back" 'White'
    Say "    0  -  Quit" 'White'
    Say ""
    switch (Read-MenuKey 'Choose 0-3' '0123') {
        '1' { Invoke-Duplicates 'A' }
        '2' { Invoke-Duplicates 'B' }
        '3' { return 'back' }
        '0' { return 'quit' }
    }
}

# =====================================================================
#  MENU
# =====================================================================
Clear-Host
$last = ''; $quit = $false
while ($true) {
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        FILE SYNC" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  Compare A and B" 'White'
    Say "    2  -  Add: A -> B" 'White'
    Say "    3  -  Add: B -> A" 'White'
    Say "    4  -  Add: A <> B" 'White'
    Say "    5  -  Replace: A -> B" 'White'
    Say "    6  -  Replace: B -> A" 'White'
    Say "    7  -  Sync" 'White'
    Say "    8  -  Duplicates" 'White'
    Say "    9  -  Settings" 'White'
    Say "    0  -  Quit" 'White'
    $choice = Read-MenuKey 'Choose 0-9' '0123456789' $last
    if ($choice -eq '0') { break }
    if ($choice -eq '9') {
        if (@(Show-Settings) -contains 'quit') { break }
        Clear-Host
        continue
    }
    try {
        switch ($choice) {
            '1' { Confirm-Folders; if (@(Invoke-Compare) -contains 'quit') { $quit = $true } }
            '2' { Confirm-Folders; Invoke-Sync $true  $false }
            '3' { Confirm-Folders; Invoke-Sync $false $true }
            '4' { Confirm-Folders; Invoke-Sync $true  $true }
            '5' { Confirm-Folders; Invoke-Update $true  $false }
            '6' { Confirm-Folders; Invoke-Update $false $true }
            '7' { Confirm-Folders; Invoke-FullSync }
            '8' { if (@(Show-Duplicates) -contains 'quit') { $quit = $true } }
        }
    } catch {
        Say "  FAILED: $($_.Exception.Message)" 'Red'
    }
    if ($quit) { break }
    $last = $choice
    Say ""
}
# Quit closes this window, not just the menu (errors above skip this and leave it open)
[Environment]::Exit(0)
