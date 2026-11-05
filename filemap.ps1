# irm https://github.com/therepos/windows/raw/main/filemap.ps1 | iex
# Purpose: List every file in a folder and its subfolders, as a text list or a clickable Excel index
# =============================================================================
#  File Map:
#   1 = List    every file (relative path, size, modified), sorted, one per line
#               -> filemap <PC> yyyyMMdd-HHmm.txt
#               for comparing two machines (e.g. paste both lists into an AI)
#   2 = Index   every file with clickable links to the file and its folder
#               -> filemap <PC> yyyyMMdd-HHmm.csv (opens in Excel)
#               mapped drives are written as \\server\share paths so links
#               work for anyone with access; over ~1 million files it splits
#               into ...-a.csv, ...-b.csv, ...
#   0 = Quit
#  Output: Downloads. Hidden and system files are included; folders that
#  cannot be opened are listed at the end. Read-only: nothing is changed,
#  no admin, no downloads.
#
#  If GitHub is blocked on the PC:
#   - Copy this whole file into Notepad, Save as type "All files", name it
#     filemap.ps1. In PowerShell, in that folder, run:
#       powershell -ExecutionPolicy Bypass -File .\filemap.ps1
#     (no admin; the bypass lasts for this one run only)
#   - If that is blocked too: open PowerShell ISE, paste into a new tab, press F5.
#   - Do not paste it into a normal PowerShell window: the lines after the
#     menu get eaten as answers to its prompts.
#
#  No PowerShell at all? Plain list from Command Prompt (cmd), into Downloads:
#     chcp 65001
#     dir "C:\Some Folder" /s /b /a-d > "%USERPROFILE%\Downloads\filelist.txt"
#   /s = include subfolders, /b = full paths only, /a-d = all files incl. hidden.
#   Drop /b to also get sizes and dates (grouped per folder, harder to compare).
#   chcp 65001 keeps non-English file names readable.
# =============================================================================

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

$Stamp   = Get-Date -Format 'yyyyMMdd-HHmm'
$MaxRows = 1048000                      # Excel limit is 1,048,576 rows per sheet

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }

function Format-Size([double]$bytes) {
    if ($bytes -ge 1GB) { return '{0:N1} GB' -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return '{0:N0} MB' -f ($bytes / 1MB) }
    return '{0:N0} KB' -f [Math]::Ceiling($bytes / 1KB)
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

function Read-TargetFolder {
    $default = Get-DownloadsFolder
    Say ""
    Say "  Which folder? Type or drag it here (a drive like Z:\ works too)." 'Cyan'
    while ($true) {
        $answer = Read-Host "  Folder [$default]"
        $path = if ([string]::IsNullOrWhiteSpace($answer)) { $default } else { [Environment]::ExpandEnvironmentVariables($answer.Trim().Trim('"').Trim("'")) }
        if ($path -match '^[A-Za-z]:$') { $path += '\' }
        if (Test-Path -LiteralPath $path -PathType Container) {
            $p = (Resolve-Path -LiteralPath $path).ProviderPath
            if ($p.Length -gt 3) { $p = $p.TrimEnd('\') }
            return $p
        }
        Say "  Cannot open '$path'. Try again." 'Red'
    }
}

# Z:\Old\Files -> \\server\share\Old\Files when Z: is a mapped network drive
function Get-DisplayRoot([string]$path) {
    if ($path -match '^([A-Za-z]):') {
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($Matches[1].ToUpper()):'" -ErrorAction SilentlyContinue
        if ($disk -and $disk.DriveType -eq 4 -and $disk.ProviderName) {
            return ($disk.ProviderName.TrimEnd('\') + $path.Substring(2)).TrimEnd('\')
        }
    }
    return $path.TrimEnd('\')
}

# Use the \\?\ form when this PC allows it, so paths over 260 characters are read quickly
# (where it is not allowed, Get-AllFiles falls back to Get-ChildItem for those folders)
function Get-WalkRoot([string]$path) {
    $long = if ($path.StartsWith('\\')) { '\\?\UNC\' + $path.Substring(2) } else { '\\?\' + $path }
    try {
        $e = (New-Object IO.DirectoryInfo $long).EnumerateFileSystemInfos().GetEnumerator()
        [void]$e.MoveNext()
        return $long
    } catch { return $path }
}

# Walk the tree folder by folder; a folder that cannot be opened is noted, not fatal.
# Outputs files (FileInfo) and empty folders (DirectoryInfo).
function Get-AllFiles([string]$Root) {
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    while ($stack.Count) {
        $dir = $stack.Pop()
        $count = 0
        try {
            # .NET refuses paths over 260 characters on most PCs; Get-ChildItem reads them fine
            try { $items = (New-Object IO.DirectoryInfo $dir).GetFileSystemInfos() }
            catch [IO.PathTooLongException] { $items = @(Get-ChildItem -LiteralPath $dir -Force -ErrorAction Stop) }
            foreach ($e in $items) {
                $count++
                $a = [int]$e.Attributes
                if ($a -band 0x10) {
                    # Junctions/links are not followed (avoids loops); OneDrive/cloud folders are
                    if (($a -band 0x400) -and -not ($a -band 0x1C0000)) { $script:Links.Add($e.FullName); continue }
                    $stack.Push($e.FullName)
                } else {
                    $e
                }
            }
            if ($count -eq 0 -and $dir -ne $Root) { New-Object IO.DirectoryInfo $dir }
        } catch {
            $script:Errors.Add([pscustomobject]@{ Path = $dir; Reason = $_.Exception.GetBaseException().Message })
        }
    }
}

function Get-Relative([string]$full) {
    if ($full.Length -le $WalkRoot.Length) { return '' }
    $full.Substring($WalkRoot.Length).TrimStart('\')
}

# Plain string split (Split-Path treats [ ] as wildcards; .NET Path fails on long paths)
function Get-Parent([string]$p) { $i = $p.LastIndexOf('\'); if ($i -lt 0) { '' } elseif ($i -eq 2 -and $p[1] -eq ':') { $p.Substring(0, 3) } else { $p.Substring(0, $i) } }

function Join-Display([string]$rel) { if ($rel) { "$DisplayRoot\$rel" } else { $DisplayRoot } }

function Show-Progress([int]$n) {
    if ($Clock.ElapsedMilliseconds -ge 1500) {
        Write-Host ("`r  {0:N0} files..." -f $n) -NoNewline -ForegroundColor DarkGray
        $Clock.Restart()
    }
}

function New-Writer([string]$path) {
    New-Object IO.StreamWriter($path, $false, (New-Object Text.UTF8Encoding($true)))
}

# =====================================================================
#  OPTION 1 - TEXT LIST
# =====================================================================
function New-FileList {
    $lines = New-Object System.Collections.Generic.List[string]
    $bytes = [long]0
    $empty = 0
    Get-AllFiles $WalkRoot | ForEach-Object {
        $rel = Get-Relative $_.FullName
        if ($_ -is [IO.DirectoryInfo]) {
            $lines.Add("$rel\`t[empty folder]`t")
            $empty++
        } else {
            $lines.Add("$rel`t$($_.Length)`t$($_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))")
            $bytes += $_.Length
        }
        Show-Progress $lines.Count
    }
    $lines.Sort([StringComparer]::OrdinalIgnoreCase)
    $files = $lines.Count - $empty

    $out = Join-Path $OutDir "filemap $env:COMPUTERNAME $Stamp.txt"
    $w = New-Writer $out
    try {
        $w.WriteLine("# File list of: $DisplayRoot")
        $w.WriteLine("# Computer:     $env:COMPUTERNAME")
        $w.WriteLine("# Created:      $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
        $w.WriteLine(("# Files:        {0:N0}  ({1}, {2:N0} bytes)" -f $files, (Format-Size $bytes), $bytes))
        $w.WriteLine(("# Empty folders: {0:N0}   Folders not readable: {1:N0}   Links not followed: {2:N0}" -f $empty, $Errors.Count, $Links.Count))
        $w.WriteLine("# Columns (tab-separated): relative path, size in bytes, last modified")
        $w.WriteLine("# To compare machines, match on relative path; same size means same content is likely.")
        $w.WriteLine("")
        foreach ($l in $lines) { $w.WriteLine($l) }
        if ($Errors.Count) {
            $w.WriteLine("")
            $w.WriteLine("# FOLDERS NOT READABLE (contents missing from this list)")
            foreach ($e in $Errors) { $w.WriteLine("$(Get-Relative $e.Path)\`t$($e.Reason)") }
        }
        if ($Links.Count) {
            $w.WriteLine("")
            $w.WriteLine("# LINKED FOLDERS NOT FOLLOWED")
            foreach ($l in $Links) { $w.WriteLine("$(Get-Relative $l)\") }
        }
    } finally { $w.Close() }

    Write-Host ""
    Finish -Paths @($out) -Files $files -Bytes $bytes
}

# =====================================================================
#  OPTION 2 - CSV INDEX WITH LINKS
# =====================================================================
# Excel opens a CSV with the PC's list separator (comma, or semicolon in many countries)
$Sep = (Get-Culture).TextInfo.ListSeparator
if (-not $Sep) { $Sep = ',' }

function ConvertTo-Cell([string]$s) {
    if ($s.IndexOfAny([char[]]"$Sep`"`r`n") -ge 0) { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

# Stop Excel turning names like 2024-01, 007 or =x into dates, numbers or formulas
function ConvertTo-TextCell([string]$s) {
    if ($s -match '^[=+\-@]' -or $s -match '^[\d\s.,:/%()+\-]+$' -or $s -match '^\d+[eE][+\-]?\d+$') {
        if ($s.Length -le 250) { $s = '="' + $s + '"' } else { $s = "'" + $s }
    }
    ConvertTo-Cell $s
}

# Excel links stop at 255 characters and break on '#'; those rows keep the full path only
function ConvertTo-LinkCell([string]$target, [string]$label) {
    if (-not $target -or $target.Length -gt 255 -or $target.Contains('#')) { return '' }
    ConvertTo-Cell "=HYPERLINK(`"$target`"$Sep`"$label`")"
}

function New-FileIndex {
    $base   = Join-Path $OutDir "filemap $env:COMPUTERNAME $Stamp"
    $header = (('Name', 'Type', 'Folder', 'Size KB', 'Modified', 'Open file', 'Open folder', 'Full path') -join $Sep)
    $paths  = New-Object System.Collections.Generic.List[string]
    $state  = @{ W = $null; Rows = 0; Part = 0 }
    $newPart = {
        if ($state.W) { $state.W.Close() }
        $state.Part++
        # One file: <base>.csv. Once a second is needed, the first becomes -a, then -b, -c, ...
        if ($state.Part -eq 2) {
            Rename-Item -LiteralPath "$base.csv" -NewName (Split-Path "$base-a.csv" -Leaf)
            $paths[0] = "$base-a.csv"
        }
        $p = if ($state.Part -eq 1) { "$base.csv" } else { "$base-$([char](96 + $state.Part)).csv" }
        $paths.Add($p)
        $state.W = New-Writer $p
        $state.W.WriteLine($header)
        $state.Rows = 0
    }
    $writeRow = {
        param([string[]]$cells)
        if ($state.Rows -ge $MaxRows) { & $newPart }
        $state.W.WriteLine($cells -join $Sep)
        $state.Rows++
    }

    $files = 0
    $bytes = [long]0
    & $newPart
    try {
        Get-AllFiles $WalkRoot | ForEach-Object {
            $rel    = Get-Relative $_.FullName
            $full   = Join-Display $rel
            $folder = Get-Parent $full
            if ($_ -is [IO.DirectoryInfo]) {
                & $writeRow @('(empty folder)', '', (ConvertTo-TextCell $rel), '', '', '',
                    (ConvertTo-LinkCell $full 'Open folder'), (ConvertTo-Cell $full))
                return
            }
            if ($full.StartsWith($OutBase, [StringComparison]::OrdinalIgnoreCase)) { return }
            $files++
            $bytes += $_.Length
            $relFolder = Get-Parent $rel
            if (-not $relFolder) { $relFolder = '.' }
            & $writeRow @(
                (ConvertTo-TextCell $_.Name),
                (ConvertTo-TextCell $_.Extension.TrimStart('.').ToLower()),
                (ConvertTo-TextCell $relFolder),
                [string][long][Math]::Ceiling($_.Length / 1KB),
                $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'),
                (ConvertTo-LinkCell $full 'Open'),
                (ConvertTo-LinkCell $folder 'Open folder'),
                (ConvertTo-Cell $full))
            Show-Progress $files
        }
        foreach ($e in $Errors) {
            $full = Join-Display (Get-Relative $e.Path)
            & $writeRow @('(cannot open folder)', '', (ConvertTo-TextCell (Get-Relative $e.Path)), '', '', '',
                (ConvertTo-LinkCell $full 'Open folder'), (ConvertTo-Cell "$full  -  $($e.Reason)"))
        }
        foreach ($l in $Links) {
            $full = Join-Display (Get-Relative $l)
            & $writeRow @('(linked folder, not followed)', '', (ConvertTo-TextCell (Get-Relative $l)), '', '', '',
                (ConvertTo-LinkCell $full 'Open folder'), (ConvertTo-Cell $full))
        }
    } finally { if ($state.W) { $state.W.Close() } }

    Write-Host ""
    Finish -Paths $paths.ToArray() -Files $files -Bytes $bytes
}

function Finish([string[]]$Paths, [int]$Files, [long]$Bytes) {
    Say ""
    Say ("  DONE - {0:N0} files, {1}" -f $Files, (Format-Size $Bytes)) 'Green'
    foreach ($p in $Paths) { Say "  Saved to: $p" 'Green' }
    if ($Errors.Count) { Say ("  {0:N0} folder(s) could not be opened - listed at the end of the file." -f $Errors.Count) 'Yellow' }
    if ($Links.Count)  { Say ("  {0:N0} linked folder(s) not followed - listed at the end of the file." -f $Links.Count) 'Yellow' }
    Say ""
}

# =====================================================================
#  MENU
# =====================================================================
Clear-Host
Say ""
Say "  =============================================" 'Cyan'
Say "        FILE MAP" 'Cyan'
Say "  =============================================" 'Cyan'
Say ""
Say "    1  -  List (text)" 'White'
Say "    2  -  Index with links (Excel)" 'White'
Say "    0  -  Quit" 'White'
Say ""
do {
    $choice = "$(Read-Host '  Choose 0-2')".Trim()
    if ($choice -notin '0', '1', '2') { Say "  Please type 1, 2 or 0." 'Red' }
} until ($choice -in '0', '1', '2')
if ($choice -eq '0') { return }

$Target      = Read-TargetFolder
$DisplayRoot = Get-DisplayRoot $Target
$WalkRoot    = Get-WalkRoot $Target
$OutDir      = Get-DownloadsFolder
$OutBase     = Join-Path $OutDir "filemap $env:COMPUTERNAME $Stamp"
$Errors      = New-Object System.Collections.Generic.List[object]
$Links       = New-Object System.Collections.Generic.List[string]
$Clock       = [Diagnostics.Stopwatch]::StartNew()

Say ""
Say "  Reading $DisplayRoot ..." 'Cyan'
Say "  Large or network folders can take a while." 'DarkGray'
switch ($choice) {
    '1' { New-FileList }
    '2' { New-FileIndex }
}
