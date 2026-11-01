#Requires -Version 7
# ===== STAGE 1: COMPLETE INVENTORY + VERIFICATION (resumable, chunked) =====
# Run with PowerShell 7 (pwsh.exe) on a PC on the LAN, using an account with full read rights.
# Portable pwsh: https://github.com/PowerShell/PowerShell/releases (zip, no admin needed)
#
#   pwsh -File .\scanfiles.ps1                                # scan everything, resume if interrupted
#   pwsh -File .\scanfiles.ps1 -Only "AStar","BW Maritime"    # rescan just these top folders from scratch
#   pwsh -File .\scanfiles.ps1 -Force                         # rescan everything from scratch
#   pwsh -File .\scanfiles.ps1 -NoPrecount                    # skip the robocopy pre-count (no ETA; verify still runs at the end)
#   pwsh -File .\scanfiles.ps1 -Exclude "~snapshot","Temp"    # top-level folders to leave out (default: snapshot/system folders)
#   pwsh -File .\scanfiles.ps1 -IncludeHidden                 # also scan hidden top-level folders (skipped by default, robocopy tends to hang on them)
#
# Output goes to .\scan_output\ next to this script:
#   <TopFolder>.csv              one CSV per top-level folder (own header, opens in Excel)
#   <TopFolder>_01.csv, _02.csv  same folder split at -RowsPerPart rows (default 250,000)
#   <TopFolder>_NN.part          the part currently being written (do not open in Excel)
#   progress.csv                 per-folder status + resume checkpoint (human readable)
#   rc_counts.csv                robocopy file/byte counts (independent engine)
#   reconcile.csv                PowerShell count vs robocopy count per folder
#   errors.txt                   TopFolder <TAB> Path <TAB> Reason
#
# Ctrl+C at any time: the current part is trimmed to the last fully-scanned directory and progress.csv
# is saved. Re-run the same command to continue from exactly there. Nothing is redone.
#
# DONE = reconcile.csv shows Status=OK for every folder AND errors.txt is empty.

param(
    [string[]]$Only = @(),
    [switch]$Force,
    [switch]$NoPrecount,
    [switch]$IncludeHidden,          # hidden top-level folders (not shown in Explorer) are skipped unless this is set
    [int]$RowsPerPart = 250000,
    [double]$SaveEverySec = 30,      # how often progress.csv is written mid-folder (crash / power-loss safety)
    [string]$Root = '\\Sgsinvapfl20\20AP0026\S\SGBRS$',
    [string[]]$Exclude = @('~snapshot', '.snapshot', '$RECYCLE.BIN', 'System Volume Information', 'DfsrPrivate')   # top-level names never scanned
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false   # robocopy exit codes are informational
if ($PSStyle) { $PSStyle.Progress.View = 'Classic' }   # minimal view overwrites console lines and garbles the log
$Sep  = [IO.Path]::DirectorySeparatorChar
$Root = $Root.TrimEnd($Sep)

$OutDir        = Join-Path $PSScriptRoot 'scan_output'
$ProgressFile  = Join-Path $OutDir 'progress.csv'
$RcFile        = Join-Path $OutDir 'rc_counts.csv'
$ReconcileFile = Join-Path $OutDir 'reconcile.csv'
$ErrLog        = Join-Path $OutDir 'errors.txt'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$Header  = '"TopFolder","FullPath","FileName","Ext","SizeBytes","Modified","Created","Depth","PathLen","Hidden"'
$Utf8Bom = [Text.UTF8Encoding]::new($true)
$ATTR_DIR     = [int][IO.FileAttributes]::Directory
$ATTR_HIDDEN  = [int][IO.FileAttributes]::Hidden
$ATTR_REPARSE = [int][IO.FileAttributes]::ReparsePoint

# ---------- helpers ----------
function Safe([string]$s)   { $s -replace '[\\/:*?"<>|]', '_' }
function N([object]$n)      { ('{0:N0}' -f [double]$n) }
function GB([object]$b)     { $g = [double]$b / 1GB; if ($g -ge 1000) { '{0:N2} TB' -f ($g / 1024) } else { '{0:N1} GB' -f $g } }
function HMS([double]$sec)  { if ($sec -lt 0 -or [double]::IsInfinity($sec) -or [double]::IsNaN($sec)) { '--:--:--' } else { [TimeSpan]::FromSeconds([Math]::Round($sec)).ToString('h\:mm\:ss') } }
function Now()              { [DateTime]::Now.ToString('yyyy-MM-dd HH:mm') }
function Say([string]$s)    { [Console]::WriteLine($s) }   # works inside finally after Ctrl+C, unlike Write-Host

function Log-Err([string]$top, [string]$path, [string]$msg) {
    [IO.File]::AppendAllText($ErrLog, ("{0}`t{1}`t{2}`r`n" -f $top, $path, ($msg -replace '\s+', ' ')), $Utf8Bom)
    $script:ErrCount++
}

function Rename-Retry([string]$from, [string]$to) {
    for ($try = 1; ; $try++) {
        try { if (Test-Path -LiteralPath $to) { Remove-Item -LiteralPath $to -Force }; Move-Item -LiteralPath $from -Destination $to -Force; return }
        catch { if ($try -ge 10) { throw }; Say ("  cannot rename {0} (close it if open in Excel), retrying..." -f (Split-Path $from -Leaf)); Start-Sleep 3 }
    }
}

# ---------- state: progress.csv (resume checkpoint) + rc_counts.csv ----------
$StateCols = 'TopFolder','Status','Files','Bytes','LongPaths','Parts','PartRows','PartPos','LastDir','Seconds','UpdatedAt'
$State = [ordered]@{}
if (Test-Path $ProgressFile) { foreach ($r in Import-Csv $ProgressFile) { $State[$r.TopFolder] = $r } }
function Save-State {   # pure .NET on purpose: cmdlets are not usable inside finally after Ctrl+C
    $sb = [Text.StringBuilder]::new()
    [void]$sb.AppendLine(($StateCols | ForEach-Object { '"' + $_ + '"' }) -join ',')
    foreach ($r in $State.Values) { [void]$sb.AppendLine(($StateCols | ForEach-Object { '"' + ([string]$r.$_).Replace('"', '""') + '"' }) -join ',') }
    $tmp = "$ProgressFile.tmp"; [IO.File]::WriteAllText($tmp, $sb.ToString(), $Utf8Bom); [IO.File]::Move($tmp, $ProgressFile, $true)
}
function New-StateRow([string]$name) {
    [pscustomobject]@{ TopFolder=$name; Status='new'; Files=0; Bytes=0; LongPaths=0; Parts=1; PartRows=0; PartPos=0; LastDir=''; Seconds=0; UpdatedAt='' }
}

$Rc = @{}
if (Test-Path $RcFile) { foreach ($r in Import-Csv $RcFile) { $Rc[$r.TopFolder] = @{ Files=[int64]$r.Files; Bytes=[int64]$r.Bytes } } }
function Save-Rc {
    $Rc.GetEnumerator() | Sort-Object Key | ForEach-Object { [pscustomobject]@{ TopFolder=$_.Key; Files=$_.Value.Files; Bytes=$_.Value.Bytes } } |
        Export-Csv $RcFile -NoTypeInformation
}
$script:RcProc = $null
function Get-RcCount($unit, [string]$label) {
    # robocopy runs as a child process writing a log (directory lines kept, file lines suppressed) so we can show
    # elapsed time and directories seen while it works, and kill it cleanly on Ctrl+C.
    $log = Join-Path $OutDir ("rc_{0}.log" -f (Safe $unit.Name))
    Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue
    $psi = [Diagnostics.ProcessStartInfo]::new('robocopy')
    foreach ($x in @($unit.Path, 'C:\__null__', '/L', '/BYTES', '/NFL', '/NJH', '/XJ', '/R:0', '/W:0', "/LOG:$log")) { $psi.ArgumentList.Add($x) }
    if ($unit.Recurse) { $psi.ArgumentList.Add('/S'); $psi.ArgumentList.Add('/E') }
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $proc = [Diagnostics.Process]::Start($psi); $script:RcProc = $proc
    $sw = [Diagnostics.Stopwatch]::StartNew(); $pos = 0L; $dirs = 0L; $last = ''
    $buf = [byte[]]::new(65536)
    while (-not $proc.WaitForExit(1000)) {
        try {
            $fs = [IO.FileStream]::new($log, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            $fs.Position = $pos
            while (($n = $fs.Read($buf, 0, $buf.Length)) -gt 0) {
                for ($i = 0; $i -lt $n; $i++) { if ($buf[$i] -eq 10) { $dirs++ } }
                $pos += $n
            }
            if ($fs.Length -gt 0) { $fs.Position = [Math]::Max(0, $fs.Length - 2048); $n = $fs.Read($buf, 0, 2048)
                $tail = [Text.Encoding]::UTF8.GetString($buf, 0, $n) -split "`n" | Where-Object { $_ -like "*$($unit.Path)*" } | Select-Object -Last 1
                if ($tail) { $last = ($tail -replace '.*' + [regex]::Escape($unit.Path), '').Trim('\', ' ') } }
            $fs.Dispose()
        } catch {}
        Write-Progress -Id 1 -Activity ("Pre-count (robocopy /L)  {0}  {1}" -f $label, $unit.Name) -Status ("elapsed {0}   {1} directories seen   now: {2}" -f (HMS $sw.Elapsed.TotalSeconds), (N $dirs), $(if ($last) { $last } else { '(top)' }))
    }
    $script:RcProc = $null
    $txt = [IO.File]::ReadAllText($log)
    $f = [regex]::Match($txt, 'Files :\s+(\d+)'); $b = [regex]::Match($txt, 'Bytes :\s+(\d+)')
    if (-not $f.Success -or -not $b.Success) { throw "robocopy gave no summary for '$($unit.Name)'. See $log" }
    Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue
    @{ Files=[int64]$f.Groups[1].Value; Bytes=[int64]$b.Groups[1].Value; Seconds=[int]$sw.Elapsed.TotalSeconds }
}

# ---------- work units: _ROOT_FILES + one per top-level folder ----------
$rootDi = [IO.DirectoryInfo]::new($Root)
if (-not $rootDi.Exists) { throw "Root not reachable: $Root" }
$tops = [Collections.Generic.List[IO.DirectoryInfo]]::new()
$excluded = [Collections.Generic.List[string]]::new()
$hiddenTops = [Collections.Generic.List[string]]::new()
foreach ($d in $rootDi.EnumerateDirectories()) {
    if ([int]$d.Attributes -band $ATTR_REPARSE) { continue }
    if ($Exclude -contains $d.Name) { $excluded.Add($d.Name); continue }
    if (-not $IncludeHidden -and ([int]$d.Attributes -band $ATTR_HIDDEN)) { $hiddenTops.Add($d.Name); continue }
    $tops.Add($d)
}
$tops.Sort([Comparison[IO.DirectoryInfo]]{ param($x, $y) [string]::CompareOrdinal($x.Name, $y.Name) })

$units = [Collections.Generic.List[object]]::new()
$units.Add([pscustomobject]@{ Name='_ROOT_FILES'; Path=$Root; Recurse=$false })
foreach ($t in $tops) { $units.Add([pscustomobject]@{ Name=$t.Name; Path=$t.FullName; Recurse=$true }) }

if ($Only.Count) {
    $missing = $Only | Where-Object { $_ -notin $units.Name }
    if ($missing) { throw "Not found under root: $($missing -join ', ')$(if ($missing | Where-Object { $_ -in $hiddenTops }) { '  (hidden folder: add -IncludeHidden)' })" }
    $units = [Collections.Generic.List[object]]@($units | Where-Object { $_.Name -in $Only })
    $Force = $true            # -Only means: rescan these from scratch
}
foreach ($u in $units) { if (-not $State.Contains($u.Name)) { $State[$u.Name] = New-StateRow $u.Name } }

Say ("scanfiles v2    root : {0}" -f $Root)
Say ("                out  : {0}" -f $OutDir)
Say ("                parts: {0} rows per CSV" -f (N $RowsPerPart))
Say ("Top folders: {0} to scan{1}{2}" -f $tops.Count,
    $(if ($hiddenTops.Count) { "  hidden, skipped (use -IncludeHidden): " + ($hiddenTops -join ', ') } else { '' }),
    $(if ($excluded.Count) { "  excluded: " + ($excluded -join ', ') } else { '' }))

# ---------- robocopy pre-count (independent engine; also gives the ETA) ----------
function Stop-Rc {   # called from finally blocks: kill a running robocopy (whole tree) and say where we stopped
    if (-not $script:RcProc) { return }
    try { if (-not $script:RcProc.HasExited) { $script:RcProc.Kill($true) } } catch {}
    $script:RcProc = $null
    Say ''
    Say "Stopped during the robocopy count. Folders already counted are cached in rc_counts.csv; re-run the same command to continue."
}
try {
if (-not $NoPrecount) {
    $need = @($units | Where-Object { $Force -or -not $Rc.ContainsKey($_.Name) })
    if ($need.Count) {
        $sw0 = [Diagnostics.Stopwatch]::StartNew()
        $k = 0
        foreach ($u in $need) {
            $k++
            $Rc[$u.Name] = Get-RcCount $u ("{0}/{1}" -f $k, $need.Count)
            Save-Rc
            Say ("  count {0,3}/{1,-3} {2,-40} {3,12} files  {4,10}  {5}" -f $k, $need.Count, $u.Name, (N $Rc[$u.Name].Files), (GB $Rc[$u.Name].Bytes), (HMS $Rc[$u.Name].Seconds))
        }
        Write-Progress -Id 1 -Activity 'Pre-count' -Completed
        Say ("Pre-count (robocopy /L) {0} folders ..... {1} files  {2}  [{3}]" -f $need.Count, (N ($units | ForEach-Object { $Rc[$_.Name].Files } | Measure-Object -Sum).Sum), (GB ($units | ForEach-Object { $Rc[$_.Name].Bytes } | Measure-Object -Sum).Sum), (HMS $sw0.Elapsed.TotalSeconds))
    } else {
        Say ("Pre-count: cached in rc_counts.csv   {0} files  {1}" -f (N ($units | ForEach-Object { $Rc[$_.Name].Files } | Measure-Object -Sum).Sum), (GB ($units | ForEach-Object { $Rc[$_.Name].Bytes } | Measure-Object -Sum).Sum))
    }
}
} finally { Stop-Rc }
$GrandTotal = 0L; foreach ($u in $units) { if ($Rc.ContainsKey($u.Name)) { $GrandTotal += $Rc[$u.Name].Files } }

# ---------- scan one unit ----------
# Traversal is depth-first, subdirectories in ordinal name order, so it is deterministic between runs.
# After every fully-written directory a checkpoint (files, bytes, part, byte offset, dir) is taken.
# On resume, directories at or before the checkpoint in that order are skipped without being listed
# (only the checkpoint's ancestors are listed), and the .part file is trimmed to the checkpoint offset.

$script:ErrCount = 0
$script:Cur = $null   # current unit context, used by the finally block on Ctrl+C

function Part-Path([string]$safe, [int]$n) { Join-Path $OutDir ("{0}_{1:d2}.part" -f $safe, $n) }
function Part-Final([string]$safe, [int]$n, [bool]$single) { Join-Path $OutDir ($(if ($single) { "{0}.csv" -f $safe } else { "{0}_{1:d2}.csv" -f $safe, $n })) }

function Open-Part($cur, [bool]$fresh) {
    $p = Part-Path $cur.Safe $cur.Part
    if ($fresh) {
        $cur.Fs = [IO.FileStream]::new($p, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::Read)
        $cur.Sw = [IO.StreamWriter]::new($cur.Fs, $Utf8Bom, 65536)
        $cur.Sw.Write($Header); $cur.Sw.Write("`r`n"); $cur.Sw.Flush()
        $cur.PartRows = 0
    } else {
        $cur.Fs = [IO.FileStream]::new($p, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::Read)
        $cur.Fs.SetLength($cur.Ck.PartPos); $cur.Fs.Position = $cur.Ck.PartPos
        $cur.Sw = [IO.StreamWriter]::new($cur.Fs, $Utf8Bom, 65536)   # no BOM is written when Position > 0
    }
}
function Close-Part($cur, [bool]$single) {   # $single: folder fit in one part -> plain <name>.csv, else <name>_NN.csv
    $cur.Sw.Flush(); $cur.Fs.SetLength($cur.Ck.PartPos); $cur.Sw.Dispose(); $cur.Sw = $null; $cur.Fs = $null
    Rename-Retry (Part-Path $cur.Safe $cur.Part) (Part-Final $cur.Safe $cur.Part $single)
}
# The checkpoint is a consistent snapshot: counters + byte offset + directory, all taken together after a
# directory is fully written. Anything written after it (a half-scanned directory) is trimmed on stop/resume.
function Checkpoint($cur, [string]$relDir) {
    $cur.Sw.Flush()
    $cur.Ck = @{ Files=$cur.Files; Bytes=$cur.Bytes; LongPaths=$cur.LongPaths; Part=$cur.Part; PartRows=$cur.PartRows; PartPos=$cur.Fs.Position; LastDir=$relDir }
}
function Write-StateFrom($cur, [string]$status) {
    $r = $State[$cur.Name]; $k = $cur.Ck
    $r.Status = $status; $r.Files = $k.Files; $r.Bytes = $k.Bytes; $r.LongPaths = $k.LongPaths
    $r.Parts = $k.Part; $r.PartRows = $k.PartRows; $r.PartPos = $k.PartPos; $r.LastDir = $k.LastDir
    $r.Seconds = [int]($cur.Seconds0 + $cur.Watch.Elapsed.TotalSeconds); $r.UpdatedAt = Now
    Save-State
}

function Scan-Unit($unit, [int]$idx, [int]$count, [int64]$doneBefore) {
    $name = $unit.Name; $safe = Safe $name
    $rec = $State[$name]
    $resume = $false
    if ($rec.Status -eq 'partial' -and -not $Force) {
        $n = [int]$rec.Parts; $pp = Part-Path $safe $n
        if (-not (Test-Path -LiteralPath $pp) -and (Test-Path -LiteralPath (Part-Final $safe $n $false))) {
            # crash after a part rollover but before the next checkpoint save: undo the rollover
            $rx = '^' + [regex]::Escape($safe) + '_(\d{2})\.(csv|part)$'
            Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Name -match $rx -and [int]$Matches[1] -gt $n } | Remove-Item -Force
            Move-Item -LiteralPath (Part-Final $safe $n $false) -Destination $pp -Force
            Say ("  {0} : reopened part {1} (crash happened after a rollover, before the next save)" -f $name, $n)
        }
        $resume = Test-Path -LiteralPath $pp
        if (-not $resume) { Say "  $name : checkpoint found but its .part file is missing, rescanning from scratch" }
    }

    $cur = [pscustomobject]@{
        Name=$name; Safe=$safe; Files=0L; Bytes=0L; LongPaths=0L; Part=1; PartRows=0
        Ck=@{ Files=0L; Bytes=0L; LongPaths=0L; Part=1; PartRows=0; PartPos=0L; LastDir='' }
        Seconds0=0.0; Watch=[Diagnostics.Stopwatch]::StartNew(); Fs=$null; Sw=$null
    }
    if ($resume) {
        $cur.Files=[int64]$rec.Files; $cur.Bytes=[int64]$rec.Bytes; $cur.LongPaths=[int64]$rec.LongPaths
        $cur.Part=[int]$rec.Parts; $cur.PartRows=[int]$rec.PartRows; $cur.Seconds0=[double]$rec.Seconds
        $cur.Ck = @{ Files=$cur.Files; Bytes=$cur.Bytes; LongPaths=$cur.LongPaths; Part=$cur.Part; PartRows=$cur.PartRows; PartPos=[int64]$rec.PartPos; LastDir=[string]$rec.LastDir }
        Say ("[{0,2}/{1}] {2,-24} resuming after {3}   ({4} files so far, part {5})" -f $idx, $count, $name, $(if ($cur.Ck.LastDir -eq '.') { '(top-level files)' } else { $cur.Ck.LastDir }), (N $cur.Files), $cur.Part)
    } else {
        # wipe anything from an earlier attempt at this folder
        $rx = '^' + [regex]::Escape($safe) + '(_\d{2})?\.(csv|part)$'
        Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Name -match $rx } | Remove-Item -Force
        if (Test-Path $ErrLog) {
            $keep = [IO.File]::ReadAllLines($ErrLog) | Where-Object { -not $_.StartsWith("$name`t") }
            [IO.File]::WriteAllLines($ErrLog, [string[]]$keep, $Utf8Bom)
        }
    }
    $script:Cur = $cur
    Open-Part $cur (-not $resume)
    if (-not $resume) { Checkpoint $cur '' }   # empty LastDir = nothing written yet
    Write-StateFrom $cur 'partial'

    $sw = $cur.Sw
    $qTop = '"' + $name + '"'
    $baseDepth = $unit.Path.Split($Sep).Count - $Root.Split($Sep).Count   # 0 for root, 1 for a top folder
    $rcTotal = if ($Rc.ContainsKey($name)) { $Rc[$name].Files } else { -1 }
    $sessionStart = $cur.Files
    $lastUi = 0.0; $lastSave = 0.0

    # resume comparison
    $resuming = $resume -and $cur.Ck.LastDir -ne ''          # '' = header only, nothing to skip
    $ckComps = if ($resuming -and $cur.Ck.LastDir -ne '.') { $cur.Ck.LastDir.Split('\') } else { [string[]]@() }

    $stack = [Collections.Generic.Stack[object]]::new()
    $stack.Push(@([IO.DirectoryInfo]::new($unit.Path), [string[]]@()))

    while ($stack.Count) {
        $di, $comps = $stack.Pop()
        $rel = if ($comps.Count) { $comps -join '\' } else { '.' }

        # 0 = skip entirely, 1 = list but don't write (ancestor of checkpoint), 2 = write
        $mode = 2
        if ($resuming) {
            $mode = 1
            $n = [Math]::Min($comps.Count, $ckComps.Count)
            for ($i = 0; $i -lt $n; $i++) {
                $c = [string]::CompareOrdinal($comps[$i], $ckComps[$i])
                if ($c -lt 0) { $mode = 0; break }
                if ($c -gt 0) { $mode = 2; break }
            }
            if ($mode -eq 1 -and $comps.Count -gt $ckComps.Count) { $mode = 2 }   # descendant of checkpoint dir
            if ($mode -eq 2) { $resuming = $false }
        }
        if ($mode -eq 0) { continue }

        try { $entries = $di.GetFileSystemInfos() }
        catch { $ex = $_.Exception; if ($ex.InnerException) { $ex = $ex.InnerException }; Log-Err $name $di.FullName $ex.Message; continue }

        $subs = [Collections.Generic.List[IO.DirectoryInfo]]::new()
        $depth = $baseDepth + $comps.Count + 1
        foreach ($e in $entries) {
            $attr = [int]$e.Attributes
            if ($attr -band $ATTR_REPARSE) { continue }                     # junctions/symlinks, same as robocopy /XJ
            if ($attr -band $ATTR_DIR) { if ($unit.Recurse) { $subs.Add($e) }; continue }
            if ($mode -ne 2) { continue }
            $full = $e.FullName
            $sw.Write($qTop); $sw.Write(',"'); $sw.Write($full); $sw.Write('","'); $sw.Write($e.Name); $sw.Write('","')
            $sw.Write($e.Extension.ToLowerInvariant()); $sw.Write('",'); $sw.Write($e.Length); $sw.Write(',')
            $sw.Write($e.LastWriteTime.ToString('yyyy-MM-dd HH:mm')); $sw.Write(','); $sw.Write($e.CreationTime.ToString('yyyy-MM-dd'))
            $sw.Write(','); $sw.Write($depth); $sw.Write(','); $sw.Write($full.Length); $sw.Write(',')
            $sw.Write($(if ($attr -band $ATTR_HIDDEN) { '1' } else { '0' })); $sw.Write("`r`n")
            $cur.Files++; $cur.Bytes += $e.Length; $cur.PartRows++
            if ($full.Length -gt 250) { $cur.LongPaths++ }
        }
        $entries = $null

        if ($mode -eq 2) {
            Checkpoint $cur $rel
            if ($cur.PartRows -ge $RowsPerPart) {
                Close-Part $cur $false
                $cur.Part++; Open-Part $cur $true; $sw = $cur.Sw
                Checkpoint $cur $rel
            }
        }

        if ($subs.Count) {
            $subs.Sort([Comparison[IO.DirectoryInfo]]{ param($x, $y) [string]::CompareOrdinal($x.Name, $y.Name) })
            for ($i = $subs.Count - 1; $i -ge 0; $i--) { $stack.Push(@($subs[$i], [string[]]($comps + $subs[$i].Name))) }
        }

        # progress UI (max 2x/sec) + periodic checkpoint save (every 30 s, survives a crash or power loss)
        $t = $cur.Watch.Elapsed.TotalSeconds
        if ($t - $lastUi -ge 0.5) {
            $lastUi = $t
            $rate = ($cur.Files - $sessionStart) / [Math]::Max($t, 0.001)
            $pct = if ($rcTotal -gt 0) { [Math]::Min(100, [int](100 * $cur.Files / $rcTotal)) } else { -1 }
            $eta = if ($rcTotal -gt 0 -and $rate -gt 0) { ($rcTotal - $cur.Files) / $rate } else { -1 }
            $st = if ($rcTotal -gt 0) { "{0}% {1} / {2} files" -f $pct, (N $cur.Files), (N $rcTotal) } else { "{0} files" -f (N $cur.Files) }
            $st += "   {0} f/s   {1}   part {2}   now: {3}" -f (N $rate), (GB $cur.Bytes), $cur.Part, $rel
            $pp = @{ Id=2; ParentId=1; Activity=("[{0}/{1}] {2}" -f $idx, $count, $name); Status=$st; PercentComplete=$pct }
            if ($eta -ge 0) { $pp.SecondsRemaining = [int][Math]::Min($eta, 1e9) }
            Write-Progress @pp
            $doneAll = $doneBefore + $cur.Files
            $ost = "{0} / {1} files   elapsed {2}   errors {3}   Ctrl+C to stop, re-run to resume" -f (N $doneAll), $(if ($GrandTotal) { N $GrandTotal } else { '?' }), (HMS $script:RunWatch.Elapsed.TotalSeconds), $script:ErrCount
            $op = @{ Id=1; Activity='overall'; Status=$ost; PercentComplete=$(if ($GrandTotal) { [Math]::Min(100, [int](100 * $doneAll / $GrandTotal)) } else { -1 }) }
            if ($GrandTotal -and $rate -gt 0) { $op.SecondsRemaining = [int][Math]::Min(($GrandTotal - $doneAll) / $rate, 1e9) }
            Write-Progress @op
        }
        if ($t - $lastSave -ge $SaveEverySec) { $lastSave = $t; Write-StateFrom $cur 'partial' }
    }

    Close-Part $cur ($cur.Part -eq 1)
    $script:Cur = $null
    Write-StateFrom $cur 'done'
    Write-Progress -Id 2 -Activity ' ' -Completed
    Say ("[{0,2}/{1}] {2,-24} done  {3,12} files  {4,10}  {5} part{6}  {7}" -f $idx, $count, $name, (N $cur.Files), (GB $cur.Bytes), $cur.Part, $(if ($cur.Part -eq 1) { ' ' } else { 's' }), (HMS ($cur.Seconds0 + $cur.Watch.Elapsed.TotalSeconds)))
    return $cur.Files
}

# ---------- main ----------
$script:RunWatch = [Diagnostics.Stopwatch]::StartNew()
$completed = $false
try {
    $todo = @($units | Where-Object { $Force -or $State[$_.Name].Status -ne 'done' })
    $skipped = $units.Count - $todo.Count
    if ($skipped) { Say ("Resuming from progress.csv: {0} of {1} folders already done, skipping." -f $skipped, $units.Count) }
    if ($Force -and -not $Only.Count) { foreach ($u in $units) { $State[$u.Name] = New-StateRow $u.Name } }

    $doneFiles = 0L; foreach ($u in $units) { if ($State[$u.Name].Status -eq 'done') { $doneFiles += [int64]$State[$u.Name].Files } }
    $i = 0
    foreach ($u in $todo) {
        $i++
        $doneFiles += Scan-Unit $u $i $todo.Count $doneFiles
    }
    Write-Progress -Id 1 -Activity 'overall' -Completed
    $completed = $true
}
finally {
    Stop-Rc
    if ($script:Cur) {
        $c = $script:Cur
        try { if ($c.Sw) { $c.Sw.Flush(); $c.Fs.SetLength($c.Ck.PartPos); $c.Sw.Dispose() } } catch {}
        try { Write-StateFrom $c 'partial' } catch {}
        Say ''
        Say ("Stopped in {0} after directory {1}" -f $c.Name, $(switch ($c.Ck.LastDir) { '.' { '(top-level files)' } '' { '(nothing yet)' } default { $_ } }))
        Say ("Saved: {0} kept at {1} files. Re-run the same command to resume." -f [IO.Path]::GetFileName((Part-Path $c.Safe $c.Ck.Part)), (N $c.Ck.Files))
    }
}
if (-not $completed) { return }

# ---------- verification: PowerShell count vs robocopy count ----------
if ($NoPrecount) {
  try {
    $k = 0
    foreach ($u in $units) {
        $k++; Write-Progress -Id 1 -Activity 'Verify (robocopy /L)' -Status ("{0}/{1}  {2}" -f $k, $units.Count, $u.Name) -PercentComplete (100 * ($k - 1) / $units.Count)
        if ($Force -or -not $Rc.ContainsKey($u.Name)) { $Rc[$u.Name] = Get-RcCount $u ("{0}/{1}" -f $k, $units.Count); Save-Rc }
    }
    Write-Progress -Id 1 -Activity 'Verify' -Completed
  } finally { Stop-Rc }
}
$rows = foreach ($r in $State.Values) {
    if ($r.Status -ne 'done' -or -not $Rc.ContainsKey($r.TopFolder)) { continue }
    $cnt = $Rc[$r.TopFolder]
    [pscustomobject]@{
        TopFolder=$r.TopFolder; PS_Files=[int64]$r.Files; RC_Files=$cnt.Files; PS_Bytes=[int64]$r.Bytes; RC_Bytes=$cnt.Bytes
        Status=$(if ([int64]$r.Files -eq $cnt.Files -and [int64]$r.Bytes -eq $cnt.Bytes) { 'OK' } else { 'MISMATCH' })
    }
}
$rows | Export-Csv $ReconcileFile -NoTypeInformation

$bad   = @($rows | Where-Object Status -ne 'OK')
$done  = @($State.Values | Where-Object Status -eq 'done')
$csvs  = @(Get-ChildItem -LiteralPath $OutDir -File -Filter *.csv | Where-Object { $_.Name -notin 'progress.csv','rc_counts.csv','reconcile.csv' })
$split = @($done | Where-Object { [int]$_.Parts -gt 1 } | ForEach-Object { "{0} into {1}" -f $_.TopFolder, $_.Parts })
$errLines = if (Test-Path $ErrLog) { @([IO.File]::ReadAllLines($ErrLog) | Where-Object { $_ }).Count } else { 0 }

Say ''
Say ("Done in {0}" -f (HMS $script:RunWatch.Elapsed.TotalSeconds))
Say ("Folders : {0}   OK {1}   MISMATCH {2}   -> reconcile.csv" -f $rows.Count, ($rows.Count - $bad.Count), $bad.Count)
Say ("Files   : {0}   {1}   long paths (>250): {2}" -f (N ($done | Measure-Object -Property Files -Sum).Sum), (GB ($done | Measure-Object -Property Bytes -Sum).Sum), (N ($done | Measure-Object -Property LongPaths -Sum).Sum))
Say ("CSV     : {0} files{1}" -f $csvs.Count, $(if ($split.Count) { "  (split: " + ($split -join '; ') + ")" } else { '' }))
if ($errLines) { Say ("Errors  : {0} -> errors.txt   fix rights, then re-run with -Only ""<folder>""" -f $errLines) } else { Say "Errors  : none" }
if ($bad.Count) { $bad | Format-Table -AutoSize | Out-String | ForEach-Object { Say $_ } }
