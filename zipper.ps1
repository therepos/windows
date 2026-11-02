# irm https://github.com/therepos/windows/raw/main/zipper.ps1 | iex
# Purpose: Zip files into a password-protected 7z (auto-split over 500 MB) or unzip one, split sets included
# =============================================================================
#  Zipper:
#   1 = Zip     files/folders -> YYYYMMDD-HHMMSS (no extension)
#               over 500 MB   -> YYYYMMDD-HHMMSS.001, .002, ...
#   2 = Unzip   archive or any part of a split set -> <name>-extracted
#   0 = Quit
#  Password: Enter = default. Output: asks for a folder, defaults to Downloads.
#  Uses an installed 7-Zip, otherwise downloads the standalone 7za.exe
#  into %LOCALAPPDATA%\zipper (no admin needed).
# =============================================================================

$DefaultPassword = 'Zip@2026!'
$SplitMB         = 500
$ToolDir         = Join-Path $env:LOCALAPPDATA 'zipper'
$ExtraUrl        = 'https://www.7-zip.org/a/7z2409-extra.7z'
$SevenZrUrl      = 'https://www.7-zip.org/a/7zr.exe'

$ProgressPreference = 'SilentlyContinue'

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }

function Format-Size([double]$bytes) {
    if ($bytes -ge 1GB) { return '{0:N1} GB' -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return '{0:N0} MB' -f ($bytes / 1MB) }
    return '{0:N0} KB' -f [Math]::Ceiling($bytes / 1KB)
}

# =====================================================================
#  7-ZIP: installed copy, else standalone 7za.exe in %LOCALAPPDATA%
# =====================================================================
function Find-SevenZip {
    foreach ($p in "$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe") {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    foreach ($name in '7z', '7za', '7zz') {
        $cmd = Get-Command $name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cmd) { return $cmd.Path }
    }
    foreach ($p in (Join-Path $ToolDir 'x64\7za.exe'), (Join-Path $ToolDir '7za.exe')) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Get-SevenZip {
    $found = Find-SevenZip
    if ($found) { return $found }

    Say ""
    Say "  7-Zip not found. Downloading the standalone version (about 2 MB, no admin)..." 'Yellow'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null
        $sevenZr = Join-Path $ToolDir '7zr.exe'
        $extra   = Join-Path $ToolDir 'extra.7z'
        Invoke-WebRequest -Uri $SevenZrUrl -OutFile $sevenZr -UseBasicParsing
        Invoke-WebRequest -Uri $ExtraUrl   -OutFile $extra   -UseBasicParsing
        & $sevenZr x $extra "-o$ToolDir" -y | Out-Null
        Remove-Item -LiteralPath $extra, $sevenZr -Force -ErrorAction SilentlyContinue
    } catch {
        Say "  Download failed: $($_.Exception.Message)" 'Red'
    }
    $exe = if ([Environment]::Is64BitOperatingSystem) { Join-Path $ToolDir 'x64\7za.exe' } else { Join-Path $ToolDir '7za.exe' }
    if (-not (Test-Path -LiteralPath $exe)) { $exe = Join-Path $ToolDir '7za.exe' }
    if (Test-Path -LiteralPath $exe) { Say "  7-Zip ready." 'Green'; return $exe }
    Say "  Could not get 7-Zip. Install it from https://www.7-zip.org and run this again." 'Red'
    return $null
}

# Quote one argument for a Windows command line (CommandLineToArgvW rules)
function ConvertTo-Arg([string]$s) {
    if ($s -and $s -notmatch '[\s"]') { return $s }
    $s = $s -replace '(\\*)"', '$1$1\"'
    $s = $s -replace '(\\+)$', '$1$1'
    return '"' + $s + '"'
}

# Run 7-Zip with its progress shown live; return exit code and the most useful error line
function Invoke-SevenZip([string[]]$Arguments) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = $SevenZip
    $all = @($Arguments[0], '-bso0', '-bsp1') + @($Arguments | Select-Object -Skip 1)
    $psi.Arguments = ($all | ForEach-Object { ConvertTo-Arg $_ }) -join ' '
    $psi.UseShellExecute       = $false
    $psi.RedirectStandardError = $true
    $p = [Diagnostics.Process]::Start($psi)
    $errTask = $p.StandardError.ReadToEndAsync()
    $p.WaitForExit()
    $lines = @($errTask.Result -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $err = ($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1)
    if (-not $err) { $err = $lines | Where-Object { $_ -match 'missing volume|unexpected end' } | Select-Object -First 1 }
    if (-not $err) { $err = $lines | Where-Object { $_ -match 'cannot|can not|^ERROR' } | Select-Object -First 1 }
    if (-not $err -and $lines) { $err = $lines[-1] }
    [pscustomobject]@{ Code = $p.ExitCode; Error = "$err" }
}

# =====================================================================
#  INPUT HELPERS
# =====================================================================
function Get-DownloadsFolder {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $raw = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'{374DE290-123F-4565-9164-39C4925E467B}'
    if ($raw) {
        $path = [Environment]::ExpandEnvironmentVariables($raw)
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return (Join-Path $env:USERPROFILE 'Downloads')
}

function Select-OutputFolder {
    $default = Get-DownloadsFolder
    while ($true) {
        Say ""
        Say "  Where should it be saved?" 'Cyan'
        $answer = Read-Host "  Folder [$default]"
        $path = if ([string]::IsNullOrWhiteSpace($answer)) { $default } else { [Environment]::ExpandEnvironmentVariables($answer.Trim().Trim('"').Trim("'")) }
        New-Item -ItemType Directory -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
        if (Test-Path -LiteralPath $path -PathType Container) {
            return (Resolve-Path -LiteralPath $path).ProviderPath
        }
        Say "  Cannot use '$path'. Try another folder." 'Red'
    }
}

function ConvertFrom-Secure($secure) {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Read-Password([switch]$Confirm) {
    Say ""
    while ($true) {
        $s = Read-Host '  Password [Enter = default]' -AsSecureString
        if ($s.Length -eq 0) { return $DefaultPassword }
        $pw = ConvertFrom-Secure $s
        if (-not $Confirm) { return $pw }
        if ((ConvertFrom-Secure (Read-Host '  Confirm password' -AsSecureString)) -ceq $pw) { return $pw }
        Say "  Passwords differ, try again." 'Red'
    }
}

# Dropping files into the console pastes them on one line, quoted when they contain spaces
function Split-Dropped([string]$line) {
    $line = $line.Trim()
    $whole = $line.Trim('"').Trim("'")
    if ($whole -and (Test-Path -LiteralPath $whole)) { return , $whole }
    [regex]::Matches($line, '"([^"]+)"|''([^'']+)''|(\S+)') | ForEach-Object {
        $_.Groups[1].Value + $_.Groups[2].Value + $_.Groups[3].Value
    }
}

function Read-DroppedPaths([string]$What, [scriptblock]$Resolve) {
    $items = New-Object System.Collections.Generic.List[string]
    Say ""
    Say "  Drop $What into this window (or type the path), then press Enter." 'Cyan'
    Say "  Press Enter on an empty line when done." 'DarkGray'
    while ($true) {
        Write-Host '  > ' -NoNewline
        $line = [Console]::ReadLine()
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($items.Count) { return $items.ToArray() }
            continue
        }
        foreach ($raw in Split-Dropped $line) {
            $p = & $Resolve $raw
            if (-not $p) { Say "    not found: $raw" 'Red' }
            elseif ($items -notcontains $p) { $items.Add($p); Say "    + $(Split-Path $p -Leaf)" 'Green' }
        }
    }
}

# =====================================================================
#  OPTION 1 - ZIP
# =====================================================================
function Invoke-Zip {
    $targets = Read-DroppedPaths 'files or folders' {
        param($raw)
        $p = [Environment]::ExpandEnvironmentVariables($raw)
        if (Test-Path -LiteralPath $p) { (Resolve-Path -LiteralPath $p).ProviderPath }
    }
    $password = Read-Password -Confirm
    $outDir   = Select-OutputFolder

    $total = 0
    foreach ($t in $targets) {
        if (Test-Path -LiteralPath $t -PathType Container) {
            $total += (Get-ChildItem -LiteralPath $t -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
        } else {
            $total += (Get-Item -LiteralPath $t -Force).Length
        }
    }
    $split = $total -gt ($SplitMB * 1MB)

    $base = Join-Path $outDir (Get-Date -Format 'yyyyMMdd-HHmmss')
    $args7z = @('a', '-t7z', "-p$password", '-mhe=on', '-mx=5')
    if ($split) { $args7z += "-v${SplitMB}m" }
    $args7z += "$base.7z"
    $args7z += $targets

    Say ""
    Say "  $($targets.Count) item(s), $(Format-Size $total) before compression." 'Gray'
    if ($split) { Say "  Over $SplitMB MB - will split into $SplitMB MB parts." 'Yellow' }
    Say "  Zipping..." 'Cyan'
    $r = Invoke-SevenZip $args7z
    if ($r.Code -ne 0) {
        Get-ChildItem -LiteralPath $outDir -Filter "$(Split-Path $base -Leaf).7z*" | Remove-Item -Force -ErrorAction SilentlyContinue
        Say "  FAILED: $($r.Error)" 'Red'
        return
    }

    # Name the result: one file -> no extension; several parts -> .001, .002, ...
    $leaf  = Split-Path $base -Leaf
    $parts = @(Get-ChildItem -LiteralPath $outDir -Filter "$leaf.7z.*" | Where-Object { $_.Name -match '\.\d{3}$' } | Sort-Object Name)
    if ($parts.Count -gt 1) {
        foreach ($f in $parts) { Rename-Item -LiteralPath $f.FullName -NewName ($leaf + $f.Extension) }
        $saved = @(Get-ChildItem -LiteralPath $outDir -Filter "$leaf.*" | Where-Object { $_.Name -match '\.\d{3}$' } | Sort-Object Name)
    } else {
        $single = if ($parts.Count -eq 1) { $parts[0].FullName } else { "$base.7z" }
        Rename-Item -LiteralPath $single -NewName $leaf
        $saved = @(Get-Item -LiteralPath $base)
    }

    Say ""
    Say "  DONE - $($saved.Count) file(s)" 'Green'
    foreach ($f in $saved) { Say ("    {0}  ({1})" -f $f.FullName, (Format-Size $f.Length)) 'Green' }
    if ($saved.Count -gt 1) { Say "  Keep all parts together in one folder to unzip." 'DarkYellow' }
}

# =====================================================================
#  OPTION 2 - UNZIP
# =====================================================================

# A dropped item -> the file to hand to 7-Zip (part .001 for split sets)
function Resolve-Archive([string]$raw) {
    $p = [Environment]::ExpandEnvironmentVariables($raw)
    foreach ($c in $p, "$p.7z", "$p.001") {
        if (Test-Path -LiteralPath $c -PathType Leaf) { $p = (Resolve-Path -LiteralPath $c).ProviderPath; break }
    }
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return $null }
    if ($p -match '^(.*)\.(\d{3})$') { return "$($Matches[1]).001" }
    return $p
}

# For a .001 path: check the set has no gaps; return a problem message or $null
function Test-SplitSet([string]$first) {
    $base = $first.Substring(0, $first.Length - 4)
    $dir  = Split-Path $base -Parent
    $leaf = Split-Path $base -Leaf
    $nums = @(Get-ChildItem -LiteralPath $dir -File |
        Where-Object { $_.Name.Length -eq $leaf.Length + 4 -and $_.Name.StartsWith("$leaf.", [StringComparison]::OrdinalIgnoreCase) -and $_.Name -match '\.(\d{3})$' } |
        ForEach-Object { [int]$_.Name.Substring($_.Name.Length - 3) } | Sort-Object)
    if (-not $nums -or $nums[0] -ne 1) { return "part 001 is missing" }
    $missing = @(1..$nums[-1] | Where-Object { $nums -notcontains $_ } | ForEach-Object { '{0:D3}' -f $_ })
    if ($missing) { return "part(s) $($missing -join ', ') missing (found up to $('{0:D3}' -f $nums[-1]))" }
    Say "    $leaf - split set of $($nums.Count) part(s)" 'DarkGray'
    return $null
}

function Invoke-Unzip {
    $archives = Read-DroppedPaths 'archives (any part of a split set is fine)' { param($raw) Resolve-Archive $raw }
    $password = Read-Password
    $outDir   = Select-OutputFolder
    Say ""

    $ok = 0
    foreach ($arc in $archives) {
        $name = Split-Path $arc -Leaf
        $stem = $name -replace '\.001$', '' -replace '\.7z$', ''
        if ($arc -match '\.001$') {
            $problem = Test-SplitSet $arc
            if ($problem) { Say "  SKIPPED $stem - $problem" 'Red'; continue }
        }
        $dest = Join-Path $outDir "$stem-extracted"
        New-Item -ItemType Directory -Path $dest -Force | Out-Null

        for ($attempt = 1; $attempt -le 4; $attempt++) {
            Say "  Extracting $stem..." 'Cyan'
            $r = Invoke-SevenZip @('x', "-p$password", '-y', "-o$dest", $arc)
            if ($r.Code -eq 0) { $ok++; Say "    done -> $dest" 'Green'; break }
            if ($r.Error -notmatch 'password' -or $attempt -eq 4) { Say "  SKIPPED $stem - $($r.Error)" 'Red'; break }
            $s = Read-Host "  Wrong password, try again ($attempt/3, Enter = skip)" -AsSecureString
            if ($s.Length -eq 0) { Say "  SKIPPED $stem" 'Red'; break }
            $password = ConvertFrom-Secure $s
        }
        if (-not (Get-ChildItem -LiteralPath $dest -Force -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue
        }
    }
    Say ""
    Say "  DONE - $ok of $($archives.Count) extracted to $outDir" $(if ($ok -eq $archives.Count) { 'Green' } else { 'Yellow' })
}

# =====================================================================
#  WINDOW LAYOUT (fails silently): this window right + on top, Explorer left
# =====================================================================
$WinApi = @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class ZipperWin {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder sb, int n);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SystemParametersInfoW(uint action, uint param, ref RECT r, uint ini);
    public static IntPtr[] ExplorerWindows() {
        var list = new List<IntPtr>();
        EnumWindows((h, l) => {
            var sb = new StringBuilder(32);
            GetClassNameW(h, sb, 32);
            if (sb.ToString() == "CabinetWClass" && IsWindowVisible(h)) list.Add(h);
            return true;
        }, IntPtr.Zero);
        return list.ToArray();
    }
}
'@

function Set-WindowLayout {
    try {
        if (-not ('ZipperWin' -as [type])) { Add-Type -TypeDefinition $WinApi -ErrorAction Stop }
        $area = New-Object ZipperWin+RECT
        [void][ZipperWin]::SystemParametersInfoW(0x30, 0, [ref]$area, 0)
        $height = $area.Bottom - $area.Top
        $width  = [Math]::Max(560, [int](($area.Right - $area.Left) * 0.30))

        # Under Windows Terminal the console is a hidden window owned by the terminal
        $console = [ZipperWin]::GetConsoleWindow()
        if ($console -ne [IntPtr]::Zero) {
            $root = [ZipperWin]::GetAncestor($console, 3)
            if ($root -ne [IntPtr]::Zero) { $console = $root }
            [void][ZipperWin]::ShowWindow($console, 9)
            [void][ZipperWin]::SetWindowPos($console, [IntPtr](-1), $area.Right - $width, $area.Top, $width, $height, 0x40)
        }

        $windows = [ZipperWin]::ExplorerWindows()
        if (-not $windows) {
            Start-Process explorer.exe (Get-DownloadsFolder)
            for ($i = 0; $i -lt 30 -and -not $windows; $i++) { Start-Sleep -Milliseconds 100; $windows = [ZipperWin]::ExplorerWindows() }
        }
        if ($windows) {
            [void][ZipperWin]::ShowWindow($windows[0], 9)
            [void][ZipperWin]::SetWindowPos($windows[0], [IntPtr]::Zero, $area.Left, $area.Top, $area.Right - $area.Left - $width, $height, 0x40)
        }
        if ($console -ne [IntPtr]::Zero) { [void][ZipperWin]::SetForegroundWindow($console) }
        return $console
    } catch { return [IntPtr]::Zero }
}

function Reset-WindowLayout($console) {
    try {
        if ($console -and $console -ne [IntPtr]::Zero) {
            [void][ZipperWin]::SetWindowPos($console, [IntPtr](-2), 0, 0, 0, 0, 0x13)   # not topmost; keep size/position
        }
    } catch {}
}

# =====================================================================
#  MENU
# =====================================================================
Clear-Host
$Console = Set-WindowLayout
try {
    Say ""
    Say "  =============================================" 'Cyan'
    Say "        ZIPPER" 'Cyan'
    Say "  =============================================" 'Cyan'
    Say ""
    Say "    1  -  Zip (password-protected, split over $SplitMB MB)" 'White'
    Say "    2  -  Unzip (split sets handled)" 'White'
    Say "    0  -  Quit" 'White'
    Say ""

    do {
        $choice = "$(Read-Host '  Choose 0-2 and press Enter')".Trim()
        if ($choice -notin '0', '1', '2') { Say "  Please type 1, 2 or 0." 'Red' }
    } until ($choice -in '0', '1', '2')

    if ($choice -ne '0') {
        $SevenZip = Get-SevenZip
        if ($SevenZip) {
            if ($choice -eq '1') { Invoke-Zip } else { Invoke-Unzip }
        }
        Say ""
        Read-Host "  Press Enter to exit" | Out-Null
    }
} finally {
    Reset-WindowLayout $Console
}
