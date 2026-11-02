# irm https://github.com/therepos/windows/raw/main/zipper.ps1 | iex
# Purpose: Zip files into a password-protected 7z (auto-split over 500 MB) or unzip one, split sets included
# =============================================================================
#  Zipper:
#   1 = Zip       files/folders -> YYYYMMDD-HHMMSS (no extension)
#                 over the split size -> YYYYMMDD-HHMMSS.001, .002, ...
#   2 = Unzip     archive or any part of a split set -> <name>-extracted
#   3 = Settings  password, save-to folder, split size (kept between runs;
#                 password stored encrypted for this Windows user), and the
#                 'zipper' shortcut, which runs a local copy in %LOCALAPPDATA%\zipper
#   0 = Quit
#  Password: asked once on first use, then saved. Defaults: Downloads, split at 500 MB.
#  Uses an installed 7-Zip, otherwise downloads the standalone 7za.exe
#  into %LOCALAPPDATA%\zipper (no admin needed).
# =============================================================================

$DefaultSplitMB  = 500
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

function ConvertFrom-Secure($secure) {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# =====================================================================
#  SETTINGS  (%LOCALAPPDATA%\zipper\settings.json)
# =====================================================================
$SettingsFile = Join-Path $ToolDir 'settings.json'

function Import-Settings {
    $s = [pscustomobject]@{ Password = $null; OutDir = $null; SplitMB = $DefaultSplitMB }
    try {
        $saved = Get-Content -LiteralPath $SettingsFile -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($saved.Password) { $s.Password = [string]$saved.Password }
        if ($saved.OutDir)   { $s.OutDir   = [string]$saved.OutDir }
        if ($null -ne $saved.SplitMB) { $s.SplitMB = [int]$saved.SplitMB }
    } catch {}
    return $s
}

function Export-Settings {
    New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null
    $Settings | ConvertTo-Json | Set-Content -LiteralPath $SettingsFile -Encoding UTF8
}

# Saved password is DPAPI-encrypted: only this Windows user on this PC can read it
function Get-Password([switch]$Confirm) {
    if ($Settings.Password) {
        try { return ConvertFrom-Secure ($Settings.Password | ConvertTo-SecureString) } catch {}
    }
    # First use: ask once, save it (encrypted) for next time
    Say ""
    Say "  Your zip password is needed once. It will be remembered." 'Cyan'
    while ($true) {
        $s = Read-Host '  Zip password' -AsSecureString
        if ($s.Length -eq 0) { continue }
        if ($Confirm -and (ConvertFrom-Secure (Read-Host '  Confirm password' -AsSecureString)) -cne (ConvertFrom-Secure $s)) {
            Say "  Passwords differ, try again." 'Red'
            continue
        }
        $Settings.Password = $s | ConvertFrom-SecureString
        Export-Settings
        return (ConvertFrom-Secure $s)
    }
}

function Get-OutputDir {
    $dir = if ($Settings.OutDir) { $Settings.OutDir } else { Get-DownloadsFolder }
    New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue | Out-Null
    if (Test-Path -LiteralPath $dir -PathType Container) { return $dir }
    Say "  Cannot use '$dir' - saving to Downloads instead." 'Yellow'
    return (Get-DownloadsFolder)
}

function Set-PasswordSetting {
    Say ""
    while ($true) {
        $s = Read-Host '  New password (blank = keep current)' -AsSecureString
        if ($s.Length -eq 0) { Say "  Password unchanged." 'Gray'; return }
        $c = Read-Host '  Confirm password' -AsSecureString
        if ((ConvertFrom-Secure $c) -ceq (ConvertFrom-Secure $s)) {
            $Settings.Password = $s | ConvertFrom-SecureString
            Say "  Password saved." 'Green'
            return
        }
        Say "  Passwords differ, try again." 'Red'
    }
}

function Set-OutDirSetting {
    Say ""
    while ($true) {
        $answer = Read-Host "  Save to folder (blank = Downloads)"
        if ([string]::IsNullOrWhiteSpace($answer)) { $Settings.OutDir = $null; Say "  Save to: Downloads" 'Green'; return }
        $path = [Environment]::ExpandEnvironmentVariables($answer.Trim().Trim('"').Trim("'"))
        New-Item -ItemType Directory -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
        if (Test-Path -LiteralPath $path -PathType Container) {
            $Settings.OutDir = (Resolve-Path -LiteralPath $path).ProviderPath
            Say "  Save to: $($Settings.OutDir)" 'Green'
            return
        }
        Say "  Cannot use '$path'. Try another folder." 'Red'
    }
}

function Set-SplitSetting {
    Say ""
    while ($true) {
        $answer = "$(Read-Host "  Split into parts of how many MB? (0 = never split, blank = $DefaultSplitMB)")".Trim()
        if (-not $answer) { $answer = "$DefaultSplitMB" }
        $mb = 0
        if ([int]::TryParse($answer, [ref]$mb) -and $mb -ge 0) {
            $Settings.SplitMB = $mb
            Say $(if ($mb -eq 0) { "  Split: never" } else { "  Split: $mb MB parts" }) 'Green'
            return
        }
        Say "  Please type a whole number, e.g. 500." 'Red'
    }
}

# Optional one-word shortcut: typing `zipper` in any new PowerShell window runs a local copy
# of this script saved in %LOCALAPPDATA%\zipper (nothing is fetched from the internet)
$ScriptFile   = $PSCommandPath
$LocalScript  = Join-Path $ToolDir 'zipper.ps1'
$ShortcutLine = 'function zipper { & "$env:LOCALAPPDATA\zipper\zipper.ps1" }'

function Test-Shortcut {
    (Test-Path -LiteralPath $PROFILE) -and (Test-Path -LiteralPath $LocalScript) -and
        (@(Get-Content -LiteralPath $PROFILE -ErrorAction SilentlyContinue) -contains $ShortcutLine)
}

function Switch-Shortcut {
    try {
        $on = Test-Shortcut
        # Remove any zipper shortcut, including the older one that downloaded the script each time
        if (Test-Path -LiteralPath $PROFILE) {
            @(Get-Content -LiteralPath $PROFILE) | Where-Object { $_ -notmatch '^function zipper ' } | Set-Content -LiteralPath $PROFILE
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
            Invoke-WebRequest -Uri 'https://github.com/therepos/windows/raw/main/zipper.ps1' -OutFile $LocalScript -UseBasicParsing
        }
        Unblock-File -LiteralPath $LocalScript -ErrorAction SilentlyContinue
        if ((Get-ExecutionPolicy -Scope CurrentUser) -in 'Undefined', 'Restricted', 'AllSigned') {
            Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
        }
        if (-not (Test-Path -LiteralPath $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }
        Add-Content -LiteralPath $PROFILE "`n$ShortcutLine"
        Say "  Shortcut added. In a new PowerShell window, just type: zipper" 'Green'
    } catch {
        Say "  Could not change the shortcut: $($_.Exception.Message)" 'Red'
    }
}

function Show-Settings {
    while ($true) {
        $pw    = if ($Settings.Password) { 'saved' } else { 'not set' }
        $out   = if ($Settings.OutDir) { $Settings.OutDir } else { "Downloads" }
        $split = if ($Settings.SplitMB -gt 0) { "$($Settings.SplitMB) MB" } else { 'never' }
        $short = if (Test-Shortcut) { 'on' } else { 'off' }
        Say ""
        Say "  ---------------- SETTINGS -------------------" 'Cyan'
        Say ""
        Say "    1  -  Set password          [$pw]" 'White'
        Say "    2  -  Set output location   [$out]" 'White'
        Say "    3  -  Set split size        [$split]" 'White'
        Say "    4  -  Shortcut 'zipper'     [$short]" 'White'
        Say "    0  -  Back" 'White'
        Say ""
        $c = "$(Read-Host '  Choose 0-4')".Trim()
        switch ($c) {
            '1' { Set-PasswordSetting; Export-Settings }
            '2' { Set-OutDirSetting;   Export-Settings }
            '3' { Set-SplitSetting;    Export-Settings }
            '4' { Switch-Shortcut }
            '0' { return }
            default { Say "  Please type 1, 2, 3, 4 or 0." 'Red' }
        }
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
    Say "  Drop $What here (one or many):" 'Cyan'
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
    $password = Get-Password -Confirm
    $outDir   = Get-OutputDir
    $SplitMB  = $Settings.SplitMB

    $total = 0
    foreach ($t in $targets) {
        if (Test-Path -LiteralPath $t -PathType Container) {
            $total += (Get-ChildItem -LiteralPath $t -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
        } else {
            $total += (Get-Item -LiteralPath $t -Force).Length
        }
    }
    $split = $SplitMB -gt 0 -and $total -gt ($SplitMB * 1MB)

    # Unique name: 7-Zip would otherwise add to an archive made in the same second
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $base  = Join-Path $outDir $stamp
    for ($n = 2; Get-ChildItem -LiteralPath $outDir -Filter "$(Split-Path $base -Leaf)*" -ErrorAction SilentlyContinue; $n++) {
        $base = Join-Path $outDir "$stamp-$n"
    }
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
    $archives = Read-DroppedPaths 'archives' { param($raw) Resolve-Archive $raw }
    $password = Get-Password
    $outDir   = Get-OutputDir
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
            $s = Read-Host "  Wrong password, try again ($attempt/3, blank = skip)" -AsSecureString
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
    $Settings = Import-Settings
    while ($true) {
        Say ""
        Say "  =============================================" 'Cyan'
        Say "        ZIPPER" 'Cyan'
        Say "  =============================================" 'Cyan'
        Say ""
        Say "    1  -  Zip" 'White'
        Say "    2  -  Unzip" 'White'
        Say "    3  -  Settings" 'White'
        Say "    0  -  Quit" 'White'
        Say ""
        $choice = "$(Read-Host '  Choose 0-3')".Trim()
        if ($choice -eq '3') { Show-Settings; Clear-Host; continue }
        if ($choice -in '0', '1', '2') { break }
        Say "  Please type 1, 2, 3 or 0." 'Red'
    }

    if ($choice -ne '0') {
        $SevenZip = Get-SevenZip
        if ($SevenZip) {
            if ($choice -eq '1') { Invoke-Zip } else { Invoke-Unzip }
        }
        Say ""
    }
} finally {
    Reset-WindowLayout $Console
}
