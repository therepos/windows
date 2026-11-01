@echo off
REM ============================================================
REM  File Finder - Double-click to run
REM  Edit settings below with Notepad (right-click > Edit)
REM ============================================================

REM -- Extract PowerShell script from this file and run it --
setlocal enabledelayedexpansion
set "TEMPPS=%TEMP%\find_files_temp.ps1"

> "!TEMPPS!" (
    set "found=0"
    for /f "usebackq delims=" %%A in ("%~f0") do (
        if !found! equ 1 echo %%A
        if "%%A"=="#POWERSHELL_START" set "found=1"
    )
)

powershell -ExecutionPolicy Bypass -NoProfile -File "!TEMPPS!"
del "!TEMPPS!" 2>nul
exit /b

#POWERSHELL_START
# ============================================================
#  USER CONFIGURATION - Change these when moving environments
# ============================================================

# LOCAL FOLDER: The OneDrive-synced root folder containing client subfolders
$ROOT = "C:\Users\LL725AE\EY\Risk IA - D. Proposals"

# KEYWORDS: Search for files containing any of these in the filename
$KEYWORDS = @("proposal", "ett", "rfq", "itt", "bq")

# FILE TYPES: Which file extensions to search (add/remove as needed)
$FILE_TYPES = @("*.pdf", "*.pptx")

# MAX RESULTS: Set to 0 for full scan, or a number to limit (e.g. 3 for testing)
$MAX_RESULTS = 0

# OUTPUT FILE: Name of the output CSV file (saved next to this script)
# For a full path, use e.g.: "C:\Users\LL725AE\Desktop\results.csv"
$OUTPUT_FILENAME = "proposals_results.csv"

# ============================================================
#  SHAREPOINT CONFIGURATION
# ============================================================
# How to set up:
#   1. Open any CLIENT FOLDER in SharePoint via Edge
#      (e.g. right-click folder in File Explorer > "View online")
#   2. Copy the full URL from Edge. It will look like:
#      https://sites.ey.com/sites/RiskIA/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FRiskIA%2FShared%20Documents%2FGeneral%2FD%2E%20Proposals%2FAGD&viewid=...
#   3. Split the URL into two parts:
#      PART A - Everything BEFORE "?id=" :
$SP_URL_BASE = "https://sites.ey.com/sites/RiskIA/Shared%20Documents/Forms/AllItems.aspx"
#      PART B - Everything AFTER "?id=" and BEFORE the client folder name (e.g. AGD):
$SP_FOLDER_PATH = "%2Fsites%2FRiskIA%2FShared%20Documents%2FGeneral%2FD%2E%20Proposals%2F"

# ============================================================
#  DO NOT EDIT BELOW THIS LINE
# ============================================================

$ErrorActionPreference = "Continue"

# Determine script location
$ScriptDir = $null
try {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
} catch { }
if (-not $ScriptDir -or $ScriptDir -eq "" -or $ScriptDir -eq $env:TEMP) {
    $ScriptDir = (Get-Location).Path
}

# Determine output path with timestamp
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$outName = [System.IO.Path]::GetFileNameWithoutExtension($OUTPUT_FILENAME)
$outExt = [System.IO.Path]::GetExtension($OUTPUT_FILENAME)
$outNameTimestamped = "${outName}_${timestamp}${outExt}"

if ([System.IO.Path]::IsPathRooted($OUTPUT_FILENAME)) {
    $outDir = [System.IO.Path]::GetDirectoryName($OUTPUT_FILENAME)
    $OutFile = Join-Path $outDir $outNameTimestamped
} else {
    $OutFile = Join-Path $ScriptDir $outNameTimestamped
}

function ConvertTo-SharePointFolderUrl {
    param([string]$LocalFolderPath)

    $relativePath = $LocalFolderPath.Replace($ROOT, "").TrimStart("\")
    $urlPath = $relativePath -replace '\\', '%2F'
    $urlPath = $urlPath -replace ' ', '%20'

    return "$SP_URL_BASE" + "?id=" + "$SP_FOLDER_PATH" + "$urlPath"
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N0} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# ============================================================
# MAIN
# ============================================================
if (-not (Test-Path $ROOT)) {
    Write-Host "[ERROR] Folder not found: $ROOT" -ForegroundColor Red
    Write-Host "  Please update the ROOT variable at the top of this script." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

$modeText = if ($MAX_RESULTS -gt 0) { "TEST MODE: First $MAX_RESULTS matches only" } else { "Full scan" }
$today = Get-Date

Write-Host ""
Write-Host ("=" * 65)
Write-Host "  Scanning    : $ROOT"
Write-Host "  Keywords    : $($KEYWORDS -join ', ')"
Write-Host "  File types  : $($FILE_TYPES -join ', ')"
Write-Host "  SharePoint  : $SP_URL_BASE"
Write-Host "  Mode        : $modeText"
Write-Host "  Output      : $OutFile"
Write-Host ("=" * 65)
Write-Host ""

# CSV header
$header = '"Filename","File Link","Folder","Subfolder","Folder (Local)","Folder (SharePoint)","Last Modified","Created","Age (Days)","File Size"'
$header | Out-File -FilePath $OutFile -Encoding UTF8

$totalClients = 0
$totalFound = 0
$totalMissing = 0
$limitReached = $false

$clientFolders = Get-ChildItem -Path $ROOT -Directory | Sort-Object Name

foreach ($clientDir in $clientFolders) {
    if ($limitReached) { break }

    $clientName = $clientDir.Name
    $totalClients++
    $clientFound = $false

    Write-Host "Scanning: $clientName ..."

    # Search for all configured file types
    $allFiles = @()
    foreach ($ft in $FILE_TYPES) {
        $found = Get-ChildItem -Path $clientDir.FullName -Filter $ft -Recurse -File -ErrorAction SilentlyContinue
        if ($found) { $allFiles += $found }
    }

    foreach ($file in $allFiles) {
        if ($limitReached) { break }

        $fname = $file.Name
        $fnameLower = $fname.ToLower()

        $matched = $false
        foreach ($kw in $KEYWORDS) {
            if ($fnameLower.Contains($kw)) {
                $matched = $true
                break
            }
        }

        if (-not $matched) { continue }

        $clientFound = $true
        $totalFound++

        $fnameClean = $fname -replace '"', '""'
        $folderPath = $file.DirectoryName
        $spFolderUrl = ConvertTo-SharePointFolderUrl -LocalFolderPath $folderPath

        # Subfolder path relative to client folder
        $subfolderPath = $folderPath.Replace($clientDir.FullName, "").TrimStart("\")
        if (-not $subfolderPath) { $subfolderPath = "(root)" }

        # File metadata
        $lastModified = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        $createdDate = $file.CreationTime.ToString("yyyy-MM-dd HH:mm")
        $ageDays = [math]::Floor(($today - $file.LastWriteTime).TotalDays)
        $fileSize = Format-FileSize -Bytes $file.Length

        $line = '"' + $fnameClean + '",' +
                '"=HYPERLINK(""' + $file.FullName + '"",""Open"")",' +
                '"' + $clientName + '",' +
                '"' + $subfolderPath + '",' +
                '"=HYPERLINK(""' + $folderPath + '"",""Open"")",' +
                '"=HYPERLINK(""' + $spFolderUrl + '"",""Open"")",' +
                '"' + $lastModified + '",' +
                '"' + $createdDate + '",' +
                '"' + $ageDays + '",' +
                '"' + $fileSize + '"'

        $line | Out-File -FilePath $OutFile -Encoding UTF8 -Append

        Write-Host "  [+] Found: $fname" -ForegroundColor Green

        if ($MAX_RESULTS -gt 0 -and $totalFound -ge $MAX_RESULTS) {
            Write-Host ""
            Write-Host "  ** Test limit of $MAX_RESULTS reached. Stopping. **" -ForegroundColor Yellow
            $limitReached = $true
            break
        }
    }

    if (-not $clientFound -and -not $limitReached) {
        $totalMissing++

        $clientFolderPath = $clientDir.FullName
        $spClientUrl = ConvertTo-SharePointFolderUrl -LocalFolderPath $clientFolderPath

        $line = '"Not Found",' +
                '"",' +
                '"' + $clientName + '",' +
                '"",' +
                '"=HYPERLINK(""' + $clientFolderPath + '"",""Open"")",' +
                '"=HYPERLINK(""' + $spClientUrl + '"",""Open"")",' +
                '"",' +
                '"",' +
                '"",' +
                '""'

        $line | Out-File -FilePath $OutFile -Encoding UTF8 -Append

        Write-Host "  [!] NO MATCH: $clientName" -ForegroundColor Yellow
    }
}

# Summary
Write-Host ""
Write-Host ("=" * 65)
Write-Host "  DONE!" -ForegroundColor Green
Write-Host "  Client folders scanned  : $totalClients"
Write-Host "  Matching files found    : $totalFound"
Write-Host "  Clients with no match   : $totalMissing"
if ($MAX_RESULTS -gt 0) {
    Write-Host ""
    Write-Host "  ** TEST MODE: Change MAX_RESULTS to 0 for full scan **" -ForegroundColor Yellow
}
Write-Host ("=" * 65)
Write-Host ""
Write-Host "  Results saved to:"
Write-Host "    $OutFile" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"