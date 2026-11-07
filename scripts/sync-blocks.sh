#!/usr/bin/env bash
# Purpose: Copy the shared blocks (say, menukey, downloads) into every PowerShell script
# =============================================================================
# Usage:
#   scripts/sync-blocks.sh           copy each block into every script that has its markers
#   scripts/sync-blocks.sh --check   report scripts out of sync, change nothing
#
# Markers:
#   # >>> <name>  …  # <<< <name>
# =============================================================================
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

# --- Blocks ------------------------------------------------------------------
BLOCK_NAMES=(say menukey downloads)

# say: coloured console output
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_say=$(cat <<'BLOCK'
function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }
BLOCK
)

# menukey: one keypress picks a menu option; Enter picks the default
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_menukey=$(cat <<'BLOCK'
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
BLOCK
)

# downloads: the user's real Downloads folder (registry Shell Folders)
# shellcheck disable=SC2034  # read via ${!var} below
BLOCK_downloads=$(cat <<'BLOCK'
function Get-DownloadsFolder {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $raw = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'{374DE290-123F-4565-9164-39C4925E467B}'
    if ($raw) {
        $path = [Environment]::ExpandEnvironmentVariables($raw)
        if (Test-Path -LiteralPath $path) { return $path }
    }
    return (Join-Path $env:USERPROFILE 'Downloads')
}
BLOCK
)

# --- Sync --------------------------------------------------------------------
CHECK=0
[[ "${1:-}" == "--check" ]] && CHECK=1

TMP_BLOCK=$(mktemp)
trap 'rm -f "$TMP_BLOCK"' EXIT

drift=0
changed=0
for f in "$ROOT"/apps/*.ps1; do
    [[ -f "$f" ]] || continue
    cp "$f" "$f.new"
    for b in "${BLOCK_NAMES[@]}"; do
        grep -qE "^# >>> $b( |$)" "$f.new" || continue
        var="BLOCK_$b"
        printf '%s\n' "${!var}" > "$TMP_BLOCK"
        awk -v blockfile="$TMP_BLOCK" -v b="$b" '
            BEGIN { n=0; while ((getline l < blockfile) > 0) blk[n++]=l; close(blockfile) }
            $0 ~ "^# >>> " b "( |$)" { print; for (i=0;i<n;i++) print blk[i]; skip=1; next }
            $0 ~ "^# <<< " b "( |$)" { skip=0; print; next }
            !skip { print }
        ' "$f.new" > "$f.tmp" && mv "$f.tmp" "$f.new"
    done

    if cmp -s "$f" "$f.new"; then
        rm -f "$f.new"
    elif [[ $CHECK -eq 1 ]]; then
        rm -f "$f.new"
        echo "DRIFT: $(basename "$f")"
        drift=1
    else
        cat "$f.new" > "$f" && rm -f "$f.new"
        echo "updated: $(basename "$f")"
        changed=1
    fi
done

if [[ $CHECK -eq 1 ]]; then
    [[ $drift -eq 0 ]] && { echo "All in sync."; exit 0; } || { echo "Out of sync — run scripts/sync-blocks.sh"; exit 1; }
fi
[[ $changed -eq 0 ]] && echo "All in sync."
exit 0
