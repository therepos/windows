#!/usr/bin/env bash
# Purpose: Check scripts against AGENTS.md and fix what is safe to fix
# =============================================================================
# Usage:
#   scripts/lint.sh                    check all scripts, change nothing
#   scripts/lint.sh --fix              also fix bootstrap, shared blocks, line endings
#   scripts/lint.sh --changed <base>   only scripts changed since <base>
# =============================================================================
set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT" || exit 1

FIX=0
BASE=""
while (( $# )); do
    case "$1" in
        --fix)     FIX=1 ;;
        --changed) BASE="${2:?--changed needs a base, e.g. origin/main}"; shift ;;
        *)         echo "Unknown option '$1'"; exit 2 ;;
    esac
    shift
done

# Files that change the rules for every script: when one changes, check everything.
FULL_CHECK_TRIGGERS=(scripts/sync-blocks.sh scripts/lint.sh .github/workflows/ci-lint.yml)

shopt -s nullglob
ALL=(apps/*.ps1 apps/*.bat apps/*.cmd apps/*.py scripts/*.sh scripts/*.py)

if [[ -n "$BASE" ]]; then
    mapfile -t CHANGED < <({ git diff --name-only --diff-filter=ACMR "$BASE"...HEAD
                             git diff --name-only --diff-filter=ACMR HEAD
                             git ls-files --others --exclude-standard; } 2>/dev/null | sort -u)
    full=0
    for t in "${FULL_CHECK_TRIGGERS[@]}"; do
        printf '%s\n' "${CHANGED[@]}" | grep -qx "$t" && { full=1; echo "[INFO] $t changed: checking every script"; }
    done
    if (( ! full )); then
        FILES=()
        for f in "${CHANGED[@]}"; do
            [[ -f "$f" ]] || continue
            for a in "${ALL[@]}"; do [[ "$f" == "$a" ]] && FILES+=("$f"); done
        done
        echo "[INFO] Checking ${#FILES[@]} changed script(s) since $BASE"
    fi
fi
if [[ -z "$BASE" ]] || (( full )); then FILES=("${ALL[@]}"); fi

REPO_RAW="https://github.com/therepos/windows/raw/main"
# ?$(Get-Random) makes GitHub serve the latest copy right after a merge (no 5-minute cache)
# shellcheck disable=SC2016
bootstrap() { printf '# irm "%s/%s?$(Get-Random)" | iex' "$REPO_RAW" "$1"; }
FENCE='# ============================================================================='
BFENCE=':: ============================================================================'

errors=0
bad() { printf '[FAIL] %s: %s\n' "$1" "$2"; errors=$((errors + 1)); }
line() { sed -n "${2}p" "$1" | tr -d '\r'; }

# --- Safe fixes --------------------------------------------------------------
if (( FIX )); then
    for f in "${FILES[@]}"; do
        case "$f" in
            *.ps1)
                l1="$(line "$f" 1)"; want="$(bootstrap "$f")"
                if [[ "$l1" == '# irm '* && "$l1" != "$want" ]]; then
                    awk -v w="$want" 'NR==1{print w; next} {print}' "$f" > "$f.tmp" && cat "$f.tmp" > "$f" && rm -f "$f.tmp"
                    echo "[FIX ] $f: bootstrap"
                fi ;;
            *.bat|*.cmd)
                if sed -n '$!p' "$f" | grep -qv $'\r$'; then
                    sed -i 's/\r$//; s/$/\r/' "$f"
                    echo "[FIX ] $f: CRLF line endings"
                fi ;;
        esac
    done
    scripts/sync-blocks.sh | sed -n 's/^updated: /[FIX ] shared blocks: /p'
fi

for tool in shellcheck pwsh; do
    command -v "$tool" >/dev/null || { echo "[FAIL] $tool not installed"; exit 1; }
done

# header_ok <file> <fence> <first-fence-line>: fence on that line; any comment
# block right after it must end with a closing fence
header_ok() {
    local f="$1" fence="$2" n="$3" prefix="${2%% *}" last
    [[ "$(line "$f" "$n")" == "$fence" ]] || { bad "$f" "line $n must be the header fence"; return; }
    last="$(tr -d '\r' < "$f" | awk -v s=$((n + 1)) -v p="$prefix" 'NR>=s { if (index($0, p) == 1) l=$0; else exit } END { print l }')"
    [[ -z "$last" || "$last" == "$fence" ]] || bad "$f" "header comment block must end with a fence"
}

# --- Layout: scripts live in apps/ with one-word names --------------------------
for f in *.ps1 *.bat *.cmd *.py *.sh; do bad "$f" "scripts belong in apps/ (or scripts/ for repo helpers), not the root"; done
for f in "${FILES[@]}"; do
    [[ "$f" == apps/* ]] || continue
    [[ "$f" =~ ^apps/[a-z0-9]+\.(ps1|bat|cmd|py)$ ]] || bad "$f" "name must be one lowercase word"
done

# --- PowerShell --------------------------------------------------------------
PS=()
for f in "${FILES[@]}"; do
    [[ "$f" == *.ps1 ]] || continue
    PS+=("$f")
    l1="$(line "$f" 1)"
    if [[ "$l1" == '# irm '* ]]; then
        [[ "$l1" == "$(bootstrap "$f")" ]] || bad "$f" "bootstrap (line 1) is wrong (scripts/lint.sh --fix)"
        { grep -q '^# >>> layout$' "$f" && grep -q '^Set-WindowRight$' "$f"; } \
            || bad "$f" "needs the layout block and a top-level 'Set-WindowRight' call before the first screen"
        if grep -q "Choose 0-" "$f"; then
            { grep -q '^# >>> shortcut$' "$f" && grep -Eq 'Show-ShortcutSettings|Switch-Shortcut' "$f"; } \
                || bad "$f" "a script with a menu needs the shortcut block and Settings > Shortcut"
        fi
    elif [[ "$l1" != '# pwsh -File '* ]]; then
        bad "$f" "line 1 must be the bootstrap ('# irm ... | iex' or '# pwsh -File ...')"
    fi
    [[ "$(line "$f" 2)" == "# Purpose: "* ]] || bad "$f" "line 2 must be '# Purpose: ...'"
    header_ok "$f" "$FENCE" 3
    LC_ALL=C grep -q '[^[:print:][:space:]]' "$f" && bad "$f" "non-ASCII characters (Windows PowerShell 5.1 misreads them)"
done
if (( ${#PS[@]} )); then
    PARSE="$(mktemp --suffix=.ps1)"
    cat > "$PARSE" <<'PWSH'
foreach ($f in $args) {
    $t = $null; $e = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path -LiteralPath $f).Path, [ref]$t, [ref]$e)
    foreach ($x in $e) { "{0}|line {1}: {2}" -f $f, $x.Extent.StartLineNumber, $x.Message }
}
PWSH
    while IFS= read -r msg; do [[ -n "$msg" ]] && bad "${msg%%|*}" "${msg#*|}"; done < <(pwsh -NoProfile -File "$PARSE" "${PS[@]}")
    rm -f "$PARSE"
fi

# --- Batch -------------------------------------------------------------------
for f in "${FILES[@]}"; do
    [[ "$f" == *.bat || "$f" == *.cmd ]] || continue
    [[ "$(line "$f" 1 | tr 'A-Z' 'a-z')" == "@echo off" ]] || bad "$f" "line 1 must be '@echo off'"
    [[ "$(line "$f" 2)" == ":: Purpose: "* ]] || bad "$f" "line 2 must be ':: Purpose: ...'"
    header_ok "$f" "$BFENCE" 3
    sed -n '$!p' "$f" | grep -qv $'\r$' && bad "$f" "needs CRLF line endings (scripts/lint.sh --fix)"
done

# --- Python ------------------------------------------------------------------
for f in "${FILES[@]}"; do
    [[ "$f" == *.py ]] || continue
    [[ "$(line "$f" 1)" == "#!/usr/bin/env python3" ]] || bad "$f" "line 1 must be '#!/usr/bin/env python3'"
    [[ "$(line "$f" 2)" == "# Purpose: "* ]] || bad "$f" "line 2 must be '# Purpose: ...'"
    header_ok "$f" "$FENCE" 3
    python3 -m py_compile "$f" 2>/dev/null || bad "$f" "python syntax error"
done
find . -name __pycache__ -type d -prune -exec rm -rf {} + 2>/dev/null

# --- Repo helpers ------------------------------------------------------------
for f in "${FILES[@]}"; do
    [[ "$f" == scripts/*.sh ]] || continue
    bash -n "$f" 2>/dev/null || bad "$f" "bash syntax error (bash -n)"
    shellcheck -S warning "$f" >/dev/null || bad "$f" "shellcheck warnings (shellcheck -S warning $f)"
    [[ "$(line "$f" 2)" == "# Purpose: "* ]] || bad "$f" "line 2 must be '# Purpose: ...'"
    [[ "$(line "$f" 4)" == "# Usage:" ]] || bad "$f" "line 4 must be '# Usage:'"
    awk 'NR>4 && /^# =/ {exit} NR>4 && /^#           *[^ ]/ {found=1} END {exit !found}' "$f" \
        && bad "$f" "Usage entries must be one line each (no wrapping)"
done

# --- Shared-block copies match the source (all scripts; fast) ----------------
scripts/sync-blocks.sh --check >/dev/null || bad "scripts/sync-blocks.sh" "shared blocks out of sync (scripts/lint.sh --fix)"

if (( errors )); then
    echo "$errors problem(s)."
    exit 1
fi
echo "[ OK ] All scripts pass."
