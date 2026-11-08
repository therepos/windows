# Agent instructions

One-click Windows scripts for a non-technical user. Scripts are run straight from GitHub
`main` with `irm ... | iex`, so `main` must always work. Proxmox tools live in
`therepos/proxmox`, MCP servers in `therepos/mcp`, Claude skills in `therepos/skills`.

## Layout

- Root: docs only (`AGENTS.md`, `CLAUDE.md`, `README.md`), plus `apps/`, `scripts/`, `.github/`.
- `apps/<name>.ps1|.bat|.cmd|.py`, flat, one script per tool; `<name>` is one lowercase word
  (the knowledge-base stages are `kbprofile`, `kbrcm`, `kbbundle`; their plan is `apps/kbplan.md`)
- `scripts/`  repo helpers, never run by users: `sync-blocks.sh` (shared blocks), `lint.sh` (checks and safe fixes),
  `docgen-pages.py` (the GitHub Pages script page, therepos.github.io/windows, built and deployed by CI; nothing is committed)

## Script format

- Prefer `.ps1` run as `irm "https://github.com/therepos/windows/raw/main/apps/<name>.ps1?$(Get-Random)" | iex`:
  nothing to download, no SmartScreen warning, no Python needed. Do not keep a `.bat`
  copy of the same script.
- Header (see Documentation): line 1 the bootstrap, line 2 `# Purpose:`, then a fence.
- Shared blocks: regions between `# >>> <name>` and `# <<< <name>` are copies of the blocks in
  `scripts/sync-blocks.sh` (`say`, `menukey`, `downloads`, `elevate`, `layout`, `shortcut`). Never edit inside the markers; edit the block
  there, then run it or `scripts/lint.sh --fix`. Per-script code goes outside. Shared look and behaviour
  belongs in a block, so changing it never touches a script's features.
- `.ps1` files are ASCII only: Windows PowerShell 5.1 misreads other characters in files without a BOM.
- Must run in Windows PowerShell 5.1 (no PowerShell 7-only syntax).
- Ask for admin only when the script needs it; otherwise run in the user's own window. Self-elevate
  with the `elevate` shared block: `if (-not (Test-Admin)) { Start-Elevated $ScriptUrl; return }`
  (`-NoExit` for one-shot scripts whose output must stay on screen). It re-runs the bootstrap (or the
  local copy, when started from a shortcut) in an admin Windows Terminal window, because Windows ignores
  the default-terminal setting for admin windows; without Terminal it uses the classic console.
- Window: every bootstrap script puts its window on the right half of the screen with the `layout` shared
  block, calling `Set-WindowRight` at top level before its first screen (after the elevate check, never in
  a hidden or scheduled run). Lint checks the call.
- Missing tools are fetched without admin into `%LOCALAPPDATA%\<app>` (e.g. 7za.exe).
- Never hard-code passwords or tokens. Ask once on first use and save them
  DPAPI-encrypted (`ConvertFrom-SecureString`) in `%LOCALAPPDATA%\<app>\settings.json`.

## Menu style (see zipper.ps1, systeminfo.ps1)

- `Clear-Host`, a cyan `====` banner with the tool name in caps, then numbered options:
  `    1  -  Zip`. Short labels, no descriptions. `0  -  Quit` is always last.
- Prompt `Choose 0-N`; one keypress picks the option, no Enter (`Read-MenuKey` in
  zipper.ps1). Never write "press Enter"; prompts with a default show it as
  `Folder [C:\Users\...\Downloads]` or `(blank = Downloads)`.
- Tools that are used repeatedly (zipper) return to the menu after each action with the
  last action as the default, shown as `Choose 0-N [1]`, so Enter repeats it. One-shot
  tools (systeminfo) run the action once, then end. No "press Enter to exit" (only the
  elevated window may need a pause).
- `0  -  Quit` closes the window (`[Environment]::Exit(0)`), not just the menu.
- Settings, when there are any, is its own submenu: `Set <thing>   [current value]`,
  `N  -  Back`, `0  -  Quit`, saved between runs. Blank keeps the current value or
  restores the default.
- Every script with a menu offers `Settings` > `Shortcut '<name>'  [on/off]` via the `shortcut` shared block
  (name from `$ScriptUrl`, which every script defines): it saves a local copy in `%LOCALAPPDATA%\<name>` and adds
  or removes `function <name> { & "$env:LOCALAPPDATA\<name>\<name>.ps1" }` in `$PROFILE` (setting CurrentUser
  execution policy to RemoteSigned if needed). Nothing is fetched when the shortcut runs; turning it off and on
  again refreshes the copy. Scripts whose only setting is the shortcut call `Show-ShortcutSettings`; others
  list `Shortcut` in their own Settings and call `Test-Shortcut` / `Switch-Shortcut`. Lint checks it.
- Output via `Say <text> <colour>`: Cyan headings, White options, Green done, Yellow
  notes, Red errors, DarkGray progress.
- Default output folder is the user's real Downloads folder (registry Shell Folders).
- File names: `<Kind> <Subject> yyyyMMdd-HHmm`, e.g. `Data System XPS-13-9310 20261102-1430.txt`;
  model names become one word with dashes.

## Documentation

- Header by type, each followed by a fence and optional sections (`# Usage:`, `# Output:`, `# Note:`)
  closed by a fence:
  - `.ps1`: line 1 the bootstrap `# irm "https://github.com/therepos/windows/raw/main/apps/<name>.ps1?$(Get-Random)" | iex`
    (or `# pwsh -File .\<name>.ps1` for scripts run locally), line 2 `# Purpose:`.
    `?$(Get-Random)` makes a copied bootstrap fetch the latest version right after a merge, past GitHub's cache.
  - `.bat` / `.cmd`: line 1 `@echo off`, line 2 `:: Purpose:`, `::` fences; CRLF line endings
  - `.py`: line 1 `#!/usr/bin/env python3`, line 2 `# Purpose:`
- `# Usage:` lists commands or menu options only: one per line, descriptions aligned in one column, never wrapped.
  Menu conventions (one keypress, `0 = Quit`) are not repeated in headers.
- Explanations (how, why, what CI does) live in AGENTS.md, not in script headers.
- Concise, standard terms (bootstrap, shared block, lint, fence); no filler words.
- Code comments only where the reason is not obvious, one line each.
- Refinement: when the owner corrects a name, wording or format, apply it everywhere it occurs in the same
  PR, record the rule here, and add a lint check when it can be checked mechanically.

## Before pushing

- `scripts/lint.sh --fix --changed origin/main` must pass. `--fix` repairs the bootstrap, the shared blocks and
  `.bat`/`.cmd` line endings; fix the rest by hand. Options combine. `--changed` checks every script when
  `sync-blocks.sh`, `lint.sh` or `ci-lint.yml` changed. Needs `pwsh` and `shellcheck`.
- CI checks changed scripts on every PR and push (every script weekly, on a manual run, or when those files
  change) and pushes safe fixes to the PR branch, never to `main`.
- Work on a branch and open a PR; merge only when lint is green.

- Scripts cannot be run here: say so, and list what the user should test on Windows.
- Keep diffs small; check quoting (`ConvertTo-Arg` for 7-Zip arguments), `-LiteralPath`
  for user paths, that menus read one key with `[Console]::ReadKey` and every other
  prompt with `Read-Host` or `[Console]::ReadLine()`, and that `Test-Path` on user text
  cannot print an error (several quoted paths on one line are "illegal characters").
- Never commit tokens, passwords, LAN IPs, serial numbers or MAC addresses.
- Do not put model names in commits, comments or docs.
