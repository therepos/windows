# Agent instructions

One-click Windows scripts for a non-technical user, run straight from GitHub `main`, so `main` must
always work. Proxmox tools: `therepos/proxmox`; MCP servers: `therepos/mcp`; Claude skills: `therepos/skills`.

## Layout

- Root: docs only (`AGENTS.md`, `CLAUDE.md`, `README.md`), plus `apps/`, `scripts/`, `.github/`.
- `apps/<name>.ps1|.bat|.cmd|.py`: flat, one script per tool, `<name>` one lowercase word
  (knowledge-base stages `kbprofile`, `kbrcm`, `kbbundle`; plan in `apps/kbplan.md`).
- `scripts/` (repo helpers, never run by users): `sync-blocks.sh` shared blocks, `lint.sh` checks and safe
  fixes, `docgen-pages.py` the script page (therepos.github.io/windows, built and deployed by CI).

## Scripts

- Prefer `.ps1` (no download, no SmartScreen, no Python); no `.bat` copy of the same tool.
- Windows PowerShell 5.1 syntax, ASCII only (5.1 misreads other characters in files without a BOM).
- Every bootstrap script defines `$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/<name>.ps1'`.
- Shared blocks: code between `# >>> <name>` and `# <<< <name>` is copied from `scripts/sync-blocks.sh`.
  Never edit inside the markers; edit the block, then run `scripts/lint.sh --fix`. Shared look and
  behaviour belongs in a block. Blocks:
  - `say`: `Say <text> <colour>` (Cyan headings, White options, Green done, Yellow notes, Red errors, DarkGray progress)
  - `menukey`: `Read-MenuKey`, one keypress per menu choice
  - `downloads`: `Get-DownloadsFolder`, the real Downloads folder (default output)
  - `elevate`: `if (-not (Test-Admin)) { Start-Elevated $ScriptUrl; return }`, only when admin is needed;
    `-NoExit` for one-shot scripts whose output must stay. Re-runs the bootstrap, or the local copy when
    started from a shortcut, in an admin Windows Terminal window (Windows ignores the default-terminal
    setting for admin windows); classic console when Terminal is missing.
  - `layout`: `Set-WindowRight` at top level before the first screen (after the elevate check, never in a
    hidden or scheduled run) puts the window on the right half of its screen.
  - `shortcut`: every script with a menu offers Settings > `Shortcut '<name>'  [on/off]`. It saves a local
    copy in `%LOCALAPPDATA%\<name>` and adds `function <name> { & "$env:LOCALAPPDATA\<name>\<name>.ps1" }`
    to `$PROFILE`; nothing is fetched when it runs, and off/on refreshes the copy. Use `Show-ShortcutSettings`
    when the shortcut is the only setting, else `Test-Shortcut` / `Switch-Shortcut` in the script's Settings.
- Per-tool files (tools, settings, logs, local copy) go in `%LOCALAPPDATA%\<name>`, fetched without admin.
- Secrets: never hard-coded; ask once, save DPAPI-encrypted (`ConvertFrom-SecureString`) in `settings.json`.
- Renaming a tool: on first run, move the old name's scheduled task, shortcut and files, then remove them.

## Menus (see zipper.ps1)

- `Clear-Host`, cyan `====` banner with the tool name in caps, options as `    1  -  Zip` (short labels),
  `0  -  Quit` last; it closes the window (`[Environment]::Exit(0)`).
- Prompt `Choose 0-N`, one keypress. Never write "press Enter"; defaults show as `[value]` or `(blank = ...)`.
- Repeated-use tools return to the menu with the last choice as default (`Choose 0-N [1]`); one-shot tools
  run once and end (only an elevated window may pause before closing).
- Settings is a submenu: `Set <thing>   [current value]`, `N  -  Back`, `0  -  Quit`, saved between runs;
  blank keeps the current value.
- Output files: `<Kind> <Subject> yyyyMMdd-HHmm`, device model names as one word with dashes.

## Headers

- `.ps1` line 1 the bootstrap `# irm "https://github.com/therepos/windows/raw/main/apps/<name>.ps1?$(Get-Random)" | iex`
  (`?$(Get-Random)` gets the latest version right after a merge, past GitHub's cache), or `# pwsh -File .\<name>.ps1`
  for scripts run locally. `.bat`/`.cmd`: `@echo off`, `::` comments, CRLF. `.py`: `#!/usr/bin/env python3`.
- Line 2 `# Purpose:` (`:: Purpose:`), then a fence; optional `# Usage:`, `# Output:`, `# Note:` closed by a fence.
- `# Usage:` lists commands or menu options only, one line each, descriptions in one column. Explanations
  live here, not in headers. Code comments only where the reason is not obvious, one line each.
- Concise standard terms (bootstrap, shared block, lint, fence).
- Refinement: when the owner corrects a name, wording or format, apply it everywhere in the same PR, record
  the rule here, and add a lint check when it can be checked mechanically.

## PowerShell pitfalls

- `"$var?"` reads a variable named `var?`: write `"${var}?..."`.
- `.BaseObject` on a hashtable reads a key; use `-is [hashtable]` / a cast for runspace results.
- Running a program (`taskkill.exe`, `winget.exe`) hides `Write-Progress` until the next call; draw progress
  as text lines instead.
- In Windows Terminal, `GetConsoleWindow()` returns a hidden window that reports visible; the real window
  is its owner (`GetAncestor(..., GA_ROOTOWNER)`).
- `Start-Process -Verb RunAs` cannot pass `;` or nested quotes through `wt.exe`: use `-EncodedCommand`.
- Never close Windows services (session 0) or relaunch helper processes; their own installer or app does.

## Before pushing

- `scripts/lint.sh --fix --changed origin/main` must pass (needs `pwsh`, `shellcheck`). `--fix` repairs the
  bootstrap, shared blocks and CRLF; fix the rest by hand. Block or lint changes re-check every script.
- CI lints every PR and push (all scripts weekly) and pushes safe fixes to the PR branch, never to `main`.
- Work on a branch and open a PR (`gh api repos/<owner>/<repo>/pulls`; GraphQL `gh pr` commands are blocked here);
  merge only when lint is green. Commit before tests that run `git checkout`, and check each edit landed
  (`git diff`) before reporting it.
- Scripts cannot be run here: say so and list what to test on Windows.
- Check quoting (`ConvertTo-Arg` for 7-Zip), `-LiteralPath` for user paths, menus via `[Console]::ReadKey`,
  other prompts via `Read-Host`, and that `Test-Path` on user text cannot print an error.
- Never commit tokens, passwords, LAN IPs, serial numbers or MAC addresses; no AI model names in commits, comments or docs.
