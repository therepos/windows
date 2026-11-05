# Agent instructions

One-click Windows scripts for a non-technical user. Scripts are run straight from GitHub
`main` with `irm ... | iex`, so `main` must always work. Proxmox tools live in
`therepos/proxmox`, MCP servers in `therepos/mcp`, Claude skills in `therepos/skills`.

## Script format

- Prefer `.ps1` run as `irm https://github.com/therepos/windows/raw/main/<name>.ps1 | iex`:
  nothing to download, no SmartScreen warning, no Python needed. Do not keep a `.bat`
  copy of the same script.
- Line 1 is that one-liner, line 2 is `# Purpose: ...`, then a short header comment
  listing the menu options and where output goes.
- Must run in Windows PowerShell 5.1 (no PowerShell 7-only syntax).
- Ask for admin only when the script needs it. Self-elevate by re-running the one-liner
  in an elevated window (`"${ScriptUrl}?$(Get-Random)"`, note the braces); otherwise run
  in the user's own window.
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
- Settings may offer `Shortcut '<name>'  [on/off]`, which adds or removes
  `function <name> { irm ... | iex }` in `$PROFILE` (setting CurrentUser execution
  policy to RemoteSigned if needed).
- Output via `Say <text> <colour>`: Cyan headings, White options, Green done, Yellow
  notes, Red errors, DarkGray progress.
- Default output folder is the user's real Downloads folder (registry Shell Folders).
- File names: `<Kind> <Subject> yyyyMMdd-HHmm`, e.g. `Data System XPS-13-9310 20261102-1430.txt`;
  model names become one word with dashes.

## Before pushing

- Scripts cannot be run here: say so, and list what the user should test on Windows.
- Keep diffs small; check quoting (`ConvertTo-Arg` for 7-Zip arguments), `-LiteralPath`
  for user paths, that menus read one key with `[Console]::ReadKey` and every other
  prompt with `Read-Host` or `[Console]::ReadLine()`, and that `Test-Path` on user text
  cannot print an error (several quoted paths on one line are "illegal characters").
- Never commit tokens, passwords, LAN IPs, serial numbers or MAC addresses.
- Do not put model names in commits, comments or docs.
