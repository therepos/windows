# windows

One-click Windows scripts.

- `.bat` / `.cmd`: download and double-click.
- `.ps1`: run the bootstrap on line 1 of the file (`irm https://github.com/therepos/windows/raw/main/<name>.ps1 | iex`) in PowerShell.
- `.py`: run with Python 3.

## One-word shortcuts (optional)

`zipper`, `filesync`, `autoupdate`: run the tool once, then Settings > Shortcut. It saves a local copy in
`%LOCALAPPDATA%\<tool>` and runs that, so nothing is downloaded each time.

`sysinfo`: paste once in PowerShell, then open a new window and type `sysinfo`:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force; if (!(Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }; Add-Content $PROFILE "`nfunction sysinfo { irm https://github.com/therepos/windows/raw/main/systeminfo.ps1 | iex }"
```

Remove them:

```powershell
(Get-Content $PROFILE) | Where-Object { $_ -notmatch '^function (zipper|sysinfo|filesync|autoupdate) ' } | Set-Content $PROFILE
```

Kept per Windows user, per PC.

## Structure

```
windows
├── AGENTS.md               # rules for AI agents
├── CLAUDE.md               # @AGENTS.md
├── .github
│   └── workflows           # lint on every PR and push
├── *.ps1 *.bat *.cmd *.py  # one script per tool
├── kb                      # knowledge-base stages (Python)
└── scripts
    ├── sync-blocks.sh      # shared blocks (say, menukey, downloads), synced into every script
    └── lint.sh             # checks scripts against AGENTS.md (--fix repairs what it can)
```
