# windows

One-click Windows scripts.

- `.bat` / `.cmd`: download and double-click.
- `.ps1`: run the `irm https://github.com/therepos/windows/raw/main/<name>.ps1 | iex` line from the file header in PowerShell.
- `.py`: run with Python 3.

## One-word shortcuts (optional)

`zipper`, `filesync`: run the tool once, then Settings > Shortcut. It saves a local copy in
`%LOCALAPPDATA%\<tool>` and runs that, so nothing is downloaded each time.

`sysinfo`: paste once in PowerShell, then open a new window and type `sysinfo`:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force; if (!(Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }; Add-Content $PROFILE "`nfunction sysinfo { irm https://github.com/therepos/windows/raw/main/systeminfo.ps1 | iex }"
```

Remove them:

```powershell
(Get-Content $PROFILE) | Where-Object { $_ -notmatch '^function (zipper|sysinfo|filesync) ' } | Set-Content $PROFILE
```

Kept per Windows user, per PC.
