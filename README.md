# windows

One-click Windows scripts.

- `.bat` / `.cmd`: download and double-click.
- `.ps1`: run the `irm https://github.com/therepos/windows/raw/main/<name>.ps1 | iex` line from the file header in PowerShell.
- `.py`: run with Python 3.

## One-word shortcuts (optional)

Paste once in PowerShell, then open a new window and type `zipper` or `sysinfo`:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force; if (!(Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }; Add-Content $PROFILE "`nfunction zipper { irm https://github.com/therepos/windows/raw/main/zipper.ps1 | iex }`nfunction sysinfo { irm https://github.com/therepos/windows/raw/main/systeminfo.ps1 | iex }"
```

Remove them:

```powershell
(Get-Content $PROFILE) | Where-Object { $_ -notmatch '^function (zipper|sysinfo) ' } | Set-Content $PROFILE
```

Kept per Windows user, per PC.
