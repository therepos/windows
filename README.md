## Usage

- `.ps1`: run the bootstrap on line 1 in PowerShell.
- `.bat` / `.cmd`: download and double-click.
- `.py`: run with Python 3.

## Structure

```
windows
├── AGENTS.md               # rules for AI agents
├── CLAUDE.md               # @AGENTS.md
├── .github
│   └── workflows
├── apps                    # one script per tool, one-word names
└── scripts
    ├── sync-blocks.sh      # standardise scripts (say, menukey, downloads)
    └── lint.sh             # diagnose and repair scripts
```
