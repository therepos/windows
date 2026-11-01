#!/usr/bin/env python3
"""
zipper.py
---------
Compress files/folders into a password-protected 7z archive,
or extract a password-protected 7z archive.
Auto-installs 7-Zip if not present (Windows / macOS / Linux).

ZIP    : output YYYYMMDD-HHMMSS (no extension)
UNZIP  : output <archive>-extracted
Password and output folder prompt with defaults (Enter = default
password 'Zip@2026!'; Downloads).

Usage
  python zipper.py                              (interactive — asks zip or unzip)
  python zipper.py --zip   <file_or_folder> ...
  python zipper.py --unzip <archive> [output_dir]
  python zipper.py --password <pw> --zip <file_or_folder> ...
  python zipper.py <file_or_folder> ...         (legacy: defaults to zip)

Interactive mode on Windows docks this window to the right edge (on top)
and Explorer (open window or Downloads) to the left. Several files can be
dropped at once. Command-line zips are saved next to the script.

Flags
  --zip / -z              compress mode
  --unzip / -u / -x       extract mode
  --password <pw> / -p    skip the password prompt and use this value
"""

import sys
import os
import subprocess
import shutil
import platform
from datetime import datetime
from pathlib import Path


DEFAULT_PASSWORD = "Zip@2026!"

# Standard install locations 7-Zip uses on Windows
WIN_PATHS = [
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe",
]


# ──────────────────────────────────────────────
# 1.  FIND 7-ZIP
# ──────────────────────────────────────────────

def find_7z() -> str | None:
    """Return path to 7z executable, or None if not found."""
    for candidate in ("7z", "7za", "7zz"):
        found = shutil.which(candidate)
        if found:
            return found
    for p in WIN_PATHS:
        if os.path.isfile(p):
            return p
    return None


# ──────────────────────────────────────────────
# 2.  AUTO-INSTALL HELPERS
# ──────────────────────────────────────────────

def _install_windows() -> str:
    """Install via winget (Win 10/11). Falls back to direct MSI download."""
    if shutil.which("winget"):
        print("[*] Installing 7-Zip via winget ...")
        result = subprocess.run(
            ["winget", "install", "-e", "--id", "7zip.7zip",
             "--accept-package-agreements", "--accept-source-agreements"],
            capture_output=False,
            text=True,
        )
        if result.returncode == 0:
            os.environ["PATH"] += r";C:\Program Files\7-Zip"
            found = find_7z()
            if found:
                return found

    # Fallback: download MSI directly
    import urllib.request, tempfile
    print("[*] winget unavailable - downloading 7-Zip MSI installer ...")
    url = "https://www.7-zip.org/a/7z2409-x64.msi"
    tmp = os.path.join(tempfile.gettempdir(), "7zip_installer.msi")
    urllib.request.urlretrieve(url, tmp)
    print("[*] Running silent install ...")
    subprocess.run(["msiexec", "/i", tmp, "/quiet", "/norestart"], check=True)
    os.environ["PATH"] += r";C:\Program Files\7-Zip"
    found = find_7z()
    if found:
        return found
    raise RuntimeError("7-Zip installation finished but executable still not found.")


def _install_macos() -> str:
    """Install via Homebrew."""
    if not shutil.which("brew"):
        raise RuntimeError(
            "Homebrew is not installed. Install it from https://brew.sh then re-run."
        )
    print("[*] Installing p7zip via Homebrew ...")
    subprocess.run(["brew", "install", "p7zip"], check=True, capture_output=False)
    found = find_7z()
    if found:
        return found
    raise RuntimeError("p7zip installed but executable not found.")


def _install_linux() -> str:
    """Try apt, dnf, pacman, zypper in order."""
    managers = [
        (["apt-get", "install", "-y", "p7zip-full"],             "apt-get"),
        (["dnf",     "install", "-y", "p7zip", "p7zip-plugins"], "dnf"),
        (["pacman",  "-S", "--noconfirm", "p7zip"],              "pacman"),
        (["zypper",  "install", "-y", "p7zip"],                  "zypper"),
    ]
    for cmd, name in managers:
        if shutil.which(cmd[0]):
            print(f"[*] Installing p7zip via {name} ...")
            subprocess.run(["sudo"] + cmd, check=True, capture_output=False)
            found = find_7z()
            if found:
                return found
            raise RuntimeError(f"{name} finished but executable not found.")
    raise RuntimeError(
        "No supported package manager found (apt, dnf, pacman, zypper). "
        "Please install p7zip manually and re-run."
    )


def ensure_7z() -> str:
    """Return path to 7z, auto-installing if necessary."""
    found = find_7z()
    if found:
        return found

    system = platform.system()
    print(f"[!] 7-Zip not found. Auto-installing for {system} ...\n")

    if system == "Windows":
        return _install_windows()
    elif system == "Darwin":
        return _install_macos()
    elif system == "Linux":
        return _install_linux()
    else:
        raise RuntimeError(
            f"Unsupported OS '{system}'. Install 7-Zip manually and re-run."
        )


# ──────────────────────────────────────────────
# 3.  INPUT HELPERS
# ──────────────────────────────────────────────

def ask(prompt: str) -> str:
    """input() that exits quietly on Ctrl+C / Ctrl+Z."""
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print()
        sys.exit(0)


def clean_path(raw: str) -> Path:
    return Path(raw.strip().strip('"').strip("'")).expanduser().resolve()


def split_dropped(line: str) -> list[str]:
    """
    Dropping several files into a console pastes them on one line,
    space separated, with quotes only around paths containing spaces.
    A line that is itself an existing path is kept whole.
    """
    line = line.strip()
    if clean_path(line).exists():
        return [line]
    import re
    return [a or b for a, b in re.findall(r'"([^"]+)"|(\S+)', line)]


def collect_targets(raw_paths: list[str]) -> list[Path]:
    """Resolve CLI paths, skip duplicates, exit if any is missing."""
    targets: list[Path] = []
    missing = []
    for raw in raw_paths:
        p = clean_path(raw)
        if not p.exists():
            missing.append(p)
        elif p not in targets:
            targets.append(p)
    if missing:
        for p in missing:
            print(f"Not found: {p}")
        sys.exit(1)
    return targets


def find_archive(raw: str) -> Path | None:
    """Accept the bare archive name or name.7z."""
    p = clean_path(raw)
    for c in (p, Path(str(p) + ".7z")):
        if c.is_file():
            return c
    return None


def prompt_paths(*, archives: bool = False) -> list[Path]:
    """Collect dropped items until Enter on an empty line."""
    items: list[Path] = []
    print("Drop archives here" if archives else "Drop files/folders here")
    while True:
        line = ask("> ")
        if not line:
            if items:
                return items
            continue
        for raw in split_dropped(line):
            p = find_archive(raw) if archives else clean_path(raw)
            if p is None or not p.exists():
                print(f"  ! not found: {raw}")
            elif p not in items:
                items.append(p)
                print(f"  + {p.name}{os.sep if p.is_dir() else ''}")
        print("Drop more, or Enter to continue")


def prompt_password(*, confirm: bool = False) -> str:
    """Hidden input; Enter = DEFAULT_PASSWORD. Custom ones are confirmed."""
    import getpass
    while True:
        try:
            pw = getpass.getpass("Password (Enter = default): ")
            if not pw:
                return DEFAULT_PASSWORD
            if not confirm or getpass.getpass("Confirm: ") == pw:
                return pw
        except (EOFError, KeyboardInterrupt):
            print()
            sys.exit(0)
        print("  ! passwords differ, try again")


def downloads_dir() -> Path:
    d = Path.home() / "Downloads"
    return d if d.is_dir() else Path.home()


def prompt_folder(default: Path) -> Path:
    raw = ask(f"Save to (Enter = {default}): ")
    return clean_path(raw) if raw else default


# ──────────────────────────────────────────────
# 4.  RUN 7-ZIP
# ──────────────────────────────────────────────

def run_7z(cmd: list[str]) -> tuple[int, str]:
    """Run 7z quietly; return (exit code, its most useful error line)."""
    r = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
    lines = [l.strip() for l in (r.stderr + "\n" + r.stdout).splitlines() if l.strip()]
    err = next((l for l in lines if "password" in l.lower()), None) \
        or next((l for l in lines if "cannot" in l.lower()), None) \
        or next((l for l in lines if l.upper().startswith("ERROR")), None) \
        or (lines[-1] if lines else "")
    return r.returncode, err


def make_archive(targets: list[Path], seven_z: str, password: str,
                 out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    output = out_dir / datetime.now().strftime("%Y%m%d-%H%M%S")   # no extension
    cmd = [
        seven_z, "a",
        "-t7z",            # 7z format
        f"-p{password}",
        "-mhe=on",         # encrypt file headers too
        "-mx=5",           # compression level 0-9
        str(output),       # destination (no .7z suffix -- intentional)
        *[str(t) for t in targets],
    ]
    print("Zipping...")
    rc, err = run_7z(cmd)
    if rc != 0:
        raise RuntimeError(err or f"7z exit code {rc}")
    # Some 7z builds silently append .7z -- strip it back off
    with_ext = Path(str(output) + ".7z")
    if with_ext.exists() and not output.exists():
        with_ext.rename(output)
    return output


def _remove_if_empty(folder: Path) -> None:
    try:
        folder.rmdir()
    except OSError:
        pass


def extract_archive(archive: Path, seven_z: str, password: str,
                    out_dir: Path, *, retry: bool = True) -> str:
    """
    Extract into out_dir; on a wrong password re-prompt up to 3 times.
    Returns the password that worked so the next archive can reuse it.
    """
    import getpass
    out_dir.mkdir(parents=True, exist_ok=True)
    for attempt in range(1, 5):
        print(f"Extracting {archive.name}...", end="", flush=True)
        rc, err = run_7z([seven_z, "x", f"-p{password}", "-y",
                          f"-o{out_dir}", str(archive)])
        if rc == 0:
            print("  done")
            return password
        print()
        if "password" not in err.lower() or not retry or attempt == 4:
            _remove_if_empty(out_dir)
            raise RuntimeError(err or f"7z exit code {rc}")
        try:
            password = getpass.getpass(f"Wrong password, try again ({attempt}/3): ")
        except (EOFError, KeyboardInterrupt):
            password = ""
        if not password:
            _remove_if_empty(out_dir)
            raise RuntimeError("aborted")
    return password


# ──────────────────────────────────────────────
# 5.  WINDOW LAYOUT  (Windows only, fails silently)
#     Zipper docked right and on top, Explorer on the left.
# ──────────────────────────────────────────────

def _win_layout() -> None:
    import ctypes
    import time
    from ctypes import wintypes

    user32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
    k32.GetConsoleWindow.restype = wintypes.HWND
    user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    user32.GetAncestor.restype = wintypes.HWND
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND,
                                    ctypes.c_int, ctypes.c_int,
                                    ctypes.c_int, ctypes.c_int, wintypes.UINT]
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.IsWindowVisible.argtypes = [wintypes.HWND]

    area = wintypes.RECT()
    user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(area), 0)  # work area
    height = area.bottom - area.top
    width = max(560, (area.right - area.left) * 30 // 100)
    SWP_SHOWWINDOW, SW_RESTORE = 0x0040, 9

    # Zipper: right edge, full height, on top. Under Windows Terminal the
    # console is a hidden pseudo window owned by the real terminal window.
    console = k32.GetConsoleWindow()
    if console:
        console = user32.GetAncestor(console, 3) or console     # GA_ROOTOWNER
        user32.ShowWindow(console, SW_RESTORE)
        user32.SetWindowPos(console, wintypes.HWND(-1),          # HWND_TOPMOST
                            area.right - width, area.top, width, height,
                            SWP_SHOWWINDOW)

    # Explorer: reuse the topmost open window, else open Downloads.
    def explorer_windows() -> list[int]:
        found: list[int] = []
        proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def cb(h, _):
            buf = ctypes.create_unicode_buffer(32)
            user32.GetClassNameW(h, buf, 32)
            if buf.value == "CabinetWClass" and user32.IsWindowVisible(h):
                found.append(h)
            return True
        user32.EnumWindows(proc(cb), 0)
        return found                     # z-order, topmost first

    windows = explorer_windows()
    if not windows:
        subprocess.Popen(["explorer", str(downloads_dir())])
        for _ in range(30):              # wait up to ~3 s for it to appear
            time.sleep(0.1)
            windows = explorer_windows()
            if windows:
                break
    if windows:
        user32.ShowWindow(windows[0], SW_RESTORE)
        user32.SetWindowPos(windows[0], None, area.left, area.top,
                            area.right - area.left - width, height,
                            SWP_SHOWWINDOW)
    if console:
        user32.SetForegroundWindow(console)


def arrange_windows() -> None:
    if platform.system() != "Windows":
        return
    try:
        _win_layout()
    except Exception:
        pass


# ──────────────────────────────────────────────
# 6.  INTERACTIVE MODE
# ──────────────────────────────────────────────

def interactive(seven_z: str) -> None:
    print("ZIPPER")
    print("  1) Zip")
    print("  2) Unzip")
    while True:
        choice = ask("> ").lower()
        if choice in ("1", "z", "zip", "2", "u", "unzip"):
            break
        print("Choose 1 or 2")
    print()

    try:
        if choice in ("1", "z", "zip"):
            targets = prompt_paths()
            password = prompt_password(confirm=True)
            out_dir = prompt_folder(downloads_dir())
            print()
            out = make_archive(targets, seven_z, password, out_dir)
            print(f"Done: {out}")
        else:
            archives = prompt_paths(archives=True)
            password = prompt_password()
            out_dir = prompt_folder(downloads_dir())
            print()
            ok = 0
            for arc in archives:
                try:
                    password = extract_archive(arc, seven_z, password,
                                               out_dir / f"{arc.stem}-extracted")
                    ok += 1
                except RuntimeError as e:
                    print(f"  ! skipped: {e}")
            print(f"\nDone: {ok} of {len(archives)} -> {out_dir}")
    except RuntimeError as e:
        print(f"Failed: {e}")
    ask("\nEnter to close")


# ──────────────────────────────────────────────
# 7.  ENTRY POINT
# ──────────────────────────────────────────────

def main():
    args = sys.argv[1:]
    mode: str | None = None
    cli_password: str | None = None

    # --- Extract --password / -p from anywhere in args ---
    cleaned: list[str] = []
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("--password", "-p"):
            if i + 1 >= len(args):
                print("--password requires a value.")
                sys.exit(1)
            cli_password = args[i + 1]
            i += 2
            continue
        if a.startswith("--password="):
            cli_password = a.split("=", 1)[1]
            i += 1
            continue
        cleaned.append(a)
        i += 1
    args = cleaned

    # --- Parse mode flag if present ---
    if args and args[0] in ("--zip", "-z"):
        mode = "zip"
        args = args[1:]
    elif args and args[0] in ("--unzip", "-u", "--extract", "-x"):
        mode = "unzip"
        args = args[1:]
    elif args and args[0] in ("--help", "-h"):
        print(__doc__)
        sys.exit(0)

    # --- No mode flag but paths given: zip them (legacy behaviour) ---
    if mode is None and args:
        mode = "zip"

    if mode is None:
        arrange_windows()

    try:
        seven_z = ensure_7z()
    except RuntimeError as e:
        print(f"Could not obtain 7-Zip: {e}")
        sys.exit(1)

    if mode is None:
        interactive(seven_z)
        return

    # --- Command-line mode: output next to the script unless given ---
    script_dir = Path(__file__).resolve().parent
    try:
        if mode == "zip":
            targets = collect_targets(args) if args else prompt_paths()
            password = cli_password or prompt_password(confirm=True)
            out = make_archive(targets, seven_z, password, script_dir)
        else:
            archive = find_archive(args[0]) if args else prompt_paths(archives=True)[0]
            if archive is None:
                print(f"Not found: {args[0]}")
                sys.exit(1)
            out = clean_path(args[1]) if len(args) >= 2 else \
                archive.parent / f"{archive.stem}-extracted"
            password = cli_password or prompt_password()
            extract_archive(archive, seven_z, password, out,
                            retry=cli_password is None)
        print(f"Done: {out}")
    except RuntimeError as e:
        print(f"Failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
