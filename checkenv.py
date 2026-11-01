# Purpose: check a Windows PC can run the offline text extractor (Python, libraries, .doc converters, long paths, share access).
#   python checkenv.py                       # report only
#   python checkenv.py --install             # also pip-install any missing Python libraries
#   python checkenv.py --share \\server\share  # also test reading the share and a long path
import sys, os, subprocess, shutil, importlib, platform, ctypes, glob

OK, BAD, WARN = "  ok   ", "  FAIL ", "  warn "
results = []
def rep(status, what, detail=""):
    results.append(status); print(f"{status}{what:42}{detail}")

# 1. Python itself
v = sys.version_info
rep(OK if v >= (3, 9) else BAD, "Python >= 3.9", f"{platform.python_version()}  {sys.executable}")
rep(OK if sys.maxsize > 2**32 else WARN, "64-bit Python", "yes" if sys.maxsize > 2**32 else "32-bit: memory limit ~2 GB, still usable")

# 2. pip and PyPI
try:
    subprocess.run([sys.executable, "-m", "pip", "--version"], check=True, capture_output=True)
    rep(OK, "pip available")
    r = subprocess.run([sys.executable, "-m", "pip", "download", "-q", "--no-deps", "-d", os.devnull if os.name != "nt" else os.environ.get("TEMP", "."), "six"],
                       capture_output=True, text=True, timeout=60)
    rep(OK if r.returncode == 0 else WARN, "PyPI reachable (pip install works)", "" if r.returncode == 0 else "no: libraries must be brought in as wheel files (pip download on another PC)")
except Exception as e:
    rep(BAD, "pip available", str(e)[:80])

# 3. Python libraries: (import name, pip name, used for)
LIBS = [
    ("docx",        "python-docx", ".docx text and tables"),
    ("openpyxl",    "openpyxl",    ".xlsx/.xlsm"),
    ("xlrd",        "xlrd",        ".xls (old Excel)"),
    ("pptx",        "python-pptx", ".pptx"),
    ("pypdf",       "pypdf",       ".pdf text"),
    ("extract_msg", "extract-msg", ".msg Outlook emails"),
    ("olefile",     "olefile",     ".doc/.ppt container access"),
    ("win32com",    "pywin32",     "Word/Excel automation for .doc/.ppt (optional)"),
]
missing = []
for mod, pipname, use in LIBS:
    try:
        importlib.import_module(mod); rep(OK, f"lib {pipname}", use)
    except Exception:
        missing.append(pipname); rep(WARN if mod == "win32com" else BAD, f"lib {pipname}", f"missing  ({use})")

if "--install" in sys.argv and missing:
    print("\ninstalling:", " ".join(missing))
    subprocess.run([sys.executable, "-m", "pip", "install", *missing])
    print("re-run checkenv.py to confirm.\n")

# 4. Converters for legacy .doc / .ppt
word = False
try:
    import win32com.client
    w = win32com.client.Dispatch("Word.Application"); w.Quit(); word = True
except Exception:
    pass
rep(OK if word else WARN, "Microsoft Word via COM", "yes" if word else "not available")
soffice = shutil.which("soffice") or next(iter(glob.glob(r"C:\Program Files*\LibreOffice*\program\soffice.exe")), None)
rep(OK if soffice else WARN, "LibreOffice (soffice)", soffice or "not found")
if not (word or soffice):
    rep(BAD, ".doc/.ppt conversion", "need Word or LibreOffice; without either, .doc files get a crude text-only fallback")

# 5. Long paths (>260 chars)
try:
    import winreg
    k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem")
    lp = winreg.QueryValueEx(k, "LongPathsEnabled")[0]
    rep(OK if lp == 1 else WARN, "LongPathsEnabled in registry", str(lp) + ("" if lp == 1 else "  (extractor will use \\\\?\\ prefix instead)"))
except Exception as e:
    rep(WARN, "LongPathsEnabled in registry", f"cannot read: {e}")

# 6. Share access, optional
if "--share" in sys.argv:
    share = sys.argv[sys.argv.index("--share") + 1].rstrip("\\")
    rep(OK if os.path.isdir(share) else BAD, "share reachable", share)
    if os.path.isdir(share):
        longest = ""
        for root, dirs, files in os.walk(share):
            for f in files:
                p = os.path.join(root, f)
                if len(p) > len(longest): longest = p
            if len(longest) > 260: break
        if longest:
            unc = "\\\\?\\UNC\\" + longest[2:]
            try:
                with open(unc, "rb") as fh: fh.read(16)
                rep(OK, f"open long path ({len(longest)} chars) with \\\\?\\ prefix")
            except Exception as e:
                rep(BAD, f"open long path ({len(longest)} chars)", str(e)[:80])

# 7. Disk space where the database will go
free = shutil.disk_usage(os.getcwd()).free / 1e9
rep(OK if free > 20 else WARN, "free disk here", f"{free:.0f} GB  (need ~10-20 GB for the text database)")

print()
if BAD in results: print("Result: FAIL items above must be fixed before the extractor will run properly.")
elif WARN in results: print("Result: usable. warn items reduce coverage or speed but do not block.")
else: print("Result: all good.")
