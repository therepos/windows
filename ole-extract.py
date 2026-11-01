import importlib, subprocess, sys, os

def ensure(pkg, import_name=None):
    try:
        return importlib.import_module(import_name or pkg)
    except ImportError:
        pip = [sys.executable, "-m", "pip", "install"]
        for flags in ([], ["--user"], ["--break-system-packages"]):
            try:
                subprocess.check_call(pip + flags + [pkg])
                return importlib.import_module(import_name or pkg)
            except Exception:
                continue
        raise ImportError(f"Could not install {pkg}")

import zipfile, io, struct, email
from email import policy
olefile = None   # loaded on first use

# ---------- type detection ----------
def zip_kind(b):
    try:
        names = zipfile.ZipFile(io.BytesIO(b)).namelist()
        if any(n.startswith("xl/")   for n in names): return "xlsx"
        if any(n.startswith("word/") for n in names): return "docx"
        if any(n.startswith("ppt/")  for n in names): return "pptx"
    except Exception:
        pass
    return "zip"

def ole_kind(b):
    try:
        o = olefile.OleFileIO(io.BytesIO(b))
        s = {"/".join(x).lower() for x in o.listdir()}
        o.close()
        if "workbook" in s or "book" in s:    return "xls"
        if "worddocument" in s:               return "doc"
        if "powerpoint document" in s:        return "ppt"
        if "__properties_version1.0" in s or any(x.startswith("__substg1.0_") for x in s):
            return "msg"
    except Exception:
        pass
    return "ole"

def sniff(b):
    if b[:4] == b"%PDF":             return "pdf"
    if b[:4] == b"PK\x03\x04":       return zip_kind(b)
    if b[:4] == b"\xd0\xcf\x11\xe0": return ole_kind(b)
    return "bin"

def trim(b, kind):
    if kind == "pdf":
        i = b.rfind(b"%%EOF"); return b[:i+5] if i != -1 else b
    if kind in ("xlsx", "docx", "pptx", "zip"):
        i = b.rfind(b"PK\x05\x06")
        if i != -1 and i + 22 <= len(b):
            clen = struct.unpack_from("<H", b, i + 20)[0]
            return b[:i + 22 + clen]
    return b

def from_packager(stream):
    try:
        p = 6
        p = stream.index(b"\x00", p) + 1
        e = stream.index(b"\x00", p); orig = stream[p:e]; p = e + 1
        p += 4
        tlen = struct.unpack_from("<I", stream, p)[0]; p += 4 + tlen
        dsize = struct.unpack_from("<I", stream, p)[0]; p += 4
        payload = stream[p:p + dsize]
        if len(payload) == dsize and sniff(payload) != "bin":
            return os.path.basename(orig.decode("latin-1", "replace").replace("\\", "/")), payload
    except Exception:
        pass
    idx = min([i for i in (stream.find(b"%PDF"), stream.find(b"PK\x03\x04"),
                           stream.find(b"\xd0\xcf\x11\xe0")) if i != -1], default=-1)
    return None, (stream[idx:] if idx != -1 else stream)

def safe_name(name):
    name = name.replace("\\", "/").split("/")[-1]
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, "_")
    name = "".join(c for c in name if ord(c) >= 32).strip().rstrip(". ")
    return name or "attachment"

def save(out, name, data):
    name = safe_name(name)
    path = os.path.join(out, name); n = 1
    while os.path.exists(path):
        stem, ext = os.path.splitext(name)
        path = os.path.join(out, f"{stem}_{n}{ext}"); n += 1
    open(path, "wb").write(data)
    return os.path.basename(path)

# ---------- Office containers (.xlsx/.docx/.pptx) ----------
def extract_office_embeddings(src, out):
    names = []
    with zipfile.ZipFile(src) as z:
        embeds = [n for n in z.namelist() if "embeddings/" in n and not n.endswith("/")]
        for n in embeds:
            raw = z.read(n)
            base = os.path.basename(n)              # e.g. oleObject1.bin
            stem = os.path.splitext(base)[0]        # oleObject1

            if not base.endswith(".bin"):           # already a real file
                names.append(save(out, base, raw)); continue

            ole = olefile.OleFileIO(io.BytesIO(raw))
            streams = {"/".join(s) for s in ole.listdir()}
            low = {s.lower() for s in streams}

            native = ("xls" if ("workbook" in low or "book" in low) else
                      "doc" if "worddocument" in low else
                      "ppt" if "powerpoint document" in low else
                      "msg" if "__properties_version1.0" in low or any(s.startswith("__substg1.0_") for s in streams) else None)
            if native:
                ole.close(); names.append(save(out, f"{stem}.{native}", raw)); continue

            sn = next((s for s in ("CONTENTS", "Package", "\x01Ole10Native", "Ole10Native") if s in streams), None)
            if not sn:
                ole.close(); names.append(save(out, f"{stem}.bin", raw)); continue
            payload = ole.openstream(sn).read(); ole.close()

            oname = None
            if "Ole10Native" in sn:
                oname, payload = from_packager(payload)   # recover true original name if present
            kind = sniff(payload)
            payload = trim(payload, kind)
            names.append(save(out, oname or f"{stem}.{kind}", payload))
    return names

# ---------- Outlook .msg (MS-OXMSG, parsed with olefile) ----------
def extract_msg_attachments(src, out):
    ole = olefile.OleFileIO(src)
    attach_dirs = sorted({e[0] for e in ole.listdir(streams=True, storages=False)
                          if e and e[0].startswith("__attach_version1.0_")})
    names = []
    for d in attach_dirs:
        def rd(tagtype):
            p = [d, f"__substg1.0_{tagtype}"]
            return ole.openstream(p).read() if ole.exists(p) else None
        data = rd("37010102")                      # PidTagAttachDataBinary
        if not data:
            continue                               # embedded message / no file bytes -> skip
        fn = None
        for tt in ("3707001F", "3704001F"):        # long, then short filename (Unicode)
            raw = rd(tt)
            if raw:
                fn = raw.decode("utf-16-le", "replace").rstrip("\x00")
                if fn:
                    break
        if not fn:
            for tt in ("3707001E", "3704001E"):    # ASCII fallbacks
                raw = rd(tt)
                if raw:
                    fn = raw.decode("latin-1", "replace").rstrip("\x00")
                    if fn:
                        break
        names.append(save(out, fn or "attachment", data))
    ole.close()
    return names

# ---------- .eml (MIME, stdlib) ----------
def extract_eml_attachments(src, out):
    with open(src, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    names = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        fn = part.get_filename()
        if fn or part.get_content_disposition() == "attachment":
            payload = part.get_payload(decode=True)
            if payload:
                names.append(save(out, fn or "attachment", payload))
    return names

# ---------- dispatcher ----------
def extract_any(src):
    out = os.path.join(os.path.dirname(os.path.abspath(src)), "extracted")
    os.makedirs(out, exist_ok=True)
    ext = os.path.splitext(src)[1].lower()
    if ext == ".eml":
        return out, extract_eml_attachments(src, out)
    if ext == ".msg":
        return out, extract_msg_attachments(src, out)
    return out, extract_office_embeddings(src, out)

# ---------- CLI ----------
def clean_path(p):
    p = p.strip()
    if len(p) >= 2 and p[0] == p[-1] and p[0] in "\"'":
        p = p[1:-1]
    return p.strip()

def handle(path):
    path = clean_path(path)
    if not os.path.isfile(path):
        print("  ! File not found:", path); return
    try:
        out, names = extract_any(path)
        if names:
            print(f"  Extracted {len(names)} file(s) to: {out}")
            for nm in names:
                print("     -", nm)
        else:
            print("  Nothing to extract (no embedded files or attachments found).")
    except zipfile.BadZipFile:
        print("  ! Unreadable file (expected .xlsx/.docx/.pptx/.msg/.eml).")
    except Exception as e:
        print("  ! Error:", e)
    print()

def main():
    global olefile
    olefile = ensure("olefile")
    if len(sys.argv) > 1:                 # file dragged onto the .py icon
        handle(sys.argv[1])
        input("Press Enter to exit...")
    else:
        print("Drag a file into this window and press Enter.")
        print("Accepts: .xlsx .docx .pptx .msg .eml")
        print("(leave blank and press Enter to quit)\n")
        while True:
            p = input("File: ").strip()
            if not p:
                break
            handle(p)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Fatal error:", e)
        input("Press Enter to exit...")