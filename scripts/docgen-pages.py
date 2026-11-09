#!/usr/bin/env python3
# Purpose: Build the GitHub Pages script page from each script's lines 1-2
# =============================================================================
# Usage:
#   python3 scripts/docgen-pages.py [out_dir]   write index.html and downloadable files (default: _site)
#   python3 scripts/docgen-pages.py --check     list scripts with no display name in NAMES
# =============================================================================

import html
import os
import shutil
import sys

OUT_DIR = sys.argv[1] if len(sys.argv) > 1 else '_site'
OUTPUT_HTML = os.path.join(OUT_DIR, 'index.html')

KINDS = {'.ps1': 'PowerShell', '.bat': 'Batch', '.cmd': 'Batch', '.py': 'Python'}
REPO_URL = 'https://github.com/therepos/windows'

# Display name per script: Title Case words, acronyms in caps (lint fails on a missing entry)
NAMES = {
    'addinexport': 'Add-in Export',
    'addinsetup': 'Add-in Setup',
    'agentprofiler': 'Agent Profiler',
    'appupdate': 'App Update',
    'checkenv': 'Check Env',
    'displayswitch': 'Display Switch',
    'extractfiles': 'Extract Files',
    'filefind': 'File Find',
    'filemap': 'File Map',
    'filesync': 'File Sync',
    'hibernate': 'Hibernate',
    'kbbundle': 'KB Bundle',
    'kbprofile': 'KB Profile',
    'kbrcm': 'KB RCM',
    'lagcheck': 'Lag Check',
    'oleextract': 'OLE Extract',
    'scanfiles': 'Scan Files',
    'smbcache': 'SMB Cache',
    'systeminfo': 'System Info',
    'usbclean': 'USB Clean',
    'win11postprocess': 'Win11 Post-Process',
    'zipper': 'Zipper',
}

def scripts():
    return [f for f in sorted(os.listdir('apps')) if os.path.splitext(f)[1] in KINDS]

def header(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        return [f.readline().strip() for _ in range(2)]

def purpose_of(line):
    text = line.split('Purpose:', 1)[1].strip() if 'Purpose:' in line else ''
    if text and not text.endswith('.'):
        text += '.'
    return text

def row(name, purpose, button):
    return '<div class="row"><b>{}</b><p>{}</p>{}</div>'.format(html.escape(name), html.escape(purpose), button)

def html_rows(run_by_command):
    rows = []
    for filename in scripts():
        stem, ext = os.path.splitext(filename)
        l1, l2 = header(os.path.join('apps', filename))
        name = NAMES.get(stem, stem[0].upper() + stem[1:])
        purpose = purpose_of(l2)
        is_cmd = l1.startswith('# irm ')
        if is_cmd != run_by_command:
            continue
        if is_cmd:
            button = '<button data-cmd="{}">Copy command</button>'.format(html.escape(l1[2:], quote=True))
        else:
            # served from the page itself, so the browser downloads it instead of showing text
            os.makedirs(os.path.join(OUT_DIR, 'apps'), exist_ok=True)
            shutil.copy(os.path.join('apps', filename), os.path.join(OUT_DIR, 'apps', filename))
            button = '<a class="btn" href="apps/{}" download>Download {}</a>'.format(html.escape(filename, quote=True), KINDS[ext])
        rows.append(row(name, purpose, button))
    return '\n'.join(rows)

PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
:root{--bg:#f6f6f4;--card:#fff;--text:#1f1f1e;--muted:#6b6a66;--line:#e2e1dc;--ok:#2e7d32}
@media (prefers-color-scheme:dark){:root{--bg:#1c1c1b;--card:#262624;--text:#ecebe6;--muted:#a3a29c;--line:#3a3a37;--ok:#7bc47f}}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:960px;margin:0 auto;padding:2rem 1rem}
h1{font-size:22px;font-weight:600;margin:0 0 .25rem}
h2{font-size:13px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin:2rem 0 .5rem}
.hint{color:var(--muted);margin:0 0 1rem}
input{width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-size:15px}
.list{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.row{display:grid;grid-template-columns:180px 1fr 170px;gap:4px 16px;align-items:center;padding:10px 14px;border-top:1px solid var(--line)}
.row:first-child{border-top:0}
.row b{font-weight:600}
.row p{margin:0;color:var(--muted);font-size:14px}
button,.btn{box-sizing:border-box;width:100%;padding:7px;text-align:center;text-decoration:none;display:block;border:1px solid var(--line);border-radius:8px;background:transparent;color:var(--text);font-size:14px;cursor:pointer}
button:hover,.btn:hover{background:var(--bg)}
button.done{color:var(--ok);border-color:var(--ok)}
footer{margin-top:2.5rem;padding-top:1rem;border-top:1px solid var(--line);color:var(--muted);font-size:14px}
footer a{color:inherit}
@media (max-width:640px){.row{grid-template-columns:1fr 170px}.row p{grid-column:1/-1;grid-row:2}}
</style>
</head>
<body>
<main>
<h1>{title}</h1>
<p class="hint">{hint}</p>
<input id="q" type="search" placeholder="Search scripts" aria-label="Search scripts">
{sections}
<footer><a href="{repo}">{repo_label}</a></footer>
</main>
<script>
function copy(text){
  if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
  var t=document.createElement('textarea'); t.value=text; document.body.appendChild(t); t.select();
  document.execCommand('copy'); document.body.removeChild(t); return Promise.resolve();
}
document.querySelectorAll('button[data-cmd]').forEach(function(b){
  b.addEventListener('click',function(){
    copy(b.dataset.cmd).then(function(){
      b.textContent='Copied'; b.classList.add('done');
      setTimeout(function(){ b.textContent='Copy command'; b.classList.remove('done'); },1800);
    });
  });
});
document.getElementById('q').addEventListener('input',function(e){
  var q=e.target.value.toLowerCase();
  document.querySelectorAll('section').forEach(function(s){
    var shown=0;
    s.querySelectorAll('.row').forEach(function(r){
      var hit=r.textContent.toLowerCase().indexOf(q)>-1;
      r.style.display=hit?'':'none'; if(hit) shown++;
    });
    s.style.display=shown?'':'none';
  });
});
</script>
</body>
</html>
"""

def section(title, rows):
    return '<section><h2>{}</h2><div class="list">\n{}\n</div></section>'.format(title, rows) if rows else ''

def generate_html():
    page = (PAGE.replace('{title}', 'Windows Scripts')
                .replace('{hint}', 'Copy command: paste into PowerShell and press Enter. '
                                   'Download: double-click the file (Python files need Python 3).')
                .replace('{sections}', section('Run in PowerShell', html_rows(True)) + '\n'
                                       + section('Download and Run', html_rows(False)))
                .replace('{repo_label}', REPO_URL.split('//', 1)[1])
                .replace('{repo}', REPO_URL))
    with open(OUTPUT_HTML, 'w') as out:
        out.write(page)

def check():
    missing = [f for f in scripts() if os.path.splitext(f)[0] not in NAMES]
    for f in missing:
        print('apps/{}: no display name in NAMES (scripts/docgen-pages.py)'.format(f))
    return 1 if missing else 0

if __name__ == "__main__":
    if OUT_DIR == '--check':
        sys.exit(check())
    os.makedirs(OUT_DIR, exist_ok=True)
    generate_html()
