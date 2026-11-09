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
    'oleextract': 'OLE Extract',
    'scanfiles': 'Scan Files',
    'smbcache': 'SMB Cache',
    'systeminfo': 'System Info',
    'usbclean': 'USB Clean',
    'win11postprocess': 'Win11 Post-Process',
    'zipper': 'Zipper',
}

# Inline icons (Lucide copy/download/check, GitHub mark), stroke/fill from currentColor
ICON_COPY = '<svg viewBox="0 0 24 24" aria-hidden="true"><rect width="14" height="14" x="8" y="8" rx="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg>'
ICON_DOWNLOAD = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>'
ICON_CHECK = '<svg viewBox="0 0 24 24" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>'
ICON_GITHUB = ('<svg class="gh" viewBox="0 0 16 16" aria-hidden="true"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59'
               '.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53'
               '.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15'
               '-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92'
               '.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38'
               'A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>')

def actions(filename, cmd):
    # the file is served from the page itself, so the browser downloads it instead of showing text
    os.makedirs(os.path.join(OUT_DIR, 'apps'), exist_ok=True)
    shutil.copy(os.path.join('apps', filename), os.path.join(OUT_DIR, 'apps', filename))
    copy = ('<button class="ic" data-cmd="{}" title="Copy command" aria-label="Copy command">{}</button>'.format(
        html.escape(cmd, quote=True), ICON_COPY) if cmd else '<span class="ic blank"></span>')
    label = 'Download ' + os.path.splitext(filename)[1]
    download = '<a class="ic" href="apps/{0}" download title="{1}" aria-label="{1}">{2}</a>'.format(
        html.escape(filename, quote=True), label, ICON_DOWNLOAD)
    return '<div class="act">{}{}</div>'.format(copy, download)

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
        button = actions(filename, l1[2:] if is_cmd else '')
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
.row{display:grid;grid-template-columns:180px 1fr auto;gap:4px 16px;align-items:center;padding:10px 14px;border-top:1px solid var(--line)}
.row:first-child{border-top:0}
.row b{font-weight:600}
.row p{margin:0;color:var(--muted);font-size:14px}
.act{display:flex;gap:6px}
.ic{box-sizing:border-box;width:36px;height:36px;display:grid;place-items:center;padding:0;border:1px solid var(--line);border-radius:8px;background:transparent;color:var(--text);cursor:pointer}
.ic:hover{background:var(--bg)}
.ic.blank{border:0;cursor:default}
.ic.done{color:var(--ok);border-color:var(--ok)}
.ic svg,.hint svg{width:18px;height:18px;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.hint svg{width:15px;height:15px;vertical-align:-2px}
footer{margin-top:2.5rem;padding-top:1rem;border-top:1px solid var(--line);display:flex;justify-content:center}
footer a{display:inline-flex;align-items:center;gap:8px;color:var(--muted);text-decoration:none;font-size:14px}
footer a:hover{color:var(--text)}
.gh{width:18px;height:18px;fill:currentColor}
@media (max-width:640px){.row{grid-template-columns:1fr auto}.row p{grid-column:1/-1;grid-row:2}}
</style>
</head>
<body>
<main>
<h1>{title}</h1>
<p class="hint">{hint}</p>
<input id="q" type="search" placeholder="Search scripts" aria-label="Search scripts">
{sections}
<footer><a href="{repo}">{github} GitHub</a></footer>
</main>
<script>
function copy(text){
  if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
  var t=document.createElement('textarea'); t.value=text; document.body.appendChild(t); t.select();
  document.execCommand('copy'); document.body.removeChild(t); return Promise.resolve();
}
document.querySelectorAll('button[data-cmd]').forEach(function(b){
  b.addEventListener('click',function(){
    if (b.classList.contains('done')) return;
    copy(b.dataset.cmd).then(function(){
      var icon=b.innerHTML; b.innerHTML='{check}'; b.classList.add('done');
      setTimeout(function(){ b.innerHTML=icon; b.classList.remove('done'); },1800);
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
                .replace('{hint}', ICON_COPY + ' Copy the command, paste it into PowerShell and press Enter. '
                                   + ICON_DOWNLOAD + ' Or download the file and double-click it (.py needs Python 3).')
                .replace('{sections}', section('Run in PowerShell', html_rows(True)) + '\n'
                                       + section('Download and Run', html_rows(False)))
                .replace('{github}', ICON_GITHUB).replace('{check}', ICON_CHECK)
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
