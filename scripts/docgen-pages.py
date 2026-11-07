#!/usr/bin/env python3
# Purpose: Build the GitHub Pages script page from each script's lines 1-2
# =============================================================================
# Usage:
#   python3 scripts/docgen-pages.py [out_dir]   write index.html and downloadable files (default: _site)
# =============================================================================

import html
import os
import shutil
import sys

OUT_DIR = sys.argv[1] if len(sys.argv) > 1 else '_site'
OUTPUT_HTML = os.path.join(OUT_DIR, 'index.html')

KINDS = {'.ps1': 'PowerShell', '.bat': 'Batch', '.cmd': 'Batch', '.py': 'Python'}

def header(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        return [f.readline().strip() for _ in range(2)]

def purpose_of(line):
    text = line.split(':', 1)[1].strip() if ':' in line else ''
    if text and not text.endswith('.'):
        text += '.'
    return text

def card(name, purpose, button):
    return '<div class="card"><b>{}</b><p>{}</p>{}</div>'.format(html.escape(name), html.escape(purpose), button)

def html_cards(run_by_command):
    cards = []
    for filename in sorted(os.listdir('apps')):
        stem, ext = os.path.splitext(filename)
        if ext not in KINDS:
            continue
        l1, l2 = header(os.path.join('apps', filename))
        name = stem[0].upper() + stem[1:]
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
        cards.append(card(name, purpose, button))
    return '\n'.join(cards)

PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Windows scripts</title>
<style>
:root{--bg:#f6f6f4;--card:#fff;--text:#1f1f1e;--muted:#6b6a66;--line:#e2e1dc;--ok:#2e7d32}
@media (prefers-color-scheme:dark){:root{--bg:#1c1c1b;--card:#262624;--text:#ecebe6;--muted:#a3a29c;--line:#3a3a37;--ok:#7bc47f}}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:960px;margin:0 auto;padding:2rem 1rem}
h1{font-size:22px;font-weight:600;margin:0 0 .25rem}
h2{font-size:16px;font-weight:600;margin:2rem 0 .75rem}
.hint{color:var(--muted);margin:0 0 1rem}
input{width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text);font-size:15px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;display:flex;flex-direction:column;gap:8px}
.card b{font-size:16px;font-weight:600}
.card p{margin:0;color:var(--muted);font-size:13px;flex:1}
button,.btn{padding:9px;text-align:center;text-decoration:none;display:block;border:1px solid var(--line);border-radius:8px;background:transparent;color:var(--text);font-size:14px;cursor:pointer}
button:hover,.btn:hover{background:var(--bg)}
button.done{color:var(--ok);border-color:var(--ok)}
</style>
</head>
<body>
<main>
<h1>Windows scripts</h1>
<p class="hint">Copy command: paste into PowerShell and press Enter. Download: double-click the file (Python files need Python 3).</p>
<input id="q" type="search" placeholder="Search scripts" aria-label="Search scripts">
<h2>Run in PowerShell</h2>
<div class="grid">
{installers}
</div>
<h2>Download and run</h2>
<div class="grid">
{tools}
</div>
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
  document.querySelectorAll('.card').forEach(function(c){
    c.style.display=c.textContent.toLowerCase().indexOf(q)>-1?'':'none';
  });
});
</script>
</body>
</html>
"""

def generate_html():
    page = (PAGE.replace('{installers}', html_cards(True))
                .replace('{tools}', html_cards(False)))
    with open(OUTPUT_HTML, 'w') as out:
        out.write(page)

if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    generate_html()


