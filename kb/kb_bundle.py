# Purpose: Stage 3. Build one markdown bundle per engagement for the AI to read, a regex draft of the registry, and ingest the AI's answers.
#   python kb_bundle.py [--db kb.db] [--years 2015-2026]      # writes bundles/<client>__<engagement>.md and registry_draft.csv
#   python kb_bundle.py --schema                               # the JSON record the AI should write per engagement into registry.jsonl
#   python kb_bundle.py --ingest registry.jsonl                # loads the JSONL into table registry
import sqlite3, argparse, re, os, json, csv
from collections import defaultdict

SCHEMA = {
    'client': 'client legal or common name', 'engagement': 'engagement folder name as in kb.db', 'year': 'YYYY',
    'industry': 'sector, e.g. healthcare, statutory board, REIT, manufacturing', 'engagement_type': 'internal audit | co-sourced IA | outsourced IA | SOX/J-SOX | review | advisory | other',
    'processes': ['list of processes / areas audited'], 'period_covered': 'e.g. Apr 2018 to Mar 2019',
    'fieldwork': 'MM-YYYY to MM-YYYY', 'report_date': 'YYYY-MM-DD of final report', 'ac_date': 'YYYY-MM-DD audit committee presentation',
    'fee': 'number, currency and whether excluding GST', 'hours_or_days': 'budgeted effort if stated', 'team': ['EY partner/manager/staff named'],
    'client_contacts': ['name, title'], 'scope_summary': '2-3 sentences', 'sources': ['doc ids used'], 'confidence': 'high | medium | low', 'notes': ''}
ap = argparse.ArgumentParser(); ap.add_argument('--db', default='kb.db'); ap.add_argument('--years', default='2010-2026')
ap.add_argument('--schema', action='store_true'); ap.add_argument('--ingest', default=''); ap.add_argument('--maxchars', type=int, default=30000); a = ap.parse_args()
if a.schema: print(json.dumps(SCHEMA, indent=2)); raise SystemExit
db = sqlite3.connect(a.db); db.row_factory = sqlite3.Row

if a.ingest:
    db.execute('drop table if exists registry')
    db.execute('create table registry(client, engagement, year, industry, engagement_type, processes, period_covered, fieldwork, report_date, ac_date, fee, hours_or_days, team, client_contacts, scope_summary, sources, confidence, notes)')
    n = 0
    for line in open(a.ingest, encoding='utf-8'):
        line = line.strip()
        if not line: continue
        r = json.loads(line); n += 1
        db.execute('insert into registry values(' + ','.join('?' * 18) + ')', [json.dumps(r.get(k), ensure_ascii=False) if isinstance(r.get(k), list) else r.get(k) for k in SCHEMA])
    db.commit(); print(f'registry: {n} rows'); raise SystemExit

y0, y1 = (int(x) for x in a.years.split('-'))
ORDER = {'Engagement letter': 0, 'Final report': 1, 'Fee': 2, 'Proposal': 3, 'Planning memo': 4, 'AC paper': 5, 'Closing': 6}
CAP = {'Engagement letter': 6000, 'Final report': 8000, 'Fee': 3000, 'Proposal': 5000, 'Planning memo': 4000, 'AC paper': 3000, 'Closing': 2000}
DATE = re.compile(r'\b(\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{4}|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})\b', re.I)
MONEY = re.compile(r'\b(?:S\$|SGD|\$|USD|RM|MYR)\s?\d[\d,]{2,}(?:\.\d+)?', re.I)

docs = db.execute("""select id, name, doctype, confidence, client, engagement, year, process from docs
                     where status='done' and client != '' and doctype in ('Engagement letter','Final report','Fee','Proposal','Planning memo','AC paper','Closing','RCM','Audit programme')""").fetchall()
eng = defaultdict(list)
for d in docs:
    try: yr = int(d['year'])
    except Exception: yr = 0
    if y0 <= yr <= y1: eng[(d['client'], d['engagement'], yr)].append(d)
os.makedirs('bundles', exist_ok=True)
def safe(s): return re.sub(r'[^\w.-]+', '_', s)[:60]
draft = []
for (client, engagement, yr), ds in sorted(eng.items(), key=lambda x: (-x[0][2], x[0][0])):
    reg = [d for d in ds if d['doctype'] in ORDER]
    procs = sorted({d['process'] for d in ds if d['process']} | {re.sub(r'\.[a-z]+$', '', d['name']) for d in ds if d['doctype'] == 'RCM'})
    if not reg: draft.append([client, engagement, yr, '; '.join(procs)[:300], len(ds), '', '', '', '', 'no registry docs']); continue
    reg.sort(key=lambda d: (ORDER[d['doctype']], -d['confidence']))
    out = [f'# {client} | {engagement} | {yr}', f'processes seen in folder: {"; ".join(procs)[:500]}', f'documents in engagement: {len(ds)}', '']
    total = 0; used = []; dates = []; money = []; seen_types = set()
    for d in reg:
        if d['doctype'] in seen_types and d['doctype'] not in ('Fee',): continue      # one per type is enough, best confidence first
        seen_types.add(d['doctype'])
        txt = '\n'.join(r[0] for r in db.execute("select content from parts where doc_id=? and kind in ('body','page','sheet','table','slide','email','body-crude') order by seq", (d['id'],)))
        if not txt.strip(): continue
        cap = CAP[d['doctype']]; snippet = txt[:cap]
        if len(txt) > cap:   # also grab the part around scope / fee / audit committee keywords
            for kw in ('scope', 'fee', 'audit committee', 'timeline', 'period'):
                m = re.search(kw, txt[cap:], re.I)
                if m: s = cap + max(0, m.start() - 300); snippet += f'\n[...]\n' + txt[s:s + 1200]
        out += [f'## {d["doctype"]}  (doc {d["id"]}, confidence {d["confidence"]}): {d["name"]}', snippet.strip(), '']
        used.append(d['id']); total += len(snippet)
        dates += DATE.findall(txt[:20000]); money += MONEY.findall(txt[:20000])
        if total > a.maxchars: break
    with open(os.path.join('bundles', f'{safe(client)}__{safe(engagement)}__{yr}.md'), 'w', encoding='utf-8') as fh: fh.write('\n'.join(out))
    uniq = lambda xs: list(dict.fromkeys(x.strip() for x in xs))[:8]
    draft.append([client, engagement, yr, '; '.join(procs)[:300], len(ds), ' '.join(str(u) for u in used), '; '.join(uniq(dates)), '; '.join(uniq(money)), total, ''])
with open('registry_draft.csv', 'w', newline='', encoding='utf-8-sig') as fh:
    w = csv.writer(fh); w.writerow(['Client', 'Engagement', 'Year', 'Processes', 'Docs', 'BundleDocIds', 'DatesFound', 'AmountsFound', 'BundleChars', 'Note']); w.writerows(draft)
print(f'{len(eng):,} engagements in {y0}-{y1}; bundles written to .\\bundles, draft to registry_draft.csv')
print('next: read bundles newest first, write one JSON line per engagement to registry.jsonl (python kb_bundle.py --schema), then --ingest registry.jsonl')
