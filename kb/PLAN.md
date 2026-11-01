# Audit knowledge base: plan and handover

Read this first in a new session. It records the objective, what exists, the decisions taken, and the next steps.

## Objective

Turn the legacy internal audit file share into a knowledge base the team can query. Three uses, in priority order:

1. **Methodology**: for a given process (revenue, procurement, payroll, ...) give the consolidated risks, controls and
   test procedures the team has used across past engagements, with counts of how many engagements used each and links
   back to the source workpapers. This closes the knowledge gap between staff and makes new joiners productive.
2. **Engagement registry**: one row per engagement with client, industry, processes audited, period covered, fieldwork
   and report dates, audit committee date, fees, team, client contacts. Feeds tender sections: value proposition, client
   references, timeline and fees.
3. **Evidence search**: full-text search over everything extracted, to answer ad hoc questions and to cite sources.

Approach: build structured layers on top of the extracted text. AI reads documents at build time; at query time an MCP
server exposes search and lookup tools. Migration of the share is NOT a goal.

## What exists

| File | What it is |
|---|---|
| `scan_output/*.csv` | Metadata catalogue of the whole share: 739,469 files, 904 GB. Built by `../scanfiles.ps1`. |
| `extraction_shortlist.xlsx` | 45,021 key documents chosen by file name (RCM, audit programme, engagement letter, proposal, fee, final report, AC paper, planning memo, walkthrough, issues, closing) with client, engagement folder, year, process hint, confidence. Built by `build_shortlist.py` (in the session scratch, logic documented below). |
| `kb.db` | SQLite, the extracted text of the shortlist. Built by `../extractfiles.py`. 40,608 documents with text, 4,251 scanned PDFs without text, 162 empty. About 3 GB. |
| `extract.log` | Run log of the extractor, one line per failed file. |

### kb.db schema

```
docs(id, path, name, ext, size, modified, doctype, confidence, client, engagement, year, yearsource, process,
     topfolder, duplicates, status, parts, chars, error, seconds, done_at, copy_s)
parts(id, doc_id, seq, kind, title, content)       kind: body | table | sheet | slide | page | email | attachment | body-crude | note
parts_fts(title, content)                           FTS5 over parts, porter tokenizer; rowid = parts.id
```

- `docs.status`: done | error | empty. Errors are almost all "no text layer" (scanned PDF).
- `docs.client` / `engagement` / `year` come from the folder path under `SGBRS Workfiles\<client>\<engagement>\...`.
  `yearsource` says where the year came from (folder, subfolder, filename, modified). `modified` is a weak guess.
- For the registry-type top folders (Engagement Letters & Final IA Reports, Proposals, Fee Schedule) `client` is empty;
  the client must be read from the document.
- Tables from Word and Excel are kept as rows: one line per row, cells joined with ` | `. RCM content lives in
  parts of kind `table` (Word) or `sheet` (Excel).
- `process` is a hint from the folder name (e.g. "N. Purchases, payables and payment"); it is empty when the folder
  was a phase name (planning, PBC, report...). One client-year can cover several processes, each with its own RCM.

### Shortlist selection logic (for reference)

File name regexes per document type, priority order: RCM, audit programme, engagement letter, proposal, fee, final
report, AC paper, planning memo, walkthrough, issues, closing. Confidence 3 = name says final/signed/issued,
2 = clear keyword, 1 = draft/version marker or weak keyword type. Same name+size elsewhere collapsed into one row.
Known gap: RCMs with generic names (D100.xls) are missed; content search in kb.db finds them (see `kb_profile.py`).

## Decisions taken

- Files may leave the network; the share is legacy. kb.db and the CSVs are the working copies.
- Old .doc files converted by LibreOffice (bulk), Word as fallback, raw text as last resort. 7 files are raw text.
- Scanned PDFs: no OCR yet. Decide after `kb_profile.py` shows how many have no text twin in the same folder.
  Rule agreed: none for scans with a text twin; Tesseract for the rest so they are searchable; model reading of page
  images only for key documents with no twin.
- AI reading happens where the database is: locally in Claude Code / Cowork with the folder open, or batch via
  `claude -p` (headless Claude Code, subscription billing). Not via API keys.
- The methodology library needs a senior's review before new joiners learn from it. The registry is factual.

## Build stages

### Stage 1: profile (script, no AI)   `python kb_profile.py`
Sanity checks and the numbers needed for decisions: coverage per client-year, scanned PDFs without a text twin,
RCMs found by content that the shortlist missed, top processes. Writes `kb_profile.xlsx`.

### Stage 2: RCM rows (script, no AI)   `python kb_rcm.py`
Parses every RCM / audit programme table into rows: engagement, process, risk, control, test procedure, plus the
source doc and table. Header detection is heuristic; rows it cannot map are kept with raw cells. Writes table
`rcm_rows` into kb.db and `rcm_rows.xlsx` for review. This is the raw material of the methodology library.

### Stage 3: registry bundles (script + AI)   `python kb_bundle.py`
For each engagement builds one compact markdown bundle (engagement letter first pages, report cover and scope, fee
sheet, proposal summary) under `bundles/<client>__<engagement>.md`, capped at ~6,000 words, plus a regex first pass
of dates and fees into `registry_draft.csv`. The AI then reads bundles and fills `registry.jsonl`, one JSON object per
engagement with the fields in `kb_bundle.py --schema`. `python kb_bundle.py --ingest` loads the JSONL into table
`registry`. Work newest years first; older engagements matter less for tenders.

### Stage 4: methodology normalisation (AI)
Group `rcm_rows` by process; the AI merges near-duplicate risks/controls/tests into canonical entries with a count of
engagements and the list of sources. Output table `methodology(process, subprocess, risk, control, test, n_engagements,
sources)`. Senior review before use. Script to be written after Stage 2 shows the real shape of the rows.

### Stage 5: serve (MCP server on Proxmox)
`apps/mcp/mcpkb/server.py` following the repo's server conventions: tools `search(query)`, `get_doc(id)`,
`engagements(filter)`, `methodology(process)`. kb.db copied to the shared folder. Team asks questions through Claude.

## Next steps, in order

1. Put kb.db, extract.log and extraction_shortlist.xlsx in this folder (or pass `--db`).
2. `python kb_profile.py` and read `kb_profile.xlsx`. Decide OCR scope from the "scanned no twin" sheet.
3. `python kb_rcm.py`, open `rcm_rows.xlsx`, check a dozen rows against the source documents, tune header patterns.
4. `python kb_bundle.py`, then start reading bundles for the newest engagements and writing `registry.jsonl`.
5. Report numbers back: engagements found, rows parsed, coverage. Then design Stage 4 on real data.

## Running scripts here

Python 3.9+, `pip install openpyxl`. All scripts take `--db kb.db` (default: kb.db in this folder) and are read-only on
`docs`/`parts`; they add their own tables. Re-running replaces those tables.
