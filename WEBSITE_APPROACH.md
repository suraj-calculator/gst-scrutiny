# Turning this into a website — approach

Goal: a single-page site, deployed on GitHub Pages, where a user uploads the raw
files they download from the GST/E-Way Bill portals, the site runs the same
merge/convert/scrutiny logic that lives in this repo today, and hands back the
final `GST_MASTER_<GSTIN>_FY<...>.xlsx` workbook. Minimal ongoing maintenance —
push to `main`, the site is live, nothing to babysit.

## Decision: no backend server. Everything runs in the browser.

GitHub Pages only serves static files — it cannot run a Python process. The
options were:

1. **Client-side**: run the actual Python (unchanged) inside the browser via
   [Pyodide](https://pyodide.org) (CPython compiled to WebAssembly). GitHub
   Pages alone is enough to host it.
2. **Backend**: keep GitHub Pages for the UI only, and run the Python on a
   hosted server the page calls over an API.

**Chosen: option 1.** Reasoning, given this is a side project meant to run on
autopilot with minimal upkeep:

- **Nothing to operate.** A backend means an account somewhere, a service that
  can go to sleep/rate-limit/expire, dependency updates, and a bill to
  notice. A static site has none of that — once GitHub Pages is turned on for
  this repo, every `git push` is the entire deploy process, forever.
- **Data never leaves the device.** This tool processes GSTINs, PANs, invoice
  detail, bank/BS-PL figures — genuinely sensitive data. Client-side means the
  files a user uploads are processed entirely in their own browser tab and
  never transmitted anywhere. With a backend, that same data would cross the
  network and sit in a server's memory/logs, which is a real compliance
  surface for a side project to be responsible for.
- **Easy for Claude (or you) to fix later.** The existing `.py` files are
  reused as-is inside Pyodide — no rewrite into another language, no new
  framework to reason about. When something breaks, it's the same familiar
  Python module, plus one small JS "adapter" layer per section (see below).
  A backend would add a second codebase, a second deploy pipeline, and a second
  place for things to fail.

**Resolved (was a risk, now tested for real):** `pdf to excel/` depends on
`pdfplumber`, which depends on `pypdfium2` (compiled/native code). Tested this
directly in a real browser: `micropip.install("pdfplumber")` fails —
`ValueError: Can't find a pure Python 3 wheel for 'pypdfium2>=5.9.0'`. Also
checked Pyodide's own curated package repository (310 packages) directly —
zero PDF-related packages of any kind (no `pypdfium2`, `pymupdf`, or
`pdfplumber` pre-built for Pyodide either). This isn't a version/pinning
problem — `pdfplumber` genuinely cannot run in Pyodide today.

**Decision**: defer the 3 PDF-sourced inputs (BO Profile, GSTR-9, GSTR-9C)
from the first working version of the site. This is a safe deferral, not a
compromise — none of the three is a mandatory input to `main gst tool`
(only GSTR-1 and GSTR-3B are), so the site is fully useful without them from
day one. Real path forward for later, if wanted: swap the PDF *reading*
primitive only — use [PDF.js](https://mozilla.github.io/pdf.js/) (pure JS,
no native deps, the same engine browsers already use to render PDFs) to
extract per-page text/word-position data in JavaScript, then hand that
already-extracted data into the existing Python parser logic running in
Pyodide, instead of asking Python to open the PDF itself. That keeps the 4
parsers' real business logic (the section-splitting, table-building work)
unchanged in Python — only the "give me this page's text" step moves to JS.
Scoped as its own later phase, not a blocker for anything else here.

## What "main gst tool" needs — full input inventory

`main gst tool/master_build.py` reads one folder and classifies every file in
it by content (`gst_core.classify_folder()` — never by filename). Only GSTR-1
and GSTR-3B are mandatory; everything else is optional and degrades to a
documented SKIPPED/INFO finding when absent. Grouped by where the file comes
from:

### A — Produced by another folder in this repo (user uploads *raw* portal downloads; site auto-runs the merge/convert step)

| Final input | Raw upload from user | Processing pipeline (this repo) |
|---|---|---|
| GSTR-1 (merged, whole FY) | per-period GSTR-1 Excel exports | `forms merger/merger-tool/gstr1/merge_gstr1.py` |
| GSTR-3B (merged) | `GSTR3B_<GSTIN>_<MMYYYY>.zip` bundles | `files out/extractor/run.py` (unzip) → `forms merger/merger-tool/gstr3b/merge_gstr3b.py` (merge) |
| GSTR-2B (merged) | per-period GSTR-2B Excel exports | `files out/alligner/complete_workbooks.py` (align sheets) → `forms merger/merger-tool/gstr2b/merge_gstr2b.py` (merge) |
| GSTR-2A (whole FY) | per-period GSTR-2A Excel exports | `forms merger/merger-tool/gstr2a/merge_r2a.py` |
| E-Invoice (merged) | per-period E-Invoice Excel exports | `forms merger/merger-tool/e invoice/merge_einv.py` |
| Outward E-Way Bill (annual) | `EWB_MIS_Report_Excel (N).xls` — outward dir | `e_way bill/convert_ewb_files.py` → `e_way bill/auto_ewb_merger.py` |
| Inward E-Way Bill (annual) | `EWB_MIS_Report_Excel (N).xls` — inward dir | same as above, inward direction |
| BO / 360° Profile | `<GSTIN>_BO_Profile_<date>.pdf` | `pdf to excel/parse_bo_profile.py` |
| GSTR-9 | GSTR-9 PDF export | `pdf to excel/parse_gstr9.py` |
| GSTR-9C | GSTR-9C PDF export | `pdf to excel/parse_gstr9c.py` |

Not wired into the main tool at all today: `pdf to excel/parse_ewb_analytics.py`
(E-Way Bill Analytics PDF → xlsx). Confirmed by grep — `main gst tool` never
reads this output. Decide before Phase 3 whether to expose it in the UI as a
standalone/informational converter, or drop it from v1.

### B — Direct upload, no processing needed (already the shape the tool wants, straight off the portal)

| Input | Format |
|---|---|
| Cash Ledger | CSV (portal export) |
| Credit Ledger | CSV (portal export) |
| Liability Register (Part I) | CSV (portal export) |
| Liability Ledger (Part II, DRC) | CSV (portal export) |
| Portal Tax Liability & ITC Comparison | Excel (portal export) |
| Table 8A | Excel (portal export) |

### C — Reference/master files (org-level, upload once, reuse across every run — not per-filing)

| Input | Notes |
|---|---|
| HSN/SAC code master | e.g. NIC e-Invoice system's `HSN_SAC.xlsx` |
| Blocked-ITC keyword master | taxpayer/org-supplied keyword→HSN mapping for Sec 17(5) screening |
| Machinery HSN master | taxpayer/org-supplied HSN list for capital-goods detection |

These three should be **saved in browser storage** (see UI section) after the
first upload so the user doesn't re-upload them every filing.

### D — Not a file at all today

| Input | Current form | What changes for the website |
|---|---|---|
| Balance Sheet / P&L | hand-typed `BS_PL_DATA` dict inside `main gst tool/bs_pl_input.py`, edited per taxpayer | Becomes an actual **web form**: ~16 line items, each a prior-FY/current-FY number pair, tagged to the GSTIN being processed. The form's values get assembled into the same dict shape at run time — `gst_checks_forensic.check_bs_pl_rules()` itself doesn't need to change. |

## Proposed site structure

```
/docs/                      ← GitHub Pages serves straight from here (Settings → Pages → main → /docs), zero build step
  index.html                ← the single page: one upload section per input above
  app.js                    ← orchestration: wires uploads → Pyodide calls → next stage → final run
  py/                       ← the actual repo's .py files, vendored/symlinked in as-is
    core/                   ← main gst tool/*.py (unmodified)
    merge/gstr1/, gstr2a/, gstr2b/, gstr3b/, einv/   ← forms merger scripts (unmodified)
    ewb/                    ← e_way bill scripts (unmodified)
    pdf/                    ← pdf to excel scripts (unmodified)
    extract_align/          ← files out scripts (unmodified)
    web_adapters.py         ← the ONLY new Python: thin functions bridging
                                "list of uploaded files" → existing folder-based
                                main()/functions → "bytes to hand back to JS"
```

Keeping every existing script byte-for-byte where possible, and putting all
new glue in one `web_adapters.py`, is deliberate: it keeps the diff between
"today's repo" and "the website" small and legible, so fixing a bug later
means either editing the one familiar module that already has a README, or
the one adapter file — never hunting through a rewritten codebase.

### Why the existing scripts can mostly run unmodified

Every merge/convert script here already works by scanning **a folder** for
files by content signature (`glob.glob`, `os.listdir`) rather than taking
explicit file arguments. Pyodide ships a real in-memory filesystem
(`pyodide.FS`), so the adapter layer's job is small and mechanical: write each
uploaded browser `File`'s bytes into a virtual folder (e.g. `/work/gstr1/`),
call the script's existing `main("/work/gstr1")` unchanged, then read whatever
output file it wrote back out of the virtual FS and hand it to JS as a
downloadable blob / pass it to the next stage. No rewriting of the actual
merge/parse/check logic.

## UI flow

One section per input category (A, B, C above), in upload order matching the
pipeline (E-Way Bill → E-Invoice → GSTR-1 → GSTR-2A → GSTR-2B → GSTR-3B →
ledgers/comparison/Table 8A → BO Profile/GSTR-9/GSTR-9C PDFs → masters →
Balance Sheet/P&L form), then one final action:

- **Category A sections auto-process on upload** — the moment files land in
  a section, the site runs that section's merge/convert step immediately in
  the background and shows a small summary (e.g. "6 months merged, 2 sheets
  aligned") plus a download link for the intermediate file. This matches your
  "just handle it" preference over a manual per-section run button — one less
  click, and the user still sees exactly what happened.
- **Category B/C sections are just upload slots** — no processing, stored
  as-is. C (the 3 masters) is remembered in `localStorage`/IndexedDB after
  first upload so returning users don't re-upload them.
- **Category D is a form**, not an upload.
- A single **"Run Full Scrutiny"** button at the bottom is enabled once
  GSTR-1 + GSTR-3B (the only mandatory inputs) are ready; it assembles
  everything processed so far into one working folder in the virtual FS and
  calls `master_build.main()` unchanged, then offers the final
  `GST_MASTER_<GSTIN>_FY<...>.xlsx` as a download.
- Every step's status (done / skipped / failed, and why) is shown inline —
  mirroring the tool's own existing philosophy of "explicit SKIPPED, never a
  silent gap" (see `main gst tool/docs/GST_360_SCRUTINY_DIRECTIVE.md`).

## Build sequence

**Phase 0 — repo hygiene (do first, before wiring anything to the web).**
Remove the two dead-code duplicates flagged earlier (`forms
merger/merger-tool/merge_gstr3b.py` top-level copy, and the stray `forms
merger/gst_merge_common.py`) so the vendoring step in Phase 2+ doesn't have to
decide which copy is real. (Not deleted yet — flagging for your go-ahead.)

**Phase 1 — de-risk PDF parsing in Pyodide.** ✅ Done — tested for real in a
browser, `pdfplumber` cannot run in Pyodide (see above). PDF inputs deferred.

**Phase 2 — skeleton + first working slice.** ✅ Done, and expanded well past
"first slice" — `docs/index.html` + `docs/app.js` is a real, working site
(not a mockup) covering 6 of the 7 upload sections end-to-end:

| Section | Real pipeline wired | Verified against real data |
|---|---|---|
| E-Way Bill (in/out) | `convert_ewb_files.py` → `auto_ewb_merger.py` | ✅ CPython + browser |
| GSTR-1 | `merge_gstr1.py` | ✅ CPython + browser |
| E-Invoice | `merge_einv.py` | ✅ CPython |
| GSTR-2B | `complete_workbooks.py` (align) → `merge_gstr2b.py` | ✅ CPython + browser |
| GSTR-3B | `run.py` (extract zips) → `merge_gstr3b.py` | ✅ CPython + browser |
| GSTR-2A | `merge_r2a.py` | Code path only — no real GSTR-2A fixtures exist in this repo to test against; UI says so |

"Verified in browser" means: real production `index.html`, real "Try with
sample data" buttons (fetch real sample bytes, run the actual adapter),
checked with browser automation — not a separate spike page. Ledgers/Table
8A/Comparison and the 3 reference masters are wired too (real byte capture,
no processing needed; masters persist across visits via `localStorage`).
Annual PDF Reports section is present but disabled with an explanation (see
Phase 1). Balance Sheet/P&L is a real form. "Run full scrutiny" stays
disabled on purpose — `main gst tool` isn't wired up yet (next).

**New finding from testing the real page**: Pyodide runs synchronously on
the main thread, so a multi-second merge **freezes the tab** (confirmed —
browser automation timed out waiting on a GSTR-2B run). Nothing broke or
raced (Pyodide's single-threaded execution naturally serializes overlapping
calls — tested by clicking two sections back-to-back, both came back
correct), but the frozen-tab UX is real and worth fixing before this goes to
real users. This is Phase 2's originally-deferred "Web Worker vs main
thread" decision (see below) — now backed by a real measurement instead of
a hypothetical.

**Also fixed while building this**: an HTML-escaping bug where a section's
description text (`GSTR3B_<GSTIN>_<MMYYYY>.zip`) was silently stripped by
the browser's HTML parser, since raw `<GSTIN>` reads as an unknown tag —
found by reading the real rendered page, not by inspection.

**Phase 3 — the 5 `forms merger` return types.** ✅ Done (gstr1, gstr2a,
gstr2b, gstr3b, einv all wired — see Phase 2 table for what's verified).

**Phase 4 — `files out`** (extractor + alligner) as pre-steps feeding the
GSTR-3B and GSTR-2B sections. ✅ Done (folded into Phase 2/3 rather than a
separate pass — `process_gstr3b`/`process_gstr2b` in `web_adapters.py`).

**Phase 5 — `pdf to excel`** parsers. Deferred per Phase 1's finding — not
started.

**Phase 6 — `main gst tool` itself.** ✅ Done. `web_adapters.process_full_scrutiny()`
assembles every merged/collected file into one folder, injects the browser's
BS/P&L form as a real `bs_pl_input.py` module, and calls
`master_build.main()` unchanged. Dependency check first (learning from the
PDF situation): all 12 core files import nothing but `openpyxl` (already
installed) plus stdlib — zero new package risk. "Run full scrutiny" is wired
for real, gated on GSTR-1 + GSTR-3B (the same requirement `master_build.py`
itself enforces), with real results (month count, HSN/fraud and flow
finding counts parsed from `master_build`'s own printed summary, never
fabricated), a full log, and a real download.

Verified twice: once under plain CPython with real matched-taxpayer sample
data (GSTR-1 + GSTR-3B, same GSTIN) — valid 87-sheet workbook, opened with
openpyxl. Once for real in the browser via the production page — ran to
completion, downloaded a real 343 KB report (207 HSN/fraud flags, 1 flow
flag, 1 rectification pair), matching the CPython run almost exactly.

**Measured, not guessed**: a full 12-month scrutiny takes ~117s under plain
CPython — the browser run took noticeably longer (WASM overhead). The "Run
full scrutiny" button now shows a live elapsed-time counter with an explicit
"1-3 minutes, keep this tab open" note instead of a plain spinner.

**Phase 6.5 — main-thread freeze.** ✅ Fixed. Confirmed by testing (a merge
call froze the tab, timing out browser automation), then fixed by moving
Pyodide entirely into `docs/worker.js` — `app.js` never touches Pyodide
directly, only postMessage. Re-verified after the fix: triggered the
heaviest call (GSTR-3B zip extraction) and confirmed the tab stayed fully
responsive (a page read that previously timed out returned instantly,
typing worked with no lag) while the worker computed in the background.

**Phase 7 — Balance Sheet/P&L form.** ✅ Done — real form in `index.html`,
now correctly keyed to `BS_PL_DATA`'s actual field names (fixed a bug where
it was keyed by display label and stored raw strings instead of parsed
numbers) and tagged with the entered GSTIN, ready for Phase 6 to consume.
Currently 6 of the ~16 line items from `bs_pl_input.py`; the rest are a
mechanical addition.

**Two more bugs found by testing the live page, both fixed**: an
`enableAllDropzones()` bug where the disabled-check matched a dropzone's own
disabled marker before it could be cleared, silently breaking real
drag-and-drop/click-to-browse for every section since the previous commit
(only the sample-data buttons still worked, since they check a different
property) — and a **serious one**: `main gst tool/bs_pl_input.py` had real
hardcoded financial figures for an actual taxpayer baked into the first
commit. Since this repo was never pushed anywhere, rewrote local git history
to strip it entirely and replaced it with the sanitized template it's
actually meant to be.

**Phase 8 — deploy.** Not started — the last remaining piece. Enable GitHub
Pages on this repo pointed at `/docs` on `main`. No GitHub Actions needed —
plain static files are enough. From here on, shipping a fix is just
`git push`.

## Open items

- ✅ Dead-code duplicates deleted (Phase 0).
- ✅ `files out/extractor/run.py`'s hardcoded path fixed, verified against
  real sample zips.
- ✅ License-expiry lock removed from every copy of `gst_merge_common.py`
  (confirmed intentional by the repo owner).
- Still open: OK to drop `parse_ewb_analytics.py` from the v1 UI (unused by
  the main tool today), or should it be exposed as a standalone converter
  section?
- Still open: any preference on how the 3 reference masters (HSN/SAC,
  blocked-ITC keywords, machinery HSN) get their *first* copy in front of a
  new user — bundled into the site as defaults, or purely user-uploaded on
  first use?
- Still open: the e-way bill `.bat` date trap
  (`e_way bill/convert_and_merge.bat`, fires 2026-09-16) — moot for the
  website itself (that `.bat` file isn't part of the web port), but worth a
  decision for the standalone local-Windows tool it belongs to.
