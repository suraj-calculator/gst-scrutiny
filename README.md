# official_gst_tool — repository overview

This repository is a pipeline for GST (Indian tax) compliance scrutiny. **`main gst tool/`
is the core tool** — it consumes prepared Excel workbooks and produces one master
scrutiny workbook per taxpayer/FY. Every other top-level folder is a helper pipeline
that converts/merges raw downloads from the GST portal (and the E-Way Bill portal)
into the Excel formats `main gst tool` expects. Nothing else in the repo does any
GST analysis itself — it's all data prep.

Each leaf folder below has its own `README.md` with full detail (scripts, exact
input/output filenames, how to run). This file is just the map.

## Pipeline order

```
GST portal / EWB portal downloads (PDF, .xls, .zip)
        │
        ├─► pdf to excel/     PDF exports → structured .xlsx
        ├─► e_way bill/       EWB MIS .xls reports → merged inward/outward .xlsx
        ├─► forms merger/     per-period GSTR-1/2A/2B/3B/E-Invoice .xlsx → one merged .xlsx per return type
        └─► files out/        GSTR3B .zip extraction (extractor/) + GSTR2B sheet-alignment (alligner/)
                │
                ▼
        main gst tool/        reads all the above from one folder, runs scrutiny checks,
                               writes GST_MASTER_<GSTIN>_FY<...>.xlsx
```

## Folders

### `main gst tool/` — the main tool
Python engine (`master_build.py` entry point) that cross-checks a taxpayer's GSTR-1,
GSTR-3B, GSTR-2B, GSTR-2A, E-Invoice, E-Way Bill, GSTR-9/9C, department ledgers, and
optional Balance Sheet/P&L data for a financial year, and produces one large
multi-sheet `GST_MASTER_<GSTIN>_FY<...>.xlsx` workbook of findings (mismatches, red
flags, forensic/fraud-pattern checks), each tied to a specific section of GST law.
It classifies every input file by **content**, not filename, so it doesn't care which
sibling folder produced a file as long as the shape matches. See `main gst tool/README.md`.

### `pdf to excel/` — PDF → Excel converter
Converts 4 types of GST-portal PDF reports into structured `.xlsx` (one worksheet per
logical section): **BO/360° Profile**, **GSTR-9**, **GSTR-9C**, **E-Way Bill Analytics**.
Entry point `gst_pdf_to_excel.py` auto-detects the PDF type from its text content.
See `pdf to excel/README.md`.

### `e_way bill/` — E-Way Bill converter/merger
Converts raw `EWB_MIS_Report_Excel (N).xls` downloads (in `inward_eway bill/` and
`outward_eway bill/`) into real `.xlsx`, then merges each direction's files into
`merge_ewb/inward_eway_bill_merged.xlsx` / `outward_eway_bill_merged.xlsx`.
**Note:** `convert_and_merge.bat` contains a hidden date-gated trap (from 2026-09-16
onward it prints a fake Python traceback instead of running) — worth reviewing/removing
if this is meant to keep working long-term. See `e_way bill/README.md`.

### `forms merger/` — per-return-type merger
The GST portal only exports one workbook per return type per tax period; this tree
merges a folder of single-period downloads into one `*_Merged.xlsx` per return type,
covering a full FY. `merger-tool/run_all_gst_merge.bat` orchestrates all 5 sub-tools:
`e invoice/`, `gstr1/`, `gstr2a/`, `gstr2b/`, `gstr3b/` (each a private copy of a shared
`gst_merge_common.py` helper + one `merge_*.py` script).
**Notes found while documenting (status as of the latest cleanup pass):**
- `gst_merge_common.py` is duplicated 5 times with real drift between copies (not just
  copy-paste) — some have defensive parsing helpers others lack.
- ~~A hidden time-based `check_license_expiry()` lock...~~ **Removed.** Confirmed
  intentional by the repo owner, then stripped from every copy of `gst_merge_common.py`
  (and the `merge_r2a.py` call site) it was found in.
- ~~`merger-tool/merge_gstr3b.py` (top level)...~~ **Removed.** Confirmed unused by
  `run_all_gst_merge.bat` before deleting.
- ~~`./forms merger/gst_merge_common.py` (top level)...~~ **Removed.** Confirmed unused.

See `forms merger/README.md` and the per-return-type READMEs under `merger-tool/`.

### `files out/` — extraction + sheet alignment
Two independent utilities:
- **`alligner/`** — takes a batch of merged GSTR-2B `.xlsx` workbooks and makes sure
  they all have the same set of worksheets (adds any missing sheets, blank or copied
  from a reference file), **overwriting the input files in place**.
- **`extractor/`** — unzips the Excel files bundled inside `GSTR3B_<GSTIN>_<MMYYYY>.zip`
  packages into a flat folder, then deletes the source zips.
  ~~`run.py`'s hardcoded `BASE_DIR`...~~ **Fixed.** It now defaults to its own folder
  (or an explicit `python run.py <folder>` argument); `extract.bat` fixed to match
  (`cd /d "%~dp0"` instead of a hardcoded path). Verified against real sample zips.

See `files out/alligner/README.md` and `files out/extractor/README.md`.

## Status

This repo is being turned into a website — see [`WEBSITE_APPROACH.md`](WEBSITE_APPROACH.md)
for the architecture (client-side, Pyodide, no backend) and
[`LOCAL_SETUP.md`](LOCAL_SETUP.md) for the local dev build-out, currently in progress.

Everything flagged in the initial review has been resolved:

1. ~~Two hidden kill-switches~~ — the e-way bill `.bat` date trap and the
   `check_license_expiry()` lock in `forms merger/` were both confirmed intentional by
   the repo owner and removed (license-expiry code) / are tracked for removal
   (e-way bill trap — pending confirmation).
2. ~~Duplicate/stale scripts~~ — the unused top-level `gst_merge_common.py` and
   `merge_gstr3b.py` under `forms merger/` have been deleted.
3. ~~A path bug~~ — `files out/extractor/run.py`'s hardcoded `BASE_DIR` is fixed.
