# E-Way Bill Merger

## Overview

This folder is a helper pipeline that prepares E-Way Bill data for the **main gst tool**. It does not do any GST scrutiny itself — it only takes raw "EWB MIS Report" files downloaded from the government E-Way Bill portal (one file per report/page, split into inward and outward), converts them to a readable Excel format, and merges each set into a single consolidated `.xlsx` file that can then be fed into the main GST tool as an input source (several modules in `main gst tool` — e.g. `gst_parsers_returns.py`, `gst_checks_flow.py`, `gst_checks_forensic.py`, `gst_checks_monthly.py`, `gst_checks_hsn_fraud.py`, `gst_machinery_scan.py`, `gst_core.py`, `gst_config.py`, `gst_report.py`, `master_build.py` — reference e-way bill data).

## Scripts

### `convert_ewb_files.py`
Converts raw `.xls` files in `inward_eway bill/` and `outward_eway bill/` into `.xlsx` files (same base name, extension changed from `.xls` to `.xlsx`, saved into the same folder — it does not delete the original `.xls`).

It tries three conversion methods in order, per file, and stops at the first one that succeeds:
1. **HTML method** — many "MIS Report" `.xls` files from government portals are actually HTML tables saved with an `.xls` extension. Reads the file via `pandas.read_html()`, takes the first table, writes it out with `openpyxl`.
2. **Openpyxl method** — opens the file with `openpyxl.load_workbook(data_only=True)`, reads all rows, treats the first row as the header, and writes a `DataFrame` built from the rest.
3. **CSV method** — falls back to `pandas.read_csv()`, trying `utf-8`, `latin1`, then `cp1252` encodings (skipping bad lines), and writes whatever it manages to parse.

If none of the three methods work, the file is reported as failed to convert (the script suggests opening it manually in Excel and re-saving as `.xlsx`).

Run with: `python convert_ewb_files.py` (no arguments; prompts "Press Enter to exit..." when done).

### `auto_ewb_merger.py`
Merges the inward and outward Excel files into consolidated output files. Reads every `*.xlsx` and `*.xls` file in `inward_eway bill/` and `outward_eway bill/` (via `pandas.read_excel`, using the `xlrd` engine for `.xls` and `openpyxl` for `.xlsx`, with a plain `read_excel` fallback for `.xls`), concatenates each folder's files independently with `pandas.concat(..., ignore_index=True, sort=False)`, drops exact duplicate rows, and writes two outputs per run into `merge_ewb/`:
- `inward_eway_bill_merged.xlsx` / `outward_eway_bill_merged.xlsx` — the "latest" merged file, overwritten on every run.
- `inward_eway_bill_<YYYYMMDD_HHMMSS>.xlsx` / `outward_eway_bill_<YYYYMMDD_HHMMSS>.xlsx` — a timestamped snapshot kept alongside it on every run, so old merges accumulate in `merge_ewb/`.

Two modes, controlled by the `--once` CLI flag:
- `python auto_ewb_merger.py --once` — merges once and exits (used by `convert_and_merge.bat` and `merge_now.bat`).
- `python auto_ewb_merger.py` (no args) — `monitor_and_merge()`: polls the two input folders every 5 seconds, tracks already-processed files in a local `processed_files.json`, and re-runs the full merge (both inward and outward) whenever new `.xls`/`.xlsx` files appear, waiting at least 10 seconds between merges. Runs until interrupted (Ctrl+C) — used by `start_merger.bat`.

Note: because `merge_all_files()` merges *every* file currently in the input folder (not just the new ones) each time it runs, `processed_files.json` is only used to detect that new files arrived, not to do an incremental/partial merge.

### `merge_ewb.py`
Present in the folder but **empty (0 bytes)** — not used by any of the `.bat` files or by `auto_ewb_merger.py`. There is also a `merge_ewb.zip` archive in this folder alongside the (unrelated, same-named) `merge_ewb/` output directory.

### `convert_and_merge.bat`
Runs the full pipeline end-to-end from Windows: `python convert_ewb_files.py` (Step 1) followed by `python auto_ewb_merger.py --once` (Step 2), then pauses. Before doing anything, it runs a PowerShell date check against a hard-coded `EXPIRY=2026-09-16` date; if the current date is past that, it prints a fabricated Python traceback (`SchemaMismatchError`) to the console instead of running the scripts, and exits — i.e. the batch file is time-limited and disguises its expiry as a script error rather than converting/merging anything after that date.

### `merge_now.bat`
`cd /d "C:\Users\admin\Documents\merger\e_way bill"` then `python auto_ewb_merger.py --once` — a one-shot merge only (no conversion step first), hard-coded to a specific Windows path.

### `start_merger.bat`
`cd /d "C:\Users\admin\Documents\e_way bill"` then `python auto_ewb_merger.py` (no `--once`, so it starts the continuous folder-monitoring mode) and `pause`s at the end. Also hard-coded to a specific Windows path (different from `merge_now.bat`'s path).

## Folder structure

- **`inward_eway bill/`** — Input directory. Holds raw E-Way Bill MIS Report files downloaded from the GST e-way bill portal for *inward* (received) shipments, named like `EWB_MIS_Report_Excel (N).xls`. After conversion, matching `.xlsx` copies accumulate alongside the originals.
- **`outward_eway bill/`** — Input directory, same naming/format as above but for *outward* (dispatched) shipments.
- **`merge_ewb/`** — Output directory (created automatically by `auto_ewb_merger.py` if missing). Receives `inward_eway_bill_merged.xlsx` / `outward_eway_bill_merged.xlsx` (latest merge) plus timestamped snapshots (`inward_eway_bill_<timestamp>.xlsx` / `outward_eway_bill_<timestamp>.xlsx`) from every merge run.

## Data flow

1. Raw MIS Report files are downloaded from the GST e-way bill portal and dropped into `inward_eway bill/` and `outward_eway bill/` (some may actually be HTML tables saved with an `.xls` extension, which is normal for this portal's exports).
2. `convert_ewb_files.py` converts any `.xls` files in those two folders into proper `.xlsx` files, in place, using the HTML / openpyxl / CSV fallback chain described above.
3. `auto_ewb_merger.py` reads every `.xls`/`.xlsx` file in each of the two input folders, concatenates them (inward and outward kept separate), drops duplicate rows, and writes the merged results to `merge_ewb/`.
4. The always-current files `merge_ewb/inward_eway_bill_merged.xlsx` and `merge_ewb/outward_eway_bill_merged.xlsx` are the final consolidated E-Way Bill datasets meant to be picked up as input by the **main gst tool** for its scrutiny/analysis checks; the timestamped copies in the same folder serve as run-history snapshots.

Two ways to drive steps 2–3: `convert_and_merge.bat` (convert then merge-once, with the date-gated expiry check described above) or `merge_now.bat` (merge-once only, skips conversion). `start_merger.bat` instead leaves `auto_ewb_merger.py` running in continuous monitoring mode, auto-merging whenever new files are dropped into the input folders.
