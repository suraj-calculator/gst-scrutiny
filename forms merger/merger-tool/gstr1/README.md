# gstr1 — GSTR-1 merger

## Overview

Merges regular (portal-downloaded, not e-invoice-drafted) GSTR-1 workbooks
— outward-supply returns, one file per tax period — into a single
consolidated workbook, one continuous month-wise block per data sheet.

## `gst_merge_common.py`

This folder's copy (307 lines, labelled "v2" in its own docstring) is the
base shared-helper set **plus** a set of defensive helpers added because a
single "reference file's sheet list + fixed header row count" assumption
broke on real filings. It has no license-expiry lock (that code is absent
from this copy). Notable contents:

- Base helpers: `find_xlsx_files`, `detect_file_type` (`EINV`/`GSTR1`/
  `GSTR3B`/`GSTR2B` only — no `R2A` branch), `fy_start_year`, `month_key`,
  `einv_period_to_key`, `warn_duplicates`, `sheet_max_data_row`,
  `write_separator`, `copy_sheet_full`.
- `find_label_cell`, `dump_region`, `value_after_label`, `label_lookup` —
  scan a sheet's top-left block for a label's text and return the value
  next to it, as a fallback when a fixed cell address no longer holds what
  is expected (e.g. because a template's rows/columns shifted).
- `robust_read_meta(ws, fixed_cells, label_fallbacks, path,
  key_validators)` — tries the fixed cell addresses first, falls back to
  label-text search per field if a validator rejects the fixed-cell value,
  and returns `None` (printing a diagnostic dump of the sheet) if a
  required field still can't be found.
- `looks_like_fy`, `looks_like_tax_period` — validators used by
  `robust_read_meta` for the `fy` and `tax_period` fields.
- `detect_header_rows(ws, ...)` — finds the header/data boundary by
  scanning for the first row containing a cell that looks like real GST
  data (matches a GSTIN pattern, a date pattern, or is numeric/date-typed),
  rather than assuming a fixed row count.

## `merge_gstr1.py`

Its own docstring documents this as a "v2 rebuild": the previous version
assumed every input file shares the same sheet set and a fixed 4-row
header, which broke when a mid-year portal template change split the
`hsn` sheet into `hsn(b2b)` and `hsn(b2c)`.

- **Input**: every `.xlsx` in the working folder (excluding the output
  file itself, `GSTR1_Merged.xlsx`) that `detect_file_type()` identifies
  as `GSTR1` (has a `Read me` sheet and a `b2cl` sheet).
- **Meta fields**, read via `robust_read_meta()` against fixed cells on the
  `Read me` sheet: `fy`=`C4`, `tax_period`=`C5`, `gstin`=`C6`,
  `legal_name`=`C7`, `arn`=`C9`, `arn_date`=`C10`, `generated_on`=`C11`,
  with label-text fallbacks (`Financial Year`/`Year`/`FY`, `Tax Period`,
  `GSTIN`, etc.) and validators on `fy`/`tax_period`. A file whose meta
  can't be resolved is skipped (its header block is dumped to the console)
  rather than aborting the whole run.
- **Sort order**: `(fy_start_year(fy), month_key(tax_period))`.
- **Sheets merged**: the **union** of every sheet name seen across all
  input files (excluding `Read me`) — not just the first file's sheets. A
  `KNOWN_SHEETS` set is used only to print a heads-up when an unfamiliar
  sheet name shows up (e.g. a future portal change), never to filter
  anything out. If a given sheet is missing from some files (e.g. the
  `hsn`/`hsn(b2b)`/`hsn(b2c)` split), those files are simply skipped for
  that sheet, with a console note.
- **Header rows**: auto-detected per sheet via `detect_header_rows()`
  (data-signature based, since GSTR-1's headers are flat rows with no
  merged cells to rely on).
- **Merge logic**: for each sheet, the header block is copied once from
  the first file that has that sheet; then for every file containing that
  sheet, a labelled separator row (`Financial Year | Tax Period | ARN |
  ARN Date | Date and Time of Generation`) is written, followed by that
  file's data rows down to the last non-empty row.
- The original `Read me` sheet from the chronologically earliest file is
  copied in full (via `copy_sheet_full`) into the output as its own
  `Read me` tab.
- **Output**: `GSTR1_Merged.xlsx`, written into the current working
  directory.

## How to run

Put `merge_gstr1.py` and this folder's `gst_merge_common.py` in the same
folder as the downloaded `GSTR1_*.xlsx` files, then run:

```
python merge_gstr1.py
```

(Also invoked automatically as part of `run_all_gst_merge.bat`.)
