# e invoice — E-Invoice merger

## Overview

Merges auto-drafted GSTR-1-from-E-Invoice workbooks (the "E-Invoice" excel
the GST portal generates from e-invoice data, one per tax period) into a
single consolidated workbook, one continuous month-wise block per data
sheet.

## `gst_merge_common.py`

This folder's copy (137 lines) is the base shared-helper set, without the
"v2" defensive header/meta-detection helpers found in the `gstr1/`,
`gstr2b/`, and `gstr3b/` copies, and without the license-expiry lock found
in the top-level and `gstr2a/` copies. It provides:

- `find_xlsx_files(folder)` — lists `*.xlsx` in a folder.
- `detect_file_type(path)` — identifies a workbook as `EINV`, `GSTR1`,
  `GSTR3B`, or `GSTR2B` by its sheet-name signature (this copy has no `R2A`
  branch). A workbook is `EINV` if it has both a `Read me` sheet and a
  `b2b, sez, de` sheet.
- `fy_start_year`, `month_key`, `einv_period_to_key` — turn a financial
  year string / tax-period string into sortable keys.
- `warn_duplicates(records)` — prints a warning if two input files resolve
  to the same sort key.
- `sheet_max_data_row(ws, min_row)` — last non-empty row in a sheet.
- `write_separator(ws_out, row_idx, text, n_cols)` — writes a bold,
  yellow-filled, merged separator row.
- `copy_sheet_full(src_ws, dst_wb, title)` — copies an entire worksheet
  (values, styles, merges, dimensions) into a new sheet.

## `merge_einv.py`

- **Input**: every `.xlsx` file directly in the current working folder
  that `detect_file_type()` identifies as `EINV` (detected by sheet
  content, not filename).
- **Meta fields**, read from each file's `Read me` sheet at fixed cells:
  `fy`=`C4`, `tax_period`=`C5`, `gstin`=`C6`, `legal_name`=`C7`,
  `date_updated_till`=`C9`. No fallback if these cells are wrong/shifted.
- **Sort order**: chronological, via `einv_period_to_key(tax_period)`,
  which expects `tax_period` as an `"MMYYYY"` string (e.g. `"042023"`).
- **Sheets merged**: a fixed list, `DATA_SHEETS = ["b2b, sez, de", "cdnr",
  "cdnur", "exp"]` (the `Read me` sheet itself is dropped). This is **not**
  a union across files — the header (first 4 rows, `HEADER_ROWS = 4`) is
  taken from the first (chronologically earliest) file only and reused for
  every sheet.
- **Merge logic**: for each of the 4 sheets, appends one block per source
  file — a labelled separator row (`Financial Year | Tax Period | Date
  Updated till`) followed by that file's data rows (row 5 through the last
  non-empty row, via `sheet_max_data_row`).
- **Output**: `EINV_Merged.xlsx`, written into the current working
  directory (i.e. the same folder the script is run from).

## How to run

Put `merge_einv.py` and this folder's `gst_merge_common.py` in the same
folder as the downloaded `EINV_*.xlsx` files, then run:

```
python merge_einv.py
```

(Also invoked automatically as part of `run_all_gst_merge.bat`.)
