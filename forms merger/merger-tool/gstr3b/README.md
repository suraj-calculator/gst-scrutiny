# gstr3b — GSTR-3B merger

## Overview

Merges GSTR-3B (summary return) workbooks into a single consolidated
workbook. Unlike the other four mergers, GSTR-3B files don't hold
transaction-level rows to append together — each source file is a single
`GSTR-3B` sheet of fixed-position summary figures — so this script instead
gives every source period its own **fully-formatted tab** in the output
workbook, ordered left to right by month.

## `gst_merge_common.py`

This folder's copy (253 lines, "v2") is identical in kind to the `gstr1/`
and `gstr2b/` copies: the base shared-helper set plus the defensive
helpers (`find_label_cell`, `dump_region`, `value_after_label`,
`label_lookup`, `robust_read_meta`, `looks_like_fy`, `looks_like_tax_period`).
It has no license-expiry lock. Note: `merge_gstr3b.py` in this folder does
**not** actually import or use `robust_read_meta()` / the label-fallback
machinery — it has its own, simpler two-column (E/G) fallback instead (see
below). The main functions it does use from this module are
`find_xlsx_files`, `detect_file_type`, `fy_start_year`, `month_key`, and
`copy_sheet_full`.

## `merge_gstr3b.py`

- **Input**: every `.xlsx` in the working folder that `detect_file_type()`
  identifies as `GSTR3B` (the workbook's only sheet is named exactly
  `GSTR-3B`).
- **Meta fields**, read via a local `_first_nonempty(ws, cells)` helper
  that checks a **list** of candidate cells per field and returns the
  first non-blank one — specifically column **E** first, then column
  **G**, for the same row: `gstin`=`E5`/`G5`, `fy`=`E6`/`G6`,
  `tax_period`=`E7`/`G7`, `arn`=`E8`/`G8`, `arn_date`=`E9`/`G9`,
  `legal_name`=`E10`/`G10`, `generated_on`=`E12`/`G12`. This handles a
  portal template revision that shifted these values from column E to
  column G in some filing periods.
- **Validation**: if both `fy` and `tax_period` are still empty after
  checking both columns, the script raises a `ValueError` (with a message
  telling the user to check cells `E6`/`G6` and `E7`/`G7` in that file's
  `GSTR-3B` sheet) rather than silently mis-sorting or crashing later with
  a less informative error.
- **Sort order**: `(fy_start_year(fy), month_key(tax_period))`.
- **Merge logic**: no row-appending — for each source file, its entire
  `GSTR-3B` sheet (values, styles, merges, column/row dimensions) is
  copied via `copy_sheet_full()` into a **new tab** of the output
  workbook, named `"{tax_period}_{fy}"` (any `/` replaced with `-`); if
  that name is already used (truncated to Excel's 31-character sheet-name
  limit) a numeric suffix (`_2`, `_3`, …) is appended. Tabs end up ordered
  left to right by the chronological sort key.
- **Output**: `GSTR3B_Merged.xlsx`, written into the current working
  directory.

### Relationship to `merger-tool/merge_gstr3b.py`

`merger-tool/merge_gstr3b.py` (one directory up, not inside a `gstrN`
folder) is an **older, simpler version of this exact script**: same
overall algorithm, same output filename and tab-naming scheme, but its
`read_meta()` only reads fixed column-E cells (`E5`…`E12`) with no
column-G fallback and no `ValueError` check for a missing `fy`/
`tax_period`. `run_all_gst_merge.bat` runs **this** folder's script, not
the top-level one — see `merger-tool/README.md` for more detail.

## How to run

Put `merge_gstr3b.py` and this folder's `gst_merge_common.py` in the same
folder as the downloaded `GSTR3B_*.xlsx` files, then run:

```
python merge_gstr3b.py
```

(Also invoked automatically as part of `run_all_gst_merge.bat`.)
