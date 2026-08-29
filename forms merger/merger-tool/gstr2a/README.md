# gstr2a — GSTR-2A / R2A merger

## Overview

Merges GSTR-2A (also referred to as "R2A" in this script — the auto-drafted,
dynamic purchase-side ITC statement) workbooks — monthly, any financial
year, any count — into a single consolidated workbook, one continuous
month-wise block per data sheet.

## `gst_merge_common.py`

This folder's copy is the base shared-helper set plus one thing not shared
with `gstr1/`, `gstr2b/`, or `gstr3b/`:

- `load_data_workbook(path)` — always loads a workbook with
  `load_workbook(path, data_only=True)`, deliberately **not**
  `read_only=True`. The docstring explains why: some GSTN-portal-generated
  R2A files carry a stale `<dimension>` tag in their XML, which
  `read_only=True` trusts blindly for `max_row` — on a real file this
  reported `max_row=30` when the true last data row was 3403, silently
  dropping ~99% of the sheet's rows with no error. This function exists so
  every data-reading `load_workbook()` call goes through one place that
  can't accidentally regress to `read_only=True`.

Other differences from the "v2" copies (`gstr1`/`gstr2b`/`gstr3b`): this
copy does **not** have the `robust_read_meta`/label-fallback/
`detect_header_rows` helpers. Its `detect_file_type()` does add an `R2A`
branch (a workbook is `R2A` if it has `Read me`, `ECO`, and `TDS` sheets).

## `merge_r2a.py`

- **Input**: every `.xlsx` in the working folder that `detect_file_type()`
  identifies as `R2A`.
- **Meta fields**, read from each file's `Read me` sheet at fixed cells —
  no label-text fallback: `gstin`=`C2`, `tax_period`=`E2`,
  `legal_name`=`C3`, `fy`=`E3`, `trade_name`=`C4`, `generated_on`=`E4`.
  Workbooks are loaded with `load_data_workbook()`, not a plain
  `load_workbook()` call.
- **Sort order**: chronological via `einv_period_to_key(tax_period)`
  (expects `"MMYYYY"`).
- **Sheets merged**: assumes every input file has the **same** sheet set,
  taken from the first (chronologically earliest) file. It does not build
  a union — it only prints a `WARNING` if another file's sheet list
  differs, without adjusting the merge for that difference.
- **Header rows**: fixed at 6 for every sheet (`HEADER_ROWS = 6`: title,
  blank, blank, section title, then 2 rows of column headers) — the
  docstring states R2A's data sheets (B2B, B2BA, CDNR, CDNRA, ECO, ECOA,
  ISD, ISDA, TDS, TDSA, TCS, IMPG, IMPG SEZ) all share this layout and,
  unlike GSTR-2B, none of them change header text per period.
- **Merge logic**: for each data sheet, the header is copied once from the
  first file, then for every file a labelled separator row (`Financial
  Year | Tax Period | GSTIN | Date of Generation`) is written followed by
  that file's data rows down to the last non-empty row.
- The `Read me` sheet from the earliest period is copied in full, as-is.
- **Output**: `R2A_Merged.xlsx`, written into the current working
  directory.

## How to run

Put `merge_r2a.py` and this folder's `gst_merge_common.py` in the same
folder as the downloaded GSTR-2A/R2A `.xlsx` files, then run:

```
python merge_r2a.py
```

(Also invoked automatically as part of `run_all_gst_merge.bat`.)
