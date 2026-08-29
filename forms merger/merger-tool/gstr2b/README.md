# gstr2b — GSTR-2B merger

## Overview

Merges GSTR-2B (static, auto-drafted monthly/quarterly ITC statement)
workbooks — monthly, quarterly, or a mix across years — into a single
consolidated workbook, one continuous block per data sheet, handling both
"detailed" and portal "summary" (`_summary.xlsx`) export variants.

## `gst_merge_common.py`

This folder's copy (253 lines, "v2") matches the `gstr1/` and `gstr3b/`
copies: the base shared-helper set plus the defensive helpers
(`find_label_cell`, `dump_region`, `value_after_label`, `label_lookup`,
`robust_read_meta`, `looks_like_fy`, `looks_like_tax_period`). It has no
license-expiry lock (removed from this copy) and its `detect_file_type()`
has no `R2A` branch (only `EINV`/`GSTR1`/`GSTR3B`/`GSTR2B`).

## `merge_gstr2b.py`

Its own docstring documents three rounds of hardening (labelled v2–v4):

- **v2**: switched from "use the first file's sheet list for everyone" to
  a union of sheets across all files, because the portal's
  `_summary.xlsx` variant only contains a subset of sheets (it drops the
  line-item detail sheets) and mixing a detailed file with a summary file
  used to crash with a `KeyError`.
- **v3**: added per-sheet header-row auto-detection (from merged cells)
  instead of a hardcoded per-sheet-name row count, to cope with sheets the
  portal added later (ECO, ECOA, IMPGA, IMPGSEZA, ITC Rejected, and six
  `(Rejected)` sheets from the IMS rollout).
- **v4**: switched `read_meta()` to `robust_read_meta()` instead of
  trusting fixed cells outright, since GSTR-3B has already shown the
  portal will shift a filing template's header layout mid-year and the
  `Read me` sheet here uses the same label + merged-value layout.

Details:

- **Input**: every `.xlsx` in the working folder that `detect_file_type()`
  identifies as `GSTR2B` (has `Read me` and `ITC Available` sheets).
- **Meta fields**, via `robust_read_meta()` against `FIXED_CELLS` on the
  `Read me` sheet: `fy`=`C4`, `tax_period`=`C5`, `gstin`=`C6`,
  `legal_name`=`C7`, `generated_on`=`C9`, with label-text fallbacks and
  validators on `fy`/`tax_period`. Files whose meta can't be resolved are
  skipped (header block dumped to console).
- **Sort order**: `(fy_start_year(fy), month_key(tax_period))` — quarterly
  labels like `"Apr-Jun"` sort by their start month.
- **Sheets merged**: union of every sheet name across all files (a
  `_summary.xlsx` period naturally has fewer sheets — missing sheets are
  simply skipped for that period, with a console note that it's "likely a
  `_summary.xlsx` period"). `KNOWN_SHEETS` is used only to print a
  heads-up for unfamiliar sheet names.
- **Two merge strategies per sheet**:
  - `REPEAT_HEADER_SHEETS = {"ITC Available", "ITC not available", "ITC
    Rejected"}` — each period's own full header (rows 1 through its
    detected header) is kept with its block, because a quarterly file's
    header shows the actual month names for that quarter (e.g.
    April/May/June), which differ from file to file.
  - All other sheets (`merge_static_header_sheet`) — header copied once
    from the first file that has the sheet; header-row count auto-detected
    via `detect_header_rows()` (the highest row touched by a merged cell
    near the top of the sheet, since the portal merges header cells for
    grouped columns but never merges data-row cells).
  - In both cases, each file's block is preceded by a labelled separator
    row (`Financial Year | Tax Period | Date of Generation`).
- The `Read me` sheet from the chronologically earliest file is copied in
  full as its own tab.
- **Output**: `GSTR2B_Merged.xlsx`, written into the current working
  directory.

## How to run

Put `merge_gstr2b.py` and this folder's `gst_merge_common.py` in the same
folder as the downloaded GSTR-2B `.xlsx` files (detailed and/or summary),
then run:

```
python merge_gstr2b.py
```

(Also invoked automatically as part of `run_all_gst_merge.bat`.)
