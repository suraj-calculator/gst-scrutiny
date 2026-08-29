# alligner

## Overview

This folder holds a small utility, `complete_workbooks.py` (run via
`run_complete_workbooks.bat`), that "aligns"/"completes" a batch of GSTR-2B
Excel workbooks so they all contain the same set of worksheets, in the same
order. It is a data-preparation helper for the main GST scrutiny tool — it
does not itself analyze GST data, it only normalizes the structure of a set
of GSTR-2B export workbooks (produced upstream, e.g. by the `extractor`
folder or the GST portal) so that downstream tooling can rely on every
period's workbook having an identical sheet layout.

The folder currently contains 12 sample/working GSTR-2B workbooks named like
`MMYYYY_<GSTIN>_GSTR2B_<downloadDate>_summary.xlsx` (and some without the
`_summary` suffix). All 12 currently have the same 22 sheets, which is
consistent with them being the *result* of a prior run of this script (see
"Ambiguity" note below).

## Script details

### Purpose

`complete_workbooks.py` takes a set of `.xlsx` GSTR-2B workbooks (typically
one per return period for the same GSTIN) where some workbooks may be
missing worksheets that others have (e.g. a period with no ISD entries might
be missing the "ISD" sheet). It computes a "master" list of all sheet names
seen across the whole set, and adds any missing sheets to each workbook so
that every workbook ends up with the identical set of sheets, in the same
order.

### Inputs

- One or more `.xlsx` file paths passed as positional arguments, **or**
  `--folder <path>` to process every `.xlsx` file in a folder (the `.bat`
  file uses this mode, pointed at the alligner folder itself).
- Optional `--master-order <file>`: force the sheet order/reference file
  instead of auto-selecting it.
- Optional `--copy-sheets "Name1,Name2"`: comma-separated list of sheet
  names whose *content* (not just an empty sheet) should be copied in when
  missing, rather than added blank. Default: `"Read me"`.
- Based on the sample data, expected input filename pattern is
  `MMYYYY_<GSTIN>_GSTR2B_<downloadDate>[_summary].xlsx`, and files live
  directly in this folder (the `.bat` script's `TARGET_FOLDER`).

### Outputs

- **The script overwrites its input files in place.** There is no separate
  output file or `_completed.xlsx` copy produced — the same filenames stay,
  only their contents (worksheets/order) are updated. The script writes to
  a temporary file (`<original>.tmp_writing.xlsx`) first and then atomically
  replaces the original via `os.replace`.
- Because it overwrites in place, running it with no source backup means
  the pre-run version of each workbook is lost. The script's own docstring
  explicitly warns: take a backup first if you need one.
- Observed sheet set across all sample workbooks (22 sheets, standard
  GSTR-2B breakdown): `Read me`, `ITC Available`, `ITC not available`,
  `ITC Rejected`, `B2B`, `B2BA`, `B2B-CDNR`, `B2B-CDNRA`, `ECO`, `ECOA`,
  `ISD`, `ISDA`, `IMPG`, `IMPGA`, `IMPGSEZ`, `IMPGSEZA`, `B2B(Rejected)`,
  `B2BA(Rejected)`, `B2B-CDNR(Rejected)`, `B2B-CDNRA(Rejected)`,
  `ECO(Rejected)`, `ECOA(Rejected)`.

### Key logic / transformations

1. **Master sheet order**: `get_master_sheet_order()` scans every input
   file's sheet names. If `--master-order` is not given, it picks the file
   with the *most* sheets as the reference/base order (first such file wins
   ties), then appends any sheet names seen in other files but not in that
   base, preserving first-seen order.
2. **Per-file completion** (`complete_workbook()`): for each file, any sheet
   name in the master order that the file doesn't already have is created.
   - If the missing sheet's name is in the `--copy-sheets` list (default
     `"Read me"`) and the reference file actually has that sheet, its full
     content is copied over: cell values, font/fill/border/alignment/number
     format/protection, column widths, row heights, and merged-cell ranges
     (`copy_sheet_content()`).
   - Otherwise the missing sheet is added completely blank.
   - **Existing sheets are never touched** — even if a `copy-sheet` (e.g.
     "Read me") already exists but is empty, it is left as-is, not
     refreshed from the reference.
3. **Reordering**: after adding missing sheets, each workbook's sheet order
   (`wb._sheets`) is rewritten to match the master order exactly.
4. **Safe write**: saved to a `.tmp_writing.xlsx` temp file, then
   `os.replace()`d over the original to avoid partial/corrupt overwrites.

### Ambiguity / notes

- All 12 sample workbooks in this folder currently have identical (22)
  sheet counts. This is consistent with the script having already been run
  once over this folder (i.e., these appear to be already-"aligned"
  output), but nothing in the files themselves proves that — it's an
  inference from the data, not something confirmed by reading the code.
- The `_summary` vs. non-`_summary` filename suffix is not something this
  script produces, checks for, or otherwise treats specially — it appears
  to simply be inherited from however the files were originally named
  before being placed in this folder (e.g. from the GST portal export or an
  earlier merge step). The script itself is filename-pattern agnostic; it
  just processes whatever `.xlsx` files it's given.
- `run_complete_workbooks.bat` prints `DONE! "_completed.xlsx" files banayi
  ja chuki hain.` (i.e. claims `_completed.xlsx` files were created), but
  the Python script does **not** create any `_completed.xlsx` files — it
  overwrites originals in place. This appears to be a stale/inaccurate
  message in the `.bat` file left over from an earlier version of the
  script's behavior.

## How to run

### Option A: via the batch file (Windows, one-click)

1. Make sure `complete_workbooks.py` and `run_complete_workbooks.bat` are
   in the same folder as the `.xlsx` files to process (or edit the
   `TARGET_FOLDER` variable at the top of the `.bat` file to point at the
   correct folder — it's currently hardcoded to
   `C:\Users\admin\Documents\merger\files out\alligner`).
2. Double-click `run_complete_workbooks.bat`.
3. The script checks for Python on `PATH` and installs `openpyxl` via `pip`
   automatically if it's missing, then runs
   `python complete_workbooks.py --folder "%TARGET_FOLDER%"` over every
   `.xlsx` file in that folder.

### Option B: directly via Python (any OS)

```bash
# Process specific files
python complete_workbooks.py file1.xlsx file2.xlsx file3.xlsx

# Process every .xlsx in a folder
python complete_workbooks.py --folder /path/to/folder

# Optional: force a specific reference/order file
python complete_workbooks.py --folder /path/to/folder --master-order reference.xlsx

# Optional: change which sheets get content-copied instead of left blank
python complete_workbooks.py --folder /path/to/folder --copy-sheets "Read me,Instructions"
```

Requires Python 3 with `openpyxl` installed (`pip install openpyxl`).

**Warning**: this modifies files in place with no automatic backup — copy
the folder first if you want to preserve the pre-run versions.
