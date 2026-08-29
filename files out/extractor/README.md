# extractor

## Overview

This folder holds a small utility, `run.py` (run via `extract.bat`), that
extracts the Excel files bundled inside a set of GSTR-3B `.zip` packages
(as downloaded from the GST portal) into a single flat output folder named
`X`, and then deletes the original `.zip` files. It is a data-preparation
helper for the main GST scrutiny tool — it turns a pile of per-period ZIP
downloads into a flat folder of ready-to-use Excel files.

The folder currently contains 13 sample `.zip` files named like
`GSTR3B_<GSTIN>_<MMYYYY>.zip` (one has a `" (2)"` suffix, suggesting a
re-download/duplicate), plus an empty `X` subfolder — this is the
extraction destination folder that `run.py` creates/uses (currently empty,
consistent with the script not having been run yet against these zips, or
the zips being re-added after a previous run).

## Script details

### Purpose

`run.py` iterates over every `.zip` file in a base folder, opens each one,
finds any Excel files inside it (`.xlsx`, `.xls`, `.xlsm`), extracts just
those files (flattened, ignoring any internal folder structure) into a
single destination subfolder called `X`, and — once all zips have been
processed — deletes all the original `.zip` files from the base folder.

### Inputs

- All files ending in `.zip` (case-insensitive) directly inside a base
  folder, passed as an optional command-line argument
  (`python run.py [folder]`) or defaulting to the folder `run.py` itself
  lives in when no argument is given. It does not recurse into subfolders.
  (Previously this was a hardcoded Windows path pointing one level up, at
  the parent `files out` folder, instead of `extractor` itself — fixed so
  the script's default now matches where it actually lives and where the
  sample zips/`X` folder sit.)
- Based on the sample data, expected input filename pattern is
  `GSTR3B_<GSTIN>_<MMYYYY>.zip`, each presumably a GSTR-3B return package
  downloaded from the GST portal containing one or more Excel files.
- `extract.bat` still `cd`s to a hardcoded Windows path before invoking
  `python run.py` — see the batch-file note under "How to run" below.

### Outputs

- A destination folder named `X` (`DEST_FOLDER_NAME = "X"`), created inside
  the base folder if it doesn't already exist (`os.makedirs(dest_dir,
  exist_ok=True)`), containing every Excel file found inside every zip,
  using each file's original inner filename.
  - If two different zips contain Excel files with the same name, later
    ones are disambiguated with a numeric suffix before the extension
    (e.g. `report.xlsx`, `report_2.xlsx`, `report_3.xlsx`, ...) so nothing
    gets silently overwritten.
- **All original `.zip` files in the base folder are deleted** once extraction
  finishes (`os.remove`), regardless of whether they contained any Excel
  files — the intent per the script's own docstring is to leave only the
  `X` folder behind.
- A summary is printed to the console: count of Excel files extracted,
  count of zips deleted, any zip that had no Excel file inside it, and a
  final listing of everything in the `X` folder.
- Zips that fail to open (`zipfile.BadZipFile`, i.e. corrupt/invalid zip)
  are skipped for extraction but — per the code — are still included in the
  later delete-all-zips loop, since that loop iterates over the same
  `zip_files` list rather than only the successfully-processed ones.

### Key logic / transformations

1. Resolve the base folder to an absolute path and ensure the `X`
   destination folder exists.
2. List every `*.zip` file (case-insensitive match) directly in the base
   folder (sorted alphabetically for processing order).
3. For each zip: open it, filter its member list to entries ending in
   `.xlsx`/`.xls`/`.xlsm` that aren't directories, and for each such member,
   extract it into `X` using only the base filename (`os.path.basename`),
   deduplicating names with a `_2`, `_3`, ... suffix as needed via a
   `used_names` set.
4. After all zips are processed, delete every zip file that was found in
   step 2 (success or failure to extract).
5. Print a summary report.

### Ambiguity / notes

- The `X` folder is currently empty even though 13 `.zip` files are still
  present in this folder — this is only consistent with the script not
  having been run yet (or having been re-seeded with fresh zips after a
  prior run cleared them out); nothing in the code or file state proves
  which. This is worth confirming with whoever owns the workflow before
  assuming "not yet run."
- **Fixed**: `run.py` previously hardcoded `BASE_DIR` to the parent
  `files out` folder rather than to `extractor` itself, so running it (or
  `extract.bat`) actually operated one level up from where the sample zips
  and the `X` folder live. It now defaults to its own folder (or an
  explicit `python run.py <folder>` argument), verified end-to-end against
  a copy of two of the real sample zips in this folder.

## How to run

### Option A: via the batch file (Windows, one-click)

1. Double-click `extract.bat`.
2. It changes directory to wherever `extract.bat` itself is saved (`%~dp0`
   — previously a hardcoded Windows path, fixed alongside `run.py`) and
   runs `python run.py` there, then pauses so you can read the output.
3. Unlike `run_complete_workbooks.bat` in the `alligner` folder,
   `extract.bat` does **not** check for Python on `PATH` or auto-install
   any dependencies — it assumes Python 3 is already installed and on
   `PATH` (the script only uses the standard library — `os`, `zipfile` — so
   no extra packages are required).

### Option B: directly via Python (any OS)

```bash
python run.py            # defaults to the folder run.py itself is in
python run.py /path/to/folder   # or point it at any folder explicitly
```

Edit the `DEST_FOLDER_NAME` constant near the top of `run.py` to change the
name of the extraction subfolder.

**Warning**: this script deletes the original `.zip` files after
extraction with no backup or confirmation prompt — keep a copy of the
downloaded zips elsewhere first if you need to preserve them.
