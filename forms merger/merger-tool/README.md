# merger-tool

Root of the actual GST-return merger tool. Each return type has its own
subfolder (a self-contained copy of `gst_merge_common.py` plus one
`merge_*.py` script); this folder additionally holds the master
orchestrator batch file.

## `run_all_gst_merge.bat` — master orchestrator

Runs the merge script for every return type, one after another, from a
single double-click on Windows. For each return type it:

1. `pushd`s into `%BASE_DIR%\<subfolder>` (e.g. `...\merger-tool\gstr1`),
2. runs `python <merge_script>.py` there (so the script's own `find_xlsx_files(".")`
   picks up the `.xlsx` files sitting in that subfolder),
3. reports `[OK]`/`[FAIL]` for that step, then `popd`s and moves to the next.

It calls, in this order:

| Subfolder | Script |
|---|---|
| `gstr1` | `merge_gstr1.py` |
| `gstr2a` | `merge_r2a.py` |
| `gstr2b` | `merge_gstr2b.py` |
| `gstr3b` | `merge_gstr3b.py` |
| `e invoice` | `merge_einv.py` |

Note: the file's own header comment says "runs all 4 merge scripts", but
the script body actually calls all 5 listed above.

`BASE_DIR` is hard-coded near the top of the file
(`C:\Users\admin\Documents\merger\forms merger\merger-tool`) and must be
edited to match wherever this tree actually lives before the batch file
will find anything. It also checks `python` is on `PATH` first and aborts
with a message if not.

## `desktop.ini`

A Windows Explorer folder-metadata file, not part of the tool.

## Sub-tools

- [`gstr1/`](gstr1/README.md) — GSTR-1
- [`gstr2a/`](gstr2a/README.md) — GSTR-2A / R2A
- [`gstr2b/`](gstr2b/README.md) — GSTR-2B
- [`gstr3b/`](gstr3b/README.md) — GSTR-3B
- [`e invoice/`](e%20invoice/README.md) — E-Invoice
