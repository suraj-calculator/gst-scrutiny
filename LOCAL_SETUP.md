# Getting it running locally — setup proposal

Scope: get the site working **on your own machine** with real Python logic
wired to the UI (the mockup you already reviewed), before touching
deployment. This is the concrete first slice of Phase 1–2 from
[`WEBSITE_APPROACH.md`](WEBSITE_APPROACH.md).

## The non-obvious parts, upfront

Four things aren't "just open the HTML file":

1. **Pyodide won't run from a plain `file://` double-click.** It loads its
   runtime and packages as WebAssembly/fetch requests, which browsers block
   under the `file://` origin. You need *any* local static file server — no
   framework, just something serving the folder over `http://localhost`.
2. **First load needs internet.** Pyodide's runtime and every package
   (`openpyxl`, `pdfplumber`, ...) are fetched from a CDN the first time,
   the same way the mockup's Google Font is. There's no fully-offline local
   dev unless we later vendor the wheel files ourselves — not needed yet.
3. **One real unknown to resolve before building more UI**: whether
   `pdfplumber` (→ its dependency `pypdfium2`, which has compiled/native
   code) actually runs in Pyodide. This has to be answered with a 20-line
   throwaway test page, not assumed either way.
4. **The scripts assume a folder on disk; the browser gives you `File`
   objects.** Every script here already scans a folder — that pattern
   actually transfers cleanly to Pyodide's in-memory filesystem (write
   uploaded bytes to a virtual folder, then call the existing function
   unchanged) — but a few scripts need small, specific fixes first (table
   below). None of this is a rewrite; it's each script's entry point taking
   a folder argument instead of hardcoding one.

## Prerequisites

- Python 3 (already on this machine) — used only to serve static files
  locally, e.g. `python3 -m http.server 8000`. No new install.
- A current Chromium- or Firefox-based browser (WebAssembly + ES modules).
- Internet access for the browser tab (CDN-fetched Pyodide + packages).
- Nothing else. No Node, no npm, no build tool, no framework — deliberately,
  per the "minimal ongoing maintenance" goal from the approach doc.

## Per-script fixes needed before wiring (found by reading the actual code)

| Script | Today | Fix needed for the web adapter |
|---|---|---|
| `files out/extractor/run.py` | Hardcoded `BASE_DIR = r"C:\Users\admin\Documents\merger\files out"`; entry point is `extract_excel_from_zips()` with no folder argument | Add a folder parameter (this was already flagged earlier as a live path bug, independent of the web port — worth fixing either way) |
| `files out/alligner/complete_workbooks.py` | Already clean: `main()` supports `--folder`, and `complete_workbook(path, ...)` is a real function taking one file at a time. Mutates files **in place**. | No logic change — the adapter just needs to read the file back out of the virtual FS after calling it, since there's no separate output filename |
| `e_way bill/auto_ewb_merger.py` | `if __name__ == "__main__":` branches on `--once` vs a `while True:` 5-second-poll continuous-watch mode | Adapter must always invoke the equivalent of `--once` — the watch-loop mode has no meaning inside a browser tab |
| `forms merger/merger-tool/*/merge_*.py` (5 scripts) | Already clean: every one is `def main(folder="."):` | No change — call directly with the virtual folder path |
| `pdf to excel/gst_pdf_to_excel.py` | `def main():` reads its input path from `sys.argv` | Needs a plain callable form, e.g. `convert(pdf_bytes) -> xlsx_bytes`, alongside the existing CLI `main()` (which stays for local/offline use) |
| `main gst tool/master_build.py` | `def main(folder="."):`, also reads `sys.argv` for the CLI case | No change — this is the last step, called on the fully-assembled virtual folder |

Everything else (the four PDF `parse_*.py` files, the shared
`gst_merge_common.py` copies, all the `gst_checks_*.py`/`gst_parsers_*.py`
modules) is called *through* the entry points above and needs no direct
change.

## Step-by-step

**Step 1 — repo hygiene (5 min).**
Before vendoring anything into the web folder, remove the two dead-code
duplicates already flagged (stale `merger-tool/merge_gstr3b.py`, stray
`forms merger/gst_merge_common.py`) so there's only one copy of each script
to point the site at. *(Not deleted yet — say the word and I'll do it.)*

**Step 2 — skeleton folder + local server.**
```
docs/
  index.html      ← reuse the mockup's HTML/CSS as the shell
  app.js          ← replaces the mockup's setTimeout-based fake processing
  py/
    core/          → symlink to main gst tool/*.py
    merge/         → symlink to forms merger/merger-tool/*/*.py
    ewb/           → symlink to e_way bill/*.py
    pdf/           → symlink to pdf to excel/*.py
    extract_align/ → symlink to files out/*/*.py
    web_adapters.py ← new file, the only genuinely new Python
```
Symlinks (not copies) keep one source of truth — editing `main gst
tool/gst_core.py` is instantly reflected in the site, no sync step.
Run `python3 -m http.server 8000` from `docs/` and confirm the mockup page
still loads at `http://localhost:8000`.

**Step 3 — prove Pyodide loads and can see the repo's own files.**
Add the Pyodide `<script>` tag, call `loadPyodide()`, and from the browser
console run something as simple as reading one vendored `.py` file's
docstring through Pyodide's FS. This confirms the local server + CDN +
symlink setup all actually work together before any real logic is involved.

**Step 4 — install `openpyxl` via micropip, sanity-check it.**
Every merge/align script depends on it. Confirm `import openpyxl` and a
trivial `Workbook()` round-trip work inside the Pyodide session.

**Step 5 — the PDF spike (do this before building more sections).**
Bare test page: `micropip.install("pdfplumber")`, then try opening the
sample PDF already sitting in the repo
(`pdf to excel/05ASQPB9012R1ZA_BO_Profile_08_08_2026.pdf`). Two outcomes:
- **Works** → proceed with `pdf to excel/` as planned in Phase 5.
- **Fails on `pypdfium2`** → note it here and decide then: drop PDF inputs
  from v1, or isolate just those 4 parsers behind one small serverless
  function later. Either way, nothing else in the plan is blocked by this.

**Step 6 — write `web_adapters.py`, test it under plain CPython first.**
Before touching Pyodide/the browser at all, write and run the adapter
functions locally with regular `python3` against the sample files already
in this repo (`e_way bill/inward_eway bill/*.xls`,
`files out/alligner/*.xlsx`, `files out/extractor/*.zip`, the sample PDF).
This is the fast iteration loop — a CPython script run/fix/rerun cycle is
seconds, a browser-reload-and-recheck-console cycle is much slower. Only
once an adapter function works correctly under CPython does it get pointed
at from `app.js` inside Pyodide.

**Step 7 — wire one real section into the mockup UI.**
Replace the E-Way Bill section's fake `setTimeout` in the mockup's `app.js`
with a real call: write dropped files into Pyodide's virtual FS at
`/work/ewb_inward/`, call the (now-fixed) `auto_ewb_merger` equivalent in
`--once` mode, read `merge_ewb/inward_eway_bill_merged.xlsx` back out, offer
it as a real download. This is the "walking skeleton" — once this one
section works end-to-end in the browser, the rest are the same pattern
repeated per Phase 3–6 of the approach doc.

**Step 8 — decide main-thread vs. Web Worker.**
Start on the main thread (simpler to build and debug — browser devtools see
everything directly). If a real merge/scrutiny run visibly freezes the tab's
UI while it works, that's the point to move the Pyodide session into a Web
Worker — a real but well-understood browser pattern, worth deferring until
it's an actual problem rather than building it speculatively.

**Step 9 — repeat Step 7's pattern per section**, in the order from the
approach doc (E-Invoice → GSTR-1 → 2A → 2B (+ alligner) → 3B (+ extractor,
after its Step-Table fix) → PDFs (contingent on Step 5) → main gst tool
itself), each one testable locally the whole way, before any deployment
work starts.

## What "done, locally" looks like

Every section in the mockup does its real job against real uploaded files,
served from `http://localhost:8000`, with no server involved anywhere —
just `python3 -m http.server` and a browser. At that point deployment is
close to a non-event: point GitHub Pages at the same `docs/` folder (Phase
8 of `WEBSITE_APPROACH.md`) and swap the symlinks for real copies at that
point, since GitHub won't dereference symlinks across the repo the same way
during a Pages build.

## Before Step 1 — your call

- OK to delete the two dead-code duplicates now, or leave them for later?
- OK to fix the `files out/extractor/run.py` hardcoded path now (small,
  independent fix), or fold it into Step 9 when that section gets wired up?
