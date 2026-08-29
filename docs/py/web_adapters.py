"""
Web adapters — the only new Python in this project's web port.

Every function here bridges "bytes the browser handed us" to the existing,
unmodified scripts in ../ewb, ../merge, ../pdf, ../extract_align, ../core —
by materialising those bytes into a real folder (in Pyodide's virtual
filesystem in the browser, or a plain temp folder under CPython for local
testing), calling the existing script's functions directly, and reading
whatever they wrote back out as bytes to hand back.

Every function here is written to run unchanged under plain CPython (for
fast local testing — see test_web_adapters.py) and under Pyodide in the
browser (see app.js). Nothing here is Pyodide-specific.
"""

import os
import sys
import glob

_HERE = os.path.dirname(os.path.abspath(__file__))
for _sub in ("ewb", "merge", "pdf", "extract_align", "core"):
    _p = os.path.join(_HERE, _sub)
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def process_ewb(inward_files, outward_files, work_dir):
    """
    inward_files / outward_files: list of (filename, bytes) tuples — the raw
    EWB_MIS_Report_Excel (N).xls downloads, one direction each.
    work_dir: an empty folder to do the work in (caller creates/cleans it).

    Returns a dict:
        {
          "inward":  {"rows": int, "output_name": str, "output_bytes": bytes} | None,
          "outward": {"rows": int, "output_name": str, "output_bytes": bytes} | None,
          "log": [str, ...],
        }
    "inward"/"outward" are None if that direction had no files.
    """
    import convert_ewb_files
    import auto_ewb_merger

    log = []
    inward_dir = os.path.join(work_dir, auto_ewb_merger.INWARD_FOLDER)
    outward_dir = os.path.join(work_dir, auto_ewb_merger.OUTWARD_FOLDER)
    os.makedirs(inward_dir, exist_ok=True)
    os.makedirs(outward_dir, exist_ok=True)

    for name, data in inward_files:
        with open(os.path.join(inward_dir, name), "wb") as f:
            f.write(data)
    for name, data in outward_files:
        with open(os.path.join(outward_dir, name), "wb") as f:
            f.write(data)

    prev_cwd = os.getcwd()
    os.chdir(work_dir)
    try:
        # Convert every raw .xls in both folders to real .xlsx first
        # (mirrors convert_ewb_files.main(), minus its input()/print noise).
        for folder in (auto_ewb_merger.INWARD_FOLDER, auto_ewb_merger.OUTWARD_FOLDER):
            for path in glob.glob(os.path.join(folder, "*.xls")):
                ok = convert_ewb_files.convert_file(path)
                log.append(f"convert {os.path.basename(path)}: {'ok' if ok else 'FAILED'}")

        auto_ewb_merger.merge_all_files()

        result = {"inward": None, "outward": None, "log": log}
        merge_dir = os.path.join(work_dir, auto_ewb_merger.MERGE_FOLDER)

        inward_out = os.path.join(merge_dir, "inward_eway_bill_merged.xlsx")
        if os.path.exists(inward_out):
            import openpyxl
            wb = openpyxl.load_workbook(inward_out, read_only=True)
            rows = wb.active.max_row - 1  # minus header
            wb.close()
            result["inward"] = {
                "rows": rows,
                "output_name": "inward_eway_bill_merged.xlsx",
                "output_bytes": _read_bytes(inward_out),
            }

        outward_out = os.path.join(merge_dir, "outward_eway_bill_merged.xlsx")
        if os.path.exists(outward_out):
            import openpyxl
            wb = openpyxl.load_workbook(outward_out, read_only=True)
            rows = wb.active.max_row - 1
            wb.close()
            result["outward"] = {
                "rows": rows,
                "output_name": "outward_eway_bill_merged.xlsx",
                "output_bytes": _read_bytes(outward_out),
            }

        return result
    finally:
        os.chdir(prev_cwd)
