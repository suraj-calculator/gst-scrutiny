"""
Local CPython smoke tests for web_adapters.py, run against the real sample
files already sitting in this repo (never committed — see .gitignore).

Not part of the shipped site. This is the fast local iteration loop from
LOCAL_SETUP.md Step 6: fix something here under plain CPython (seconds per
run) before pointing app.js at the same function inside Pyodide (much
slower to iterate on, since it's a full browser reload each time).

Requires: pip install pandas openpyxl lxml html5lib xlrd
(a throwaway venv is fine — this script is never shipped)

Usage:
    python3 test_web_adapters.py
"""
import glob
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))  # docs/py -> docs -> repo root
sys.path.insert(0, HERE)

import web_adapters  # noqa: E402


def load_dir(d):
    out = []
    for p in sorted(glob.glob(os.path.join(d, "*.xls"))):
        with open(p, "rb") as f:
            out.append((os.path.basename(p), f.read()))
    return out


def test_process_ewb():
    inward_files = load_dir(os.path.join(REPO, "e_way bill", "inward_eway bill"))
    outward_files = load_dir(os.path.join(REPO, "e_way bill", "outward_eway bill"))
    print(f"loaded {len(inward_files)} inward .xls, {len(outward_files)} outward .xls")

    work_dir = os.path.join(HERE, "_test_work_ewb")
    shutil.rmtree(work_dir, ignore_errors=True)
    os.makedirs(work_dir)
    try:
        result = web_adapters.process_ewb(inward_files, outward_files, work_dir)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    for direction in ("inward", "outward"):
        r = result[direction]
        assert r is not None, f"expected {direction} output"
        assert r["rows"] > 0, f"expected rows in {direction} output"
        print(f"{direction}: {r['rows']} rows -> {r['output_name']} ({len(r['output_bytes'])} bytes)")

    print("test_process_ewb: PASSED")


if __name__ == "__main__":
    test_process_ewb()
