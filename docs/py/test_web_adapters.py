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


def load_dir(d, pattern="*.xls"):
    out = []
    for p in sorted(glob.glob(os.path.join(d, pattern))):
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


def test_process_merge():
    # Real per-period fixtures, one glob pattern each so the cross-type
    # *_Merged.xlsx files already sitting in these folders (see forms
    # merger/README.md's "Observed stray output files") aren't picked up.
    # gstr2b/gstr3b are NOT tested here — they need their real two-step
    # pipelines (align, extract), covered by test_process_gstr2b/gstr3b.
    cases = [
        ("gstr1", os.path.join(REPO, "forms merger/merger-tool/gstr1"), "GSTR1_*_Inv_1.xlsx"),
        ("einv", os.path.join(REPO, "forms merger/merger-tool/e invoice"), "EINV_*.xlsx"),
    ]
    for kind, folder, pattern in cases:
        files = load_dir(folder, pattern)
        assert files, f"no fixture files found for {kind} in {folder}"
        work_dir = os.path.join(HERE, f"_test_work_{kind}")
        shutil.rmtree(work_dir, ignore_errors=True)
        try:
            result = web_adapters.process_merge(kind, files, work_dir)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)
        assert result is not None, f"{kind}: expected output, got None ({len(files)} input files)"
        print(f"{kind}: {len(files)} input files -> {result['output_name']} ({len(result['output_bytes'])} bytes)")

    print("test_process_merge: PASSED (gstr1, einv)")
    print("NOTE: gstr2a/R2A has no sample fixture files in this repo — not exercised here.")


def test_process_gstr3b():
    # Real portal .zip bundles — exercises the extract-then-merge pipeline.
    files = load_dir(os.path.join(REPO, "files out/extractor"), "GSTR3B_*.zip")
    assert files, "no gstr3b zip fixtures found"
    work_dir = os.path.join(HERE, "_test_work_gstr3b_full")
    shutil.rmtree(work_dir, ignore_errors=True)
    try:
        result = web_adapters.process_gstr3b(files, work_dir)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    assert result is not None, f"expected output from {len(files)} zip fixtures"
    print(f"gstr3b (from zips): {len(files)} zips -> {result['output_name']} ({len(result['output_bytes'])} bytes)")
    print("test_process_gstr3b: PASSED")


def test_process_gstr2b():
    # Real per-period raw exports — exercises the align-then-merge pipeline.
    files = load_dir(os.path.join(REPO, "files out/alligner"), "*.xlsx")
    assert files, "no gstr2b alignment fixtures found"
    work_dir = os.path.join(HERE, "_test_work_gstr2b_full")
    shutil.rmtree(work_dir, ignore_errors=True)
    try:
        result = web_adapters.process_gstr2b(files, work_dir)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    assert result is not None, f"expected output from {len(files)} alignment fixtures"
    print(f"gstr2b (aligned): {len(files)} files -> {result['output_name']} ({len(result['output_bytes'])} bytes)")
    print("test_process_gstr2b: PASSED")


if __name__ == "__main__":
    test_process_ewb()
    test_process_merge()
    test_process_gstr3b()
    test_process_gstr2b()
