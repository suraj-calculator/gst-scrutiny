#!/usr/bin/env python3
"""
Full End-to-End Pipeline Runner
Runs the entire GST scrutiny tool flow:
1. Unzips & merges GSTR-1, GSTR-2A, GSTR-2B, GSTR-3B, E-Invoice
2. Stages all merged files, E-Way Bill files, Ledgers, and Master sheets
3. Executes master_build.py to produce the Master Scrutiny Workbook
4. Copies all results and intermediate workbooks into the output/ folder
"""
import glob
import os
import shutil
import sys
import time
import zipfile

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(REPO_ROOT, "input_files")
OUTPUT_DIR = os.path.join(REPO_ROOT, "output")
STAGING_DIR = os.path.join(REPO_ROOT, "_staging")
MAIN_TOOL_DIR = os.path.join(REPO_ROOT, "main gst tool")
FORMS_MERGER_DIR = os.path.join(REPO_ROOT, "forms merger", "merger-tool")


def log(msg):
    t = time.strftime("%H:%M:%S")
    print(f"[{t}] {msg}", flush=True)


def extract_zips(src_dir, dst_dir, prefix_with_zipname=False):
    os.makedirs(dst_dir, exist_ok=True)
    count = 0
    for fname in sorted(os.listdir(src_dir)):
        if fname.lower().endswith(".zip"):
            zpath = os.path.join(src_dir, fname)
            zstem = os.path.splitext(fname)[0]
            with zipfile.ZipFile(zpath, "r") as zf:
                for member in zf.namelist():
                    if member.lower().endswith((".xlsx", ".xls", ".xlsm")) and not member.endswith("/"):
                        out_name = f"{zstem}_{os.path.basename(member)}" if prefix_with_zipname else os.path.basename(member)
                        out_path = os.path.join(dst_dir, out_name)
                        with zf.open(member) as src, open(out_path, "wb") as dst:
                            dst.write(src.read())
                        count += 1
    return count


def run_merge_script(tool_subfolder, script_name, work_dir, out_filename):
    script_dir = os.path.join(FORMS_MERGER_DIR, tool_subfolder)
    shutil.copy(os.path.join(script_dir, "gst_merge_common.py"), work_dir)
    shutil.copy(os.path.join(script_dir, script_name), work_dir)

    for mod in list(sys.modules.keys()):
        if mod in ("gst_merge_common", script_name.replace(".py", "")):
            del sys.modules[mod]

    prev_sys_path = list(sys.path)
    if work_dir not in sys.path:
        sys.path.insert(0, work_dir)

    prev_cwd = os.getcwd()
    os.chdir(work_dir)
    try:
        mod = __import__(script_name.replace(".py", ""))
        mod.main(".")
    finally:
        os.chdir(prev_cwd)
        sys.path = prev_sys_path

    out_path = os.path.join(work_dir, out_filename)
    if not os.path.exists(out_path):
        raise RuntimeError(f"Merge failed: {out_filename} was not created in {work_dir}")
    return out_path


def main():
    start_time = time.time()
    log("=" * 60)
    log("STARTING FULL GST SCRUTINY PIPELINE")
    log(f"Input directory:  {INPUT_DIR}")
    log(f"Output directory: {OUTPUT_DIR}")
    log("=" * 60)

    # Clean / prepare output and staging directories
    shutil.rmtree(STAGING_DIR, ignore_errors=True)
    os.makedirs(STAGING_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. GSTR-1
    log("\n>>> Step 1/6: Processing GSTR-1...")
    g1_stage = os.path.join(STAGING_DIR, "gstr1")
    n_g1 = extract_zips(os.path.join(INPUT_DIR, "gstr 1"), g1_stage)
    log(f"  Extracted {n_g1} files. Merging...")
    g1_merged = run_merge_script("gstr1", "merge_gstr1.py", g1_stage, "GSTR1_Merged.xlsx")
    log(f"  GSTR-1 Merged OK: {os.path.getsize(g1_merged):,} bytes")

    # 2. GSTR-2A (R2A)
    log("\n>>> Step 2/6: Processing GSTR-2A (R2A)...")
    r2a_stage = os.path.join(STAGING_DIR, "r2a")
    n_r2a = extract_zips(os.path.join(INPUT_DIR, "2a"), r2a_stage)
    log(f"  Extracted {n_r2a} files. Merging...")
    r2a_merged = run_merge_script("gstr2a", "merge_r2a.py", r2a_stage, "R2A_Merged.xlsx")
    log(f"  GSTR-2A Merged OK: {os.path.getsize(r2a_merged):,} bytes")

    # 3. GSTR-2B
    log("\n>>> Step 3/6: Processing GSTR-2B...")
    g2b_stage = os.path.join(STAGING_DIR, "gstr2b")
    os.makedirs(g2b_stage, exist_ok=True)
    g2b_files = glob.glob(os.path.join(INPUT_DIR, "2b", "*.xlsx"))
    for f in g2b_files:
        shutil.copy(f, g2b_stage)
    log(f"  Copied {len(g2b_files)} files. Merging...")
    g2b_merged = run_merge_script("gstr2b", "merge_gstr2b.py", g2b_stage, "GSTR2B_Merged.xlsx")
    log(f"  GSTR-2B Merged OK: {os.path.getsize(g2b_merged):,} bytes")

    # 4. GSTR-3B
    log("\n>>> Step 4/6: Processing GSTR-3B...")
    g3b_stage = os.path.join(STAGING_DIR, "gstr3b")
    n_g3b = extract_zips(os.path.join(INPUT_DIR, "3b"), g3b_stage)
    log(f"  Extracted {n_g3b} files. Merging...")
    g3b_merged = run_merge_script("gstr3b", "merge_gstr3b.py", g3b_stage, "GSTR3B_Merged.xlsx")
    log(f"  GSTR-3B Merged OK: {os.path.getsize(g3b_merged):,} bytes")

    # 5. E-Invoice
    log("\n>>> Step 5/6: Processing E-Invoice...")
    einv_stage = os.path.join(STAGING_DIR, "einv")
    n_einv = extract_zips(os.path.join(INPUT_DIR, "e inv"), einv_stage, prefix_with_zipname=True)
    log(f"  Extracted {n_einv} files. Merging...")
    einv_merged = run_merge_script("e invoice", "merge_einv.py", einv_stage, "EINV_Merged.xlsx")
    log(f"  E-Invoice Merged OK: {os.path.getsize(einv_merged):,} bytes")

    # 6. Assemble Master Folder
    log("\n>>> Step 6/6: Assembling all inputs and running Master Scrutiny Engine...")
    master_stage = os.path.join(STAGING_DIR, "master_run")
    os.makedirs(master_stage, exist_ok=True)

    # Copy 5 merged returns to master folder & output
    merged_files = [g1_merged, r2a_merged, g2b_merged, g3b_merged, einv_merged]
    for m in merged_files:
        shutil.copy(m, master_stage)
        shutil.copy(m, OUTPUT_DIR)

    # Copy EWB files from input_files
    for ewb_name in ["inward_eway_bill_merged.xlsx", "outward_eway_bill_merged.xlsx"]:
        src = os.path.join(INPUT_DIR, ewb_name)
        if os.path.exists(src):
            shutil.copy(src, master_stage)
            shutil.copy(src, OUTPUT_DIR)
            log(f"  Copied EWB: {ewb_name}")

    # Copy all root-level Excel and CSV files from input_files
    for f in sorted(glob.glob(os.path.join(INPUT_DIR, "*"))):
        if os.path.isfile(f) and f.lower().endswith((".xlsx", ".csv", ".xlsm")):
            basename = os.path.basename(f)
            if not basename.startswith("inward_") and not basename.startswith("outward_"):
                shutil.copy(f, master_stage)
                log(f"  Copied input file: {basename}")

    # Run master_build.py
    log("\n>>> Executing master_build.py...")
    if MAIN_TOOL_DIR not in sys.path:
        sys.path.insert(0, MAIN_TOOL_DIR)

    prev_cwd = os.getcwd()
    os.chdir(OUTPUT_DIR)
    try:
        import master_build
        master_build.main(master_stage)
    finally:
        os.chdir(prev_cwd)

    # Find the generated master file in OUTPUT_DIR
    master_files = glob.glob(os.path.join(OUTPUT_DIR, "GST_MASTER_*.xlsx"))
    if not master_files:
        # Check if saved in master_stage or elsewhere
        staged_masters = glob.glob(os.path.join(master_stage, "GST_MASTER_*.xlsx"))
        if staged_masters:
            for sm in staged_masters:
                shutil.copy(sm, OUTPUT_DIR)
            master_files = glob.glob(os.path.join(OUTPUT_DIR, "GST_MASTER_*.xlsx"))

    # Cleanup staging directory
    shutil.rmtree(STAGING_DIR, ignore_errors=True)

    elapsed = time.time() - start_time
    log("\n" + "=" * 60)
    log(f"PIPELINE COMPLETED SUCCESSFULLY in {elapsed:.1f}s")
    log("Files in output directory:")
    for out_f in sorted(os.listdir(OUTPUT_DIR)):
        fp = os.path.join(OUTPUT_DIR, out_f)
        log(f"  {out_f} ({os.path.getsize(fp):,} bytes)")
    log("=" * 60)


if __name__ == "__main__":
    main()
