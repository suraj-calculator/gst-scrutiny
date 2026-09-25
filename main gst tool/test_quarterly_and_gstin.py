#!/usr/bin/env python3
"""
Regression tests for (1) the spellings a quarterly return's Tax Period may carry and (2) the taxpayer-GSTIN fallback chain.

Run:  python3 test_quarterly_and_gstin.py
"""
import importlib.util
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import gst_core as mpu   # noqa: E402


def check(name, cond, detail=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    return cond


def main():
    ok = True
    q4 = ["Jan-24", "Feb-24", "Mar-24"]
    for tp in ("Jan-Mar", "Jan - Mar", "January-March", "January to March", "Jan-Mar 2024", "Quarter 4", "Q4", "Qtr-4", "Quarter 4 (Jan-Mar)", "(Jan - Mar)"):
        ok &= check(f"Tax Period {tp!r} -> Jan-24, Feb-24, Mar-24", mpu.months_for_tax_period("2023-24", tp) == q4)
    ok &= check("Q1 / Q2 / Q3", mpu.months_for_tax_period("2023-24", "Q1")[0] == "Apr-23" and mpu.months_for_tax_period("2023-24", "Quarter 2")[0] == "Jul-23"
                and mpu.months_for_tax_period("2023-24", "Q3")[0] == "Oct-23")
    ok &= check("existing spellings unchanged: month name, 'Mar', numeric MMYYYY", mpu.months_for_tax_period("2023-24", "April") == ["Apr-23"] and
                mpu.months_for_tax_period("2023-24", "Mar") == ["Mar-24"] and mpu.months_for_tax_period("2022-23", "042022") == ["Apr-22"])
    for bad in ("Bogus", "Jan-Feb", "Quarter 5", "Feb-Apr"):
        try:
            mpu.months_for_tax_period("2023-24", bad)
            ok &= check(f"{bad!r} is still rejected", False)
        except mpu.PeriodParseError:
            ok &= check(f"{bad!r} is still rejected", True)
    spec = importlib.util.spec_from_file_location("gmc", os.path.join(HERE, "..", "forms merger", "merger-tool", "gstr3b", "gst_merge_common.py"))
    gmc = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gmc)
    ok &= check("3B merger: 'Quarter 4' / 'Q4' sort as the quarter's first month (Jan = 10), 'Jan - Mar' as before",
                gmc.month_key("Quarter 4") == 10 and gmc.month_key("Q4") == 10 and gmc.month_key("Jan - Mar") == 10 and gmc.month_key("Apr-Jun") == 1 and gmc.month_key("Q1") == 1)

    tmp = tempfile.mkdtemp()
    try:
        led = os.path.join(tmp, "Electronic_Cash_Ledger.csv")
        with open(led, "w") as f:
            f.write("Electronic Cash Ledger\n\n,,,,,,GSTIN,05AAAAA0001A1Z1,\n,,,,,,From,01/04/2025,\n")
        wb = openpyxl.Workbook()
        wb.active.title = "Demographic Details"
        wb.active.append(["GSTIN", "05AAAAA0001A1Z1"])
        bo = os.path.join(tmp, "bo.xlsx")
        wb.save(bo)
        wb = openpyxl.Workbook()
        wb.active.append(["Goods and Services Tax", "GSTIN:   29ZZZZZ9999Z1Z9"])
        pc = os.path.join(tmp, "pc.xlsx")
        wb.save(pc)
        ok &= check("GSTIN read from a ledger CSV header", mpu._gstin_from_headers([led]) == "05AAAAA0001A1Z1")
        ok &= check("GSTIN read from a BO Profile 'GSTIN' label / value pair", mpu._gstin_from_headers([bo]) == "05AAAAA0001A1Z1")
        ok &= check("GSTIN read from a 'GSTIN:   <no>' title cell", mpu._gstin_from_headers([pc]) == "29ZZZZZ9999Z1Z9")
        ok &= check("most frequent GSTIN wins across files", mpu._gstin_from_headers([led, bo, pc]) == "05AAAAA0001A1Z1")
        ok &= check("nothing found / unreadable file -> None, no exception", mpu._gstin_from_headers([os.path.join(tmp, "missing.xlsx")]) is None)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
