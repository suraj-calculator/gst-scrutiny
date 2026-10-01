#!/usr/bin/env python3
"""
Regression tests for QUARTERLY (QRMP) support: a quarter is analysed as ONE period labelled by its last month
(gst_core.PERIOD_UNITS), every reader answers that label with the quarter's TOTAL, and nothing is spread over months.

Self-contained (reuses the small synthetic fixtures of the E-Invoice / E-Way-Bill / GSTR-2B tests; no real data).
Run:  python3 test_quarterly_units.py
"""
import datetime as dt
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import gst_core as mpu                          # noqa: E402
import gst_checks_forensic as fc                # noqa: E402
import gst_parsers_returns as pr                # noqa: E402
import einv_adapter as einv                     # noqa: E402
import gstr2b_adapter as g2b                    # noqa: E402
import test_einv_adapter as te                  # noqa: E402
import test_gstr2b_adapter as t2b               # noqa: E402


def check(name, cond, detail=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    return cond


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)                                           # adapters write _canonical next to the cwd
    try:
        # ---- the registry
        mpu.set_period_units({})
        ok &= check("no units: every label is its own period", mpu.unit_months("Jun-25") == ["Jun-25"] and mpu.unit_of("Jun-25") is None
                    and mpu.period_title("Jun-25") == "Jun-25")
        mpu.set_period_units({"Jun-25": ["Apr-25", "May-25", "Jun-25"], "Sep-25": ["Sep-25"]})
        ok &= check("a quarter: the anchor answers with all three months; a one-month 'unit' is ignored",
                    mpu.unit_months("Jun-25") == ["Apr-25", "May-25", "Jun-25"] and mpu.unit_months("Sep-25") == ["Sep-25"]
                    and mpu.unit_of("May-25") == ("Jun-25", ["Apr-25", "May-25", "Jun-25"]))
        ok &= check("sheet titles say it is a quarter", "QUARTER Apr-25 to Jun-25" in mpu.period_title("Jun-25"))
        mpu.PARTIAL_UNITS["Jun-25"] = ["Apr-25"]
        ok &= check("an incomplete quarter is marked INCOMPLETE INPUT in its title", "INCOMPLETE INPUT: Apr-25" in mpu.period_title("Jun-25"))
        mpu.PARTIAL_UNITS.clear()

        # ---- E-Invoice: the quarter's label returns Apr + May (+ Jun) exactly
        raw = te.fixture(os.path.join(tmp, "einv.xlsx"))
        mpu.set_period_units({})
        einv.clear_cache()
        parts = [einv.parse_einv(raw, m) for m in ("Apr-25", "May-25", "Jun-25")]
        n_lines = sum(len(einv.read_einv_lines(raw, m)) for m in ("Apr-25", "May-25", "Jun-25"))
        mpu.set_period_units({"Jun-25": ["Apr-25", "May-25", "Jun-25"]})
        einv.clear_cache()
        q = einv.parse_einv(raw, "Jun-25")
        ok &= check("E-Invoice: quarter taxable / IGST / count = sum of its months",
                    abs(q["taxable"] - sum(p["taxable"] for p in parts)) < 1e-9 and abs(q["IGST"] - sum(p["IGST"] for p in parts)) < 1e-9
                    and q["count"] == sum(p["count"] for p in parts) and q["taxable"] > 0, str(q["taxable"]))
        ok &= check("E-Invoice: the quarter's invoice lines are all three months' lines",
                    len(einv.read_einv_lines(raw, "Jun-25")) == n_lines > 0)

        # ---- GSTR-2B: invoice rows of the quarter = rows of its months
        mpu.set_period_units({})
        r2b = t2b.make_fixture(os.path.join(tmp, "g2b.xlsx"))
        g2b.clear_cache()
        pa, pm = g2b.parse_month(r2b, "Apr-24"), g2b.parse_month(r2b, "May-24")
        mpu.set_period_units({"May-24": ["Apr-24", "May-24"]})
        g2b.clear_cache()
        pq = g2b.parse_month(r2b, "May-24")
        ok &= check("GSTR-2B: the quarter's invoice rows = April's + May's, none twice",
                    len(pq["b2b"]) == len(pa["b2b"]) + len(pm["b2b"]) and len(pq["b2b"]) > 0, str(len(pq["b2b"])))

        # ---- e-way bills filter by the quarter
        rows = [{"month": "Apr-25", "v": 1}, {"month": "May-25", "v": 2}, {"month": "Jun-25", "v": 3}, {"month": "Jul-25", "v": 4}]
        mpu.set_period_units({"Jun-25": ["Apr-25", "May-25", "Jun-25"]})
        ok &= check("E-way bills: the quarter's label returns Apr-Jun bills only",
                    [r["v"] for r in pr.filter_by_month(rows, "Jun-25")] == [1, 2, 3] and [r["v"] for r in pr.filter_by_month(rows, "Jul-25")] == [4])
        mpu.set_period_units({})
        ok &= check("...and a monthly filer is unaffected", [r["v"] for r in pr.filter_by_month(rows, "Jun-25")] == [3])

        # ---- QRMP due dates (GSTR-1: 13th; GSTR-3B: 22nd for Category X states, 24th for the rest)
        q_end = dt.date(2025, 9, 1)
        ok &= check("QRMP GSTR-1 is due on the 13th after the quarter, monthly on the 11th",
                    fc.due_date_gstr1(q_end, True) == dt.date(2025, 10, 13) and fc.due_date_gstr1(q_end, False) == dt.date(2025, 10, 11))
        ok &= check("QRMP GSTR-3B: Uttarakhand / Delhi / UP (Category Y) 24th; Maharashtra / Gujarat / Karnataka / Tamil Nadu (Category X) 22nd",
                    all(fc.due_date_gstr3b(q_end, True, c + "AAAAA0000A1Z5") == dt.date(2025, 10, 24) for c in ("05", "07", "09", "19"))
                    and all(fc.due_date_gstr3b(q_end, True, c + "AAAAA0000A1Z5") == dt.date(2025, 10, 22) for c in ("24", "27", "29", "33")))
        ok &= check("monthly GSTR-3B stays on the 20th", fc.due_date_gstr3b(q_end, False, "05AAAAA0000A1Z5") == dt.date(2025, 10, 20))
    finally:
        mpu.set_period_units({})
        os.chdir(cwd)
    print("ALL PASSED" if ok else "SOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
