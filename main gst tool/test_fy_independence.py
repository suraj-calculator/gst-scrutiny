#!/usr/bin/env python3
"""
Regression tests: the HSN / fraud checks must not depend on which financial year is scrutinised.

Background: gst_checks_hsn_fraud.py used to hard-code FY 2022-23 (a month list, a "March 2023" test,
BO FY "2022-23", three 2022-23 holidays). For any other year, month ordering became arbitrary (and
changed run to run) and checks #8 / #12 reported "normal"/PASS whatever the data said.

Run:  python3 test_fy_independence.py
"""
import datetime as dt
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import gst_checks_hsn_fraud as h   # noqa: E402

FY2425 = ["Apr-24", "May-24", "Jun-24", "Jul-24", "Aug-24", "Sep-24", "Oct-24", "Nov-24", "Dec-24", "Jan-25",
          "Feb-25", "Mar-25"]


def check(name, cond, detail=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    return cond


def inv(no, d, tax, gstin="G1"):
    return dict(sheet="b2b, sez, de_inv", invno=no, invdate=d, taxable=tax, gstin=gstin)


def main():
    ok = True

    # ---- month key: chronological for any year, unparseable last, works across FY boundaries
    ok &= check("FY 2024-25 labels sort chronologically", sorted(reversed(FY2425), key=h._month_key) == FY2425)
    ok &= check("FY 2022-23 labels sort chronologically",
                sorted(["Mar-23", "Apr-22", "Jan-23", "Dec-22"], key=h._month_key) ==
                ["Apr-22", "Dec-22", "Jan-23", "Mar-23"])
    ok &= check("two FYs sort across the boundary",
                sorted(["Apr-24", "Mar-24", "Mar-25", "Jan-24"], key=h._month_key) ==
                ["Jan-24", "Mar-24", "Apr-24", "Mar-25"])
    ok &= check("unparseable label sorts last", sorted(["junk", "Apr-24", ""], key=h._month_key)[0] == "Apr-24")
    ok &= check("month gap is calendar months (also across a year end)",
                h._month_gap("Mar-25", "Oct-24") == 5 and h._month_gap("Apr-24", "Nov-23") == 5)
    ok &= check("month gap with an unparseable label is 0 (as before)", h._month_gap("Mar-25", "junk") == 0)
    ok &= check("FY of month", (h._fy_of_month("Apr-24"), h._fy_of_month("Mar-25"), h._fy_of_month("Jan-23"),
                                h._fy_of_month("x")) == ("2024-25", "2024-25", "2022-23", None))

    # ---- determinism: same input, different hash seeds -> identical findings (this was the real bug)
    probe = (
        "import sys, json; sys.path.insert(0, %r); import gst_checks_hsn_fraud as h\n"
        "M = %r\n"
        "hsn = {m: [dict(hsn='8708' if i == 0 else ('8703' if i < 4 else '9999'), taxable=1000.0 + i, rate=28.0, "
        "desc='x', qty=1)] for i, m in enumerate(M)}\n"
        "hsn = dict(reversed(list(hsn.items())))\n"
        "fs = h.check_hsn_drift(hsn)\n"
        "print(json.dumps([[f.ref, f.severity, f.detail] for f in fs]))\n" % (HERE, FY2425))
    outs = []
    for seed in ("0", "1", "2", "3"):
        r = subprocess.run([sys.executable, "-c", probe], capture_output=True, text=True,
                           env=dict(os.environ, PYTHONHASHSEED=seed))
        outs.append(r.stdout.strip())
    ok &= check("drift check: identical findings under 4 different hash seeds", len(set(outs)) == 1 and outs[0] != "",
                str([o[:80] for o in outs]))
    first_new = json.loads(outs[0])
    ok &= check("drift check: April (first month) is not reported as 'new mid-year HSN'",
                not any("Apr-24" in f[2] and f[0] == "#27" for f in first_new))

    # ---- #8 year-end concentration for FY 2024-25 (used to report 0% -> "normal" whatever the data)
    g1 = {m: [inv(f"I{i}", dt.date(2024 + (i >= 9), (i + 3) % 12 + 1, 5), 1000.0)] for i, m in enumerate(FY2425[:-1])}
    g1["Mar-25"] = [inv("BIG", dt.date(2025, 3, 25), 90000.0)]
    f8 = h.check_year_end_dumping(g1)[0]
    ok &= check("#8 flags a March-2025 concentration", f8.severity == h.FLAG, f8.detail[:100])

    # ---- #12 cross-FY shift for 31-Mar-2025
    g1["Mar-25"].append(inv("X1", dt.date(2025, 3, 31), 500.0))
    ewb = [dict(docno="X1", ewbno="E9", ewbdate=dt.date(2025, 4, 1))]
    f12 = h.check_cross_fy_shift(g1, ewb)
    ok &= check("#12 finds 31-Mar-2025 invoice with an EWB dated 1-Apr-2025",
                f12[0].severity == h.FLAG and "31-Mar-2025" in f12[0].detail)
    ok &= check("#12 PASS text names the real year",
                "31-Mar-2025" in h.check_cross_fy_shift(g1, [])[0].detail)
    ok &= check("#12 with no March in the data says so (INFO, not a false PASS)",
                h.check_cross_fy_shift({"Apr-24": []}, [])[0].severity == h.INFO)

    # ---- #25 reads the BO row of the FY actually scrutinised
    bo = {"financial_by_fy": {"2024-25": dict(turnover=100.0, taxable_turnover=80.0)}}
    f3b = {m: dict(itc_reversed_rule42_43=0.0) for m in FY2425}
    f25 = h.check_exempt_turnover_rule42(bo, f3b)
    ok &= check("#25 uses the FY 2024-25 BO row", len(f25) == 1 and "FY2024-25" in f25[0].detail
                and f25[0].severity == h.REVW, f25[0].detail[:100])
    ok &= check("#25 says so when that FY's BO row is missing",
                "FY2023-24 row not found" in h.check_exempt_turnover_rule42(bo, {"Apr-23": dict(itc_reversed_rule42_43=0.0)})[0].detail)

    # ---- #37 the same fixed holidays in every year
    days = [dict(ewbdate=dt.date(2024, 8, 15)), dict(ewbdate=dt.date(2025, 1, 26)), dict(ewbdate=dt.date(2024, 10, 2)),
            dict(ewbdate=dt.date(2024, 9, 3))]
    f37 = h.check_sunday_holiday_ewb(days, [])[0]
    ok &= check("#37 counts 15-Aug-2024, 26-Jan-2025 and 2-Oct-2024 as holidays", "3 of 4" in f37.detail, f37.detail[:90])

    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
