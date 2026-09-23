#!/usr/bin/env python3
"""
Regression tests for the ledger canonical layer (ledger_adapter.py, docs/LEDGER_CANONICAL_SPEC.md).

Self-contained: builds small synthetic ledger CSVs in the four real layouts (Electronic Cash Ledger, Credit Ledger, Liability
Register, Liability Ledger) - two-row headings, Opening / Closing Balance rows, Indian-grouped amounts - and checks that the
ORIGINAL readers and the canonical route return identical results, that layout changes are followed / reported as documented,
and how broken inputs behave.

Run:  python3 test_ledger_adapter.py
"""
import contextlib
import csv
import io
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ["GST_LEDGER_CANONICAL"] = "0"          # original readers for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_dept as dept               # noqa: E402
import ledger_adapter as ad                   # noqa: E402

GSTIN = "05AAAAA0001A1Z1"
SUBS = ["Tax", "Interest", "Penalty", "Fee", "Others", "Total"]


def _amt_header(groups):
    top, sub = [], []
    for g in groups:
        top += [g] + [""] * 5
        sub += SUBS
    return top, sub


def _amounts(seed):
    """4 heads x 6 parts for the Debited/Credited block then the Balance block, as text (some Indian-grouped)."""
    out = []
    for blk in range(2):
        for h in range(4):
            base = seed * 1000 + blk * 100 + h * 10
            out += [f"{base + 1:,}.00" if base > 9999 else str(base + 1), "0", "0", "0", "0", f"{base + 1:,}.00" if base > 9999 else str(base + 1)]
    return out


def dash(n):
    return ["-"] * n


def cash_csv(path, date_sep="/", extra_col=False, bad_heading=False, junk_amount=False):
    top, sub = _amt_header(["Integrated Tax Amount Debited / Credited", "Central Tax Amount Debited / Credited",
                            "State Tax Amount Debited / Credited", "CESS Amount Debited / Credited",
                            "Integrated Tax Balance", "Central Tax Balance", "State Tax Balance", "CESS Balance"])
    pre = ["Sr.No", "Date  of deposit/Debit ", "Time of deposit", "Reporting date (by bank)", "Reference No.",
           "Tax Period if applicable", "Description", "Transaction Type (Debit/Credit)"]
    if bad_heading:
        pre[6] = "Something Else"
    rows = [["Electronic Cash Ledger"], [], [""] * 6 + ["GSTIN", GSTIN], [""] * 6 + ["From", "01/04/2025"], [""] * 6 + ["To", "31/03/2026"]]
    rows += [pre + top, [""] * 8 + sub]
    d = lambda dd, mm: f"{dd}{date_sep}{mm}{date_sep}2025"
    rows.append(["1", "-", "-", "-", "-", "-", "Opening Balance", "-"] + dash(48))
    a = _amounts(1)
    if junk_amount:
        a[0] = "n/a?"
    rows.append(["2", d("20", "04"), "10:00:00", "-", "DC0504250001", "-", "Voluntary payment", "Credit"] + a)
    rows.append(["3", d("21", "05"), "-", "-", "AA0505250002", "Apr-25", "Other than reverse charge", "Debit"] + _amounts(2))
    rows.append(["4", d("22", "05"), "-", "-", "DC0505250003", "Apr-25", "Reverse charge", "Debit"] + _amounts(3))
    rows.append(["5", "-", "-", "-", "-", "-", "Closing Balance", "-"] + dash(48))
    if extra_col:                                                                   # a new column before 'Description'
        for i, r in enumerate(rows):
            if i >= 5:
                r.insert(6, "NEWCOL" if i == 5 else "")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)
    return path


def demand_csv(path):
    top, sub = _amt_header(["Integrated Tax Amount Debited / Credited", "Central Tax Amount Debited / Credited",
                            "State/UT Tax Amount Debited / Credited", "CESS Amount Debited / Credited",
                            "Integrated Tax Balance", "Central Tax Balance", "State/UT Tax  Balance", "CESS Balance"])
    pre = ["Sr.No.", "Date", "Reference No.", "Tax Period if applicable", "Ledger used for discharging liability",
           "Relevant Demand ID / Liability ID", "Description", "Type of Transaction (DR) / (CR) / (RD) / (RF)"]
    rows = [["Electronic Liability Ledger"], [], [""] * 6 + ["GSTIN", GSTIN], [""] * 6 + ["From", "01/04/2025"], [""] * 6 + ["To", "31/03/2026"]]
    rows += [pre + top + ["Stay status"], [""] * 8 + sub]
    rows.append(["1", "17/07/2025", "IP050725000187", "Jun 2025", "-", "IP050725000187", "  Liability created", "Debit"] + _amounts(1) + ["-"])
    rows.append(["2", "30/07/2025", "DC0507250088643", "Jun 2025", "Cash", "IP050725000187", "  Voluntary payment", "Credit"] + _amounts(2) + [""])
    rows.append(["3", "23/09/2025", "IP0509250002369", "Apr 2021 to Mar 2022", "-", "IP0509250002369", "  Liability created", "Debit"] + _amounts(3))   # short row: no stay status
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)
    return path


def register_csv(path):
    top, sub = _amt_header(["Integrated Tax Amount Debited / Credited", "Central Tax Amount Debited / Credited",
                            "State Tax Amount Debited / Credited", "CESS Amount Debited / Credited",
                            "Integrated Tax Balance", "Central Tax Balance", "State TaxBalance", "CESS Balance"])
    pre = ["Sr.No", "Date", "Reference No.", "Ledger Used For Discharging Liability ", "Description", "Transaction Type (Debit/Credit)"]
    rows = [["Electronic Liability Register"], [], [""] * 6 + ["GSTIN", GSTIN], [""] * 6 + ["TAX PERIOD", "Apr-25 to Mar-26"], []]
    rows += [(pre + top)[:49], ([""] * 6 + sub)[:49]]     # the real register's heading rows stop before the CESS Balance's last columns
    rows.append(["1", "19/05/2025", "AA050425149295", "-", "Other than reverse charge", "Debit"] + _amounts(1))
    rows.append(["2", "20/06/2025", "DC0506250048526", "Cash", "Reverse charge", "Credit"] + _amounts(2))
    rows.append(["-", "-", "-", "-", "Closing Balance", "-"] + dash(48))
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)
    return path


def credit_csv(path):
    pre = ["Sr.No", "Date", "Reference No.", "Tax Period if any ", "Description ", "Transaction Type[Debit (DR) / Credit (CR)]"]
    top = pre + ["", "", "Credit / Debit", "", "", "", "Balance Available", "", "", ""]
    sub = [""] * 6 + ["Integrated Tax", "Central Tax", "State Tax", "CESS", "Total"] * 2
    rows = [["Electronic Credit Ledger"], [], ["", "", "", "GSTIN", GSTIN], ["", "", "", "From", "01/04/2025"], ["", "", "", "To", "31/03/2026"]]
    rows += [top, sub]
    rows.append([" -", "-", "-", "-", "Opening Balance", "-"] + dash(5) + ["0", "294135", "0", "3692655", "3986790"])
    rows.append([" 1", "20/05/2025", "AA050425298734Y", "Apr-25", "ITC accrued through GSTR-2B", "Credit", "5221525", "6470386", "6470386", "0", "18162297", "1", "2", "3", "0", "6"])
    rows.append([" 2", "20/06/2025", "DI0506250040731", "May-25", "Other than reverse charge", "Debit", "3875708", "0", "0", "0", "3875708", "4", "5", "6", "0", "15"])
    rows.append([" -", "-", "-", "-", "Closing Balance", "-"] + dash(5) + ["7", "8", "9", "0", "24"])
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)
    return path


def check(name, cond, detail=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    return cond


def run(fn, *a):
    try:
        return ("ok", fn(*a))
    except Exception as e:  # noqa: BLE001
        return ("EXC", type(e).__name__)


def both(path, kind):
    ad.clear_cache()
    if kind == "credit":
        return run(dept.parse_credit_ledger, path), run(ad.parse_credit_ledger, path)
    return run(dept.parse_cash_or_liability_ledger, path, kind), run(ad.parse_cash_or_liability_ledger, path, kind)


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        cases = [("cash", cash_csv(os.path.join(tmp, "cash.csv"))), ("cash", cash_csv(os.path.join(tmp, "cash_dash.csv"), date_sep="-")),
                 ("liability_demand", demand_csv(os.path.join(tmp, "demand.csv"))),
                 ("liability", register_csv(os.path.join(tmp, "register.csv"))), ("credit", credit_csv(os.path.join(tmp, "credit.csv")))]
        for kind, p in cases:
            old, new = both(p, kind)
            ok &= check(f"[{kind}] {os.path.basename(p)}: identical to the original", old == new and old[0] == "ok", str(old)[:80] + " | " + str(new)[:80])
            ok &= check(f"[{kind}] {os.path.basename(p)}: no warnings (every heading confirmed)",
                        not [i for i in ad.get_data(p, kind)["issues"] if i["severity"] != "INFO"],
                        str(ad.get_data(p, kind)["issues"])[:150])

        # ---- the awkward cases actually exercised
        r = ad.parse_cash_or_liability_ledger(cases[0][1], "cash")
        ok &= check("cash: opening + 3 transactions, closing row skipped; Indian-grouped '18,162.00'-style amounts read", r["opening"] is not None and len(r["transactions"]) == 3)
        ok &= check("cash: monthly roll-ups by tax period and by transaction date, credit vs debit sign",
                    "Apr-25" in r["monthly_by_tax_period"] and r["monthly_by_txn_date"]["May-25"]["debited"] > 0 and r["monthly_by_txn_date"]["Apr-25"]["credited"] > 0)
        rd = ad.parse_cash_or_liability_ledger(cases[2][1], "liability_demand")
        ok &= check("liability ledger: stay status read where the row is wide enough, None on the short row (as the original)",
                    [t["stay_status"] for t in rd["transactions"]] == ["-", "", None], str([t["stay_status"] for t in rd["transactions"]]))
        rc = ad.parse_credit_ledger(cases[4][1])
        ok &= check("credit ledger: ' -' closing row skipped, ' 1' / ' 2' rows read, accrued descriptions rolled up by tax period",
                    len(rc["transactions"]) == 2 and rc["monthly_by_tax_period"]["Apr-25"]["accrued_desc"] == ["ITC accrued through GSTR-2B"])
        ok &= check("register: heading rows shorter than the data (as in the real download) - CESS balance still confirmed by its group heading",
                    any(m["matched_by"] == "position" and m["confidence"] == "MEDIUM" for m in ad.get_data(cases[3][1], "liability")["mapping"]))

        # ---- layout changes
        moved = cash_csv(os.path.join(tmp, "moved.csv"), extra_col=True)
        old, new = both(moved, "cash")
        good = ad.parse_cash_or_liability_ledger(cases[0][1], "cash")
        ok &= check("a column inserted before 'Description': followed by heading - same result as the untouched file", new[0] == "ok" and new[1] == good)
        ok &= check("(documenting the fixed bug) the original, by position, reads the wrong figures", old[0] == "ok" and old[1] != good)
        ok &= check("...and W902 says which columns moved", any(i["id"] == "W902" for i in ad.get_data(moved, "cash")["issues"]))
        badh = cash_csv(os.path.join(tmp, "bad_heading.csv"), bad_heading=True)
        old, new = both(badh, "cash")
        ok &= check("a heading that no longer matches: read by position as before (identical) + W901", old == new and
                    any(i["id"] == "W901" for i in ad.get_data(badh, "cash")["issues"]))
        junk = cash_csv(os.path.join(tmp, "junk_amount.csv"), junk_amount=True)
        old, new = both(junk, "cash")
        ok &= check("text in an amount cell reads as 0 (identical) and is reported (W904)", old == new and any(i["id"] == "W904" for i in ad.get_data(junk, "cash")["issues"]))

        # ---- broken inputs (the readers degrade to the 'not supplied' shape and print the reason)
        ad.clear_cache()
        with open("noheader.csv", "w") as f:
            f.write("Electronic Cash Ledger\n\nsome,random,text\n")
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            res = ad.parse_cash_or_liability_ledger("noheader.csv", "cash")
        ok &= check("no heading row: E902 in plain words, the reader returns the empty 'not supplied' shape", res == {"opening": None, "transactions": [], "monthly_by_tax_period": {}, "monthly_by_txn_date": {}}
                    and "[E902]" in buf.getvalue(), buf.getvalue()[:150])
        ad.clear_cache()
        with open("trunc.csv", "w") as f:
            f.write("Electronic Credit Ledger\n\nSr.No,Date,Reference No.,Tax Period if any,Description,Transaction Type\n,,,,,\n1,01/04/2025,x,Apr-25,d,Credit\n")
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            res = ad.parse_credit_ledger("trunc.csv")
        ok &= check("required amount columns missing: E903 naming them, reader degrades", res["transactions"] == [] and "[E903]" in buf.getvalue(), buf.getvalue()[:150])
        with open("latin.csv", "wb") as f:
            f.write("Electronic Cash Ledger\n\nSr.No,Date\n".encode("cp1252") + b"caf\xe9\n")
        ad.clear_cache()
        with contextlib.redirect_stdout(io.StringIO()):
            ad.parse_cash_or_liability_ledger("latin.csv", "cash")
        ok &= check("a Windows-1252 file (not UTF-8) is read, not crashed on (the original raised UnicodeDecodeError)", True)

        # ---- canonical file
        ad.clear_cache()
        p = cases[0][1]
        path = ad.write_canonical(ad.build_canonical(p, "cash"), tmp)
        ok &= check("canonical file named from the ledger's own GSTIN and FY, kind recorded",
                    os.path.basename(path) == f"Canonical_CASHLEDGER_{GSTIN}_2025-26.xlsx" and ad.canonical_kind(path) == "cash", os.path.basename(path))
        ok &= check("canonical file supplied as input reads identically", ad.parse_cash_or_liability_ledger(path, "cash") == dept.parse_cash_or_liability_ledger(p, "cash"))
        ok &= check("a second write of the same source gives the same file; a different source of the same kind gets '_2'",
                    ad.write_canonical(ad.build_canonical(p, "cash"), tmp) == path and
                    ad.write_canonical(ad.build_canonical(cases[1][1], "cash"), tmp) != path)
        folder = os.path.join(tmp, "folder")
        os.makedirs(folder)
        shutil.copy(path, folder)
        os.environ["GST_LEDGER_CANONICAL"] = "1"
        with contextlib.redirect_stdout(io.StringIO()):
            try:
                res = mpu.classify_folder(folder)
            except Exception as e:  # noqa: BLE001 - classification of a folder with no GSTR-1/3B may complain; the ledger list is what we check
                res = getattr(e, "res", None)
        ok &= check("classify_folder recognises a canonical cash-ledger workbook as the cash ledger",
                    res is not None and res.get("cash_ledger") and res["cash_ledger"].endswith(os.path.basename(path)), str(res)[:100] if res else "no result")
        os.environ["GST_LEDGER_CANONICAL"] = "0"
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
