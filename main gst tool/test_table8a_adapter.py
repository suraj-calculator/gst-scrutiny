#!/usr/bin/env python3
"""
Regression tests for the Table 8A canonical layer (table8a_adapter.py, docs/TABLE8A_CANONICAL_SPEC.md).

Self-contained: builds synthetic Table 8A workbooks in the real layout (title, two-row heading, B2B and CDNR / B2B-CDNR sheets, a leading
'Tax Period' column or not) and checks that the ORIGINAL parse_table_8a() and the canonical route return identical results, and how
broken layouts / files are reported.

Run:  python3 test_table8a_adapter.py
"""
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ["GST_T8A_CANONICAL"] = "0"          # original reader for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_dept as dept               # noqa: E402
import table8a_adapter as ad                  # noqa: E402

G1, G2, G3 = "07BBBBB0002B1Z2", "09CCCCC0003C1Z3", "29DDDDD0004D1Z4"


def b2b_top(period_col):
    top = ["GSTIN of supplier", "Trade/Legal name", "Invoice Details", None, None, None, "Place of supply", "Supply Attract Reverse Charge",
           "Rate(%)", "Taxable value (₹)", "Tax amount", None, None, None, "GSTR-1/IFF/GSTR-5 Filing Period", "GSTR-1/IFF/GSTR-5 Filing Date",
           "ITC Availability", "Reason"]
    sub = [None, None, "Invoice number", "Invoice type", "Invoice Date", "Invoice Value(₹)", None, None, None, None, "Integrated tax(₹)",
           "Central tax(₹)", "State/UT tax(₹)", "Cess(₹)", None, None, None, None]
    if period_col:
        top, sub = ["Tax Period"] + top, [None] + sub
    return top, sub


def inv(gstin, no, itc, reason=None, igst=0.0, cgst=90.0, sgst=90.0, period=None):
    r = [gstin, "Some Supplier", no, "R", "25/04/2023", 1180.0, "Uttarakhand", "N", 18.0, 1000.0, igst, cgst, sgst, 0.0, "APR-23", "11/05/2023", itc, reason]
    return ([period] + r) if period else r


def make(path, period_col=False, cdnr_name="CDNR", drop_gstin_head=False, no_header=False, drop_cdnr_gstin=False):
    wb = openpyxl.Workbook()
    wb.active.title = "Read me"
    wb.active.append(["Goods and Services Tax - GSTR 2A"])
    wb.active.append(["Financial Year", "2023-24"])
    wb.active.append(["GSTIN", "05AAAAA0001A1Z1"])
    wb.active.append(["Legal Name", "SAMPLE TAXPAYER"])
    ws = wb.create_sheet("B2B")
    top, sub = b2b_top(period_col)
    if drop_gstin_head:
        sub[1 if period_col else 0] = "Supplier ID"     # the sub-heading wins, so the column is no longer found - but the heading ROW still is
    ws.append(["   "])
    ws.append([])
    ws.append([])
    ws.append(["Taxable inward supplies received from registered persons"])
    if not no_header:
        ws.append(top)
        ws.append(sub)
    else:
        ws.append(["Something else"])
    p = "Apr-23" if period_col else None
    for r in (inv(G1, "A1", "Yes", period=p), inv(G2, "A2", "No", "Supplier not filed", period=p),
              inv(G3, "A3", "No", "Supplier not filed", period=p), inv("BADGSTIN", "A4", "Yes", period=p),
              inv(G1, "A5", "yes", igst=180.0, cgst=0.0, sgst=0.0, period=p), inv(G2, "A6", "No", period=p)):
        ws.append(r)
    wc = wb.create_sheet(cdnr_name)
    ctop = ["GSTIN of supplier", "Trade/Legal name", "Credit note/Debit note details", None, None, None, None, "Place of supply",
            "Supply attract reverse charge", "Rate(%)", "Taxable Value (₹)", "Tax Amount", None, None, None, "ITC Availability"]
    csub = [None, None, "Note number", "Note type", "Note supply type", "Note date", "Note value (₹)", None, None, None, None,
            "Integrated Tax(₹)", "Central Tax(₹)", "State/UT Tax(₹)", "Cess(₹)", None]
    if drop_cdnr_gstin:
        csub[0] = "Supplier ID"
    for _ in range(3):
        wc.append([])
    wc.append(["Debit/Credit notes"])
    wc.append(ctop)
    wc.append(csub)
    wc.append([G1, "Some Supplier", "CN1", "C", "R", "30/04/2023", 500.0, "Uttarakhand", "N", 18.0, 400.0, 0.0, 36.0, 36.0, 0.0, "Yes"])
    wb.save(path)
    return path


def check(name, cond, detail=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    return cond


def run(fn, *a):
    try:
        return ("ok", fn(*a))
    except Exception as e:  # noqa: BLE001
        return ("EXC", type(e).__name__)


def both(p):
    ad.clear_cache()
    return run(dept.parse_table_8a, p), run(ad.parse_table_8a, p)


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        for label, kw in (("no Tax Period column", {}), ("leading Tax Period column", {"period_col": True}), ("B2B-CDNR sheet name", {"cdnr_name": "B2B-CDNR"})):
            p = make(os.path.join(tmp, f"t8a_{abs(hash(label))}.xlsx"), **kw)
            old, new = both(p)
            ok &= check(f"[{label}] identical to the original", old == new and old[0] == "ok", str(old)[:70] + " | " + str(new)[:70])
            ok &= check(f"[{label}] no warnings", not [i for i in ad.get_data(p)["issues"] if i["severity"] != "INFO"])
        p = make(os.path.join(tmp, "std.xlsx"), period_col=True)
        r = ad.parse_table_8a(p)
        ok &= check("only GSTIN-shaped rows are data (the bad-GSTIN row and the wrapped sub-heading row are dropped): 5 invoices, 1 note", len(r["b2b"]) == 5 and len(r["cdnr"]) == 1, str(len(r["b2b"])))
        ok &= check("period column read when the export has one", r["b2b"][0]["period"] == "Apr-23")
        ok &= check("totals: only 'ITC availability = Yes' (case-insensitive) invoices: 2 rows, CGST 90, IGST 180",
                    r["totals"]["b2b_yes_rows"] == 2 and r["totals"]["cgst"] == 90.0 and r["totals"]["igst"] == 180.0, str(r["totals"]))
        ok &= check("'No' rows bucketed by reason, blank shown as (blank)", r["totals"]["no_reason_breakdown"] == {"Supplier not filed": 2, "(blank)": 1}, str(r["totals"]["no_reason_breakdown"]))
        ok &= check("dates kept as in the source", r["b2b"][0]["invdate"] == "25/04/2023" and r["cdnr"][0]["note_date"] == "30/04/2023")
        ok &= check("Read me GSTIN / FY in the canonical META",
                    dict(ad.get_data(p)["meta"])["recipient_gstin"] == "05AAAAA0001A1Z1" and dict(ad.get_data(p)["meta"])["fy"] == "2023-24")

        # ---- broken layouts (identical reasons to the original) + plain warnings
        for label, kw, frag, wid in (("GSTIN column heading changed", {"drop_gstin_head": True}, "could not locate the 'GSTIN of supplier' column", "W1102"),
                                     ("no heading row at all", {"no_header": True}, "no 'GSTIN of supplier' header row located", "W1101"),
                                     ("CDNR GSTIN heading changed", {"drop_cdnr_gstin": True}, "'CDNR' sheet: could not locate", "W1102")):
            p2 = make(os.path.join(tmp, f"bad_{wid}_{len(frag)}.xlsx"), **kw)
            old, new = both(p2)
            ok &= check(f"{label}: same result and reason as the original, and reported ({wid})", old == new and frag in (new[1]["reason"] or "") and
                        any(i["id"] == wid for i in ad.get_data(p2)["issues"]), str(new[1]["reason"]) if new[0] == "ok" else str(new))
        with open("junk.xlsx", "w") as f:
            f.write("not an excel file")
        ad.clear_cache()
        rj = ad.parse_table_8a("junk.xlsx")
        ok &= check("unreadable file: available=False, 'Could not read Table 8A workbook' (E1101 in plain words)", not rj["available"] and "Could not read Table 8A workbook" in rj["reason"] and "[E1101]" in rj["reason"], str(rj["reason"])[:100])
        ok &= check("path not supplied: the original's exact reason", ad.parse_table_8a(None)["reason"] == "Table 8A not supplied for this taxpayer/FY.")

        # ---- canonical file
        ad.clear_cache()
        path = ad.write_canonical(ad.build_canonical(p if False else os.path.join(tmp, "std.xlsx")), tmp)
        ok &= check("canonical file named from the Read me GSTIN and FY", os.path.basename(path) == "Canonical_TABLE8A_05AAAAA0001A1Z1_2023-24.xlsx", os.path.basename(path))
        ok &= check("canonical file supplied as input reads identically", ad.parse_table_8a(path) == dept.parse_table_8a(os.path.join(tmp, "std.xlsx")))
        wb = openpyxl.load_workbook(path, read_only=True)
        ok &= check("sheets: META, T8A_B2B, T8A_CDNR, MAPPING_REPORT, ISSUES", wb.sheetnames == ["META", "T8A_B2B", "T8A_CDNR", "MAPPING_REPORT", "ISSUES"], str(wb.sheetnames))
        wb.close()
        folder = os.path.join(tmp, "folder")
        os.makedirs(folder)
        shutil.copy(path, folder)
        os.environ["GST_T8A_CANONICAL"] = "1"
        import contextlib
        import io
        with contextlib.redirect_stdout(io.StringIO()):
            try:
                res = mpu.classify_folder(folder)
            except Exception as e:  # noqa: BLE001
                res = getattr(e, "res", None)
        ok &= check("classify_folder recognises a canonical Table 8A workbook", res is not None and res.get("table8a_files") == [os.path.join(folder, os.path.basename(path))], str(res.get("table8a_files") if res else None))
        os.environ["GST_T8A_CANONICAL"] = "0"
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
