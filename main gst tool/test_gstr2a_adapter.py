#!/usr/bin/env python3
"""
Regression tests for the GSTR-2A canonical layer (gstr2a_adapter.py, docs/GSTR2A_CANONICAL_SPEC.md).

Self-contained: builds small synthetic merged GSTR-2A workbooks in the real layout (wrapped two-row headings, a period
banner per month, rate-lines plus '-Total' rollup rows, repeated sub-heading rows on the amendment sheets, all five
parsed sheets) and checks that the ORIGINAL parse_r2a_excel() and the canonical route return identical results, that
layout changes are followed / reported as documented, and the deliberate differences.

Run:  python3 test_gstr2a_adapter.py
"""
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ["GST_R2A_CANONICAL"] = "0"          # original reader for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_dept as dept               # noqa: E402
import gstr2a_adapter as ad                   # noqa: E402

G1, G2, G3 = "07BBBBB0002B1Z2", "09CCCCC0003C1Z3", "29DDDDD0004D1Z4"
TAXG = ["Tax Amount", None, None, None]
B2B_G = ["GSTIN of supplier", "Trade/Legal name of the Supplier", "Invoice details", None, None, None, "Place of supply",
         "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Tax Amount", None, None, None,
         "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period",
         "GSTR-3B Filing Status", "Amendment made, if any", "Tax Period in which Amended", "Effective date of cancellation",
         "Source", "IRN", "IRN date"]
B2B_S = [None, None, "Invoice number", "Invoice type", "Invoice Date", "Invoice Value (₹)", None, None, None, None,
         "Integrated Tax  (₹)", "Central Tax (₹)", "State/UT tax (₹)", "Cess  (₹)"]
CDNR_G = ["GSTIN of Supplier", "Trade/Legal name of the supplier", "Credit note/Debit note details", None, None, None, None,
          "Place of supply", "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Tax Amount", None, None, None,
          "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period",
          "GSTR-3B Filing Status", "Amendment made, if any", "Tax Period in which Amended", "Effective date of cancellation",
          "Source", "IRN", "IRN date"]
CDNR_S = [None, None, "Note type", "Note number", "Note Supply type ", "Note  date", "Note Value (₹)", None, None, None, None,
          "Integrated Tax (₹)", "Central Tax (₹)", "State Tax (₹)", "Cess Amount (₹)"]
B2BA_A = ["Original details", None, "Revised details"]
B2BA_G = ["Invoice number", "Invoice Date", "GSTIN of Supplier", "Trade/Legal name of the supplier", "Invoice details", None, None,
          None, "Place of supply", "Supply Attract Reverse Charge", "Rate (%)", "Taxable Value (₹)", "Tax Amount", None, None, None,
          "GSTR-1/IFF/GSTR-1A/5 Filing Status", "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period",
          "GSTR-3B Filing Status", "Effective date of cancellation", "Amendment made, if any",
          "Original tax period in which reported "]
B2BA_S = [None, None, None, None, "Invoice type", "Invoice number", "Invoice Date", "Invoice Value (₹)", None, None, None, None,
          "Integrated Tax  (₹)", "Central Tax (₹)", "State/UT tax (₹)", "Cess  (₹)"]
CDNRA_A = ["Original details", None, None, "Revised details"]
CDNRA_G = ["Note type", "Note Number", "Note date", "GSTIN of Supplier", "Trade/Legal name of the supplier",
           "Credit note/Debit note details", None, None, None, None, "Place of supply", "Supply Attract Reverse Charge", "Rate (%)",
           "Taxable Value (₹)", "Tax Amount", None, None, None, "GSTR-1/IFF/GSTR-1A/5 Filing Status",
           "GSTR-1/IFF/GSTR-1A/5 Filing Date", "GSTR-1/IFF/GSTR-1A/5 Filing Period", "GSTR-3B Filing Status",
           "Amendment made, if any", "Tax Period in which reported earlier", "Effective date of cancellation"]
CDNRA_S = [None, None, None, None, None, "Note type", "Note Number ", "Note Supply type ", "Note date", "Note Value (₹)", None, None,
           None, None, "Integrated Tax (₹)", "Central Tax (₹)", "State Tax (₹)", "Cess Amount (₹)"]
ISD_G = ["Eligibility of ITC", "GSTIN of ISD", "Trade/Legal name of the ISD", "ISD Document type", "ISD Invoice number",
         "ISD Invoice date", "ISD credit note number", "ISD credit note date", "Original Invoice Number",
         "Original invoice date", "Input tax distribution by ISD", None, None, None, "ISD GSTR-6 Filing status",
         "Amendment made, if any", "Tax Period in which Amended"]
ISD_S = [None] * 10 + ["Integrated Tax (₹)", "Central Tax (₹)", "State/UT Tax (₹)", "Cess (₹)"]


def banner(tp, fy="2025-26"):
    return f"Financial Year: {fy}  |  Tax Period: {tp}  |  Date Updated till: 03-Aug-2026"


def add_sheet(wb, name, title, above, head, sub, blocks, repeat_sub=False):
    ws = wb.create_sheet(name)
    ws.append(["Goods and Services Tax - GSTR 2A", title])
    ws.append([])
    ws.append([])
    ws.append(["   "])
    if above:
        # keep the heading row on the same row index the real files have: [group row above the heading row]
        ws.append(above)
    ws.append(head)
    if sub and not repeat_sub:
        ws.append(sub)
    for b, rows in blocks:
        ws.append([b])
        if sub and repeat_sub:
            ws.append(sub)
        for r in rows:
            ws.append(list(r))
    return ws


def b2b_row(gstin, invno, rate_line=True, total=True, taxable=1000.0, igst=180.0, invtype="R", date="24-04-2025"):
    tail = ["Y", "10-May-25", "Apr-25", "Y", None, None, None, "E-Invoice", "irn" + invno, date]
    rows = []
    if rate_line:
        rows.append([gstin, "Some Supplier", invno, invtype, date, taxable + igst, "Uttarakhand", "N", 18, taxable, igst, 0, 0, 0] + tail)
    if total:
        rows.append([gstin, "Some Supplier", invno + "-Total", invtype, date, taxable + igst, "Uttarakhand", "N", 18, taxable, igst, 0, 0, 0] + tail)
    return rows


def fixture(path, moved_col=False, bad_heading=False, alias_cdnr=False, drop_total=False, bad_gstin=False,
            narrow_b2b=False, include_isd=True):
    wb = openpyxl.Workbook()
    wb.active.title = "Read me"
    wb.active.append([None, "Taxpayer's GSTIN", "05AAAAA0001A1Z1", "Tax period", "042025"])
    wb.active.append([None, "Legal name", "SAMPLE TAXPAYER", "Financial year", "2025-26"])
    b2b_g, b2b_s = list(B2B_G), list(B2B_S)
    rows_apr = b2b_row(G1, "INV1") + b2b_row(G2, "INV2", taxable=500.0, igst=90.0, invtype="SEZWP")
    rows_may = b2b_row(G1, "INV3", taxable=2000.0, igst=360.0)
    if drop_total:
        rows_apr = [r for r in rows_apr if not str(r[2]).endswith("INV1-Total")]      # INV1 loses its rollup row
    if bad_gstin:
        rows_may.append(["NOTAGSTIN", "Bad Row", "INV9-Total", "R", "01-05-2025", 1, "X", "N", 18, 1, 0, 0, 0, 0])
    if moved_col:                                                                       # a new column before 'Taxable Value'
        b2b_g.insert(9, "Some New Column")
        b2b_s.insert(9, None)
        rows_apr = [r[:9] + ["NEW"] + r[9:] for r in rows_apr]
        rows_may = [r[:9] + ["NEW"] + r[9:] for r in rows_may]
    if bad_heading:
        b2b_g[9] = "Something Else Entirely"
    if narrow_b2b:
        b2b_g, b2b_s = b2b_g[:10], b2b_s[:10]
        rows_apr = [r[:10] for r in rows_apr]
        rows_may = [r[:10] for r in rows_may]
    add_sheet(wb, "B2B", "B2B invoices", None, b2b_g, b2b_s, [(banner("042025"), rows_apr), (banner("052025"), rows_may)])
    cn = [G1, "Some Supplier", "Credit note", "CN1", "R", "28-04-2025", 500.0, "Uttarakhand", "N", 18, 400.0, 0, 36, 36, 0,
          "Y", "24-May-25", "Apr-25", "Y", None, None, None, "E-Invoice", "irnCN1", "28-04-2025"]
    cn_total = list(cn)
    cn_total[3] = "CN1-Total"
    add_sheet(wb, "B2B-CDNR" if alias_cdnr else "CDNR", "CDN", None, CDNR_G, CDNR_S,
              [(banner("042025"), [cn, cn_total]), (banner("052025"), [])])
    ba = ["INV0", "01-03-2025", G3, "Amended Supplier", "R", "INV0A", "05-04-2025", 1180.0, "Uttarakhand", "N", 18, 1000.0, 180.0, 0, 0, 0,
          "Y", "10-May-25", "Apr-25", "Y", None, "A", "Mar-25"]
    ba_total = list(ba)
    ba_total[5] = "INV0A-Total"
    add_sheet(wb, "B2BA", "B2BA", B2BA_A, B2BA_G, B2BA_S, [(banner("042025"), [ba, ba_total]), (banner("052025"), [])], repeat_sub=True)
    ca = ["Credit note", "CN0", "01-03-2025", G2, "Other Supplier", "Credit note", "CN0A", "R", "05-04-2025", 300.0, "Uttarakhand", "N",
          18, 250.0, 0, 22.5, 22.5, 0, "Y", "10-May-25", "Apr-25", "Y", "A", "Mar-25", None]
    ca_total = list(ca)
    ca_total[6] = "CN0A-Total"
    add_sheet(wb, "CDNRA", "CDNRA", CDNRA_A, CDNRA_G, CDNRA_S, [(banner("042025"), [ca, ca_total])], repeat_sub=True)
    if include_isd:
        isd = ["Yes", G3, "Some ISD", "Invoice", "ISD/1", "10-04-2025", None, None, "ORIG/1", "01-04-2025", 100.0, 50.0, 50.0, 0.0, "Y", None, None]
        add_sheet(wb, "ISD", "ISD", None, ISD_G, ISD_S, [(banner("Apr-Jun"), [isd])])
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


def compare(raw):
    ad.clear_cache()
    old, new = run(dept.parse_r2a_excel, raw), run(ad.parse_r2a_excel, raw)
    if old == new:
        return []
    diffs = []
    if old[0] == "ok" and new[0] == "ok":
        for k in old[1]:
            if old[1][k] != new[1][k]:
                diffs.append((k, str(old[1][k])[:160], str(new[1][k])[:160]))
    else:
        diffs.append(("result", str(old)[:160], str(new)[:160]))
    return diffs


def _exc_text(fn, *a):
    try:
        fn(*a)
        return ""
    except Exception as e:  # noqa: BLE001
        return str(e)


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        raw = fixture(os.path.join(tmp, "r2a.xlsx"))
        d = compare(raw)
        ok &= check("standard layout: identical to the original for every table", not d, str(d[:2]))
        r = ad.parse_r2a_excel(raw)
        ok &= check("available, five tables read, months from the banners (quarterly ISD banner fans out to 3 months)",
                    r["available"] and r["months_present"] == {"Apr-25", "May-25", "Jun-25"}, str(r["months_present"]))
        ok &= check("B2B: only the '-Total' row of each document kept, suffix stripped (INV1, INV2 in April)",
                    sorted(x["invno"] for x in r["b2b"]["Apr-25"]) == ["INV1", "INV2"] and r["b2b"]["Apr-25"][0]["taxable"] == 1000.0)
        ok &= check("invoice type normalised (SEZWP), date normalised to ISO",
                    r["b2b"]["Apr-25"][1]["invtype"] == "SEZWP" and r["b2b"]["Apr-25"][0]["invdate"] == "2025-04-24")
        ok &= check("B2BA: repeated sub-heading row after the banner does not become data; revised + original numbers read",
                    len(r["b2ba"]["Apr-25"]) == 1 and r["b2ba"]["Apr-25"][0]["invno"] == "INV0A" and r["b2ba"]["Apr-25"][0]["orig_invno"] == "INV0")
        ok &= check("CDNRA and ISD (no real sample in the original) read by the same rules",
                    r["cdnra"]["Apr-25"][0]["note_no"] == "CN0A" and len(r["isd"]["Jun-25"]) == 1 and r["isd"]["Jun-25"][0]["igst"] == 100.0)
        ok &= check("empty month block registered as an empty list, not missing", r["b2ba"]["May-25"] == [] and r["cdnr"]["May-25"] == [])
        m = ad.get_data(raw)["mapping"]
        ok &= check("MAPPING_REPORT: every field found at its usual column, confirmed by heading",
                    all(x["status"] == "OK" and x["matched_by"] == "position" and x["confidence"] == "HIGH" for x in m), str([x for x in m if x["matched_by"] != "position"][:1]))
        ok &= check("no warnings on the standard layout", not [i for i in ad.get_data(raw)["issues"] if i["severity"] != "INFO"])

        # ---- documented differences / error handling
        ad.clear_cache()
        drop = fixture(os.path.join(tmp, "drop_total.xlsx"), drop_total=True)
        ok &= check("a document with no '-Total' row: identical to the original, and listed in total_row_missing", not compare(drop) and
                    ad.parse_r2a_excel(drop)["total_row_missing"] == {"B2B": ["Apr-25: INV1"]}, str(ad.parse_r2a_excel(drop)["total_row_missing"]))
        bad_g = fixture(os.path.join(tmp, "bad_gstin.xlsx"), bad_gstin=True)
        ok &= check("a row whose GSTIN is not GSTIN-shaped is dropped, identical to the original", not compare(bad_g))

        moved = fixture(os.path.join(tmp, "moved.xlsx"), moved_col=True)
        ad.clear_cache()
        rm = ad.parse_r2a_excel(moved)
        ok &= check("a new column inserted before 'Taxable Value': followed by heading - values still right",
                    rm["b2b"]["Apr-25"][0]["taxable"] == 1000.0 and rm["b2b"]["Apr-25"][0]["igst"] == 180.0, str(rm["b2b"]["Apr-25"][0]))
        old_m = run(dept.parse_r2a_excel, moved)
        ok &= check("(documenting the fixed bug) the original, reading by position, gets the wrong figures", old_m[0] == "ok" and old_m[1]["b2b"]["Apr-25"][0]["taxable"] != 1000.0,
                    str(old_m)[:100])
        ok &= check("...and W703 says which columns moved, in plain words",
                    any(i["id"] == "W703" and "found by heading" in i["message"] for i in ad.get_data(moved)["issues"]))

        badh = fixture(os.path.join(tmp, "bad_heading.xlsx"), bad_heading=True)
        ad.clear_cache()
        ok &= check("a heading that no longer matches: read by position exactly as before (identical) + W704 warning", not compare(badh) and
                    any(i["id"] == "W704" for i in ad.get_data(badh)["issues"]))

        alias = fixture(os.path.join(tmp, "alias.xlsx"), alias_cdnr=True)
        ad.clear_cache()
        ok &= check("'B2B-CDNR' sheet name (accepted by the classifier) is now read; the original silently read no credit notes",
                    len(ad.parse_r2a_excel(alias)["cdnr"]["Apr-25"]) == 1 and dept.parse_r2a_excel(alias)["cdnr"] == {})

        narrow = fixture(os.path.join(tmp, "narrow.xlsx"), narrow_b2b=True)
        ad.clear_cache()
        rn = ad.parse_r2a_excel(narrow)
        ok &= check("E701: required columns missing (sheet cut short) -> that sheet reads as empty, plain message, other sheets unaffected",
                    rn["available"] and rn["b2b"] == {} and len(rn["cdnr"]["Apr-25"]) == 1 and
                    any(i["id"] == "E701" and "required column" in i["message"] for i in ad.get_data(narrow)["issues"]))

        # ---- canonical file
        ad.clear_cache()
        path = ad.write_canonical(ad.build_canonical(raw), tmp)
        ok &= check("canonical file recognised, named from the Read me GSTIN and FY",
                    ad.is_canonical_file(path) and os.path.basename(path) == "Canonical_GSTR2A_05AAAAA0001A1Z1_2025-26.xlsx", os.path.basename(path))
        ok &= check("canonical file supplied as input reads identically", ad.parse_r2a_excel(path) == dept.parse_r2a_excel(raw))
        wb = openpyxl.load_workbook(path, read_only=True)
        ok &= check("sheets: META, RETURNS, COVERAGE, five R2A_ tables, MAPPING_REPORT, ISSUES",
                    wb.sheetnames == ["META", "RETURNS", "COVERAGE", "R2A_B2B", "R2A_CDNR", "R2A_B2BA", "R2A_CDNRA", "R2A_ISD", "MAPPING_REPORT", "ISSUES"], str(wb.sheetnames))
        wb.close()

        # ---- failure paths
        with open("junk.xlsx", "w") as f:
            f.write("not an excel file")
        ad.clear_cache()
        rj = ad.parse_r2a_excel("junk.xlsx")
        ok &= check("unreadable file: available=False with a plain reason (E702), like the original", not rj["available"] and "Could not read GSTR-2A workbook" in rj["reason"] and "[E702]" in rj["reason"], rj["reason"])
        wb = openpyxl.Workbook()
        wb.active.title = "Summary"
        wb.save("wrong.xlsx")
        ad.clear_cache()
        rw = ad.parse_r2a_excel("wrong.xlsx")
        ok &= check("no 2A sheet at all: the original's exact reason", not rw["available"] and rw["reason"] == dept.parse_r2a_excel("wrong.xlsx")["reason"], str(rw["reason"]))
        ok &= check("path not supplied: the original's exact reason", ad.parse_r2a_excel(None)["reason"] == "GSTR-2A not supplied for this taxpayer/FY.")
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
