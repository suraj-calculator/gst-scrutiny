#!/usr/bin/env python3
"""
Regression tests for the BO Profile canonical layer (boprofile_adapter.py, docs/BOPROFILE_CANONICAL_SPEC.md).

Self-contained: builds a synthetic BO Profile workbook in the real layout (title row, blank row, heading row; the FY-keyed tables, the
two-row BIFA heading, the INTRA / INTER / TOTAL ITC tables, the flat lists incl. DRC payments) and checks that the ORIGINAL
parse_bo_profile() and the canonical route return identical results, plus how layout changes / broken files are reported.

Run:  python3 test_boprofile_adapter.py
"""
import contextlib
import io
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ["GST_BO_CANONICAL"] = "0"          # original reader for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_dept as dept               # noqa: E402
import boprofile_adapter as ad                # noqa: E402

FYS = ["2024-25", "2023-24"]


def sheet(wb, name, title, header, rows, sub=None, group=None):
    ws = wb.create_sheet(name)
    ws.append([title])
    ws.append([])
    if group:
        ws.append(group)
    ws.append(header)
    if sub:
        ws.append(sub)
    for r in rows:
        ws.append(r)
    return ws


def make(path, truncate_names=False, move_itc=False, drop_col=False, no_header_sheet=None, with_lists=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Demographic Details"
    ws.append(["Demographic Details"])
    ws.append([])
    for k, v in (("GSTIN", "05AAAAA0001A1Z1"), ("Legal Name", "SAMPLE TAXPAYER"), ("Trade Name", "M/s SAMPLE"), ("Status", "Active"), ("Empty one", None)):
        ws.append([k, v])
    fin_head = ["Financial Year", "Turnover", "Taxable Turnover", "Export Turnover", "Total tax liability", "Tax paid by ITC", "Tax paid in cash",
                "ITC availed", "Cash Utilization %", "ITC Utilization %", "Effective Tax Rate %", "ITC reversal"]
    fin_rows = [[fy, "1200.5", "1100.25", "0", "300", "250.5", "49.5", "260", "0.0002", "97.5%", "0.27", "1.5"] for fy in FYS]
    if drop_col:
        i = fin_head.index("Export Turnover")
        fin_head.pop(i)
        fin_rows = [r[:i] + r[i + 1:] for r in fin_rows]
    sheet(wb, "Financial Information", "Financial Information", fin_head, fin_rows)
    sheet(wb, "BIFA Specific Information", "BIFA Specific Information",
          ["Financial Year", "Liability as per GSTR1", "Liability as per GSTR3B", "Difference Liability", "ITC Availed in R3B",
           "ITC Accrued in R2B/R2A", "Difference ITC", "Supply value as per GSTR3B", "Liability as per EWB"],
          [[fy, "10", "9.5", "0.5", "8", "7.5", "0.5", "100", "12"] for fy in FYS],
          group=[None, "Supply Tax Deficit", None, None, "Excess ITC claimed", None, None, "Mismatch", "EWB mismatch"])
    grp = [None, "INTRA STATE", None, "INTER STATE", None, "TOTAL", None]
    itc_sub = ["Financial Year", "No.of beneficiaries", "ITC Passed", "No.of beneficiaries", "ITC Passed", "No.of beneficiaries", "ITC Passed"]
    itc_rows = [[fy, "5", "1.5", "3", "2.5", "8", "4.0"] for fy in FYS]
    if move_itc:
        itc_rows = [[r[0], "0"] + r[1:] for r in itc_rows]
        itc_sub = ["Financial Year", "Extra"] + itc_sub[1:]
        grp = [None, None] + grp[1:]
    wsi = wb.create_sheet("ITC Passed On")
    wsi.append(["ITC Passed On"])
    wsi.append([])
    wsi.append(grp)
    wsi.append(itc_sub)
    for r in itc_rows:
        wsi.append(r)
    wsr = wb.create_sheet("ITC Received")
    wsr.append(["ITC Received"])
    wsr.append([])
    wsr.append(["", "INTRA STATE", "", "INTER STATE", "", "TOTAL", ""])
    wsr.append(["Financial Year", "No.of suppliers", "ITC Received", "No.of suppliers", "ITC Received", "No.of suppliers", "ITC Received"])
    for fy in FYS:
        wsr.append([fy, "4", "1.0", "2", "1.1", "6", "2.1"])
    sheet(wb, "EWB Related Information", "EWB", ["Financial Year", "Inward Supply", "Inward Supply Tax", "Number of Inward EWB", "Outward Supply",
                                                 "Outward Supply Tax", "Number of Outward EWB", "Ratio"],
          [[fy, "50", "9", "40", "60", "10", "45", "0.8"] for fy in FYS])
    sheet(wb, "E-Invoice Related Information", "EINV", ["Financial Year", "Count of Active", "Count of Cancelled", "Active Assessable Value", "Active Taxes",
                                                        "Cancelled Assessable Value", "Cancelled Taxes"], [[fy, "10", "1", "100", "18", "5", "0.9"] for fy in FYS])
    sheet(wb, "Refund Details", "Refund", ["Financial Year", "Total Claimed Amount", "Total Rejected Amount", "Total Sanctioned Amount"],
          [[fy, "1", "0.2", "0.8"] for fy in FYS])
    nm = (lambda s: s[:22]) if truncate_names else (lambda s: s)
    if with_lists:
        sheet(wb, nm("Top 10 Beneficiaries based on I"), "Top", ["Receiver GSTIN", "Trade Name", "Registration Start Date", "Status", "Overall Risk Score", "ITC Passed On"],
              [["07BBBBB0002B1Z2", "Ben One", "01/01/2020", "Active", 12, 3.5], ["09CCCCC0003C1Z3", "Ben Two", "01/02/2020", "Cancelled", 55, 1.5], [None] * 6])
        sheet(wb, "Top 10 Suppliers based on ITC R", "Top", ["Supplier GSTIN", "Trade Name", "Registration Start Date", "Status", "Overall Risk Score", "ITC Received"],
              [["07BBBBB0002B1Z2", "Sup One", "01/01/2019", "Active", 3, 9.5]])
        sheet(wb, "ITC Received from Related - Can", "Related", ["Financial Year", "Supplier GSTIN", "Trade Name", "Related Parameter", "Status", "Cancellation Date",
                                                                "Cancellation Reason", "Total ITC Received"], [["2023-24", "07BBBBB0002B1Z2", "Rel", "Address", "Cancelled", "01/03/2024", "Not filing", 0.5]])
        sheet(wb, "ITC Passed On to Related - Canc", "Related", ["Financial Year", "Receiver GSTIN", "Trade Name", "Related Parameter", "Status", "Cancellation Date",
                                                                "Cancellation Reason", "Total ITC Passed"], [["2023-24", "09CCCCC0003C1Z3", "Rel2", "Address", "Cancelled", "01/03/2024", "Not filing", 0.25]])
        sheet(wb, "DRC Payment Information", "DRC", ["Source ID", "Transaction Type Description", "Transaction Date", "Payment Method", "CGST Tax", "SGST Tax", "IGST Tax", "CESS", "Other than Tax"],
              [["DC01", "Voluntary payment", "01/06/2024", "Cash", "1,000", "1,000", "0", "0", "10"], ["DC02", "Demand order", "02/06/2024", "Credit", 5, "n/a", None, "-", 1]])
        sheet(wb, "Appeal Information", "Appeals", ["Financial Year", "ARN", "Case Filing Date", "Case Status", "Case Type", "Tax Officer Name", "Order Number"], [["2023-24", "AA1", "01/01/2024", "Open", "Appeal", "Officer", "O1"]])
        sheet(wb, "Case Information", "Cases", ["ARN / Case ID", "Reference ID", "Action Date", "Tax Period", "Case Type", "Case Status", "SCN Date", "SCN Number", "CGST Tax", "SGST Tax", "IGST Tax", "CESS"],
              [["C1", "R1", "01/01/2024", "Apr-23", "Scrutiny", "Open", "01/02/2024", "S1", 1, 1, 0, 0]])
        sheet(wb, "Transfer Information", "Transfers", ["Recommend Date", "Source Case ID", "Target Case ID", "Source Module Name", "Target Module Name", "Status"], [["01/01/2024", "C1", "C2", "A", "B", "Done"]])
    wb.create_sheet("Places of Business").append(["not used by any check"])
    if no_header_sheet:
        w = wb[no_header_sheet]
        w.delete_rows(1, w.max_row)
        w.append(["only one cell"])
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
    return run(dept.parse_bo_profile, p), run(ad.parse_bo_profile, p)


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        p = make(os.path.join(tmp, "bo.xlsx"))
        old, new = both(p)
        ok &= check("standard workbook: identical to the original (all 16 tables)", old == new and old[0] == "ok", str([k for k in old[1] if old[1][k] != new[1][k]]) if old[0] == "ok" and new[0] == "ok" else str(new)[:80])
        ok &= check("no warnings on the standard workbook", not [i for i in ad.get_data(p)["issues"] if i["severity"] != "INFO"], str([i for i in ad.get_data(p)["issues"] if i["severity"] != "INFO"])[:150])
        r = ad.parse_bo_profile(p)
        ok &= check("demographic rows (blank value skipped), GSTIN / names", r["self_gstin"] == "05AAAAA0001A1Z1" and "Empty one" not in r["demographic"] and r["trade_name"] == "M/s SAMPLE")
        ok &= check("FY tables read as numbers, '97.5%' -> 97.5, keyed by FY", r["financial_by_fy"]["2024-25"]["itc_utilization_pct"] == 97.5 and r["financial_by_fy"]["2024-25"]["turnover"] == 1200.5)
        ok &= check("BIFA two-row heading, ITC passed / received TOTAL group, EWB plain-supply columns",
                    r["bifa_by_fy"]["2023-24"]["diff_itc"] == 0.5 and r["itc_passed_by_fy"]["2024-25"] == {"total_beneficiaries": 8.0, "total_itc": 4.0}
                    and r["itc_received_by_fy"]["2024-25"]["total_suppliers"] == 6.0 and r["ewb_by_fy"]["2024-25"]["inward_supply"] == 50.0 and r["ewb_by_fy"]["2024-25"]["outward_supply"] == 60.0)
        ok &= check("lists: blank row skipped, values raw", len(r["top_beneficiaries"]) == 2 and r["top_beneficiaries"][0]["risk"] == 12)
        ok &= check("DRC: amounts cleaned ('1,000' -> 1000, 'n/a' / '-' / blank -> 0) and total summed", r["drc_payments"][0]["total"] == 2010.0 and r["drc_payments"][1]["total"] == 6.0 and r["drc_payments"][1]["type"] == "Demand order", str(r["drc_payments"]))
        ok &= check("unused sheets are listed as unused (INFO), not silently ignored", any(i["id"] == "I1203" and "Places of Business" in i["message"] for i in ad.get_data(p)["issues"]))

        # ---- layout changes
        ad.clear_cache()
        mv = make(os.path.join(tmp, "moved.xlsx"), move_itc=True)
        old, new = both(mv)
        ok &= check("ITC table with a new column before the TOTAL pair: followed by heading (right figures) + W1202",
                    new[1]["itc_passed_by_fy"]["2024-25"] == {"total_beneficiaries": 8.0, "total_itc": 4.0} and any(i["id"] == "W1202" for i in ad.get_data(mv)["issues"]))
        ok &= check("(documenting the fixed bug) the original, by fixed position, reads the wrong pair", old[1]["itc_passed_by_fy"]["2024-25"] != {"total_beneficiaries": 8.0, "total_itc": 4.0})
        dc = make(os.path.join(tmp, "dropcol.xlsx"), drop_col=True)
        old, new = both(dc)
        ok &= check("a column missing from the source (Export Turnover): identical to the original (None for it) + reported (I1201)", old == new and new[1]["financial_by_fy"]["2024-25"]["export_turnover"] is None
                    and any(i["id"] == "I1201" and "export_turnover" in i["message"] for i in ad.get_data(dc)["issues"]))
        nh = make(os.path.join(tmp, "nohead.xlsx"), no_header_sheet="Refund Details")
        old, new = both(nh)
        ok &= check("a sheet with no heading row: identical to the original (empty) + W1201", old == new and new[1]["refund_by_fy"] == {} and any(i["id"] == "W1201" for i in ad.get_data(nh)["issues"]))
        tr = make(os.path.join(tmp, "trunc.xlsx"), truncate_names=True)
        old, new = both(tr)
        ok &= check("a sheet name truncated differently ('Top 10 Beneficiaries based'): now read by its leading words; the original found nothing",
                    len(new[1]["top_beneficiaries"]) == 2 and old[1]["top_beneficiaries"] == [])
        nl = make(os.path.join(tmp, "nolists.xlsx"), with_lists=False)
        old, new = both(nl)
        ok &= check("workbook without the list sheets: identical, absent sheets reported as INFO", old == new and any(i["id"] == "I1202" for i in ad.get_data(nl)["issues"]))

        # ---- broken files
        with open("junk.xlsx", "w") as f:
            f.write("not an excel file")
        ad.clear_cache()
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rj = ad.parse_bo_profile("junk.xlsx")
        ok &= check("unreadable file: the original's empty shape (no crash) with the plain reason printed (E1201)", rj == dept.parse_bo_profile("junk.xlsx") and "[E1201]" in buf.getvalue(), buf.getvalue()[:100])
        ok &= check("path not supplied: the original's empty shape", ad.parse_bo_profile(None) == dept.parse_bo_profile(None))

        # ---- canonical file
        ad.clear_cache()
        path = ad.write_canonical(ad.build_canonical(p), tmp)
        ok &= check("canonical file named from the Demographic GSTIN and the FYs found", os.path.basename(path) == "Canonical_BOPROFILE_05AAAAA0001A1Z1_2023-24_2024-25.xlsx", os.path.basename(path))
        ok &= check("canonical file supplied as input reads identically", ad.parse_bo_profile(path) == dept.parse_bo_profile(p))
        wb = openpyxl.load_workbook(path, read_only=True)
        ok &= check("one canonical sheet per table + META, MAPPING_REPORT, ISSUES", len(wb.sheetnames) == 19 and wb.sheetnames[0] == "META" and "BO_DRC" in wb.sheetnames, str(wb.sheetnames))
        wb.close()
        folder = os.path.join(tmp, "folder")
        os.makedirs(folder)
        shutil.copy(path, folder)
        os.environ["GST_BO_CANONICAL"] = "1"
        with contextlib.redirect_stdout(io.StringIO()):
            try:
                res = mpu.classify_folder(folder)
            except Exception as e:  # noqa: BLE001
                res = getattr(e, "res", None)
        ok &= check("classify_folder recognises a canonical BO Profile workbook", res is not None and res.get("bo_profile") == os.path.join(folder, os.path.basename(path)), str(res.get("bo_profile") if res else None))
        os.environ["GST_BO_CANONICAL"] = "0"
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
