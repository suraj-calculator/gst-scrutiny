#!/usr/bin/env python3
"""
Regression tests for the GSTR-9 / GSTR-9C canonical layer (gstr9_adapter.py, docs/GSTR9_CANONICAL_SPEC.md).

Self-contained: builds synthetic GSTR-9 and GSTR-9C workbooks in the real layout (one sheet per Item / Part, names truncated to 31
characters, label in column B, figures in fixed columns) and checks that the ORIGINAL parse_gstr9() / parse_gstr9c() and the canonical
route return identical results - including every 'notes' line - and how missing sheets / rows and broken files are reported.

Run:  python3 test_gstr9_adapter.py
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
os.environ["GST_R9_CANONICAL"] = "0"          # original readers for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_dept as dept               # noqa: E402
import gstr9_adapter as ad                    # noqa: E402

HEAD = [["Sr.No", "Nature of Supplies", "Taxable Value(₹)", "(Amount in ₹ in all tables)"], [None, None, None, "Central Tax(₹)"], [None, "1", "2", "3"]]


def add(wb, name, rows, head=True):
    ws = wb.create_sheet(name)
    ws.append([name])
    ws.append([])
    if head:
        for r in HEAD:
            ws.append(r)
    for r in rows:
        ws.append(r)
    return ws


def make9(path, arn="AA050324442575B", drop=None, no_row=None, item19=True, str_num=False):
    wb = openpyxl.Workbook()
    wb.active.title = "Part I - Basic Details (Items 1"
    p1 = wb.active
    p1.append(["Part I - Basic Details"])
    p1.append([])
    for r in (["1. Financial Year", "2023-24"], ["2. GSTIN", "05AAAAA0001A1Z1"], ["3(a). Legal name of the registered person", "SAMPLE TAXPAYER"],
              ["3(b). Trade name", "M/s SAMPLE"], ["3(c). ARN", arn], ["3(d). Date of Filing", "31-12-2024"]):
        p1.append(r)
    n = (lambda v: f"{v:,.2f}") if str_num else (lambda v: v)
    add(wb, "Item 4 - Advances & Outward+Inw", [
        ["A", "Supplies made to un-registered persons (B2C)", n(1000.0), 90.0, 90.0, 0.0, 0.0],
        ["B", "Supplies made to registered persons (B2B)", n(5000.0), 450.0, 450.0, 100.0, 5.0],
        ["I", "Credit Notes issued in respect of transactions specified in (B) above (-)", 200.0, 18.0, 18.0, 0.0, 0.0],
        ["G", "Inward supplies on which tax is to be paid on reverse charge basis", 294000.0, 26460.0, 26460.0, 0.0, 0.0],
        ["N", "Total Turnover on which tax is to be paid (H + M)", 0, 0, 0, 555.5, 0]])
    add(wb, "Item 5 - Outward Supplies (Tax", [
        ["A", "Zero rated supply (Export) without payment of tax", 0], ["B", "Supply to SEZs without payment of tax", 0],
        ["D", "Exempted", 12.5], ["E", "Nil Rated", 0], ["F", "Non-GST supply (includes 'no supply')", "-"]])
    add(wb, "Item 6 - ITC Availed During the", [
        ["B", "Total amount of input tax credit availed through Form GSTR-3B (sum total of Table 4A)", None, 100.0, 100.0, 50.0, 1.0],
        ["H", "Total ITC availed (B + F + G above)", None, 100.0, 100.0, 50.0, 1.0]])
    add(wb, "Item 7 - ITC Reversed & Ineligi", [["K", "Total ITC Reversed (A to K above)", 1.0, 2.0, 3.0, 4.0]])
    add(wb, "Item 8 - Other ITC Related Info", [
        ["A", "ITC as per GSTR-2B", 10.0, 20.0, 30.0, 0.0], ["B", "ITC as per sum total of 6(B) and 6(H) above", 9.0, 19.0, 29.0, 0.0],
        ["D", "Difference [A-(B+C)]", 1.0, 1.0, 1.0, 0.0]])
    add(wb, "Item 9 - Tax Paid", [
        ["", "Integrated Tax", 111.0, 110.0], ["", "Central Tax", 222.0, 220.0], ["", "State/UT Tax", 333.0, 330.0],
        ["", "Interest", 5.0, 4.0], ["", "Late fee", 7.0, 6.0]])
    add(wb, "Part V - Transactions Declared", [
        ["13", "ITC availed for the previous financial year", None, 11.0, 12.0, 13.0, 14.0],
        ["12", "Reversal of ITC availed during previous financial year", None, 1.0, 2.0, 3.0, 4.0]])
    if item19:
        add(wb, "Item 19 - Late Fee", [["A", "Central Tax", 100.0, 90.0], ["B", "State Tax", 50.0, 40.0], ["C", "total", 999.0, 999.0]])
    if drop:
        del wb[[s for s in wb.sheetnames if s.startswith(drop)][0]]
    if no_row:
        ws = wb[[s for s in wb.sheetnames if s.startswith(no_row[0])][0]]
        for row in ws.iter_rows():
            if row[1].value and no_row[1] in str(row[1].value).lower():
                row[1].value = "renamed"
    wb.save(path)
    return path


def make9c(path, zero_exempt=False, reasons=True, drop=None):
    wb = openpyxl.Workbook()
    wb.active.title = "Part I - Basic Details (Items 1"
    p1 = wb.active
    p1.append(["Part I"])
    for k, v in (("Financial Year", "2023-24"), ("GSTIN", "05AAAAA0001A1Z1"), ("Legal Name", "SAMPLE TAXPAYER"), ("ARN", "AA0503251"), ("ARN Date", "01-01-2025")):
        p1.append([None, k, v])
    add(wb, "Item 5 - Reconciliation of Gros", [
        ["A", "Turnover (including exports) as per audited Annual Financial Statement", None, 1000.0],
        ["I", "Annual turnover after adjustments as above", None, 950.0], ["J", "Turnover as declared in Annual Return (GSTR9)", None, 940.0],
        ["K", "Un-Reconciled turnover (I - J)", None, 10.0]])
    add(wb, "Item 6 - Reasons for Turnover", [["A", None, "Reason one: 10.00 timing"], ["B", None, "Reason two"], [1, None, "numeric first cell is skipped"]] if reasons else [])
    add(wb, "Item 7 - Reconciliation of Taxa", [
        ["E", "Taxable turnover as per adjustments above", 900.0], ["F", "Value of exempted, nil rated, non-GST supplies", 0.0 if zero_exempt else 12.5],
        ["G", "Taxable turnover as per liability declared in Annual Return", 895.0]])
    add(wb, "Item 8 - Reasons for Taxable", [["A", None, "Taxable reason"]])
    add(wb, "Item 9 - Rate-wise Tax Liabilit", [
        ["T", "Total amount to be paid as per tables above", None, 100.0, 100.0, 50.0, 1.0],
        ["U", "Total amount paid as declared in Annual Return (GSTR 9)", None, 99.0, 99.0, 49.0, 1.0]])
    add(wb, "Item 12 - Reconciliation of Net", [
        ["A", "ITC availed as per audited Annual Financial Statement", 500.0], ["B", "ITC booked in earlier Financial Years claimed in current Financial Year", 5.0],
        ["C", "ITC booked in current Financial Year to be claimed in subsequent Financial Years", 6.0], ["D", "ITC claimed in Annual Return (GSTR9)", 490.0],
        ["E", "Un-reconciled ITC", 10.0]])
    add(wb, "Item 13 - Reasons for Un-reconc", [["A", None, "ITC reason 5.00"]])
    if drop:
        del wb[[s for s in wb.sheetnames if s.startswith(drop)][0]]
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


def both(path, form):
    ad.clear_cache()
    if form == "gstr9":
        return run(dept.parse_gstr9, path), run(ad.parse_gstr9, path)
    return run(dept.parse_gstr9c, path), run(ad.parse_gstr9c, path)


def diff(old, new):
    return [k for k in old[1] if old[1][k] != new[1][k]] if old[0] == "ok" and new[0] == "ok" else str((old, new))[:120]


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        p9 = make9(os.path.join(tmp, "r9.xlsx"))
        old, new = both(p9, "gstr9")
        ok &= check("[GSTR-9] standard workbook: identical to the original (every key, incl. notes)", old == new and old[0] == "ok", str(diff(old, new)))
        ok &= check("[GSTR-9] no warnings on the standard workbook", not [i for i in ad.get_data(p9, "gstr9")["issues"] if i["severity"] != "INFO"])
        r = ad.parse_gstr9(p9)
        ok &= check("Part I: FY, GSTIN, ARN, filing date; filed (not system draft)", r["fy"] == "2023-24" and r["arn"] == "AA050324442575B" and r["is_system_draft"] is False and r["date_of_filing"] == "31-12-2024")
        ok &= check("Table 4B / 4I / RCM / net tax read from their labels and columns", r["table4_b2b_taxable"] == 5000.0 and r["table4_b2b_igst"] == 100.0 and r["table4_cn_cgst"] == 18.0
                    and r["table4_rcm_inward_taxable"] == 294000.0 and r["table4_net_supplies_tax_igst"] == 555.5)
        ok &= check("Table 5: '-' reads 0, all-zero flag false because Exempted is 12.5", r["table5_nongst"] == 0.0 and r["table5_exempted"] == 12.5 and r["table5_all_zero"] is False)
        ok &= check("Tables 6A / 6 / 7 / 8: sums of the right columns", r["table6a_total"] == 251.0 and r["table6_total_itc_availed"] == 251.0 and r["table7_total_reversed"] == 10.0
                    and r["table8_itc_as_per_2b"] == 60.0 and r["table8_difference"] == 3.0)
        ok &= check("Item 19 (Central + State late fee, letters A / B only) supersedes Item 9's late-fee row", r["table9_late_fee_payable"] == 150.0 and r["table9_late_fee_paid"] == 130.0
                    and r["table9_interest_payable"] == 5.0 and r["table9_liability_cgst"] == 222.0, str((r["table9_late_fee_payable"], r["table9_late_fee_paid"])))
        ok &= check("Table 13 / 12 (previous-FY ITC) read", r["table13_itc_igst"] == 13.0 and r["table12_itc_reversed_cess"] == 4.0)
        sn = make9(os.path.join(tmp, "r9_text.xlsx"), str_num=True)
        old, new = both(sn, "gstr9")
        ok &= check("amounts stored as text ('5,000.00') read as numbers, identical", old == new and new[1]["table4_b2b_taxable"] == 5000.0)
        no_arn = make9(os.path.join(tmp, "r9_draft.xlsx"), arn=None)
        old, new = both(no_arn, "gstr9")
        ok &= check("no ARN: system-draft note, identical", old == new and new[1]["is_system_draft"] is True and any("System Drafted" in n for n in new[1]["notes"]))
        no19 = make9(os.path.join(tmp, "r9_no19.xlsx"), item19=False)
        old, new = both(no19, "gstr9")
        ok &= check("no Item 19 sheet: Item 9's late-fee row stands, identical", old == new and new[1]["table9_late_fee_payable"] == 7.0)
        d5 = make9(os.path.join(tmp, "r9_no_p5.xlsx"), drop="Part V")
        old, new = both(d5, "gstr9")
        ok &= check("Part V sheet missing: identical, same note, reported as W1302 with the sheet's name", old == new and any("Part V" in n for n in new[1]["notes"]) and
                    any(i["id"] == "W1302" for i in ad.get_data(d5, "gstr9")["issues"]))
        nr = make9(os.path.join(tmp, "r9_no_row.xlsx"), no_row=("Item 4", "supplies made to registered persons"))
        old, new = both(nr, "gstr9")
        ok &= check("Table 4B row not found: identical (None + note) and W1303 says which label was looked for", old == new and new[1]["table4_b2b_taxable"] is None and
                    any(i["id"] == "W1303" and "supplies made to registered persons" in i["message"] for i in ad.get_data(nr, "gstr9")["issues"]))

        p9c = make9c(os.path.join(tmp, "r9c.xlsx"))
        old, new = both(p9c, "gstr9c")
        ok &= check("[GSTR-9C] standard workbook: identical to the original (every key, incl. notes)", old == new and old[0] == "ok", str(diff(old, new)))
        ok &= check("[GSTR-9C] no warnings", not [i for i in ad.get_data(p9c, "gstr9c")["issues"] if i["severity"] != "INFO"])
        rc = ad.parse_gstr9c(p9c)
        ok &= check("Part I, Items 5 / 7 / 9 / 12 read", rc["fy"] == "2023-24" and rc["arn_date"] == "01-01-2025" and rc["turnover_unreconciled"] == 10.0 and rc["exempt_nil_nongst_adjustment"] == 12.5
                    and rc["tax_payable_total"] == 251.0 and rc["itc_unreconciled"] == 10.0 and rc["itc_booked_now_claimed_later"] == 6.0)
        ok &= check("the auditor's reason lines (alphabetic first cell only) kept in order", rc["turnover_diff_reasons"] == ["Reason one: 10.00 timing", "Reason two"] and rc["itc_diff_reasons"] == ["ITC reason 5.00"], str(rc["turnover_diff_reasons"]))
        z = make9c(os.path.join(tmp, "r9c_zero.xlsx"), zero_exempt=True)
        old, new = both(z, "gstr9c")
        ok &= check("Table 7B zero: the R13 note is added, identical", old == new and any("Table 7B" in n for n in new[1]["notes"]))
        dd = make9c(os.path.join(tmp, "r9c_no12.xlsx"), drop="Item 12")
        old, new = both(dd, "gstr9c")
        ok &= check("Item 12 sheet missing: identical, reported (W1302)", old == new and any(i["id"] == "W1302" for i in ad.get_data(dd, "gstr9c")["issues"]))

        # ---- broken / absent files
        with open("junk.xlsx", "w") as f:
            f.write("not an excel file")
        ad.clear_cache()
        rj = ad.parse_gstr9("junk.xlsx")
        ok &= check("unreadable GSTR-9: available=False, 'Could not read GSTR-9 workbook' with E1301 in plain words", not rj["available"] and "Could not read GSTR-9 workbook" in rj["reason"] and "[E1301]" in rj["reason"], str(rj["reason"])[:90])
        ok &= check("unreadable GSTR-9C likewise", "Could not read GSTR-9C workbook" in ad.parse_gstr9c("junk.xlsx")["reason"])
        ok &= check("not supplied: the original's exact reasons", ad.parse_gstr9(None)["reason"] == "GSTR-9 not supplied for this taxpayer/FY." and ad.parse_gstr9c(None)["reason"] == "GSTR-9C not supplied for this taxpayer/FY.")

        # ---- canonical files
        ad.clear_cache()
        c9 = ad.write_canonical(ad.build_canonical(p9, "gstr9"), tmp)
        c9c = ad.write_canonical(ad.build_canonical(p9c, "gstr9c"), tmp)
        ok &= check("canonical files named from the return's own GSTIN and FY, form recorded",
                    os.path.basename(c9) == "Canonical_GSTR9_05AAAAA0001A1Z1_2023-24.xlsx" and os.path.basename(c9c) == "Canonical_GSTR9C_05AAAAA0001A1Z1_2023-24.xlsx"
                    and ad.canonical_form(c9) == "gstr9" and ad.canonical_form(c9c) == "gstr9c", os.path.basename(c9))
        ok &= check("canonical files supplied as input read identically", ad.parse_gstr9(c9) == dept.parse_gstr9(p9) and ad.parse_gstr9c(c9c) == dept.parse_gstr9c(p9c))
        wb = openpyxl.load_workbook(c9, read_only=True)
        ok &= check("sheets: META, R9_KV, R9_FACTS, R9_LISTS, MAPPING_REPORT, ISSUES", wb.sheetnames == ["META", "R9_KV", "R9_FACTS", "R9_LISTS", "MAPPING_REPORT", "ISSUES"], str(wb.sheetnames))
        wb.close()
        folder = os.path.join(tmp, "folder")
        os.makedirs(folder)
        shutil.copy(c9, folder)
        shutil.copy(c9c, folder)
        os.environ["GST_R9_CANONICAL"] = "1"
        with contextlib.redirect_stdout(io.StringIO()):
            try:
                res = mpu.classify_folder(folder)
            except Exception as e:  # noqa: BLE001
                res = getattr(e, "res", None)
        ok &= check("classify_folder puts a canonical GSTR-9 and a canonical GSTR-9C in their own lists",
                    res is not None and [os.path.basename(f) for f in res["gstr9_files"]] == [os.path.basename(c9)] and [os.path.basename(f) for f in res["gstr9c_files"]] == [os.path.basename(c9c)],
                    str((res["gstr9_files"], res["gstr9c_files"]) if res else None))
        os.environ["GST_R9_CANONICAL"] = "0"
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
