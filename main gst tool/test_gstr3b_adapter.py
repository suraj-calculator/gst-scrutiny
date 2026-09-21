#!/usr/bin/env python3
"""
Regression tests for the GSTR-3B canonical layer (gstr3b_adapter.py, docs/GSTR3B_CANONICAL_SPEC.md).

Self-contained: builds small synthetic 3B workbooks in BOTH form wordings (before / after Aug-2022) and
checks that the ORIGINAL readers and the canonical route agree, that the documented errors and warnings
appear, and that the one intended difference (4B(1) read from where it sits, whatever its wording) is
what changed. Where the repo's sample per-period 3B files exist locally they are also used.

Run:  python3 test_gstr3b_adapter.py
"""
import datetime as dt
import glob
import os
import shutil
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ["GST_3B_CANONICAL"] = "0"         # original readers for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_returns as pr              # noqa: E402
import gst_checks_hsn_fraud as hf             # noqa: E402
import gst_checks_forensic as fo              # noqa: E402
import gstr3b_adapter as ad                   # noqa: E402

GSTIN = "05AAAAA0001A1Z1"


def sheet_rows(fy, period, layout, k=1.0, arn_date="20/05/2024", drop=(), tp_text=None):
    """One month's 3B sheet as rows of 12 cells, laid out like the real form (labels in column B,
    figures from column E). `layout` = 'new' (Aug-2022 onward) or 'old'. `drop` = line prefixes to omit."""
    def r(label=None, *vals, col=1):
        row = [None] * 12
        if label is not None:
            row[col] = label
        for i, v in enumerate(vals):
            row[4 + i] = v
        return row
    b1 = ("(1) As per rules 38,42 & 43 of CGST Rules and section 17(5)" if layout == "new"
          else "(1) As per rules 42 & 43 of CGST Rules")
    rows = [r(), r("Form GSTR-3B\n[See rule 61(5)]"), r(), r(),
            r("GSTIN", GSTIN), r("Year", fy), r("Tax Period", tp_text or period), r("ARN", "AA0505240000001"),
            r("Date of ARN", arn_date), r("Legal name of the registered person ", "TEST TRADERS"),
            r("Trade name, if any", "M/s TEST"), r("Date and Time of Generation", "01/09/2026, 10:00:00."), r(),
            r("3.1 Details of Outward Supplies and inward supplies liable to reverse charge"), r(),
            r("Nature of Supplies", "Total Taxable value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"),
            r("(a) Outward Taxable  supplies  (other than zero rated, nil rated and exempted)", 1000 * k, 10 * k,
              90 * k, 90 * k, 5 * k),
            r("(b) Outward Taxable  supplies  (zero rated )", 0, 0, None, None, 0),
            r("(c) Other Outward Taxable  supplies (Nil rated, exempted)", 7 * k),
            r("(d) Inward supplies (liable to reverse charge)", 11 * k, 1 * k, 2 * k, 3 * k, 0),
            r("(e) Non-GST Outward supplies", 3 * k), r(), r("4. Eligible ITC"), r(),
            r("Details", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"),
            r("A. ITC Available (whether in full or part)"),
            r("(1) Import of goods", 0, 0, 0, 0), r("(2) Import of services", 0, 0, 0, 0),
            r("(3) Inward supplies liable to reverse charge (other than 1 & 2 above)", 1 * k, 2 * k, 3 * k, 0),
            r("(4) Inward supplies from ISD", 4 * k, 0, 0, 0),
            r("(5) All other ITC", 500 * k, 60 * k, 60 * k, 9 * k),
            r("B. ITC Reversed"), r(b1, 12 * k, 13 * k, 14 * k, 0), r("(2) Others", 20 * k, 1 * k, 1 * k, 0),
            r("C. Net ITC available (A-B)", 469 * k, 47 * k, 47 * k, 9 * k)]
    if layout == "old":
        rows += [r("(D) Ineligibe ITC"), r("(1) As per section 17(5)", 2 * k, 0, 0, 0), r("(2) Others", 0, 0, 0, 0)]
    else:
        rows += [r("(D) Other Details "), r("(1) ITC reclaimed which was reversed under Table 4(B)(2) in earlier "
                                            "tax period", 6 * k, 0, 0, 0)]
    keep = []
    for row in rows:
        label = next((c for c in row if isinstance(c, str)), "")
        if any(label.strip().startswith(d) for d in drop):
            continue
        keep.append(row)
    return keep


MONTHS = [("Apr", "2024-25"), ("May", "2024-25"), ("Jun", "2024-25")]


def make_workbook(path, layout, drop_by_month=None, dup=False, extra_sheet=False, tp_override=None, only=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for i, (mon, fy) in enumerate(MONTHS):
        if only and mon not in only:
            continue
        drop = (drop_by_month or {}).get(mon, ())
        ws = wb.create_sheet(f"{mon}_{fy}")
        for row in sheet_rows(fy, mon, layout, k=1.0 + i, drop=drop,
                              tp_text=(tp_override or {}).get(mon)):
            ws.append(row)
    if dup:
        ws = wb.create_sheet("Apr_2024-25_2")
        for row in sheet_rows("2024-25", "Apr", layout, k=9.0):
            ws.append(row)
    if extra_sheet:
        wb.create_sheet("Help").append(["not a return"])
    wb.save(path)
    return path


def check(name, cond, detail=""):
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    return cond


def exc(fn):
    try:
        fn()
        return None
    except mpu.PeriodParseError as e:
        return str(e)


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        for layout in ("new", "old"):
            raw = make_workbook(os.path.join(tmp, f"3b_{layout}.xlsx"), layout)
            ad.clear_cache()
            same_all = True
            for mon in ("Apr-24", "May-24", "Jun-24"):
                same_all &= pr.parse_gstr3b(raw, mon) == ad.parse_gstr3b(raw, mon)
            ok &= check(f"[{layout} form] parse_gstr3b identical for every month", same_all)
            ok &= check(f"[{layout} form] ARN dates identical", fo.gstr3b_arn_dates_by_month(raw) ==
                        ad.arn_dates_by_month(raw))
            ok &= check(f"[{layout} form] months identical", mpu._gstr3b_months(raw) == ad.months(raw))
            old_f, new_f = hf._gstr3b_month_fields(raw, "May-24"), ad.month_fields(raw, "May-24")
            same = {k: old_f[k] == new_f[k] for k in old_f}
            if layout == "old":
                ok &= check("[old form] fraud-check fields identical (incl. 4B(1) and 4D 17(5))", all(same.values()),
                            str(same))
            else:
                ok &= check("[new form] fraud-check fields identical except rule42_43 (intended fix)",
                            all(v for k, v in same.items() if k != "itc_reversed_rule42_43") and
                            not same["itc_reversed_rule42_43"], str(same))
                ok &= check("[new form] rule42_43 now reads the real 4B(1) figure (12 x 2 = 24), legacy read 0",
                            new_f["itc_reversed_rule42_43"] == 24.0 and old_f["itc_reversed_rule42_43"] == 0.0)
        raw = os.path.join(tmp, "3b_new.xlsx")

        # ---- canonical workbook
        ad.clear_cache()
        data = ad.build_canonical(raw)
        out = ad.write_canonical(data, os.path.join(tmp, "_canonical"))
        ok &= check("canonical file recognised", ad.is_canonical_file(out))
        d = ad.read_canonical(out)
        ok &= check("RETURNS has one row per month with ARN and GSTIN",
                    [r["month"] for r in d["returns"]] == ["Apr-24", "May-24", "Jun-24"] and
                    all(r["gstin"] == GSTIN and r["arn"] for r in d["returns"]))
        ok &= check("LINES names every recognised line for a month",
                    {ln["code"] for ln in d["lines"] if ln["month"] == "May-24"} ==
                    {"3.1a", "3.1b", "3.1c", "3.1d", "3.1e", "4A3", "4A4", "4A5", "4B1", "4B2", "4C"})
        ln = [x for x in d["lines"] if x["month"] == "May-24" and x["code"] == "3.1a"][0]
        ok &= check("LINES columns follow the form headings (taxable / igst / cgst / sgst / cess)",
                    (ln["taxable"], ln["igst"], ln["cgst"], ln["sgst"], ln["cess"]) == (2000.0, 20.0, 180.0, 180.0, 10.0))
        ln = [x for x in d["lines"] if x["month"] == "May-24" and x["code"] == "3.1b"][0]
        ok &= check("a line with blank cells keeps them blank in its column (not shifted)",
                    (ln["taxable"], ln["igst"], ln["cgst"], ln["cess"]) == (0.0, 0.0, None, 0.0))
        c = mpu.classify_folder(os.path.dirname(out))
        ok &= check("canonical file is classified as GSTR-3B and gives its months",
                    c["gstr3b_merged"] is not None and len(c["gstr3b_months"]) == 3)

        # ---- mutation: one changed value in the canonical file changes the reader
        wb = openpyxl.load_workbook(out)
        ws = wb["LINES"]
        hdr = [x.value for x in ws[1]]
        for row in ws.iter_rows(min_row=2):
            if row[hdr.index("month")].value == "Apr-24" and row[hdr.index("code")].value == "4A5":
                row[hdr.index("values_in_row_order")].value = "1.0;2.0;3.0;4.0"
        wb.save("mutant.xlsx")
        ad.clear_cache()
        ok &= check("mutation: editing the canonical file changes what the engine reads",
                    ad.parse_gstr3b("mutant.xlsx", "Apr-24")["4A5"] == [1.0, 2.0, 3.0, 4.0])

        # ---- a REQUIRED line missing in one month: that month is refused loudly, the others still work
        raw2 = make_workbook(os.path.join(tmp, "3b_req.xlsx"), "new", drop_by_month={"May": ("(5) All other ITC",)})
        ad.clear_cache()
        e = exc(lambda: ad.parse_gstr3b(raw2, "May-24"))
        ok &= check("required line 4A5 missing in May -> May raises [E302] naming the line", bool(e) and "[E302]" in e
                    and "4A5" in e, str(e)[:120])
        ok &= check("... while April and June still work",
                    exc(lambda: ad.parse_gstr3b(raw2, "Apr-24")) is None and exc(lambda: ad.parse_gstr3b(raw2, "Jun-24")) is None)
        ok &= check("... whereas the original reader silently carried on with 4A5 absent (zeros)",
                    "4A5" not in pr.parse_gstr3b(raw2, "May-24"))

        # ---- an OPTIONAL line missing: warning, key absent
        raw3 = make_workbook(os.path.join(tmp, "3b_opt.xlsx"), "new", drop_by_month={"Jun": ("(4) Inward supplies from ISD",)})
        ad.clear_cache()
        d3 = ad.build_canonical(raw3)
        ok &= check("optional line 4A4 missing in June -> W303 warning, line absent",
                    any(i["id"] == "W303" and i["field"] == "4A4" and "Jun-24" in i["periods_affected"] for i in d3["issues"])
                    and "4A4" not in ad.parse_gstr3b(raw3, "Jun-24"))

        # ---- duplicate month sheet, a non-3B sheet
        raw4 = make_workbook(os.path.join(tmp, "3b_dup.xlsx"), "new", dup=True, extra_sheet=True)
        ad.clear_cache()
        d4 = ad.build_canonical(raw4)
        ids = {i["id"] for i in d4["issues"]}
        ok &= check("duplicate April sheet -> W301, first sheet used; helper sheet -> I401", {"W301", "I401"} <= ids, str(ids))
        ok &= check("first April sheet wins", ad.parse_gstr3b(raw4, "Apr-24")["3.1a"][0] == 1000.0)

        # ---- quarterly period: legacy behaviour kept (months 2 and 3 are refused by parse_gstr3b)
        raw5 = make_workbook(os.path.join(tmp, "3b_q.xlsx"), "new", tp_override={"Apr": "Apr-Jun"}, only=("Apr",))
        ad.clear_cache()
        ok &= check("quarterly 3B: months() covers Apr-Jun, first month parses, later months refused like before",
                    ad.months(raw5) >= {"Apr-24", "May-24", "Jun-24"} and exc(lambda: ad.parse_gstr3b(raw5, "Apr-24")) is None
                    and exc(lambda: pr.parse_gstr3b(raw5, "May-24")) is not None)

        # ---- errors
        open("corrupt.xlsx", "wb").write(open(raw, "rb").read()[:1500])
        ad.clear_cache()
        ok &= check("E304 corrupt file", (exc(lambda: ad.build_canonical("corrupt.xlsx")) or "").startswith("[E304]"))
        wb = openpyxl.Workbook()
        wb.active.append(["nothing here"])
        wb.save("nothreeb.xlsx")
        ad.clear_cache()
        ok &= check("E301 no sheet with Year / Tax Period", (exc(lambda: ad.build_canonical("nothreeb.xlsx")) or "").startswith("[E301]"))

        # ---- the repo's real per-period sample files (only if present locally)
        files = sorted(glob.glob(os.path.join(HERE, "..", "docs", "sample_data", "gstr3b", "GSTR3B_0*_*.xlsx")))
        merged = os.path.join(HERE, "..", "docs", "sample_data", "gstr3b", "GSTR3B_Merged.xlsx")
        if len(files) >= 12 and os.path.exists(merged):
            ad.clear_cache()
            dper = ad.build_canonical(files)
            ad.clear_cache()
            dmer = ad.build_canonical(merged)
            key = lambda dd: sorted((ln["month"], ln["code"], ln["values_in_row_order"]) for ln in dd["lines"])  # noqa: E731
            ok &= check("12 raw per-period files and the merged workbook give the same canonical lines",
                        key(dper) == key(dmer))
        else:
            print("[SKIP] sample per-period 3B files not present locally")
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
