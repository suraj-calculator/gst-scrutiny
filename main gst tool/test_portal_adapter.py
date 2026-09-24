#!/usr/bin/env python3
"""
Regression tests for the portal-comparison canonical layer (portal_adapter.py, docs/PORTAL_COMPARISON_CANONICAL_SPEC.md).

Self-contained: builds synthetic 'Tax liability and ITC comparison' reports in both real layouts (the newer two-row heading and the older
three-row heading with blank ' ' cells for the not-yet-existing 'considering reversal' block) and checks that the ORIGINAL
parse_portal_comparison() and the canonical route return identical results, and how layout changes / broken files are reported.

Run:  python3 test_portal_adapter.py
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
os.environ["GST_PORTAL_CANONICAL"] = "0"          # original reader for the legacy side

import gst_core as mpu                        # noqa: E402
import gst_parsers_dept as dept               # noqa: E402
import portal_adapter as ad                   # noqa: E402

SUB = ["As per GSTR-1/IFF/GSTR-1A", "As per GSTR-3B", "Shortfall (-)/ Excess (+) in liability", "Cumulative Shortfall (-)/ Excess (+) in liability",
       "Cumulative Shortfall (-)/ Excess (+) in liability as percentage (%)", "As per GSTR-3B", "As per GSTR-2B", "Shortfall (-)/ Excess (+) in ITC",
       "Cumulative Shortfall (-)/ Excess (+) in ITC", "Cumulative Shortfall (-)/ Excess (+) in ITC as percentage (%)", "As per GSTR-3B",
       "Shortfall (-)/ Excess (+) in ITC", "Cumulative Shortfall (-)/ Excess (+) in ITC", "Cumulative Shortfall (-)/ Excess (+) in ITC as percentage (%)"]


def make(path, layout="new", extra_col=False, bad_heading=False, no_sheet=False, short_sub=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Other" if no_sheet else "Comparison Summary"
    ws.append(["Goods and Services Tax - Tax liability and ITC comparison"])
    ws.append([])
    ws.append([])
    ws.append(["GSTIN:   05AAAAA0001A1Z1", None, None, "Legal name:   SAMPLE TAXPAYER", None, None, None, "Report generated at:   20/04/2026"])
    ws.append(["Trade name:   M/S SAMPLE", None, None, "Financial Year:   2025-26"])
    ws.append([])
    ws.append(["Tax liability and ITC summary"])
    g1 = ["Tax Period", "Tax liability as per GSTR-1/IFF/GSTR-1A and as per GSTR-3B"] + [None] * 4 + ["ITC claimed in GSTR-3B (excluding ITC Reversal) and accrued as per GSTR-2B"] + [None] * 4 + \
         ["ITC claimed in GSTR-3B (Considering ITC Reversal) and accrued as per GSTR-2B"] + [None] * 3
    sub = [None] + list(SUB)
    if layout == "old":
        g1[6] = "ITC claimed in GSTR-3B and accrued as per GSTR-2B"
        g1[11] = "ITC claimed in GSTR-3B (Considering ITC Reversal)"
        mid = [None] * 6 + ["ITC claimed in GSTR-3B(excluding ITC Reversal)"] + [None] * 8
    if bad_heading:
        sub[2] = "Something entirely different"
    if extra_col:
        g1.insert(1, "New column")
        sub.insert(1, "New sub")
    ws.append(g1)
    if layout == "old":
        ws.append(mid)
    ws.append(sub if not short_sub else sub[:3])
    rows = [["Apr-25", 100.0, 100.0, 0, -5, -1.5, 50.0, 60.0, -10.0, -10.0, -0.5, 49.0, -11.0, -11.0, -0.4],
            ["May-25", "1,000.50", 900, -100.5, -105.5, -2, 70, 70, 0, -10, 0, 69, -1, -12, -0.3],
            ["Jun-25", 0, 0, 0, 0, 0, 5, 5, 0, -10, 0, None, None, None, None]]
    if layout == "old":
        rows = [r[:11] + [" "] * 4 for r in rows]
    for r in rows:
        ws.append(r[:1] + ["X"] + r[1:] if extra_col else r)
    ws.append(["Total", 1100.5, 1000, -100.5, -105.5, -2, 125, 135, -10, -10, 0, 118, -12, -12, -0.7])
    wb.create_sheet("Tax liability").append(["not used"])
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
    return run(dept.parse_portal_comparison, p), run(ad.parse_portal_comparison, p)


def main():
    ok = True
    tmp = tempfile.mkdtemp()
    cwd = os.getcwd()
    os.chdir(tmp)
    try:
        for layout in ("new", "old"):
            p = make(os.path.join(tmp, f"pc_{layout}.xlsx"), layout=layout)
            old, new = both(p)
            ok &= check(f"[{layout} layout] identical to the original", old == new and old[0] == "ok" and len(old[1]) == 3, str(old)[:80] + " | " + str(new)[:80])
            ok &= check(f"[{layout} layout] every heading confirmed, no warnings", not [i for i in ad.get_data(p)["issues"] if i["severity"] != "INFO"],
                        str([i["id"] for i in ad.get_data(p)["issues"] if i["severity"] != "INFO"]))
        r = ad.parse_portal_comparison(os.path.join(tmp, "pc_new.xlsx"))
        ok &= check("period rows only (the Total row is not a period); text amount '1,000.50' read as a number", list(r) == ["Apr-25", "May-25", "Jun-25"] and r["May-25"]["gstr1_liability"] == 1000.5)
        ok &= check("'considering reversal' block: number, or None when the cell is blank (Jun-25)", r["Apr-25"]["itc_3b_adj"] == 49.0 and r["Jun-25"]["itc_3b_adj"] is None and r["Jun-25"]["cum_diff_itc_adj"] is None)
        ro = ad.parse_portal_comparison(os.path.join(tmp, "pc_old.xlsx"))
        ok &= check("older report (blank ' ' cells for that block): None, like the original", all(v["itc_3b_adj"] is None for v in ro.values()))
        meta = dict(ad.get_data(os.path.join(tmp, "pc_new.xlsx"))["meta"])
        ok &= check("META: GSTIN, legal name and FY from the title block", meta["recipient_gstin"] == "05AAAAA0001A1Z1" and meta["fy"] == "2025-26" and meta["legal_name"] == "SAMPLE TAXPAYER", str(meta))

        # ---- layout changes
        mv = make(os.path.join(tmp, "moved.xlsx"), extra_col=True)
        old, new = both(mv)
        good = ad.parse_portal_comparison(os.path.join(tmp, "pc_new.xlsx"))
        ok &= check("a column inserted after 'Tax Period': followed by heading, same figures as the untouched file", new[0] == "ok" and new[1] == good)
        ok &= check("(documenting the fixed bug) the original, by position, reads the wrong columns", old[0] == "ok" and old[1] != good)
        ok &= check("...and W1402 says which columns moved", any(i["id"] == "W1402" for i in ad.get_data(mv)["issues"]))
        bh = make(os.path.join(tmp, "bad_heading.xlsx"), bad_heading=True)
        old, new = both(bh)
        ok &= check("a heading that no longer matches: read by position as before (identical) + W1403", old == new and any(i["id"] == "W1403" for i in ad.get_data(bh)["issues"]))

        # ---- broken files
        ns = make(os.path.join(tmp, "nosheet.xlsx"), no_sheet=True)
        old, new = both(ns)
        ok &= check("no 'Comparison Summary' sheet: {} like the original, W1401 says so", old == new == ("ok", {}) and any(i["id"] == "W1401" for i in ad.get_data(ns)["issues"]), str(new))
        with open("junk.xlsx", "w") as f:
            f.write("not an excel file")
        ad.clear_cache()
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rj = ad.parse_portal_comparison("junk.xlsx")
        ok &= check("unreadable file: {} (the 'not supplied' shape) with the plain reason printed (E1401); the original crashed", rj == {} and "[E1401]" in buf.getvalue() and run(dept.parse_portal_comparison, "junk.xlsx")[0] == "EXC", buf.getvalue()[:90])
        with open("plain.xlsx", "wb") as f:
            f.write(b"")
        wb = openpyxl.Workbook()
        wb.active.title = "Comparison Summary"
        wb.active.append(["Tax Period", "a"])
        wb.active.append([None, "b"])
        wb.active.append(["Apr-25", 1])
        wb.save("short.xlsx")
        ad.clear_cache()
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rs = ad.parse_portal_comparison("short.xlsx")
        ok &= check("required columns missing (report cut short): E1403 naming them, reads as {}", rs == {} and "[E1403]" in buf.getvalue(), buf.getvalue()[:100])

        # ---- canonical file
        ad.clear_cache()
        p = os.path.join(tmp, "pc_new.xlsx")
        path = ad.write_canonical(ad.build_canonical(p), tmp)
        ok &= check("canonical file named from the report's own GSTIN and FY", os.path.basename(path) == "Canonical_PORTALCOMPARISON_05AAAAA0001A1Z1_2025-26.xlsx", os.path.basename(path))
        ok &= check("canonical file supplied as input reads identically", ad.parse_portal_comparison(path) == dept.parse_portal_comparison(p))
        folder = os.path.join(tmp, "folder")
        os.makedirs(folder)
        shutil.copy(path, folder)
        os.environ["GST_PORTAL_CANONICAL"] = "1"
        with contextlib.redirect_stdout(io.StringIO()):
            try:
                res = mpu.classify_folder(folder)
            except Exception as e:  # noqa: BLE001
                res = getattr(e, "res", None)
        ok &= check("classify_folder recognises a canonical portal-comparison workbook", res is not None and [os.path.basename(f) for f in res["portal_comparison_files"]] == [os.path.basename(path)])
        os.environ["GST_PORTAL_CANONICAL"] = "0"
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nALL PASSED" if ok else "\nSOME TESTS FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
