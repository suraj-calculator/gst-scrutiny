"""Plain-script tests for gst_credit_reversal (run: python3 test_credit_reversal.py)."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gst_credit_reversal as cr

CSV = """Electronic Credit Reversal and Re-claimed Statement

,,,GSTIN,05AAAFB8782C1ZO,
,,,Legal Name,X,
,,,From,24/08/2024,
,,,To,31/03/2025
S.No.,Date,Reference No.,Return Period,Description,ITC Claimed (Table 4A(5)) (Rs),,,,ITC Reversal (Table 4B(2)) (Rs),,,,ITC Reclaimed (Table 4D(1)) (Rs),,,,Closing Balance (Rs),,,,
,,,,,Integrated Tax,Central Tax,State/UT Tax,Cess,Integrated Tax,Central Tax,State/UT Tax,Cess,Integrated Tax,Central Tax,State/UT Tax,Cess,Integrated Tax,Central Tax,State/UT Tax,Cess
1,-,-,-,Opening Balance,-,-,-,-,-,-,-,-,-,-,-,-,100,0,0,0
2,20/09/2024,A,Aug-24,Filing of Form GSTR-3B,9,9,9,9,0,0,0,0,100,0,0,0,0,0,0,0,
3,20/10/2024,B,Sep-24,Filing of Form GSTR-3B,9,9,9,9,0,50,50,7,0,0,0,0,0,50,50,7,
4,-,-,-,Closing Balance,-,-,-,-,-,-,-,-,-,-,-,-,0,50,50,7
"""
def run():
    d = tempfile.mkdtemp()
    p = os.path.join(d, "s.csv"); open(p, "w").write(CSV)
    st = cr.parse_credit_reversal_statement(p)
    assert st["available"] and st["gstin"] == "05AAAFB8782C1ZO" and len(st["rows"]) == 2, st
    assert st["opening"] == [100, 0, 0, 0] and st["closing"] == [0, 50, 50, 7]
    assert cr.check_statement_arithmetic(st) == []
    months = ["Aug-24", "Sep-24", "Oct-24", "Mar-25"]
    rec = {"D1_kind": "reclaim", "available": True}
    ex = {"Aug-24": dict(rec, D1=[100, 0, 0, 0]), "Sep-24": dict(rec, D1=[0] * 4),
          "Oct-24": dict(rec, D1=[0] * 4), "Mar-25": dict(rec, D1=[0] * 4)}
    g = {"Aug-24": {"4B2": [0] * 4}, "Sep-24": {"4B2": [0, 50, 50, 7]}, "Oct-24": {"4B2": [0] * 4},
         "Mar-25": {"4B2": [5, 0, 0, 0]}}
    r = cr.reconcile(st, g, ex, months)
    assert all(l["status"] in ("OK", "OUTSIDE WINDOW") for l in r["lines"]), r["lines"]
    assert dict(r["periods"])["Mar-25"] == "OUTSIDE STATEMENT WINDOW"
    # 3B shows a different CGST reversal and a reclaim the portal never recorded -> two mismatches
    g["Sep-24"]["4B2"] = [0, 60, 50, 7]
    ex["Oct-24"]["D1"] = [0, 0, 0, 30]
    bad = [(l["period"], l["head"]) for l in cr.reconcile(st, g, ex, months)["lines"] if l["status"] == "MISMATCH"]
    assert bad == [("Sep-24", "CGST"), ("Oct-24", "Cess")], bad
    # old-form 4(D)(1) is the s.17(5) line: never compared with the reclaim column
    ex["Oct-24"] = {"D1_kind": "17(5)", "D1": [0, 0, 0, 30], "available": True}
    bad = [(l["period"], l["head"]) for l in cr.reconcile(st, g, ex, months)["lines"] if l["status"] == "MISMATCH"]
    assert bad == [("Sep-24", "CGST")], bad
    # broken arithmetic is reported
    st["rows"][1]["closing"][1] = 99
    assert cr.check_statement_arithmetic(st)
    print("test_credit_reversal: OK")
run()
