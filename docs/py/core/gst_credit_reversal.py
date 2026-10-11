"""Electronic Credit Reversal and Re-claimed Statement -- parser and GSTR-3B reconciliation.

The portal statement lists, per GSTR-3B filing, the ITC reversed under Table 4(B)(2) (eligible to
re-claim) and the ITC re-claimed under Table 4(D)(1), with a running closing balance per tax head.
It is the department's own record of the same two lines the taxpayer typed into GSTR-3B, so the
two must agree period by period and head by head.

Reconciled here (per tax period x head: IGST, CGST, SGST/UTGST, Cess):
    3B 4(B)(2)  vs  statement "ITC Reversal (Table 4B(2))"
    3B 4(D)(1)  vs  statement "ITC Reclaimed (Table 4D(1))"   (current form only; in the pre-Aug-2022
                                                               form 4(D)(1) is the s.17(5) line, a
                                                               different thing, so it is not compared)
and the statement's own arithmetic: opening + reversal - reclaimed = closing, row by row.

A tax period whose return is filed after the statement's "To" date cannot appear in it; such a
period is reported as outside the statement window, never as a mismatch.
"""

import csv
import re
import datetime as _dt

HEADS = ("IGST", "CGST", "SGST/UTGST", "Cess")
TOL = 1.0   # Rs; the portal rounds to whole rupees

_MON = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def _num(x):
    s = str(x).strip().replace(",", "")
    if s in ("", "-"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _date(s):
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$", str(s))
    return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def _period(s):
    """'Aug-24' -> 'Aug-24' normalised; anything else -> None."""
    m = re.match(r"^\s*([A-Za-z]{3})[a-z]*[- ]+(\d{2}|\d{4})\s*$", str(s))
    if not m or m.group(1).lower() not in _MON:
        return None
    return f"{m.group(1).title()}-{m.group(2)[-2:]}"


def parse_credit_reversal_statement(path):
    """-> dict(available, gstin, from_date, to_date, opening[4], closing[4], rows=[...], issues=[...]).
    Each row: dict(sno, date, ref, period, desc, claimed[4], reversed[4], reclaimed[4], closing[4]).
    Content-based: columns are located by the two header rows, not by fixed position."""
    out = dict(available=False, gstin=None, from_date=None, to_date=None,
               opening=[0.0] * 4, closing=[0.0] * 4, rows=[], issues=[])
    try:
        raw = list(csv.reader(open(path, encoding="utf-8-sig", newline="")))
    except OSError as exc:
        out["issues"].append(f"could not open: {exc}")
        return out
    hdr = None
    for i, r in enumerate(raw):
        cells = [c.strip() for c in r]
        low = " ".join(cells).lower()
        if "gstin" in [c.lower() for c in cells]:
            k = [c.lower() for c in cells].index("gstin")
            if k + 1 < len(cells):
                out["gstin"] = cells[k + 1] or None
        elif cells and cells[0:4] and any(c.lower() == "from" for c in cells):
            out["from_date"] = _date(cells[[c.lower() for c in cells].index("from") + 1])
        elif any(c.lower() == "to" for c in cells):
            k = [c.lower() for c in cells].index("to")
            if k + 1 < len(cells):
                out["to_date"] = _date(cells[k + 1])
        if "return period" in low and "description" in low:
            hdr = i
            break
    if hdr is None:
        out["issues"].append("header row (Return Period / Description) not found")
        return out
    # Group start columns come from the top header row: 'ITC Claimed', 'ITC Reversal', 'ITC Reclaimed',
    # 'Closing Balance'; each group is four tax-head columns wide.
    top = [c.strip().lower() for c in raw[hdr]]

    def start(pat):
        for j, c in enumerate(top):
            if re.search(pat, c):
                return j
        return None

    cols = dict(claimed=start(r"itc claimed"), reversed=start(r"itc reversal"),
                reclaimed=start(r"itc reclaimed"), closing=start(r"closing balance"))
    if any(v is None for v in cols.values()):
        out["issues"].append(f"statement group columns not all found: {cols}")
        return out
    di = top.index("description") if "description" in top else 4
    for r in raw[hdr + 1:]:
        cells = [c.strip() for c in r]
        if len(cells) <= di or not any(cells):
            continue
        desc = cells[di]
        if not desc or desc.lower().startswith("integrated") or cells[0] == "":
            continue    # the head-name sub-header row, or blank

        def grp(name):
            c0 = cols[name]
            return [_num(cells[c0 + k]) if c0 + k < len(cells) else 0.0 for k in range(4)]

        dl = desc.lower()
        if dl.startswith("opening"):
            out["opening"] = grp("closing")
        elif dl.startswith("closing"):
            out["closing"] = grp("closing")
        else:
            out["rows"].append(dict(sno=cells[0], date=_date(cells[1]), ref=cells[2],
                                    period=_period(cells[3]), desc=desc, claimed=grp("claimed"),
                                    reversed=grp("reversed"), reclaimed=grp("reclaimed"),
                                    closing=grp("closing")))
    out["available"] = bool(out["rows"])
    if not out["rows"]:
        out["issues"].append("no filing rows found")
    return out


def check_statement_arithmetic(st):
    """opening + reversal - reclaimed = closing, row by row and head by head. -> [issue strings]."""
    bad, prev = [], list(st["opening"])
    for r in st["rows"]:
        for k in range(4):
            exp = prev[k] + r["reversed"][k] - r["reclaimed"][k]
            if abs(exp - r["closing"][k]) > TOL:
                bad.append(f"{r['period'] or r['ref']} {HEADS[k]}: opening {prev[k]:,.2f} + reversal "
                           f"{r['reversed'][k]:,.2f} - reclaimed {r['reclaimed'][k]:,.2f} = {exp:,.2f}, "
                           f"statement closing says {r['closing'][k]:,.2f}")
        prev = list(r["closing"])
    return bad


def reconcile(st, g3b_by_month, extra_by_month, months):
    """-> dict(lines=[...], periods=[...], arithmetic=[...]).
    A line is one tax period x head with a non-zero figure on either side."""
    by_period = {}
    for r in st["rows"]:
        if r["period"]:
            acc = by_period.setdefault(r["period"], dict(rev=[0.0] * 4, rec=[0.0] * 4))
            for k in range(4):
                acc["rev"][k] += r["reversed"][k]
                acc["rec"][k] += r["reclaimed"][k]
    to = st.get("to_date")
    lines, periods = [], []
    for m in months:
        g = g3b_by_month.get(m, {}) or {}
        ex = extra_by_month.get(m, {}) or {}
        b2 = [float(x or 0) for x in (g.get("4B2") or [0, 0, 0, 0])[:4]]
        d1_new = ex.get("D1_kind") == "reclaim"
        d1 = [float(x or 0) for x in (ex.get("D1") or [0, 0, 0, 0])[:4]] if d1_new else None
        stm = by_period.get(m)
        # Is this period's return filed after the statement's To date?  The 3B for month M is filed in
        # M+1, so a period is outside the window when its own month is the To-date month or later.
        outside = False
        if stm is None and to is not None:
            mo, yy = _MON[m[:3].lower()], 2000 + int(m[-2:])
            outside = (yy, mo) >= (to.year, to.month)
        status = "OUTSIDE STATEMENT WINDOW" if outside else "COMPARED"
        any_gap = False
        for k in range(4):
            sr = stm["rev"][k] if stm else 0.0
            sc = stm["rec"][k] if stm else 0.0
            rev_gap = b2[k] - sr
            rec_gap = (d1[k] - sc) if d1 is not None else None
            if not (b2[k] or sr or (d1 and d1[k]) or sc):
                continue
            ok = outside or (abs(rev_gap) <= TOL and (rec_gap is None or abs(rec_gap) <= TOL))
            if not outside and not ok:
                any_gap = True
            lines.append(dict(period=m, head=HEADS[k], g3b_rev=b2[k], st_rev=None if outside else sr,
                              rev_gap=None if outside else rev_gap, g3b_rec=None if d1 is None else d1[k],
                              st_rec=None if outside else sc, rec_gap=None if outside else rec_gap,
                              status="OUTSIDE WINDOW" if outside else ("OK" if ok else "MISMATCH")))
        if not outside and stm is None and not any(b2) and not (d1 and any(d1)):
            status = "NIL BOTH SIDES"
        elif any_gap:
            status = "MISMATCH"
        periods.append((m, status))
    return dict(lines=lines, periods=periods, arithmetic=check_statement_arithmetic(st))
