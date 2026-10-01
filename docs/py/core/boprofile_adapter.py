#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BO PROFILE CANONICAL LAYER  (spec: docs/BOPROFILE_CANONICAL_SPEC.md)
======================================================================
Only this module knows what the GST department's "BO Profile" (360-degree taxpayer profile) Excel export looks like.

    raw BO Profile  --build_canonical()-->  canonical data
    data            --write_canonical()-->  Canonical_BOPROFILE_<GSTIN>_<FY(s)>.xlsx
    .xlsx           --read_canonical()  -->  data
    data            --parse_bo_profile() -->  exactly what gst_parsers_dept.parse_bo_profile() returns

The export has ~35 sheets; the engine reads 16 of them (the rest - places of business, bank accounts, members, ... - are not used by any
check and are listed as such in the ISSUES sheet): the demographic label/value rows; the tables keyed by Financial Year (financial
information, BIFA, ITC passed on / received, e-way bill, e-invoice, refund); and the flat lists (top beneficiaries / suppliers, ITC from /
to related cancelled parties, DRC payments, appeals, cases, transfers). Each becomes one canonical sheet with canonical column names and
the source row number. As the original reader did, every column is found by its HEADING (first column whose heading contains all the
fragments in ledger-style 'match' lists in boprofile_mapping.json); a column that cannot be found is now REPORTED instead of silently
reading as blank. BO amounts are in lakhs, exactly as in the source.
"""

import datetime as _dt
import os
import re

import openpyxl
from openpyxl.utils import get_column_letter

import canonical_common as cc
import gst_core as mpu

SCHEMA_VERSION = "BO-1.0"
ADAPTER_VERSION = "1.0.0"
_HERE = os.path.dirname(os.path.abspath(__file__))
_MAPPING_FILE = os.path.join(_HERE, "boprofile_mapping.json")
LEAD_COLS = ["src_row"]
MAP_COLS = ["canonical_sheet", "canonical_field", "raw_sheet", "raw_header_found", "raw_column", "matched_by", "confidence",
            "status", "rows_filled", "rows_blank", "note"]
ISSUE_COLS = cc.ISSUE_COLS
_FY_RE = re.compile(r"^\d{4}-\d{2,4}$")
_CONTEXT = {"gstin": None}


def set_context(gstin=None):
    if gstin:
        _CONTEXT["gstin"] = gstin


def _tables():
    return cc.load_json(_MAPPING_FILE)["tables"]


def _canon_cols(tname):
    t = _tables()[tname]
    if t["kind"] == "kv":
        return ["src_row", "label", "value"]
    lead = ["src_row", "fy"] if t["kind"] in ("fy", "bifa", "itc") else ["src_row"]
    return lead + list(t["fields"])


def _num_bo(s):
    """Legacy gst_parsers_dept._num_bo: numbers as text with a trailing '%' or commas; blank / '-' / NA -> 0."""
    if s is None:
        return 0.0
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).replace(",", "").replace("%", "").strip()
    if s in ("", "-", "NA", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _s(v):
    return str(v or "").strip()


def _f(v):
    return None if v is None else float(v)


def _header_row(rows, max_scan=6):
    """(1-based row number, header strings) of the first of the top rows with >= 2 non-empty cells (legacy _bxl_header_row)."""
    for i, row in enumerate(rows[:max_scan], start=1):
        if len([c for c in row if c not in (None, "")]) >= 2:
            return i, [str(c).strip() if c is not None else "" for c in row]
    return None, []


def _col(header, frags):
    for i, h in enumerate(header):
        low = h.lower()
        if all(f.lower() in low for f in frags):
            return i
    return None


def _sheet_name(wb, tdef):
    if tdef["sheet"] in wb.sheetnames:
        return tdef["sheet"]
    pre = tdef.get("prefix")
    if pre:
        return next((n for n in wb.sheetnames if n.lower().startswith(pre.lower())), None)
    return None


# ----------------------------------------------------------------------------------------
# build
# ----------------------------------------------------------------------------------------
def _fy_rows(rows, hdr_row, fy_col, cols, fields, plain, first_row):
    out = []
    for ridx, row in enumerate(rows[first_row - 1:], start=first_row):
        if not row or (fy_col >= len(row)) or row[fy_col] is None:
            continue
        fy = str(row[fy_col]).strip()
        if not _FY_RE.match(fy):
            continue
        rec = dict(src_row=ridx, fy=fy)
        for f, spec in fields.items():
            if "plain_supply" in spec:
                ci = plain[spec["plain_supply"]] if len(plain) >= 2 else None
            else:
                ci = cols[f]
            rec[f] = _num_bo(row[ci]) if (ci is not None and ci < len(row)) else None
        out.append(rec)
    return out


def build_canonical(path):
    """Convert one BO Profile workbook into canonical data. A file that cannot be opened raises PeriodParseError ('[E1201] ...')."""
    if isinstance(path, (list, tuple)):
        path = path[0]
    ctx = cc.IssueLog("BO Profile", [os.path.basename(path)])
    try:
        wb = openpyxl.load_workbook(path, data_only=True)       # full (not read-only) load: government files can carry a wrong <dimension>
    except Exception as e:  # noqa: BLE001
        ctx.fail("E1201", f"BO Profile \"{os.path.basename(path)}\": cannot open the file ({type(e).__name__}: {e}). It may be corrupt, "
                 f"password-protected or not an Excel file.", skipped="every check that reads the BO Profile", action="Re-download the file.")
    tables = _tables()
    table_rows = {t["canon_sheet"]: [] for t in tables.values()}
    mapping, absent, used_sheets, no_header = [], [], set(), []
    demo = {}
    for tname, tdef in tables.items():
        sheet = _sheet_name(wb, tdef)
        canon = tdef["canon_sheet"]
        if sheet is None:
            absent.append(tdef["sheet"])
            continue
        used_sheets.add(sheet)
        rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
        kind = tdef["kind"]
        if kind == "kv":
            for ridx, row in enumerate(rows, start=1):
                if row and row[0] and len(row) > 1 and row[1] not in (None, ""):
                    table_rows[canon].append(dict(src_row=ridx, label=str(row[0]).strip(), value=row[1]))
                    demo[str(row[0]).strip()] = row[1]
            continue
        hdr_row, header = _header_row(rows)
        if hdr_row is None:
            no_header.append(sheet)
            ctx.issue("W1201", "WARNING", sheet, "", "all",
                      f"BO Profile sheet '{sheet}' has no recognisable heading row in its top rows, so it reads as empty.",
                      f"the {sheet} figures in every check that uses them", "Check that this is the standard BO Profile export.")
            continue
        fields = tdef["fields"]
        cols, how, plain = {}, {}, []
        if kind == "bifa":
            sub = [str(c).strip() if c else "" for c in rows[hdr_row]] if hdr_row < len(rows) else []
            combined = [(sub[i] if i < len(sub) and sub[i] else (header[i] if i < len(header) else ""))
                        for i in range(max(len(header), len(sub)))]
            label_row, fy_col, first_row = combined, (_col(combined, ["financial year"]) or 0), hdr_row + 2
        else:
            label_row, first_row = header, hdr_row + 1
            fy_col = (_col(header, ["financial year"]) if kind == "fy" else None)
            if kind == "fy" and fy_col is None:
                fy_col = 0
            if kind == "itc":
                fy_col = 0
        if kind == "itc":
            grp, cur = [], ""
            for c in range(max(len(header), 7)):
                v = header[c] if c < len(header) else ""
                if v:
                    cur = cc.clean(v)
                grp.append(cur)
            sub_clean = [cc.clean(rows[hdr_row][c]) if (hdr_row < len(rows) and c < len(rows[hdr_row]) and rows[hdr_row][c] is not None) else ""
                         for c in range(len(grp))]
            lab = [grp[c] + sub_clean[c] for c in range(len(grp))]
            used = set()
            for f, spec in fields.items():
                p = spec["pos"]
                fits = lambda c: c < len(lab) and any(all(t in lab[c] for t in alt) for alt in spec["match"])
                if fits(p) and p not in used:
                    cols[f], how[f] = p, ("position", "HIGH", "")
                    used.add(p)
                else:
                    cand = [c for c in range(len(lab)) if fits(c) and c not in used]
                    if len(cand) == 1:
                        cols[f], how[f] = cand[0], ("heading", "MEDIUM", f"expected at column {get_column_letter(p + 1)}")
                        used.add(cand[0])
                        ctx.issue("W1202", "WARNING", sheet, f, "all",
                                  f"BO Profile sheet '{sheet}': '{f}' is not at its usual column {get_column_letter(p + 1)}; found by "
                                  f"heading at column {get_column_letter(cand[0] + 1)} and read from there.")
                    else:
                        cols[f], how[f] = p, ("position", "LOW", "heading at this column could not be confirmed")
                        ctx.issue("W1203", "WARNING", sheet, f, "all",
                                  f"BO Profile sheet '{sheet}': cannot confirm the column for '{f}' from its heading (column "
                                  f"{get_column_letter(p + 1)} reads '{lab[p] if p < len(lab) else ''}'); read by position as before - "
                                  f"check the figures.")
            display = lab
        else:
            for f, spec in fields.items():
                if "plain_supply" in spec:
                    continue
                cols[f] = _col(label_row, spec["match"][0])
                how[f] = ("heading", "HIGH", "") if cols[f] is not None else ("", "", "")
            if kind in ("fy", "bifa") and any("plain_supply" in s for s in fields.values()):
                plain = [i for i, h in enumerate(label_row) if "supply" in h.lower() and "tax" not in h.lower()]
            display = label_row
        for f, spec in fields.items():
            if "plain_supply" in spec:
                c = plain[spec["plain_supply"]] if len(plain) >= 2 else None
                how[f] = ("heading", "HIGH", "column containing 'supply' but not 'tax'") if c is not None else ("", "", "")
            else:
                c = cols[f]
            mapping.append(dict(canonical_sheet=canon, canonical_field=f, raw_sheet=sheet,
                                raw_header_found=(display[c] if (c is not None and c < len(display)) else ""),
                                raw_column=get_column_letter(c + 1) if c is not None else "", matched_by=how[f][0],
                                confidence=how[f][1], status="OK" if c is not None else "MISSING", rows_filled=0, rows_blank=0,
                                note=how[f][2]))
        miss = [m["canonical_field"] for m in mapping if m["canonical_sheet"] == canon and m["status"] == "MISSING"]
        if miss:
            ctx.issue("I1201", "INFO", sheet, ", ".join(miss), "all",
                      f"BO Profile sheet '{sheet}': column(s) {miss} not found; they read as blank where a check uses them.")
        if kind in ("fy", "bifa", "itc"):
            if kind == "itc":
                rws = []
                for ridx, row in enumerate(rows[first_row - 1:], start=first_row):
                    if not row or row[0] is None or not _FY_RE.match(str(row[0]).strip()):
                        continue
                    rec = dict(src_row=ridx, fy=str(row[0]).strip())
                    for f, spec in fields.items():
                        ci = cols[f]
                        rec[f] = _num_bo(row[ci]) if ci < len(row) else None
                    rws.append(rec)
                table_rows[canon] = rws
            else:
                table_rows[canon] = _fy_rows(rows, hdr_row, fy_col, cols, fields, plain, first_row)
        else:                                                    # flat list
            for ridx, row in enumerate(rows[hdr_row:], start=hdr_row + 1):
                if not row or all(c is None for c in row):
                    continue
                rec = dict(src_row=ridx)
                for f in fields:
                    ci = cols[f]
                    rec[f] = row[ci] if (ci is not None and ci < len(row)) else None
                table_rows[canon].append(rec)
        rowsn = table_rows[canon]
        for m in mapping:
            if m["canonical_sheet"] == canon:
                m["rows_filled"] = sum(1 for r in rowsn if r.get(m["canonical_field"]) not in (None, ""))
                m["rows_blank"] = len(rowsn) - m["rows_filled"]
    unread = [n for n in wb.sheetnames if n not in used_sheets]
    wb.close()
    if absent:
        ctx.issue("I1202", "INFO", "", "", "all", f"Sheets the engine can use that are not in this file (read as empty): {absent}.")
    if unread:
        ctx.issue("I1203", "INFO", "", "", "all", f"Sheets not used by any check ({len(unread)}): {unread}.")
    fys = sorted({r["fy"] for r in table_rows[tables["financial"]["canon_sheet"]]})
    n_err = sum(1 for i in ctx.issues if i["severity"] == "ERROR")
    n_warn = sum(1 for i in ctx.issues if i["severity"] == "WARNING")
    meta = [("schema_version", SCHEMA_VERSION), ("recipient_gstin", _CONTEXT["gstin"] or _s(demo.get("GSTIN"))),
            ("legal_name", _s(demo.get("Legal Name"))), ("trade_name", _s(demo.get("Trade Name"))),
            ("fy", ", ".join(fys)), ("source_files", "; ".join(ctx.source_files)), ("adapter_version", ADAPTER_VERSION),
            ("created_on", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("sheets_present", "; ".join(sorted(used_sheets))), ("unread_sheets", "; ".join(unread)),
            ("status", "OK_WITH_ERRORS" if n_err else ("OK_WITH_WARNINGS" if n_warn else "OK"))]
    return dict(meta=meta, tables=table_rows, mapping=mapping, issues=ctx.issues)


# ----------------------------------------------------------------------------------------
# write / read
# ----------------------------------------------------------------------------------------
_WRITTEN = {}


def write_canonical(data, out_dir):
    m = dict(data["meta"])
    fys = [x for x in str(m.get("fy") or "").split(", ") if x]
    if len(fys) > 2:                                   # a BO Profile lists ~10 years: name the file by the span, not every year
        m["fy"] = f"{fys[0]}_to_{fys[-1]}"
    name = cc.canonical_filename("BOPROFILE", list(m.items()))
    src = dict(data["meta"]).get("source_files")
    base, ext = os.path.splitext(name)
    n = 1
    while _WRITTEN.get(os.path.abspath(os.path.join(out_dir, name)), src) != src:
        n += 1
        name = f"{base}_{n}{ext}"
    sheets = [(t["canon_sheet"], _canon_cols(k), data["tables"][t["canon_sheet"]]) for k, t in _tables().items()]
    sheets += [("MAPPING_REPORT", MAP_COLS, data["mapping"]), ("ISSUES", ISSUE_COLS, data["issues"])]
    path = cc.write_workbook(out_dir, name, data["meta"], sheets, max_col_width=28)
    _WRITTEN[os.path.abspath(path)] = src
    return path


def is_canonical_file(path):
    return cc.is_canonical(path, "BO-", ("META", "BO_DEMOGRAPHIC", "BO_FINANCIAL", "MAPPING_REPORT"))


def read_canonical(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return dict(meta=cc.read_meta(wb),
                    tables={t["canon_sheet"]: cc.read_table(wb, t["canon_sheet"], _canon_cols(k)) for k, t in _tables().items()},
                    mapping=cc.read_table(wb, "MAPPING_REPORT", MAP_COLS), issues=cc.read_table(wb, "ISSUES", ISSUE_COLS))
    finally:
        wb.close()


_SRC = cc.CanonicalCache("BO Profile", is_canonical_file, read_canonical, build_canonical, write_canonical,
                         lambda d: dict(meta=dict(d["meta"])))


def get_data(path, out_dir=None):
    return _SRC.get(path, out_dir)


def clear_cache():
    _SRC.clear()
    _WRITTEN.clear()


# ----------------------------------------------------------------------------------------
# engine-facing reader
# ----------------------------------------------------------------------------------------
def _empty():
    return dict(self_gstin=None, legal_name=None, trade_name=None, demographic={}, financial_by_fy={}, bifa_by_fy={},
                itc_passed_by_fy={}, itc_received_by_fy={}, ewb_by_fy={}, einv_by_fy={}, refund_by_fy={},
                top_beneficiaries=[], top_suppliers=[], related_itc_received=[], related_itc_passed=[],
                drc_payments=[], appeals=[], cases=[], transfers=[])


def parse_bo_profile(path):
    """Legacy gst_parsers_dept.parse_bo_profile(path)."""
    out = _empty()
    if not path or not os.path.exists(path):
        return out
    try:
        data = get_data(path)
    except mpu.PeriodParseError as ex:
        print(f"[error] BO Profile could not be read -> its checks show N/A: {ex}")
        return out
    T = _tables()
    tabs = data["tables"]
    ok = {(m["canonical_sheet"], m["canonical_field"]): m["status"] == "OK" for m in data["mapping"]}

    def rows(name):
        return tabs[T[name]["canon_sheet"]]

    demo = {}
    for r in rows("demographic"):
        demo[r["label"]] = r["value"]
    if demo or tabs[T["demographic"]["canon_sheet"]]:
        out["demographic"] = demo
        out["self_gstin"], out["legal_name"], out["trade_name"] = demo.get("GSTIN"), demo.get("Legal Name"), demo.get("Trade Name")

    def fy_table(name, keys):
        d = {}
        for r in rows(name):
            d[r["fy"]] = {k: _f(r[k]) for k in keys}
        return d

    out["financial_by_fy"] = fy_table("financial", T["financial"]["fields"])
    out["bifa_by_fy"] = fy_table("bifa", T["bifa"]["fields"])
    out["itc_passed_by_fy"] = fy_table("itc_passed", T["itc_passed"]["fields"])
    out["itc_received_by_fy"] = fy_table("itc_received", T["itc_received"]["fields"])
    plain_ok = ok.get((T["ewb"]["canon_sheet"], "inward_supply"))
    ewb_keys = [k for k in T["ewb"]["fields"] if k not in ("inward_supply", "outward_supply")]
    for r in rows("ewb"):
        rec = {k: _f(r[k]) for k in ewb_keys}
        if plain_ok:
            rec["inward_supply"], rec["outward_supply"] = _f(r["inward_supply"]), _f(r["outward_supply"])
        out["ewb_by_fy"][r["fy"]] = rec
    out["einv_by_fy"] = fy_table("einv", T["einv"]["fields"])
    out["refund_by_fy"] = fy_table("refund", T["refund"]["fields"])
    for name, key in (("top_beneficiaries", "top_beneficiaries"), ("top_suppliers", "top_suppliers"),
                      ("related_itc_received", "related_itc_received"), ("related_itc_passed", "related_itc_passed"),
                      ("appeals", "appeals"), ("cases", "cases"), ("transfers", "transfers")):
        flds = list(T[name]["fields"])
        out[key] = [{k: r[k] for k in flds} for r in rows(name)]
    drc = [{k: r[k] for k in T["drc_payments"]["fields"]} for r in rows("drc_payments")]
    for r in drc:
        r["ref"] = r["source_id"]
        r["transaction_date"] = r["date"]
        r["type"] = r["description"]
        r["total"] = sum(_num_bo(r.get(k)) for k in ("cgst", "sgst", "igst", "cess", "other"))
        for k in ("cgst", "sgst", "igst", "cess", "other"):
            r[k] = _num_bo(r.get(k))
    out["drc_payments"] = drc
    return out
