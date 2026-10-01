#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TABLE 8A CANONICAL LAYER  (spec: docs/TABLE8A_CANONICAL_SPEC.md)
=================================================================
Only this module knows what the government Table 8A workbook (GSTR-2A-style "inward supplies as per GSTR-1 of suppliers", the source of
the GSTR-9 Table 8A ITC figure) looks like.

    raw Table 8A   --build_canonical()-->  canonical data
    data           --write_canonical()-->  Canonical_TABLE8A_<GSTIN>_<FY>.xlsx
    .xlsx          --read_canonical()  -->  data
    data           --parse_table_8a()  -->  exactly what gst_parsers_dept.parse_table_8a() returns

The export has a title, a two-row heading directly below the row that contains 'GSTIN of supplier', and one row per invoice (sheet B2B) /
note (sheet CDNR or the newer B2B-CDNR). The number of columns differs between exports, so - exactly as the original reader did - every
column is located by its heading. The canonical file keeps every row under the heading with canonical column names (numbers converted,
text and dates exactly as in the source); the reader keeps the original's business logic (only GSTIN-shaped rows are data, totals of the
'ITC availability = Yes' invoices, the breakdown of the 'No' rows by reason) over those rows.
"""

import datetime as _dt
import os
import re
from collections import Counter

import openpyxl
from openpyxl.utils import get_column_letter

import canonical_common as cc
import gst_core as mpu

SCHEMA_VERSION = "T8A-1.0"
ADAPTER_VERSION = "1.0.0"
_HERE = os.path.dirname(os.path.abspath(__file__))
_MAPPING_FILE = os.path.join(_HERE, "table8a_mapping.json")
LEAD_COLS = ["src_row"]
MAP_COLS = ["canonical_sheet", "canonical_field", "raw_sheet", "raw_header_found", "raw_column", "matched_by", "confidence",
            "status", "rows_filled", "rows_blank", "note"]
ISSUE_COLS = cc.ISSUE_COLS
_GSTIN_RE = re.compile(r"^\d{2}[A-Z]{5}\d{4}[A-Z]\dZ[A-Z\d]$")
_CONTEXT = {"gstin": None}


def set_context(gstin=None):
    if gstin:
        _CONTEXT["gstin"] = gstin


def _tables():
    return cc.load_json(_MAPPING_FILE)["tables"]


def _canon_cols(tname):
    return LEAD_COLS + list(_tables()[tname]["fields"]) + ["extra"]


def _s(v):
    return str(v or "").strip()


def _f(v):
    return 0.0 if v is None else float(v)


# ----------------------------------------------------------------------------------------
# raw workbook
# ----------------------------------------------------------------------------------------
def _find_header_row(rows, must_contain, max_scan=15):
    """1-based number of the first of the top rows with a cell containing `must_contain` (as the original)."""
    for i, row in enumerate(rows[:max_scan], start=1):
        if any(must_contain.lower() in (str(c).strip() if c else "").lower() for c in row):
            return i
    return None


def _labels(rows, hdr):
    """One label per column: the sub-heading (row below) if it has one, else the group heading (as the original)."""
    top = [str(c).strip() if c else "" for c in rows[hdr - 1]]
    sub = [str(c).strip() if c else "" for c in rows[hdr]] if hdr < len(rows) else []
    n = max(len(top), len(sub))
    return [(sub[i] if i < len(sub) and sub[i] else (top[i] if i < len(top) else "")) for i in range(n)]


def _col(labels, alts):
    """First column whose label contains ALL fragments of the first alternative that hits (the original's _t8a_col chain)."""
    for frags in alts:
        for i, lab in enumerate(labels):
            if all(fr.lower() in lab.lower() for fr in frags):
                return i
    return None


def _readme(wb):
    out = {}
    name = next((n for n in wb.sheetnames if n.lower() == "read me"), None)
    if name:
        for row in wb[name].iter_rows(min_row=1, max_row=15, values_only=True):
            cells = [str(c).strip() for c in row if c not in (None, "")]
            if len(cells) >= 2:
                key = cells[0].lower()
                for lab, k in (("financial year", "fy"), ("gstin", "gstin"), ("legal name", "legal_name")):
                    if key == lab and k not in out:
                        out[k] = cells[1]
    return out


def build_canonical(path):
    """Convert one Table 8A workbook into canonical data. A file that cannot be opened raises PeriodParseError ('[E1101] ...')."""
    if isinstance(path, (list, tuple)):
        path = path[0]
    ctx = cc.IssueLog("Table 8A", [os.path.basename(path)])
    try:
        wb = openpyxl.load_workbook(path, data_only=True)      # full (not read-only) load: government files can carry a wrong <dimension>
    except Exception as e:  # noqa: BLE001
        ctx.fail("E1101", f"Table 8A \"{os.path.basename(path)}\": cannot open the file ({type(e).__name__}: {e}). It may be corrupt, "
                 f"password-protected or not an Excel file.", skipped="every check that reads Table 8A", action="Re-download the file.")
    tables = _tables()
    table_rows = {t["canon_sheet"]: [] for t in tables.values()}
    mapping, meta_status, meta_sheet = [], {}, {}
    readme = _readme(wb)
    for tname, tdef in tables.items():
        sheet = next((n for n in tdef["raw_names"] if n in wb.sheetnames), None)
        meta_sheet[tname] = sheet or ""
        if sheet is None:
            meta_status[tname] = "absent"
            continue
        rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
        hdr = _find_header_row(rows, "GSTIN of supplier")
        if hdr is None:
            meta_status[tname] = "no_header"
            ctx.issue("W1101", "WARNING", sheet, "", "all",
                      f"Table 8A sheet '{sheet}' is present but no row containing 'GSTIN of supplier' was found in its top rows, so the sheet reads as empty.",
                      f"the {sheet} figures in every Table 8A check", "Check that this is the standard Table 8A export.")
            continue
        labels = _labels(rows, hdr)
        cols = {f: _col(labels, spec["match"]) for f, spec in tdef["fields"].items()}
        mapped = {c for c in cols.values() if c is not None}
        for f, spec in tdef["fields"].items():
            c = cols[f]
            mapping.append(dict(canonical_sheet=tdef["canon_sheet"], canonical_field=f, raw_sheet=sheet,
                                raw_header_found=labels[c] if c is not None else "",
                                raw_column=get_column_letter(c + 1) if c is not None else "", matched_by="heading" if c is not None else "",
                                confidence="HIGH" if c is not None else "", status="OK" if c is not None else "MISSING",
                                rows_filled=0, rows_blank=0, note=""))
        if cols["gstin"] is None:
            meta_status[tname] = "no_gstin_col"
            ctx.issue("W1102", "WARNING", sheet, "gstin", "all",
                      f"Table 8A sheet '{sheet}': the 'GSTIN of supplier' column could not be located (headings found: "
                      f"{sorted(set(l for l in labels if l))[:25]}), so the sheet reads as empty.",
                      f"the {sheet} figures in every Table 8A check", "Add the new heading to table8a_mapping.json.")
            continue
        meta_status[tname] = "ok"
        miss = [f for f, c in cols.items() if c is None]
        if miss:
            ctx.issue("I1101", "INFO", sheet, ", ".join(miss), "all",
                      f"Table 8A sheet '{sheet}': column(s) {miss} not found; they read as blank/0 where a check uses them.")
        fill = {f: [0, 0] for f in tdef["fields"]}
        for ridx, r in enumerate(rows[hdr:], start=hdr + 1):
            if not any(c not in (None, "") for c in r):
                continue
            rec = dict(src_row=ridx)
            for f, spec in tdef["fields"].items():
                c = cols[f]
                v = r[c] if (c is not None and c < len(r)) else None
                fill[f][0 if v not in (None, "") else 1] += 1
                rec[f] = cc.num(v) if spec["type"] == "num" else v
            rec["extra"] = "; ".join(f"{labels[i] or get_column_letter(i + 1)}={r[i]}" for i in range(min(len(r), len(labels)))
                                     if i not in mapped and r[i] not in (None, ""))
            table_rows[tdef["canon_sheet"]].append(rec)
        for m in mapping:
            if m["canonical_sheet"] == tdef["canon_sheet"]:
                m["rows_filled"], m["rows_blank"] = fill[m["canonical_field"]]
    wb.close()
    if all(s == "absent" for s in meta_status.values()):
        ctx.issue("W1103", "WARNING", "", "", "all", "Neither a B2B nor a CDNR / B2B-CDNR sheet was found: this does not look like a Table 8A export.")
    n_err = sum(1 for i in ctx.issues if i["severity"] == "ERROR")
    n_warn = sum(1 for i in ctx.issues if i["severity"] == "WARNING")
    meta = [("schema_version", SCHEMA_VERSION), ("recipient_gstin", _CONTEXT["gstin"] or readme.get("gstin", "")),
            ("legal_name", readme.get("legal_name", "")), ("fy", readme.get("fy", "")),
            ("source_files", "; ".join(ctx.source_files)), ("adapter_version", ADAPTER_VERSION),
            ("created_on", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("b2b_sheet", meta_sheet["b2b"]), ("b2b_status", meta_status["b2b"]),
            ("cdnr_sheet", meta_sheet["cdnr"]), ("cdnr_status", meta_status["cdnr"]),
            ("status", "OK_WITH_ERRORS" if n_err else ("OK_WITH_WARNINGS" if n_warn else "OK"))]
    return dict(meta=meta, tables=table_rows, mapping=mapping, issues=ctx.issues)


# ----------------------------------------------------------------------------------------
# write / read the canonical workbook
# ----------------------------------------------------------------------------------------
_WRITTEN = {}


def write_canonical(data, out_dir):
    name = cc.canonical_filename("TABLE8A", data["meta"])
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
    return cc.is_canonical(path, "T8A-", ("META", "T8A_B2B", "T8A_CDNR", "MAPPING_REPORT"))


def read_canonical(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return dict(meta=cc.read_meta(wb),
                    tables={t["canon_sheet"]: cc.read_table(wb, t["canon_sheet"], _canon_cols(k)) for k, t in _tables().items()},
                    mapping=cc.read_table(wb, "MAPPING_REPORT", MAP_COLS), issues=cc.read_table(wb, "ISSUES", ISSUE_COLS))
    finally:
        wb.close()


_SRC = cc.CanonicalCache("Table 8A", is_canonical_file, read_canonical, build_canonical, write_canonical,
                         lambda d: dict(meta=dict(d["meta"])))


def get_data(path, out_dir=None):
    return _SRC.get(path, out_dir)


def clear_cache():
    _SRC.clear()
    _WRITTEN.clear()


# ----------------------------------------------------------------------------------------
# engine-facing reader
# ----------------------------------------------------------------------------------------
def parse_table_8a(path):
    """Legacy gst_parsers_dept.parse_table_8a(path)."""
    out = dict(available=False, reason=None, b2b=[], cdnr=[], totals={})
    if not path or not os.path.exists(path):
        out["reason"] = "Table 8A not supplied for this taxpayer/FY."
        return out
    try:
        data = get_data(path)
    except mpu.PeriodParseError as ex:
        out["reason"] = f"Could not read Table 8A workbook: {ex}"
        return out
    meta = data["_index"]["meta"]
    tabs = _tables()

    def genuine(rows):
        return [r for r in rows if r["gstin"] and _GSTIN_RE.match(str(r["gstin"]).strip())]

    if meta["b2b_status"] == "no_header":
        out["reason"] = (out["reason"] or "") + " 'B2B' sheet found but no 'GSTIN of supplier' header row located."
    elif meta["b2b_status"] == "no_gstin_col":
        out["reason"] = (out["reason"] or "") + " 'B2B' sheet: could not locate the 'GSTIN of supplier' column."
    elif meta["b2b_status"] == "ok":
        for r in genuine(data["tables"][tabs["b2b"]["canon_sheet"]]):
            out["b2b"].append(dict(
                period=_s(r["period"]), gstin=str(r["gstin"]).strip(), supplier=_s(r["supplier"]), invno=_s(r["invno"]),
                invtype=_s(r["invtype"]), invdate=r["invdate"], invval=_f(r["invval"]), pos=_s(r["pos"]), rcm=_s(r["rcm"]),
                rate=_f(r["rate"]), taxable=_f(r["taxable"]), igst=_f(r["igst"]), cgst=_f(r["cgst"]), sgst=_f(r["sgst"]),
                cess=_f(r["cess"]), supplier_filing_date=r["supplier_filing_date"], itc_available=_s(r["itc_available"]),
                reason_not_available=_s(r["reason_not_available"])))
    if meta["cdnr_status"] == "no_gstin_col":
        out["reason"] = (out["reason"] or "") + f" {meta['cdnr_sheet']!r} sheet: could not locate the 'GSTIN of supplier' column."
    elif meta["cdnr_status"] == "ok":
        for r in genuine(data["tables"][tabs["cdnr"]["canon_sheet"]]):
            out["cdnr"].append(dict(
                period=None, gstin=str(r["gstin"]).strip(), supplier=_s(r["supplier"]), note_type=_s(r["note_type"]),
                supply_type=_s(r["supply_type"]), note_no=_s(r["note_no"]), note_date=r["note_date"], note_val=_f(r["note_val"]),
                pos=_s(r["pos"]), rate=_f(r["rate"]), taxable=_f(r["taxable"]), igst=_f(r["igst"]), cgst=_f(r["cgst"]),
                sgst=_f(r["sgst"]), cess=_f(r["cess"]), itc_available=_s(r["itc_available"])))

    yes_b2b = [r for r in out["b2b"] if r["itc_available"].upper() == "YES"]
    out["totals"] = dict(b2b_rows=len(out["b2b"]), b2b_yes_rows=len(yes_b2b), cgst=sum(r["cgst"] for r in yes_b2b),
                         sgst=sum(r["sgst"] for r in yes_b2b), igst=sum(r["igst"] for r in yes_b2b),
                         cess=sum(r["cess"] for r in yes_b2b))
    out["totals"]["total"] = out["totals"]["cgst"] + out["totals"]["sgst"] + out["totals"]["igst"] + out["totals"]["cess"]
    no_b2b = [r for r in out["b2b"] if r["itc_available"].upper() != "YES"]
    out["totals"]["no_reason_breakdown"] = dict(Counter(r["reason_not_available"] or "(blank)" for r in no_b2b))
    out["available"] = True
    return out
