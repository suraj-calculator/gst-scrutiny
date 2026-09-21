#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GSTR-2B CANONICAL LAYER  (spec: docs/GSTR2B_CANONICAL_SPEC.md)
================================================================
Only this module knows what a raw GSTR-2B workbook looks like.

    raw / merged GSTR-2B  --build_canonical()-->  canonical 2B data
    canonical 2B data     --write_canonical()-->  Canonical_GSTR2B_<GSTIN>_<FY>.xlsx
    canonical .xlsx       --read_canonical()  -->  canonical 2B data
    canonical 2B data     --parse_month()     -->  exactly what gst_parsers_returns.parse_2b_excel
                                                    returns today (so no check changes)

Rules (spec section 1): every row carries its period; blank means "not available", never
invented; every value is traceable to its source sheet/row; nothing is guessed silently
(MAPPING_REPORT, ISSUES); column spellings live in gstr2b_mapping.json, not here.

ERROR-level problems are raised as gst_core.PeriodParseError with a full, plain message, which
the engine already turns into "GSTR-2B not supplied (<reason>)" for the affected checks.
"""

import datetime as _dt
import json
import os
import re

import openpyxl
from openpyxl.utils import get_column_letter

import canonical_common as cc
import gst_core as mpu

SCHEMA_VERSION = "2B-1.0"
ADAPTER_VERSION = "1.0.0"
_HERE = os.path.dirname(os.path.abspath(__file__))
_MAPPING_FILE = os.path.join(_HERE, "gstr2b_mapping.json")

B2B_COLS = ["period", "months_covered", "block_no", "gstin", "supplier", "invno", "invtype", "date", "invval",
            "pos", "rcm", "rate", "taxable", "igst", "cgst", "sgst", "cess", "itc_avail", "itc_avail_reason",
            "einv_source", "irn", "filed_period", "src_sheet", "src_row", "row_shift", "arith_check", "extra"]
CDNR_COLS = ["period", "months_covered", "block_no", "gstin", "supplier", "note", "ntype", "supplytype", "date",
             "noteval", "pos", "rate", "taxable", "igst", "cgst", "sgst", "cess", "itc_avail", "itc_avail_reason",
             "einv_source", "irn", "filed_period", "src_sheet", "src_row", "row_shift", "arith_check", "extra"]
B2BA_COLS = ["period", "orig_invno", "orig_date", "gstin", "supplier", "invno", "invtype", "date", "invval", "pos",
             "rcm", "rate", "taxable", "igst", "cgst", "sgst", "cess", "itc_avail", "itc_avail_reason",
             "src_sheet", "src_row", "extra"]
CDNRA_COLS = ["period", "orig_note", "gstin", "supplier", "note", "ntype", "supplytype", "date", "noteval", "pos",
              "rate", "taxable", "igst", "cgst", "sgst", "cess", "src_sheet", "src_row", "extra"]
ITC_COLS = ["period", "months_covered", "all_other_igst", "all_other_cgst", "all_other_sgst", "all_other_cess",
            "rcm_igst", "rcm_cgst", "rcm_sgst", "rcm_cess", "cn_igst", "cn_cgst", "cn_sgst", "cn_cess",
            "qtr_total_mismatch", "summary_source"]
MAP_COLS = ["canonical_sheet", "canonical_field", "raw_sheet", "raw_header_found", "raw_column", "matched_by",
            "confidence", "status", "rows_filled", "rows_blank", "note"]
ISSUE_COLS = cc.ISSUE_COLS

_AMOUNT_FIELDS = {"invval", "noteval", "taxable", "igst", "cgst", "sgst", "cess"}
_GSTIN_RE = re.compile(r"^[0-9A-Z]{15}$")
_PERIOD_TAG_RE = re.compile(r"^[A-Za-z]{3}'\d{2}$")
_FILENAME_RE = re.compile(r"(?:^|_)(\d{2})(\d{4})_([0-9A-Za-z]{15})_GSTR2B")
_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

_CONTEXT = {"gstin": None, "fy": None}


# ----------------------------------------------------------------------------------------
# small helpers
# ----------------------------------------------------------------------------------------
def set_context(gstin=None, fy=None):
    """Optional: the engine tells us the taxpayer's GSTIN / FY (a raw 2B often carries neither)."""
    if gstin:
        _CONTEXT["gstin"] = gstin
    if fy:
        _CONTEXT["fy"] = fy


def _mapping():
    return cc.load_json(_MAPPING_FILE)


_clean, _text, _num, _blank = cc.clean, cc.text, cc.num, cc.is_blank


def _parse_date(v):
    """dd/mm/yyyy text (or a native date) -> date; anything else -> None (caller keeps the text)."""
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    m = re.match(r"^\s*(\d{2})/(\d{2})/(\d{4})\s*$", str(v or ""))
    if not m:
        return None
    try:
        return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _norm_period_tag(tag):
    """"May'22" -> 'May-22'. Raises PeriodParseError when unreadable."""
    s = _text(tag)
    m = re.match(r"^([A-Za-z]{3})'(\d{2})$", s)
    if not m:
        raise mpu.PeriodParseError(f"Unrecognised per-row GSTR-2B period tag: {tag!r}")
    return f"{m.group(1).title()}-{m.group(2)}"


def _period_tag_from_date(v):
    d = _parse_date(v)
    return f"{_MONTH_ABBR[d.month - 1]}'{str(d.year)[2:]}" if d else None


def _Ctx(source_files):
    """Issues + mapping report + fill counts for one build (shared implementation: canonical_common)."""
    return cc.IssueLog("GSTR-2B", source_files)


def _fail(ctx, iid, sheet, field, message, skipped="", action=""):
    ctx.fail(iid, message, sheet=sheet, field=field, skipped=skipped, action=action)


# ----------------------------------------------------------------------------------------
# header detection (alias table -> contains); content inference is deliberately NOT in v1
# ----------------------------------------------------------------------------------------
def _first_data_index(rows):
    """Index of the first period-marker row or first row whose column A is a GSTIN."""
    for i, r in enumerate(rows):
        if not r:
            continue
        if mpu.is_marker_row(r):
            return i
        a = r[0]
        if isinstance(a, str) and _GSTIN_RE.match(a.strip()):
            return i
    return None


def _header_map(rows, first_data, k_above):
    """{clean header text: [col, ...]} in (row, col) discovery order, plus {col: original text}."""
    hmap, texts = {}, {}
    picked, r = [], first_data - 1
    # the k_above nearest NON-EMPTY rows above the data (a stray blank row must not hide the header)
    while r >= 0 and len(picked) < k_above and first_data - r <= 12:
        if r < len(rows) and rows[r] and any(c not in (None, "") for c in rows[r]):
            picked.append(r)
        r -= 1
    for ridx in sorted(picked):
        for i, c in enumerate(rows[ridx]):
            if c:
                hmap.setdefault(_clean(c), []).append(i)
                t = _text(c)
                if t and t not in texts.setdefault(i, []):
                    texts[i].append(t)
    return hmap, {i: " / ".join(t) for i, t in texts.items()}


def _resolve(hmap, spec, order):
    """-> (col or None, matched_by, confidence)"""
    nth = spec.get("nth", 0)
    for pos, name in enumerate(spec.get("exact", [])):
        idxs = hmap.get(_clean(name))
        if idxs and nth < len(idxs):
            return idxs[nth], ("exact" if pos == 0 else "alias"), "HIGH"
    for needle in spec.get("contains", []):
        nc = _clean(needle)
        if order == "first_seen" and nth == 0:
            for hdr, idxs in hmap.items():
                if nc in hdr:
                    return idxs[0], "contains", "MEDIUM"
        else:
            matches = sorted(i for hdr, idxs in hmap.items() if nc in hdr for i in idxs)
            if nth < len(matches):
                return matches[nth], "contains", "MEDIUM"
    return None, "", ""


def _resolve_sheet(ctx, rows, sheet_key, raw_sheet, first_data):
    """Map every canonical field of one sheet type to a raw column; record MAPPING_REPORT rows."""
    spec_sheet = _mapping()["sheets"][sheet_key]
    hmap, texts = _header_map(rows, first_data, spec_sheet["header_rows_above_data"])
    cols, missing_required = {}, []
    for field, spec in spec_sheet["fields"].items():
        col, how, conf = _resolve(hmap, spec, spec_sheet.get("contains_order", "first_seen"))
        cols[field] = col
        if col is None:
            status = "MISSING"
            if spec.get("required"):
                missing_required.append(field)
        else:
            status = "OK"
        ctx.mapping.append(dict(canonical_sheet=sheet_key, canonical_field=field, raw_sheet=raw_sheet,
                                raw_header_found=texts.get(col, "") if col is not None else "",
                                raw_column=get_column_letter(col + 1) if col is not None else "",
                                matched_by=how, confidence=conf, status=status, rows_filled=0, rows_blank=0,
                                note=""))
    return cols, texts, hmap, missing_required, spec_sheet


def _find_sheet(wb, raw_names):
    lower = {n.lower(): n for n in wb.sheetnames}
    for n in raw_names:
        if n.lower() in lower:
            return lower[n.lower()]
    return None


# ----------------------------------------------------------------------------------------
# period resolution for files with no in-sheet markers
# ----------------------------------------------------------------------------------------
def _period_from_readme_or_filename(wb, path):
    """(fy, tax_period_label) from 'Read me' or the portal filename; (None, None) if neither."""
    if "Read me" in wb.sheetnames:
        fy = tp = None
        for i, row in enumerate(wb["Read me"].iter_rows(values_only=True)):
            if i > 25:
                break
            cells = [str(c).strip() for c in row if c not in (None, "")]
            if len(cells) >= 2:
                key = cells[0].lower()
                if key.startswith("financial year") or key == "year":
                    fy = fy or cells[1]
                elif key == "tax period":
                    tp = tp or cells[1]
        if fy and tp:
            return fy, tp
    m = _FILENAME_RE.search(os.path.basename(path))
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            fy_start = year if month >= 4 else year - 1
            return f"{fy_start}-{str((fy_start + 1) % 100).zfill(2)}", _MONTH_ABBR[month - 1]
    return None, None


# ----------------------------------------------------------------------------------------
# blocks
# ----------------------------------------------------------------------------------------
def _blocks(rows, first_data, implicit_labels):
    """[(labels, start, end)] of period blocks. A sheet without markers is one implicit block."""
    markers = []
    for i, r in enumerate(rows):
        if r and mpu.is_marker_row(r):
            _, _, labels = mpu.parse_marker_text(str(r[0]))
            markers.append((i, labels))
    if markers:
        out = []
        for n, (i, labels) in enumerate(markers):
            end = markers[n + 1][0] if n + 1 < len(markers) else len(rows)
            out.append((labels, i + 1, end))
        return out
    return [(implicit_labels, first_data, len(rows))]


def _label(labels):
    return labels[0] if len(labels) == 1 else f"{labels[0]}..{labels[-1]}"


# ----------------------------------------------------------------------------------------
# B2B / CDNR lines
# ----------------------------------------------------------------------------------------
def _build_lines(ctx, wb, sheet_key, path, implicit_labels):
    raw_sheet = _find_sheet(wb, _mapping()["sheets"][sheet_key]["raw_names"])
    if raw_sheet is None:
        return None, []
    rows = list(wb[raw_sheet].iter_rows(values_only=True))
    first = _first_data_index(rows)
    if first is None:
        return raw_sheet, []
    cols, texts, hmap, missing, spec_sheet = _resolve_sheet(ctx, rows, sheet_key, raw_sheet, first)
    if missing:
        fields_spec = _mapping()["sheets"][sheet_key]["fields"]
        names = [(fields_spec[m].get("exact") or [fields_spec[m].get("label") or m])[0] for m in missing]
        _fail(ctx, "E201", sheet_key, ", ".join(missing),
              f"GSTR-2B \"{os.path.basename(path)}\", sheet '{raw_sheet}': required field(s) {missing} not "
              f"found. Looked for: {names}. Headers found: {sorted(set(texts.values()))}. Periods affected: all. "
              f"GSTR-2B is treated as NOT SUPPLIED; checks skipped: Comparison C/D, ITC 3B vs 2B, "
              f"EWB-in vs 2B, GSTR-2A vs 2B.",
              "every check that reads GSTR-2B",
              "Give the file as downloaded from the portal, or add the new spelling to gstr2b_mapping.json.")
    c = cols
    if c["period"] is None and c["rate"] is None:
        _fail(ctx, "E201", sheet_key, "period / rate",
              f"GSTR-2B \"{os.path.basename(path)}\", sheet '{raw_sheet}': neither the supplier filing-period "
              f"column nor 'Rate(%)' was found, so misaligned rows cannot be detected. Headers found: "
              f"{sorted(set(texts.values()))}.", "every check that reads GSTR-2B",
              "Give the file as downloaded from the portal, or add the new spelling to gstr2b_mapping.json.")
    if c["rate"] is None:
        ctx.issue("W307", "WARNING", sheet_key, "rate", "all",
                  f"Field 'Rate(%)' not found on sheet '{raw_sheet}'. Rate is left blank (the engine reads it as "
                  f"0.0); the supplier filing-period column is used to detect misaligned rows instead.",
                  "nothing (rate is display-only)", "None needed unless a check starts using the GST rate.")
    is_cdnr = sheet_key == "CDNR"
    date_col = c["date"]
    mapped_cols = {v for v in c.values() if v is not None}
    unmapped = {i: t for i, t in texts.items() if i not in mapped_cols}
    rate_anchor = c["rate"] if c["rate"] is not None else c["period"]

    out, seen_by_period, block_no_by_period = [], {}, {}
    n_shift = n_mismatch = n_dup = n_unparse = 0
    for labels, start, end in _blocks(rows, first, implicit_labels):
        if not labels:
            _fail(ctx, "E202", sheet_key, "period",
                  f"GSTR-2B \"{os.path.basename(path)}\": cannot determine the tax period. No period banner in "
                  f"sheet '{raw_sheet}', no usable 'Read me', and the filename does not match "
                  f"'<ts>_<MMYYYY>_<GSTIN>_GSTR2B_<DDMMYYYY>_<n>.xlsx'. Rename the file to the portal name.",
                  "every check that reads GSTR-2B", "Rename the file to the portal name or use the merged file.")
        period = _label(labels)
        block_no_by_period[period] = block_no_by_period.get(period, 0) + 1
        block_no = block_no_by_period[period]
        seen_here, seen_before = set(), seen_by_period.setdefault(period, set())
        for ridx in range(start, end):
            r = rows[ridx]
            if not any(r) or not r[0] or mpu.is_marker_row(r):
                continue
            key = tuple(r)
            if key in seen_before:
                n_dup += 1
                continue
            seen_here.add(key)
            shift = _row_shift(r, c["period"], c["rate"])
            eff = lambda col: (col + shift) if col is not None else None   # noqa: E731

            def g(field, shifted=False):
                col = c.get(field)
                col = eff(col) if shifted else col
                v = r[col] if (col is not None and col < len(r)) else None
                ctx.count(sheet_key, field, not _blank(v))
                return _text(v)

            def gn(field, shifted=True):
                col = c.get(field)
                col = eff(col) if shifted else col
                v = r[col] if (col is not None and col < len(r)) else None
                ctx.count(sheet_key, field, not _blank(v))
                return _num(v)

            filed_period = None
            if c["period"] is not None:
                v = r[c["period"] + shift] if (c["period"] + shift) < len(r) else None
                ctx.count(sheet_key, "period", not _blank(v))
                if is_cdnr and not (isinstance(v, str) and _PERIOD_TAG_RE.match(v.strip())):
                    v = _period_tag_from_date(r[date_col] if date_col is not None and date_col < len(r) else None)
                try:
                    filed_period = _norm_period_tag(v) if v is not None else None
                except mpu.PeriodParseError:
                    filed_period = None
            if filed_period is None and is_cdnr:
                n_unparse += 1
            row = dict(period=period, months_covered=";".join(labels), block_no=block_no,
                       gstin=g("gstin"), supplier=g("supplier"))
            if is_cdnr:
                row.update(note=g("note"), ntype=g("ntype"), supplytype=g("supplytype"))
            else:
                row.update(invno=g("invno"), invtype=g("invtype"))
            dcell = r[date_col] if (date_col is not None and date_col < len(r)) else None
            ctx.count(sheet_key, "date", not _blank(dcell))
            row["date"] = _parse_date(dcell) or _text(dcell)
            if is_cdnr:
                row["noteval"] = gn("noteval", shifted=False)
            else:
                row["invval"] = gn("invval", shifted=False)
            row["pos"] = g("pos")
            if not is_cdnr:
                row["rcm"] = g("rcm")
            if shift:
                row["rate"] = None
            elif c["rate"] is not None:
                row["rate"] = gn("rate", shifted=False)
            else:
                row["rate"] = None
            for f in ("taxable", "igst", "cgst", "sgst", "cess"):
                row[f] = gn(f)
            row["itc_avail"] = g("itc_avail", shifted=True) if c["itc_avail"] is not None else ""
            row["itc_avail_reason"] = g("reason", shifted=True) if c["reason"] is not None else ""
            row["einv_source"] = g("source", shifted=True) if c["source"] is not None else ""
            row["irn"] = g("irn", shifted=True) if c["irn"] is not None else ""
            row["filed_period"] = filed_period
            row["src_sheet"] = raw_sheet
            row["src_row"] = ridx + 1
            row["row_shift"] = shift
            total = row["taxable"] + row["igst"] + row["cgst"] + row["sgst"] + row["cess"]
            val = row["noteval"] if is_cdnr else row["invval"]
            if val == 0.0:
                row["arith_check"] = "NA"
            elif is_cdnr:
                row["arith_check"] = "OK" if abs(val - total) <= 1.0 else "MISMATCH"
            else:
                row["arith_check"] = "OK" if val + 1.0 >= total else "MISMATCH"
            if row["arith_check"] == "MISMATCH":
                n_mismatch += 1
            if shift:
                n_shift += 1
            extras = []
            for h, t in unmapped.items():
                hh = h + shift if (rate_anchor is not None and h >= rate_anchor) else h
                v = r[hh] if 0 <= hh < len(r) else None
                if not _blank(v):
                    extras.append(f"{t}={_text(v)}")
            row["extra"] = "; ".join(extras)
            out.append(row)
        seen_before |= seen_here
    if n_shift:
        ctx.issue("W303", "WARNING", sheet_key, "rate", "several",
                  f"{n_shift} of {len(out)} {sheet_key} rows had a blank Rate and were column-shifted; "
                  f"corrected (row_shift=-1). Rate is blank for those rows.", "", "")
    if n_dup:
        ctx.issue("W304", "WARNING", sheet_key, "", "several",
                  f"{n_dup} {sheet_key} rows were byte-identical to a row of an earlier block of the same "
                  f"period (same download merged twice) and were skipped.", "", "")
    extra_blocks = {p: n for p, n in block_no_by_period.items() if n > 1}
    if extra_blocks:
        ctx.issue("W304", "WARNING", sheet_key, "", ", ".join(extra_blocks),
                  f"{sheet_key}: period(s) came in several downloads {extra_blocks}; all kept "
                  f"(block_no column).", "", "")
    if n_mismatch:
        ctx.issue("W305", "WARNING", sheet_key, "arith_check", "several",
                  f"{n_mismatch} {sheet_key} rows are internally inconsistent as reported by the supplier "
                  f"(document value does not cover taxable value + taxes; see column arith_check = MISMATCH). "
                  f"Values are kept exactly as reported; the columns are correctly aligned.",
                  "nothing (informational)", "Review those documents with the supplier if the amounts matter.")
    if n_unparse:
        ctx.issue("W306", "WARNING", sheet_key, "filed_period", "several",
                  f"{n_unparse} CDNR rows had an unreadable filing-period tag and note date; filed_period "
                  f"left blank.", "", "")
    return raw_sheet, out


def _row_shift(r, c_period, c_rate):
    """0 normally, -1 when the row's Rate cell is blank/implausible (every later column sits one
    position earlier). Same rule as gst_parsers_returns._2b_row_rate_shift."""
    import gst_parsers_returns as pr
    if c_rate is not None and c_rate < len(r):
        return 0 if pr._looks_like_gst_rate(r[c_rate]) else -1
    if c_period is None:
        return 0
    v = r[c_period] if c_period < len(r) else None
    if isinstance(v, str) and _PERIOD_TAG_RE.match(v.strip()):
        return 0
    if c_period > 0:
        vp = r[c_period - 1] if (c_period - 1) < len(r) else None
        if isinstance(vp, str) and _PERIOD_TAG_RE.match(vp.strip()):
            return -1
    return 0


# ----------------------------------------------------------------------------------------
# B2BA / CDNRA amendment lines (degrade, never abort the 2B)
# ----------------------------------------------------------------------------------------
def _build_amendments(ctx, wb, sheet_key, path):
    raw_sheet = _find_sheet(wb, _mapping()["sheets"][sheet_key]["raw_names"])
    if raw_sheet is None:
        ctx.issue("I402", "INFO", sheet_key, "", "all",
                  f"Sheet {sheet_key} is not in the file; nothing to net for that amendment type.")
        return None, [], False
    rows = list(wb[raw_sheet].iter_rows(values_only=True))
    first = _first_data_index(rows)
    if first is None:
        return raw_sheet, [], True
    cols, texts, hmap, missing, spec_sheet = _resolve_sheet(ctx, rows, sheet_key, raw_sheet, first)
    if missing:
        ctx.issue("W301", "WARNING", sheet_key, ", ".join(missing), "all",
                  f"'{raw_sheet}' has no readable header (looked for the original-invoice / GSTIN / filing-period "
                  f"columns; found headers: {sorted(set(texts.values())) or 'none'}). Amendments are NOT applied: "
                  f"amended invoices/notes are taken as-is from the base sheet.",
                  f"{sheet_key} amendment netting",
                  "Re-merge the monthly files (the merge tool now keeps the header) or supply a file whose "
                  f"'{raw_sheet}' sheet has its header rows.")
        return raw_sheet, [], False
    c, out = cols, []
    is_note = sheet_key == "CDNRA"
    orig_key = "orig_note" if is_note else "orig_invno"
    mapped = {v for v in c.values() if v is not None}
    unmapped = {i: t for i, t in texts.items() if i not in mapped}
    try:
        for ridx in range(first, len(rows)):
            r = rows[ridx]
            if not any(r) or not (c["gstin"] < len(r) and r[c["gstin"]]) or mpu.is_marker_row(r):
                continue

            def g(field):
                col = c.get(field)
                v = r[col] if (col is not None and col < len(r)) else None
                ctx.count(sheet_key, field, not _blank(v))
                return _text(v)

            def gn(field):
                col = c.get(field)
                v = r[col] if (col is not None and col < len(r)) else None
                ctx.count(sheet_key, field, not _blank(v))
                return _num(v)

            period = _norm_period_tag(r[c["period"]] if c["period"] < len(r) else None)
            orig = g(orig_key)
            odate = r[c["orig_date"]] if (c.get("orig_date") is not None and c["orig_date"] < len(r)) else None
            row = dict(period=period)
            row[orig_key] = orig
            if not is_note:
                row["orig_date"] = _parse_date(odate) or _text(odate)
            row["gstin"], row["supplier"] = g("gstin"), g("supplier")
            rev = g("note" if is_note else "invno")
            row["note" if is_note else "invno"] = rev or orig.upper()
            if is_note:
                nt = g("ntype") if c.get("ntype") is not None else g("ntype_orig")
                row["ntype"], row["supplytype"] = nt, g("supplytype")
            else:
                row["invtype"] = g("invtype")
            d = r[c["date"]] if (c.get("date") is not None and c["date"] < len(r)) else None
            row["date"] = _parse_date(d) or _text(d)
            if is_note:
                row["noteval"] = gn("noteval")
            else:
                row["invval"] = gn("invval")
            row["pos"] = g("pos")
            if not is_note:
                row["rcm"] = g("rcm")
            row["rate"] = gn("rate")
            for f in ("taxable", "igst", "cgst", "sgst", "cess"):
                row[f] = gn(f)
            if not is_note:
                row["itc_avail"], row["itc_avail_reason"] = g("itc_avail"), g("reason")
            row["src_sheet"], row["src_row"] = raw_sheet, ridx + 1
            extras = [f"{t}={_text(r[h])}" for h, t in unmapped.items() if h < len(r) and not _blank(r[h])]
            row["extra"] = "; ".join(extras)
            out.append(row)
    except mpu.PeriodParseError as e:
        ctx.issue("W301", "WARNING", sheet_key, "period", "all",
                  f"'{raw_sheet}': {e}. Amendments are NOT applied; amended rows are taken as-is from the base "
                  f"sheet.", f"{sheet_key} amendment netting", "")
        return raw_sheet, [], False
    return raw_sheet, out, True


# ----------------------------------------------------------------------------------------
# ITC Available control totals
# ----------------------------------------------------------------------------------------
def _build_itc(ctx, wb, implicit=None):
    raw_sheet = _find_sheet(wb, _mapping()["sheets"]["ITC"]["raw_names"])
    if raw_sheet is None:
        ctx.issue("W302", "WARNING", "ITC_SUMMARY", "", "all",
                  "'ITC Available' sheet not in the file. Control totals are recomputed from B2B/B2B-CDNR rows; "
                  "checks F7a and R14's 2B total are skipped.", "F7a, R14 2B total", "Download the full GSTR-2B.")
        return None, []
    import gst_parsers_returns as pr
    rows = list(wb[raw_sheet].iter_rows(values_only=True))
    months = []
    for r in rows:
        if r and mpu.is_marker_row(r):
            _, _, labels = mpu.parse_marker_text(str(r[0]))
            for m in labels:
                if m not in months:
                    months.append(m)
    out = []
    if not months and implicit:
        # A raw per-period download has no marker rows: the whole sheet belongs to this file's own
        # period(s). A quarterly file lays its month groups side by side, exactly as in a merged one.
        for gi, m in enumerate(implicit):
            s_ = pr._summary_from_block(rows, group_index=gi, group_count=len(implicit))
            out.append(dict(period=m, months_covered=m,
                            all_other_igst=s_["ITC_all_other_IGST"], all_other_cgst=s_["ITC_all_other_CGST"],
                            all_other_sgst=s_["ITC_all_other_SGST"], all_other_cess=s_["ITC_all_other_CESS"],
                            rcm_igst=s_["ITC_rcm_IGST"], rcm_cgst=s_["ITC_rcm_CGST"], rcm_sgst=s_["ITC_rcm_SGST"],
                            rcm_cess=s_["ITC_rcm_CESS"], cn_igst=s_["CN_IGST"], cn_cgst=s_["CN_CGST"],
                            cn_sgst=s_["CN_SGST"], cn_cess=s_["CN_CESS"],
                            qtr_total_mismatch=s_["_qtr_total_mismatch"] or "", summary_source="ITC Available sheet"))
        return raw_sheet, out
    if not months:
        ctx.issue("W302", "WARNING", "ITC_SUMMARY", "", "all",
                  "'ITC Available' sheet has no period banners and the file's period is unknown; control totals "
                  "not read.", "F7a, R14 2B total", "Give the file as downloaded from the portal.")
        return raw_sheet, out
    for m in months:
        start, end, gi, gc = mpu.find_block_and_index_for_month(rows, m)
        s = pr._summary_from_block(rows[start:end], group_index=gi, group_count=gc)
        out.append(dict(period=m, months_covered=m,
                        all_other_igst=s["ITC_all_other_IGST"], all_other_cgst=s["ITC_all_other_CGST"],
                        all_other_sgst=s["ITC_all_other_SGST"], all_other_cess=s["ITC_all_other_CESS"],
                        rcm_igst=s["ITC_rcm_IGST"], rcm_cgst=s["ITC_rcm_CGST"], rcm_sgst=s["ITC_rcm_SGST"],
                        rcm_cess=s["ITC_rcm_CESS"], cn_igst=s["CN_IGST"], cn_cgst=s["CN_CGST"],
                        cn_sgst=s["CN_SGST"], cn_cess=s["CN_CESS"],
                        qtr_total_mismatch=s["_qtr_total_mismatch"] or "", summary_source="ITC Available sheet"))
    return raw_sheet, out


# ----------------------------------------------------------------------------------------
# build
# ----------------------------------------------------------------------------------------
def build_canonical(raw_paths):
    """Convert one or more raw/merged GSTR-2B workbooks into canonical data (a dict of tables).
    Raises PeriodParseError (message starts with the issue id) on ERROR-level problems."""
    if isinstance(raw_paths, str):
        raw_paths = [raw_paths]
    ctx = _Ctx([os.path.basename(p) for p in raw_paths])
    b2b, cdnr, b2ba, cdnra, itc = [], [], [], [], []
    fy_seen, gstin_seen, unread = set(), None, []
    layout = "merged-with-banners"
    b2ba_ok = cdnra_ok = False
    for path in raw_paths:
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception as e:  # noqa: BLE001 - any open failure is the same user-facing problem
            _fail(ctx, "E204", "", "", f"GSTR-2B \"{os.path.basename(path)}\": cannot open the file "
                  f"({type(e).__name__}: {e}). It may be corrupt, password-protected or not an Excel file.",
                  "every check that reads GSTR-2B", "Re-download the file.")
        try:
            names = wb.sheetnames
            if not (_find_sheet(wb, _mapping()["sheets"]["B2B"]["raw_names"]) or
                    _find_sheet(wb, _mapping()["sheets"]["CDNR"]["raw_names"])):
                _fail(ctx, "E203", "", "", f"GSTR-2B \"{os.path.basename(path)}\": no GSTR-2B sheet recognised "
                      f"(expected 'B2B' or 'B2B-CDNR'; found: {names}).", "every check that reads GSTR-2B",
                      "Check that this is a GSTR-2B workbook.")
            implicit, gstin_in_file = [], None
            fy, tp = _period_from_readme_or_filename(wb, path)
            if "Read me" in names:
                for row in wb["Read me"].iter_rows(values_only=True, max_row=25):
                    cells = [str(c).strip() for c in row if c not in (None, "")]
                    if len(cells) >= 2 and cells[0].lower() == "gstin" and _GSTIN_RE.match(cells[1].upper()):
                        gstin_in_file = cells[1].upper()
                        break
                layout = "full-export-with-readme"
            gstin_seen = gstin_seen or gstin_in_file
            if fy and tp:
                try:
                    implicit = mpu.months_for_tax_period(fy, tp)
                except mpu.PeriodParseError:
                    implicit = []
            # sheets
            sh_b2b, lines_b2b = _build_lines(ctx, wb, "B2B", path, implicit)
            sh_cdnr, lines_cdnr = _build_lines(ctx, wb, "CDNR", path, implicit)
            b2b += lines_b2b
            cdnr += lines_cdnr
            sh_a, lines_a, ok_a = _build_amendments(ctx, wb, "B2BA", path)
            sh_ca, lines_ca, ok_ca = _build_amendments(ctx, wb, "CDNRA", path)
            b2ba += lines_a
            cdnra += lines_ca
            b2ba_ok, cdnra_ok = b2ba_ok or ok_a, cdnra_ok or ok_ca
            sh_itc, lines_itc = _build_itc(ctx, wb, implicit)
            itc += lines_itc
            used = {s for s in (sh_b2b, sh_cdnr, sh_a, sh_ca, sh_itc) if s}
            unread += [n for n in names if n not in used and n != "Read me"]
        finally:
            wb.close()

    periods, months = [], []
    for r in b2b + cdnr:
        if r["period"] not in periods:
            periods.append(r["period"])
        for m in r["months_covered"].split(";"):
            if m not in months:
                months.append(m)
    if not months:
        _fail(ctx, "E202", "", "", f"GSTR-2B \"{', '.join(ctx.source_files)}\": no period could be determined "
              f"(no period banners, no 'Read me', filename not in portal format).",
              "every check that reads GSTR-2B", "Rename the file to the portal name or use the merged file.")
    for m in months:
        yy = int(m[-2:])
        mon = m[:3]
        start = 2000 + yy if mon not in ("Jan", "Feb", "Mar") else 2000 + yy - 1
        fy_seen.add(f"{start}-{str((start + 1) % 100).zfill(2)}")
    fy = ", ".join(sorted(fy_seen))
    extra_blocks = {}
    for r in b2b + cdnr:
        extra_blocks[r["period"]] = max(extra_blocks.get(r["period"], 1), r["block_no"])
    extra_blocks = {p: n for p, n in extra_blocks.items() if n > 1}
    if unread:
        ctx.issue("I401", "INFO", "", "", "all", f"Sheets not used: {sorted(set(unread))}.")
    n_warn = sum(1 for i in ctx.issues if i["severity"] == "WARNING")
    # finish mapping report counts
    for m in ctx.mapping:
        f = ctx.fill.get((m["canonical_sheet"], m["canonical_field"]))
        if f:
            m["rows_filled"], m["rows_blank"] = f
    gstin = _CONTEXT["gstin"] or gstin_seen or ""
    if not gstin:
        ctx.issue("W309", "WARNING", "META", "recipient_gstin", "all",
                  "Recipient GSTIN not in the file (no 'Read me') and not supplied by the caller; META left blank.")
    meta = [("schema_version", SCHEMA_VERSION), ("recipient_gstin", gstin), ("fy", fy),
            ("source_files", "; ".join(ctx.source_files)), ("source_layout", layout),
            ("adapter_version", ADAPTER_VERSION), ("created_on", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("periods_present", "; ".join(periods)),
            ("months_present", ";".join(months)),
            ("periods_with_extra_blocks", "; ".join(f"{p}: {n} blocks" for p, n in extra_blocks.items())),
            ("summary_available", "Y" if itc else "N"),
            ("b2ba_available", "Y" if b2ba_ok else "N"), ("cdnra_available", "Y" if cdnra_ok else "N"),
            ("unread_sheets", "; ".join(sorted(set(unread)))),
            ("status", "OK_WITH_WARNINGS" if n_warn else "OK")]
    return dict(meta=meta, b2b=b2b, cdnr=cdnr, b2ba=b2ba, cdnra=cdnra, itc=itc, mapping=ctx.mapping,
                issues=ctx.issues)


# ----------------------------------------------------------------------------------------
# write / read the canonical workbook
# ----------------------------------------------------------------------------------------
def canonical_filename(data):
    return cc.canonical_filename("GSTR2B", data["meta"])


def write_canonical(data, out_dir):
    return cc.write_workbook(out_dir, canonical_filename(data), data["meta"], [
        ("B2B", B2B_COLS, data["b2b"]), ("CDNR", CDNR_COLS, data["cdnr"]), ("B2BA", B2BA_COLS, data["b2ba"]),
        ("CDNRA", CDNRA_COLS, data["cdnra"]), ("ITC_SUMMARY", ITC_COLS, data["itc"]),
        ("MAPPING_REPORT", MAP_COLS, data["mapping"]), ("ISSUES", ISSUE_COLS, data["issues"])])


def is_canonical_file(path):
    return cc.is_canonical(path, "2B-", ("META", "B2B"))


def read_canonical(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return dict(meta=cc.read_meta(wb), b2b=cc.read_table(wb, "B2B", B2B_COLS),
                    cdnr=cc.read_table(wb, "CDNR", CDNR_COLS), b2ba=cc.read_table(wb, "B2BA", B2BA_COLS),
                    cdnra=cc.read_table(wb, "CDNRA", CDNRA_COLS), itc=cc.read_table(wb, "ITC_SUMMARY", ITC_COLS),
                    mapping=cc.read_table(wb, "MAPPING_REPORT", MAP_COLS),
                    issues=cc.read_table(wb, "ISSUES", ISSUE_COLS))
    finally:
        wb.close()


def months_from_canonical(path):
    """Set of 'Mon-YY' months a canonical file covers (used by classify_folder)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for r in wb["META"].iter_rows(min_row=2, values_only=True):
            if r and r[0] == "months_present":
                return {m for m in str(r[1] or "").split(";") if m}
    finally:
        wb.close()
    return set()


# ----------------------------------------------------------------------------------------
# engine-facing reader: identical shape to gst_parsers_returns.parse_2b_excel
# ----------------------------------------------------------------------------------------
_SRC = cc.CanonicalCache("GSTR-2B", is_canonical_file, read_canonical, build_canonical, write_canonical,
                         lambda data: _index(data))


def get_data(path, out_dir=None):
    return _SRC.get(path, out_dir)


def clear_cache():
    _SRC.clear()


def _index(data):
    meta = dict(data["meta"])
    months = {m for m in str(meta.get("months_present") or "").split(";") if m}
    sup_inv = {(r["gstin"] or "", str(r["orig_invno"] or "").strip().upper())
               for r in data["b2ba"] if r["gstin"] and r["orig_invno"]}
    sup_note = {(r["gstin"] or "", str(r["orig_note"] or "").strip().upper())
                for r in data["cdnra"] if r["gstin"] and r["orig_note"]}
    return dict(meta=meta, months=months, sup_inv=sup_inv, sup_note=sup_note,
                itc={r["period"]: r for r in data["itc"]})


def _s(v):
    return "" if v is None else str(v)


def _f(v):
    return 0.0 if v is None else float(v)


def _d(v):
    if isinstance(v, _dt.datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, _dt.date):
        return v.strftime("%d/%m/%Y")
    return _s(v)


def _b2b_dict(r):
    return dict(gstin=_s(r["gstin"]), supplier=_s(r["supplier"]), invno=_s(r["invno"]), invtype=_s(r["invtype"]),
                date=_d(r["date"]), invval=_f(r["invval"]), pos=_s(r["pos"]), rcm=_s(r["rcm"]),
                rate=(None if r["row_shift"] else _f(r["rate"])), taxable=_f(r["taxable"]), igst=_f(r["igst"]),
                cgst=_f(r["cgst"]), sgst=_f(r["sgst"]), cess=_f(r["cess"]), itc_avail=_s(r["itc_avail"]),
                itc_avail_reason=_s(r["itc_avail_reason"]), einv_source=_s(r["einv_source"]), irn=_s(r["irn"]),
                filed_period=(r["filed_period"] or None), via_amendment=False)


def _cdnr_dict(r):
    return dict(gstin=_s(r["gstin"]), supplier=_s(r["supplier"]), note=_s(r["note"]), ntype=_s(r["ntype"]),
                supplytype=_s(r["supplytype"]), date=_d(r["date"]), noteval=_f(r["noteval"]), pos=_s(r["pos"]),
                rate=(None if r["row_shift"] else _f(r["rate"])), taxable=_f(r["taxable"]), igst=_f(r["igst"]),
                cgst=_f(r["cgst"]), sgst=_f(r["sgst"]), cess=_f(r["cess"]), itc_avail=_s(r["itc_avail"]),
                itc_avail_reason=_s(r["itc_avail_reason"]), einv_source=_s(r["einv_source"]), irn=_s(r["irn"]),
                filed_period=(r["filed_period"] or None), via_amendment=False)


def _b2ba_dict(r):
    return dict(gstin=_s(r["gstin"]), supplier=_s(r["supplier"]), invno=_s(r["invno"]), invtype=_s(r["invtype"]),
                date=_d(r["date"]), invval=_f(r["invval"]), pos=_s(r["pos"]), rcm=_s(r["rcm"]),
                rate=_f(r["rate"]), taxable=_f(r["taxable"]), igst=_f(r["igst"]), cgst=_f(r["cgst"]),
                sgst=_f(r["sgst"]), cess=_f(r["cess"]), itc_avail=_s(r["itc_avail"]),
                itc_avail_reason=_s(r["itc_avail_reason"]), via_amendment=True,
                original_invno=_s(r["orig_invno"]).strip().upper())


def _cdnra_dict(r):
    return dict(gstin=_s(r["gstin"]), supplier=_s(r["supplier"]), note=_s(r["note"]), ntype=_s(r["ntype"]),
                supplytype=_s(r["supplytype"]), date=_d(r["date"]), noteval=_f(r["noteval"]), pos=_s(r["pos"]),
                rate=_f(r["rate"]), taxable=_f(r["taxable"]), igst=_f(r["igst"]), cgst=_f(r["cgst"]),
                sgst=_f(r["sgst"]), cess=_f(r["cess"]), via_amendment=True,
                original_note=_s(r["orig_note"]).strip().upper())


_NO_SUMMARY_REASON = ("'ITC Available' sheet not present in this GSTR-2B export -- summary "
                      "control-total not computable from this file; invoice-level B2B/B2B-CDNR "
                      "figures below are unaffected and are what every real check in this tool uses.")


def parse_month(path, month):
    """Drop-in replacement for gst_parsers_returns.parse_2b_excel(path, month)."""
    data = get_data(path)
    ix = data["_index"]
    if month not in ix["months"]:
        raise mpu.PeriodParseError(
            f"Month {month!r} not covered by any period marker in this GSTR-2B. "
            f"Months present: {sorted(ix['months'])}")
    if ix["meta"].get("summary_available") != "Y":
        summary = dict(
            ITC_all_other_IGST=0.0, ITC_all_other_CGST=0.0, ITC_all_other_SGST=0.0, ITC_all_other_CESS=0.0,
            ITC_rcm_IGST=0.0, ITC_rcm_CGST=0.0, ITC_rcm_SGST=0.0, ITC_rcm_CESS=0.0,
            CN_IGST=0.0, CN_CGST=0.0, CN_SGST=0.0, CN_CESS=0.0,
            _qtr_total_mismatch=None, available=False, _reason=_NO_SUMMARY_REASON)
    else:
        row = ix["itc"].get(month)
        if row is None:
            raise mpu.PeriodParseError(
                f"Month {month!r} not covered by any period marker in the 'ITC Available' sheet. "
                f"Months covered: {sorted(ix['itc'])}")
        summary = dict(
            ITC_all_other_IGST=_f(row["all_other_igst"]), ITC_all_other_CGST=_f(row["all_other_cgst"]),
            ITC_all_other_SGST=_f(row["all_other_sgst"]), ITC_all_other_CESS=_f(row["all_other_cess"]),
            ITC_rcm_IGST=_f(row["rcm_igst"]), ITC_rcm_CGST=_f(row["rcm_cgst"]),
            ITC_rcm_SGST=_f(row["rcm_sgst"]), ITC_rcm_CESS=_f(row["rcm_cess"]),
            CN_IGST=_f(row["cn_igst"]), CN_CGST=_f(row["cn_cgst"]), CN_SGST=_f(row["cn_sgst"]),
            CN_CESS=_f(row["cn_cess"]), _qtr_total_mismatch=(row["qtr_total_mismatch"] or None))
    b2b = []
    for r in data["b2b"]:
        if month not in str(r["months_covered"]).split(";"):
            continue
        if (_s(r["gstin"]), _s(r["invno"]).strip().upper()) in ix["sup_inv"]:
            continue
        b2b.append(_b2b_dict(r))
    b2b.extend(_b2ba_dict(r) for r in data["b2ba"] if r["period"] == month)
    cdnr, skipped = [], 0
    for r in data["cdnr"]:
        if month not in str(r["months_covered"]).split(";"):
            continue
        if not r["filed_period"]:
            skipped += 1
        if (_s(r["gstin"]), _s(r["note"]).strip().upper()) in ix["sup_note"]:
            continue
        cdnr.append(_cdnr_dict(r))
    cdnr.extend(_cdnra_dict(r) for r in data["cdnra"] if r["period"] == month)
    summary.setdefault("available", True)
    summary["cdnr_skipped_unparseable_this_month"] = skipped
    warns = [i["message"] for i in data["issues"] if i["id"] == "W301"]
    return dict(summary=summary, b2b=b2b, cdnr=cdnr, available=True, amendment_warnings=warns,
                canonical_issues=[i for i in data["issues"] if i["severity"] in ("ERROR", "WARNING")])
