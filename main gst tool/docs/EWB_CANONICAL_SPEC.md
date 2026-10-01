# E-Way Bill canonical format - Spec v1 (as built, 2026-09-23)

Fifth source converted (after GSTR-2B, GSTR-3B, GSTR-1 and E-Invoice; shared rules in GSTR2B_CANONICAL_SPEC.md,
shared code in `canonical_common.py`). Switch: `gst_config.EWB_USE_CANONICAL` (ON; `GST_EWB_CANONICAL=0` forces the
original reader).

## Why e-way bills are different
Every other source converted so far is a portal export with period banners (`Financial Year | Tax Period`) that
say which month a row belongs to. An e-way bill export has **no period markers at all**: it is one flat sheet,
one header row, one row per e-way bill (the merger script — `auto_ewb_merger.py` — just concatenates every
uploaded monthly MIS report and drops exact duplicates). The month a row belongs to is derived entirely from
its own `EWB No. & Dt.` cell, e.g. `'301546430758 - 10/01/2023 17:01:00'`.

There is also no single "GSTR-1"-style engine function for this source: **direction (inward vs outward) is not
in the file** — `gst_core.classify_folder` decides it per file, by parsing every e-way-bill-shaped candidate file
and checking whether the taxpayer's own GSTIN appears mostly in the `From` or the `To` column. This canonical
layer does not change that: it only converts what one raw file is turned into, exactly as before.

## What the engine reads
Only `gst_parsers_returns.parse_annual_ewb(path)`, which returns a list of dicts (`ewbno, ewbdate, ewbtime, month,
docno, docdate, from_gstin, from_name, to_gstin, to_name, from_place, to_place, assess, taxval, hsn, hsn_desc,
vehicle, rate`) — used directly by `gst_core.classify_folder` (direction detection) and by `master_build.py`,
which then hands the same list to every EWB-consuming check via `filter_by_month`. (An older, simpler reader,
`gst_checks_monthly.parse_ewb`, exists but is only used by the standalone `gst_report.py`, which nothing in the
actual product calls — not converted.)

## The canonical workbook `Canonical_EWB_<GSTIN>_<file-stem>.xlsx`
Unlike the other four sources, the filename carries the **source file's own stem** (`inward_eway_bill_merged` /
`outward_eway_bill_merged` / ...) instead of an FY, because nothing in the file itself distinguishes one file from
another and two different files must never overwrite each other's canonical output.

| Sheet | Content |
|---|---|
| META | schema_version `EWB-1.0`, GSTIN (see "GSTIN is always UNKNOWN" below), FY range and date range (computed from the rows, informational only), row count, unread sheets, status |
| EWB_ROWS | one row per e-way bill: `src_row` + the 13 canonical fields, kept as **raw source text** (see below) + `extra` (unmapped columns) |
| MAPPING_REPORT | per field: heading found, column, how matched, OK / MISSING / NOT_IN_THIS_FILE, rows filled / blank |
| ISSUES | problems in plain words |

**Deliberately keeps the raw combined text, not the parsed date/number.** `EWB No. & Dt.` and `Doc No. & Dt.` are
stored exactly as the source has them (e.g. `'301546430758 - 10/01/2023 17:01:00'`), and the split into
number/date/time happens in the reader (`parse_annual_ewb`), the same way the original always did it. This is the
one place this canonical layer differs from the other four (which keep raw source values too, but never derive a
new value from a combined cell) — it means a portal wording change only ever needs a new heading alias in
`ewb_mapping.json`, never a change to the parsing regex, and it sidesteps any date/time round-tripping risk through
the canonical file's own write/read cycle entirely.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E601 | ERROR | the file cannot be opened (corrupt / password-protected / not Excel) |
| E602 | ERROR | no sheet has the e-way bill signature headings (`EWB No.`, `From GSTIN & Name`, `To GSTIN & Name`) — the same test `gst_core.classify_folder` already uses, so a file that reaches this adapter is never rejected here for that reason |
| E603 | ERROR | a REQUIRED column (`EWB No.`, `EWB No. & Dt.`, `Doc No. & Dt.`, `From GSTIN & Name`, `To GSTIN & Name`, `Assess Val.`, `Tax Val.`) is missing — message lists the headings that were found |
| W601 | WARNING | N row(s) have an `EWB No. & Dt.` that could not be read as a date — kept in the file, but excluded from every month-based check (previously silent) |
| I601 | INFO | optional columns not found (`From Place & Pin`, `To Place & Pin`, `HSN Code`, `HSN Desc.`, `Latest Vehicle No.`; `TAX RATE` is `absent_ok` — real exports never have it) |
| I602 | INFO | sheets not used |

## GSTIN is always "UNKNOWN" in the canonical filename/META, and that's expected
Every other source can name its GSTIN from its own content (a `Read me` sheet, or an in-sheet field) or from
context set once the taxpayer's GSTIN is known from another file. An e-way bill export has neither — and
`self_gstin` is not just "not yet known" here, it is **computed from these very files** (the From/To GSTIN
frequency count in `gst_core.classify_folder`), a genuine chicken-and-egg: the file must be parsed once before
its own GSTIN is known, and by the time it is, the canonical file has already been built and cached under
`UNKNOWN`. This is cosmetic only (the META field and the download filename) — every row's content, and every
check's output, is unaffected; confirmed identical in the proof below regardless.

## Deliberate differences from the original reader
None. This conversion changes no behaviour — every case in the proof below (including the header-repeat-row skip,
the blank-row skip and an unparsable date) matches the original exactly. The one improvement is visibility: an
unparsable `EWB No. & Dt.` used to silently drop that row from every month-based check with no trace; it now also
raises W601 so it can be found without reading the raw file.

## Proof (2026-09-23)
* Reader level, 6 real files (2 taxpayers x inward/outward, one taxpayer's inward file also reused as a third
  "pooled multi-year" case): `parse_annual_ewb` + `filter_by_month` on every month present — 0 differences across
  16,864 total rows.
* Mutation test (a literal header-repeat data row inserted, a fully blank row, an unparsable `EWB No. & Dt.`):
  0 differences; W601 correctly raised for the unparsable date.
* `gst_core.classify_folder` direction detection: same `self_gstin`, same inward/outward file assignment, with the
  switch on or off.
* Full engine, old route vs canonical route, fixed hash seed: 598,063 cells (sample taxpayer B, FY 25-26) and
  1,576,926 cells (sample taxpayer A, FY 23-24): 0 differing cells.
* `test_ewb_adapter.py`: 23 checks (combined-text split, both directions, canonical file as input, filename
  collision handling, all three error IDs, header-repeat row, blank row, unparsable date, absent-column handling).
