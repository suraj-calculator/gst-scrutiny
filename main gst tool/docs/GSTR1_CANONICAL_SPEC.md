# GSTR-1 canonical format - Spec v1 (as built, 2026-09-21)

Third source converted (after GSTR-2B and GSTR-3B; shared rules in GSTR2B_CANONICAL_SPEC.md, shared code in
`canonical_common.py`). Switch: `gst_config.GSTR1_USE_CANONICAL` (ON; `GST_1_CANONICAL=0` forces the original
readers).

## Why GSTR-1 is different
The merged workbook has ~27 tabs, all with the same shape: a title, ONE heading row, then a banner row per month
(`Financial Year | Tax Period | ARN | ARN Date`) with that month's rows under it. The engine reads 13 of them through
**13 separate readers** in 6 files, each with its own business logic: forward-filling the blank identity cells on the
continuation rate-lines of a multi-rate invoice (each reader with a slightly different notion of "primary row"),
counting invoices, netting amendments, picking `hsn` versus `hsn(b2b)`+`hsn(b2c)` per month, ...

So the canonical file keeps each table's rows **as the source has them** (numbers converted, blank = 0; identifiers, text
and dates untouched) with canonical column names and the period on every row, and each reader runs its own logic over
those rows. That is what makes the result identical to the original readers.

## The canonical workbook `Canonical_GSTR1_<GSTIN>_<FY>.xlsx`
| Sheet | Content |
|---|---|
| META | schema_version `1-1.0`, GSTIN, legal name, FY, months, tabs present / not read, the `Read me` values, status |
| RETURNS | one row per month: fy, tax period, **ARN, ARN date** (read from the banner), first banner text |
| COVERAGE | one row per (tab, month) that has a banner, with its row count; a tab that cannot be used carries one row with month `*` and the reason |
| B2B_LINES, B2CL_LINES, B2CS_LINES, EXP_LINES, EXEMP_LINES, CDNR_LINES, CDNUR_LINES, B2BA_LINES, CDNRA_LINES, HSN_LINES, DOCS_LINES | source rows: `period, months_covered, src_sheet, src_row, _any` + the table's canonical fields + `extra` (unmapped columns as `heading=value`) |
| MAPPING_REPORT | per tab and field: heading found, column, how matched, OK / MISSING / NOT_IN_THIS_TABLE, rows filled/blank |
| ISSUES | problems in plain words |

`_any` = 1 if the source row had any truthy cell: several readers skip rows by that test, and rows of all zeros differ
from empty rows. `HSN_LINES` carries every row of `hsn`, `hsn(b2b)` and `hsn(b2c)` with its own `src_sheet`. Field names and
accepted heading spellings are in `gstr1_mapping.json`.

## The 13 readers served (all drop-in, same return shape)
`gst_parsers_returns`: parse_gstr1, read_gstr1_hsn_all_months, parse_b2ba, parse_cdnra, parse_docs -
`gst_checks_monthly`: read_gstr1_lines, read_gstr1_invoices - `gst_checks_hsn_fraud`: _cdnr_rows_by_month,
_b2cl_rows_by_month - `gst_checks_forensic`: gstr1_arn_dates_by_month - `gst_checks_flow`: read_gstr1_b2b_ff,
read_gstr1_b2c - `gst_core`: _gstr1_months, _read_me_gstin_and_name.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E101 | ERROR | a REQUIRED column of the b2b tab (taxable value and the four tax heads) not found: raised as `CanonicalLayoutError` for every month (loud; the original crashed with an unexplained TypeError) |
| E102 / E103 / E104 | ERROR | a tab has no heading row / an unreadable period banner / data before the first banner: that tab is unusable, reported as a gap like the original |
| E105 / E106 / E107 | ERROR | file cannot be opened / no `b2b, sez, de_inv` tab / that tab has no recognisable headings |
| W101 | WARNING | a tab has no period banners at all |
| W105 | WARNING | none of a tab's expected headings found: its figures read as blank/0 |
| I103 | INFO | optional columns of a tab not found |
| I401 | INFO | tabs no check uses (e.g. b2cla, expa, at, eco, ...) |

## Deliberate differences from the original readers
1. **`Non-GST Supplies` is its own column.** The original located it with "first heading containing 'non-gst'", which is
   `Exempted (Other than Nil rated/non-GST supply)`, so `nongst_taxable` silently repeated the exempted figure. No visible effect
   on the samples (both are zero there); different whenever the two columns differ.
2. A GSTR-1 whose Table 12 exists only as `hsn(b2b)` / `hsn(b2c)` is now recognised as GSTR-1 (the classifier used to
   require a tab literally named `hsn`, so such a file was silently dropped and the run stopped with "No merged GSTR-1").
3. The heading row is the row above the first banner that matches the most known columns, so a stray row cannot pass for it.
4. Speed: the workbook is opened once, not once per reader per month (a full run on a 12-month file dropped from 7 min to under 2 min).

## Proof (2026-09-21)
* Reader level, 3 real GSTR-1 workbooks (sample taxpayers A FY 23-24 and 24-25, B FY 25-26 with split HSN tabs): all 13 readers x 12
  months (102 comparisons each): 0 differences.
* Full engine, old route vs canonical route, fixed hash seed: 850,444 cells and 598,063 cells: 0 differing cells.
* `test_gstr1_adapter.py`: 25 checks (multi-rate continuation rows, three HSN tab styles, quarterly banner, ARN dates,
  canonical file as input, editing it changes the engine's input, renamed required column, corrupt file, stray row, ...).
