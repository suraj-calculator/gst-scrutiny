# E-Invoice canonical format - Spec v1 (as built, 2026-09-21)

Fourth source converted (after GSTR-2B, GSTR-3B and GSTR-1; shared rules in GSTR2B_CANONICAL_SPEC.md, shared code in
`canonical_common.py`). Switch: `gst_config.EINV_USE_CANONICAL` (ON; `GST_EINV_CANONICAL=0` forces the original readers).

## What the engine reads
Only the tab **`b2b, sez, de`** of the merged E-Invoice workbook (the `cdnr`, `cdnur`, `exp` tabs are not used by any
check and are listed as "not used" in the ISSUES sheet). The tab has a title, ONE heading row (row 4), a banner row per month
(`Financial Year | Tax Period | Date Updated till`) and that month's invoices under it. Three readers use it:

| Original reader | Used by | What it returns |
|---|---|---|
| `parse_einv(path, month)` | monthly totals, "E-Invoice vs GSTR-1", cancelled e-invoice findings | totals per tax head, count, error count, per-line map, cancelled lines |
| `read_einv_lines(path, month)` | line-level E-Invoice vs GSTR-1 checks, late IRN | one dict per invoice line |
| `read_einv_invoices(path, month)` | e-way bill vs e-invoice checks | totals per invoice number |

plus `gst_core._einv_months` (which months the file covers).

## The canonical workbook `Canonical_EINV_<GSTIN>_<FY>.xlsx`
| Sheet | Content |
|---|---|
| META | schema_version `EINV-1.0`, GSTIN (the E-Invoice download states none, so it is taken from the taxpayer's other files), FY, months present, unused tabs, status |
| RETURNS | one row per month: fy, tax period, banner text |
| COVERAGE | one row per month with the number of invoice rows placed in it |
| EINV_B2B | one row per source row: `period` (the month it belongs to), `months_covered` (what its banner covered), `src_row`, `_any`, then the 21 canonical fields, then `extra` (unmapped columns as `heading=value`) |
| MAPPING_REPORT | per field: heading found, column, how matched, OK / MISSING / NOT_IN_PORTAL_FILE, rows filled / blank |
| ISSUES | problems in plain words |

Numbers are converted (blank = 0); identifiers, text and dates stay exactly as in the source. Field names and accepted
heading spellings are in `einv_mapping.json`; headings are compared ignoring case, spaces and punctuation.
Two different E-Invoice files of the same GSTIN/FY in one run get `_2`, `_3` ... so the second never overwrites the first.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E501 | ERROR | the file cannot be opened (corrupt / password-protected / not Excel) |
| E502 | ERROR | no `b2b, sez, de` tab: this does not look like an E-Invoice file |
| E503 | ERROR | no period banner, or no recognisable column headings above the first banner |
| E504 | ERROR | a REQUIRED column (invoice number, taxable value, IGST, CGST, SGST) is missing - the message lists the headings that were found |
| E505 / E506 | ERROR | a data row before any banner / a banner that cannot be understood |
| W501 | WARNING | a quarterly banner: rows were placed in the month of their own invoice date |
| W502 | WARNING | rows under a quarterly banner with no usable date inside that quarter: counted in no month |
| I501 | INFO | optional columns not found (with the consequence, e.g. no status column = cancelled e-invoices cannot be identified) |
| I502 | INFO | tabs no check uses |

An ERROR rejects the file: the engine prints "E-Invoice file ... could not be read for month coverage: [E50x] ..." and
the E-Invoice checks are skipped (E-Invoice is an optional input) - it is never read half-way.

## Deliberate differences from the original readers
1. **Headings are matched ignoring case.** The original looked for `Invoice Value` exactly; a portal file spelling it
   `Invoice value` (the FY 25-26 sample) read as 0 in `read_einv_invoices`. Now read. No visible effect on that sample's
   report.
2. **Quarterly banner.** The original copied the whole quarter into each of its three months (every invoice counted three
   times). Now each row goes to the month of its own invoice date (IRN date as the fallback); undatable rows are reported (W502).
   Monthly banners (all real samples) behave exactly as before.
3. **Cancellation date.** The original looked for headings such as `Cancel date`; the portal's
   `GSTR-1 auto-population/ deletion upon cancellation date` was never matched, so it stays unmapped for `cancel_date` on purpose
   (it is filled for VALID invoices too, so it is not a cancellation date). It is kept in the canonical file as `autopop_date`.
4. Speed: the workbook is opened once, not once per reader per month (the 16,851-row sample: reader-level comparison of
   old vs new dropped from minutes of re-reading to an 11 s conversion plus instant lookups).

## Proof (2026-09-21)
* Reader level, 2 real E-Invoice workbooks (sample taxpayer A FY 23-24: 16,851 rows / 10 months; sample taxpayer B FY 25-26:
  50 rows / 11 months): all 3 readers x every month + a month not in the file + `months`: 0 differences, except the
  documented `Invoice value` heading (difference 1) which appears on sample B only.
* Mutation test (cancelled rows, auto-population errors, blank identity row, a `Cancel date` heading): 0 differences.
* Full engine, old route vs canonical route, fixed hash seed: 1,132,616 cells (A, FY 23-24) and 598,063 cells (B, FY 25-26):
  0 differing cells.
* `test_einv_adapter.py`: 26 checks (cancelled lines, error rows, empty and uncovered months, real date cells, numeric
  MMYYYY period, alias headings, canonical file as input, name clash, all error IDs, quarterly banner).
