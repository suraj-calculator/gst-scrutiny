# Quarterly (QRMP) filers - how the tool treats them

A quarterly filer files **one** GSTR-3B (and one GSTR-1) per quarter. The scrutiny therefore analyses each quarter as **one period**,
never as three months.

## Rule
* A quarter is a *period unit*: an anchor label (the quarter's **last** month, e.g. `Jun-25`) plus its months
  (`Apr-25, May-25, Jun-25`). Registry: `gst_core.PERIOD_UNITS` (`set_period_units`, `unit_months`, `unit_of`).
* `master_build` detects units from the GSTR-3B (`gstr3b_adapter.quarter_units`) and runs the per-period pipeline only for the anchors.
  A monthly filer has no units - nothing changes for them (checked: byte-identical output on two monthly sample taxpayers).
* Every reader answers the anchor label with the **quarter's total**: GSTR-1 (invoices, credit notes, amendments, HSN, documents),
  E-Invoice, GSTR-2B (ITC summary = sum of the months' groups; invoice and note rows of all three months), e-way bills, GSTR-2A.
* The quarterly GSTR-3B is served **only under the anchor**; the original served it under the first month, so it was compared with one
  month of sales.
* Sheets say so: `Period: Jun-25 (QUARTER Apr-25 to Jun-25, quarterly return)`. Late fee / interest / due dates use the QRMP dates:
  GSTR-1 13th, GSTR-3B 22nd (Category X states) or 24th (all other states / UTs) of the month after the quarter.

## Where the rows of a quarterly *source file* go (GSTR-1, GSTR-2A, E-Invoice)
A quarterly banner used to copy every row into all three months (quarter counted three times). Now each row goes to the month of its
**own date**; tables with no date (B2CS, exempt, HSN, documents) go to the quarter's last month. Warnings W106 (GSTR-1),
W705 (GSTR-2A), W501 (E-Invoice) say so. Quarter totals are exact (proved: quarter = sum of the three monthly results, cell by cell,
on a synthetic quarterly build of a real monthly sample).

## Incomplete input - no fake mismatches
If the GSTR-1 lacks a month of a quarter, the sales comparison would show a difference caused by missing data. The run prints a warning,
the quarter's sheets are titled `INCOMPLETE INPUT: <months> missing from the GSTR-1 file...`; a quarter whose last month is missing is skipped
(a lone month cannot be compared with a quarter's return).

## Mergers
E-Invoice, GSTR-2A and GSTR-3B mergers accept `Apr-Jun`, `Jan - Mar`, `January-March`, `Quarter 4`, `Q4` ... (`gst_merge_common.period_key`).

## Known limits
* GSTR-2B files that carry quarterly banners over-count in a *monthly* run (pre-existing); the quarter-level figure is the correct one.
* Month-over-month checks (e.g. 4(B)(2) trend) compare quarter with quarter for a quarterly filer.

Tests: `test_quarterly_units.py`, `test_quarterly_and_gstin.py`, the quarterly cases in `test_gstr1/2a/3b_adapter.py`.
