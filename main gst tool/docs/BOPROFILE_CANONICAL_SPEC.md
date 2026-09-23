# BO Profile canonical format - Spec v1 (as built, 2026-09-23)

Ninth source converted (shared rules in GSTR2B_CANONICAL_SPEC.md, shared code in `canonical_common.py`). Switch: `gst_config.BOPROFILE_USE_CANONICAL` (ON;
`GST_BO_CANONICAL=0` forces the original reader).

## What the engine reads
One function: `gst_parsers_dept.parse_bo_profile(path)` (called once per run via `master_build._safe_parse_bo`), returning `self_gstin, legal_name, trade_name,
demographic` and, per Financial Year, `financial_by_fy, bifa_by_fy, itc_passed_by_fy, itc_received_by_fy, ewb_by_fy, einv_by_fy, refund_by_fy`, plus the lists
`top_beneficiaries, top_suppliers, related_itc_received, related_itc_passed, drc_payments, appeals, cases, transfers`. The BIFA comparisons, the BS/PL
rules (DRC payments), the flow checks (related cancelled parties, top counterparties) and the annual-turnover cross-checks read it.

## What the file looks like
The GST department's "BO Profile" (360-degree profile) export has about 35 sheets; the engine reads 16 of them. Each is a title, a blank row and a heading row
(BIFA: two rows; ITC Passed On / Received: three merged groups INTRA / INTER / TOTAL). BO amounts are in **lakhs**, and are stored as text in the cells
(`'1200.5'`, `'97.5%'`). The original already located every column **by its heading** (the first column whose heading contains all the fragments in a list),
never by position - except the ITC Passed On / Received TOTAL pair, which it took from fixed columns 5 and 6.

## The canonical workbook `Canonical_BOPROFILE_<GSTIN>_<FY>.xlsx` (with more than two Financial Years, `<first>_to_<last>`)
| Sheet | Content |
|---|---|
| META | schema_version `BO-1.0`, GSTIN / legal name / trade name (Demographic Details), the Financial Years found, sheets read and not read, status |
| BO_DEMOGRAPHIC | label / value rows |
| BO_FINANCIAL, BO_BIFA, BO_ITC_PASSED, BO_ITC_RECEIVED, BO_EWB, BO_EINV, BO_REFUND | one row per Financial Year: `fy` + numeric fields (a column that could not be found is blank, not 0 - as the original returned `None`) |
| BO_TOP_BENEFICIARIES, BO_TOP_SUPPLIERS, BO_RELATED_RECEIVED, BO_RELATED_PASSED, BO_DRC, BO_APPEALS, BO_CASES, BO_TRANSFERS | one row per source row, values raw |
| MAPPING_REPORT | per table and field: heading found, column, how matched, status, rows filled / blank |
| ISSUES | problems in plain words, including the sheets the engine does not use |

Heading fragments, the sheet names and the fixed columns of the ITC pair are in `boprofile_mapping.json`. The file is loaded whole (not in read-only mode):
government exports can carry a wrong sheet dimension, which the fast mode trusts and silently truncates. The reader keeps the original's logic (an FY row is only
a row whose first cell looks like `2023-24`, later duplicate FYs overwrite, the DRC rows get `ref` / `type` / `total` derived, ...).

## The ITC Passed On / Received pair
The original took the TOTAL group's two columns from fixed positions 5 and 6. The converter checks the heading (group `TOTAL` + `beneficiaries` / `suppliers` /
`ITC`) at that position: confirmed -> used; not there but exactly one other column fits -> used, **W1202**; no column fits -> the original position, **W1203**
("cannot confirm ... check the figures").

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E1201 | ERROR | the file cannot be opened - the reader returns the empty shape the original returned for an unreadable file, and prints the plain message |
| W1201 | WARNING | a sheet has no recognisable heading row: it reads as empty |
| W1202 / W1203 | WARNING | an ITC column found by heading at a new position / not confirmable from its heading |
| I1201 | INFO | fields whose column was not found (they read as blank) - previously silent |
| I1202 | INFO | sheets the engine can use that are not in this file |
| I1203 | INFO | sheets in the file that no check uses |

## Deliberate differences from the original reader
1. **Sheet names are also accepted by their leading words** (`Top 10 Beneficiaries`, `Top 10 Suppliers`, `ITC Received from Related`, `ITC Passed On to Related`).
   The portal truncates long names to 31 characters and the original needed the exact truncated spelling, silently reading nothing otherwise.
2. The ITC pair follows its heading when a column has been added (W1202).
3. An unreadable file prints a reason (the original returned the empty shape without a word).

## Proof (2026-09-23)
* Reader level, every real BO Profile workbook found on disk - **7 distinct files** (4 taxpayers, 6 to 11 Financial Years each, up to 22 related-party rows,
  DRC / appeals / cases where present): identical to the original for every one, every heading found, no warnings.
* Full engine, old route vs canonical route, fixed hash seed: 598,063 cells (sample taxpayer B, FY 25-26; BO Profile only - it has no Table 8A) and 1,576,926 cells (sample taxpayer A, FY 23-24; BO Profile and Table 8A): 0 differing cells.
* `test_boprofile_adapter.py`: 20 checks on a synthetic workbook with all 16 tables - identical results, `%` and comma cleaning, the DRC rules, blank rows, a moved
  ITC column (W1202, and the original's wrong pair documented), a missing column (None, reported), a sheet with no heading row, a differently truncated sheet name,
  a workbook without the list sheets, an unreadable file, canonical file as input, classification.
