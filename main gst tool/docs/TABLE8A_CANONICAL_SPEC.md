# Table 8A canonical format - Spec v1 (as built, 2026-09-23)

Eighth source converted (shared rules in GSTR2B_CANONICAL_SPEC.md, shared code in `canonical_common.py`). Switch:
`gst_config.TABLE8A_USE_CANONICAL` (ON; `GST_T8A_CANONICAL=0` forces the original reader).

## What the engine reads
One function: `gst_parsers_dept.parse_table_8a(path)` (called once per run from `master_build`), returning `available, reason, b2b, cdnr, totals`.
The four-way ITC check (`check_four_way_itc`) and the 2A/2B/8A reconciliation read the `totals` (ITC of the invoices whose ITC availability is "Yes") and
the invoice / note rows.

## What the file looks like
The government Table 8A workbook: a `Read me` sheet (Financial Year, GSTIN, Legal Name), a `B2B` sheet and a credit-note sheet named `CDNR` **or** (newer
exports) `B2B-CDNR`, each with a title, a two-row heading directly below the row that contains `GSTIN of supplier`, then one row per invoice / note.
The number of columns differs between exports (one real export has a leading `Tax Period` column, another has none), so the original already found every
column **by its heading**, never by position. The converter keeps exactly that method (a field's column is the first column whose label - its sub-heading if
it has one, else its group heading - contains all the fragments of the first matching alternative in `table8a_mapping.json`).

## The canonical workbook `Canonical_TABLE8A_<GSTIN>_<FY>.xlsx`
| Sheet | Content |
|---|---|
| META | schema_version `T8A-1.0`, GSTIN / legal name / FY from the `Read me` sheet, which sheet was read for B2B and for credit notes and how it went (`ok`, `no_header`, `no_gstin_col`, `absent`) |
| T8A_B2B, T8A_CDNR | source rows under the heading: `src_row` + canonical fields (numbers converted, text and dates exactly as in the source) + `extra` (unmapped columns as `heading=value`) |
| MAPPING_REPORT | per table and field: heading found, column, status, rows filled / blank |
| ISSUES | problems in plain words |

The file is loaded whole (not in Excel's fast read-only mode): government exports can carry a wrong sheet dimension, which the fast mode trusts and silently
truncates. The reader keeps the original's business logic over the rows: only GSTIN-shaped rows are data (this also drops the wrapped sub-heading row), totals
of the "ITC availability = Yes" invoices (CGST / SGST / IGST / Cess), and the breakdown of the "No" invoices by reason.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E1101 | ERROR | the file cannot be opened - `available = False`, reason "Could not read Table 8A workbook: [E1101] ..." (as before, in plain words) |
| W1101 | WARNING | a sheet is present but no row containing `GSTIN of supplier` was found: the sheet reads as empty (the original's reason text is kept exactly) |
| W1102 | WARNING | the heading row was found but the `GSTIN of supplier` column could not be located: the sheet reads as empty (reason text kept exactly) |
| W1103 | WARNING | neither a B2B nor a CDNR / B2B-CDNR sheet: not a Table 8A export |
| I1101 | INFO | other columns not found - they read as blank / 0 where a check uses them (previously silent) |

## Deliberate differences from the original reader
None in results. The visibility is new: every failure the original recorded only as a `reason` string is now also a numbered issue in the canonical file, and
columns that were silently not found are listed.

## Proof (2026-09-23)
* Reader level, every workbook on disk that the Table 8A reader could be pointed at - **46 distinct files**: the 3 real Table 8A exports (2 taxpayers; up to
  15,867 B2B and 1,827 CDNR rows) and 43 GSTR-2B workbooks (a very different layout, fed to the Table 8A reader as a stress test): result identical for every one,
  no warnings.
* Full engine, old route vs canonical route, fixed hash seed: 1,576,926 cells (sample taxpayer A, FY 23-24, the run that has a Table 8A): 0 differing cells (the FY 25-26 run has none, so it exercises only the "not supplied" path: 598,063 cells, 0 differing).
* `test_table8a_adapter.py`: 21 checks - both column layouts, the `B2B-CDNR` name, wrapped heading row / bad-GSTIN row filtering, Yes / No totals, reason
  buckets, a changed GSTIN heading, no heading row, unreadable file, canonical file as input, classification.
