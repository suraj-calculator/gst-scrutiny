# Ledger CSV canonical format - Spec v1 (as built, 2026-09-23)

Seventh source group converted (after GSTR-2B, GSTR-3B, GSTR-1, E-Invoice, E-Way Bill and GSTR-2A; shared rules in
GSTR2B_CANONICAL_SPEC.md, shared code in `canonical_common.py`). One adapter, `ledger_adapter.py`, for the **four ledger CSV
downloads**. Switch: `gst_config.LEDGER_USE_CANONICAL` (ON; `GST_LEDGER_CANONICAL=0` forces the original readers).
(The portal's "Tax liability and ITC comparison" Excel report is a different, small file and is not part of this group.)

| Portal report | kind | Original reader | Canonical file |
|---|---|---|---|
| Electronic Cash Ledger | `cash` | `parse_cash_or_liability_ledger(path, "cash")` | `Canonical_CASHLEDGER_<GSTIN>_<FY>.xlsx` |
| Electronic Credit Ledger | `credit` | `parse_credit_ledger(path)` | `Canonical_CREDITLEDGER_...` |
| Electronic Liability Register (Part I) | `liability` | `parse_cash_or_liability_ledger(path, "liability")` | `Canonical_LIABREGISTER_...` |
| Electronic Liability Ledger (Part II, DRC) | `liability_demand` | `parse_cash_or_liability_ledger(path, "liability_demand")` | `Canonical_LIABLEDGER_...` |

Which CSV is which is still decided by `gst_core.classify_folder` from the report's own title line, unchanged; a canonical ledger
workbook supplied in place of the CSV is recognised by its META and slotted into the right list by the kind recorded there.

## The file being read
Every ledger CSV: a title line, `GSTIN` / `From` / `To` lines (the register has `TAX PERIOD` instead), a two-row heading (a group heading row -
e.g. `Integrated Tax Amount Debited / Credited` above six columns Tax / Interest / Penalty / Fee / Others / Total - and a sub-heading row; the row
starts `Sr.No`), then one row per transaction with an `Opening Balance` row and a `Closing Balance` row. Amounts are text, sometimes in the
portal's Indian grouping (`3,06,53,037.00`). The four reports have four different layouts (6 to 8 leading text columns; the credit ledger has one
block of five amounts plus a balance block instead of six parts per head; the liability ledger has a trailing `Stay status`).

## The canonical workbook
| Sheet | Content |
|---|---|
| META | schema_version `LED-1.0`, `ledger_kind`, the file's title, **GSTIN, From / To, FY - all read from the file itself** (so unlike e-way bills the name is always right), row count, status |
| LEDGER_ROWS | one row per source row under the heading: `src_row`, `src_width` (how many cells the source row had - the original's "row too short" tests need it) + the canonical fields + `extra` (any unmapped column as `heading=value`) |
| MAPPING_REPORT | per field: heading found, column, how matched, confidence, rows filled / blank |
| ISSUES | problems in plain words |

Text fields (date, time, reference, tax period, description, transaction type, ledger used, demand id, stay status) are kept exactly as in the
source. **Every amount is converted to a number**, all six parts of all four heads (IGST, CGST, SGST, CESS) in both the Debited/Credited block
(`amt_<head>_<part>`) and the Balance block (`bal_<head>_<part>`) - the engine reads only tax / total / balance total, but the whole ledger is now
in the canonical file. The credit ledger keeps its own five + five (`amt_igst`...`amt_total`, `bal_igst`...`bal_total`). Field names, legacy column
positions and heading fragments are in `ledger_mapping.json`.

The reader keeps the original's business logic over these rows: which rows are transactions (numeric `Sr.No`, or `-` / the Opening Balance row for
the debit/credit ledgers; the credit ledger accepts only numeric `Sr.No` or the Opening Balance row), the Opening row, skipping the Closing row,
credit = +, everything else = -, the roll-ups by tax period and by transaction date, and the credit ledger's accrued-description sets.

## Columns are found by heading, with the original position as the fallback
The original took every column by fixed position. Each field is now located by its heading first:

| Situation | What happens | Reported as |
|---|---|---|
| the heading at the usual position matches | read from there (every real file so far) | MAPPING_REPORT `matched_by = position`, `HIGH` |
| the sub-heading row is shorter than the data (the real Liability Register's is) | the group heading confirms the column | `MEDIUM`, note in the mapping report |
| the heading moved and exactly one unused column matches | read from where it is now | **W902** |
| no column matches | read by position as the original did | **W901** |
| a required column does not exist | the file is rejected | **E903** |

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E901 | ERROR | the file cannot be opened / read |
| E902 | ERROR | no heading row (the row starting `Sr.No`) in the first 30 lines - not a ledger download |
| E903 | ERROR | a required column (date, description, transaction type, or a debit/credit / balance amount the engine reads) could not be located |
| W901 / W902 | WARNING | a column that cannot be confirmed from its heading / that has moved (see above) |
| W903 | WARNING | the file's title does not match the kind it is being read as |
| W904 | WARNING | N amount cell(s) hold text that is not a number and read as 0 (as the original did, silently) |
| I901 | INFO | the file was not UTF-8 and was read as Windows-1252 (the original raised an unexplained UnicodeDecodeError) |
| I902 | INFO | no `Opening Balance` row (normal for the Liability Register and Liability Ledger) |

An ERROR does not stop the run: the reader prints the plain message and returns the empty shape the engine already uses for a ledger that was not
supplied, so that ledger's checks show N/A. (The original crashed the whole run with a Python traceback on such a file.)

## Deliberate differences from the original readers
None in results. The differences are visibility and robustness: a moved column is followed and reported instead of silently misread (W902), a
Windows-1252 file is read, an unusable ledger degrades to "not supplied" with a plain reason instead of crashing, and text in an amount cell is reported.

## Proof (2026-09-23)
* Reader level, every distinct real ledger CSV found in the sample build folders (6 files covering all 4 kinds, 2 taxpayers, FY 23-24 and 25-26;
  24 to 63 transactions each): result identical to the original for every one, every field confirmed by heading, no warnings.
* Full engine, old route vs canonical route, fixed hash seed: 598,063 cells (sample taxpayer B, FY 25-26: cash, credit and liability ledger) and 1,576,926 cells (sample taxpayer A, FY 23-24: cash, credit and liability register): 0 differing cells.
* `test_ledger_adapter.py`: 27 checks - a synthetic file in each of the four real layouts (both date separators, Indian grouping, opening/closing rows,
  the short Liability Register heading, a short `Stay status` row), a moved column (W902, and the original's wrong figures documented), an unmatched
  heading (W901, identical), junk in an amount cell (W904), no heading row (E902), missing amount columns (E903), Windows-1252 file, canonical file as
  input, collision handling and classification of a canonical ledger workbook.
