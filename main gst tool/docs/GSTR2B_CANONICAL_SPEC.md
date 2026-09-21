# GSTR-2B canonical format - Spec v1

Status: **v1 implemented** (2026-09-20) behind a switch (`gst_config.GSTR2B_USE_CANONICAL`, default off until you flip it). Sections 3-5 below reflect what was built; where the build differs from the first draft it is marked *(as built)*.
Scope: GSTR-2B only (first source of the "conversion layer"). GSTR-1, GSTR-3B, 2A, E-Invoice, e-way bills, ledgers follow one at a time once this one is proven.

## 1. Purpose and rules

The scrutiny checks must never read a raw portal workbook again for GSTR-2B. An **adapter** converts whatever the portal gave us (monthly file, merged file, with or without `Read me`, any column order) into one fixed Excel file - the **canonical 2B file** - and the engine reads only that.

1. One workbook per recipient GSTIN + FY: `Canonical_GSTR2B_<GSTIN>_<FY>.xlsx`.
2. Same sheets, same column names, same order, for every taxpayer, year and portal layout.
3. Every data row carries its period (see 3.2, `period`).
4. **Blank means "not available". It is never turned into 0.** (One exception, amount cells that exist but are empty - see Open decision 1.)
5. Every value can be traced to its source sheet and row.
6. The adapter never guesses silently: every inference is listed in `MAPPING_REPORT` with a status, every problem in `ISSUES` in plain words.
7. Raw columns the engine does not know are kept, not dropped (`extra` column), so a later change can promote them without re-reading the raw file.
8. Only the adapter knows portal layouts. Column spellings live in a mapping table (JSON), not in code.

## 2. The seam we keep (parity contract)

All engine code reaches GSTR-2B through two functions. The new reader must return **exactly** what they return today, so no check changes.

| Entry point | Called from | Returns |
|---|---|---|
| `parse_2b_excel(path, month)` | master_build (per-month loop, FY 2B index), gst_checks_flow, gst_checks_hsn_fraud | `dict(summary, b2b, cdnr, available, amendment_warnings)` |
| `summary_for_month(path, month)` | master_build, gst_checks_monthly | summary dict + `available`, `_summary_available`, `_source`, `_file`, `_lines` (the parse result), `_reason` when unavailable |

- `b2b` / `cdnr` = list of row dicts with the keys in 3.2 / 3.3.
- Rows are chosen by month: a row belongs to month M if M is covered by the period block it sits under. A **quarterly** block (Apr-Jun) therefore serves all three months - this is today's behaviour and is kept.
- Amended rows (B2BA / CDNRA) are spliced into the month of their own filing period; the superseded original is dropped wherever it sits (today's behaviour). The reader sets `via_amendment = False` on B2B/CDNR rows and `via_amendment = True` plus `original_invno` / `original_note` on spliced amendment rows, exactly as today.
- `summary["cdnr_skipped_unparseable_this_month"]` (a count) is derived by the reader as the number of CDNR rows in that month whose `filed_period` is blank - today such rows are still included, just counted.

## 3. Workbook layout

Types: T = text, N = number, D = date, I = integer. "Req" = adapter must be able to fill it, else ERROR (see section 5).

### 3.1 META (two columns: key, value)

| Key | Meaning |
|---|---|
| schema_version | `2B-1.0` |
| recipient_gstin | The taxpayer's own GSTIN. From `Read me` or filename if present, else supplied by the caller (from GSTR-1/3B). Missing => WARNING. |
| fy | e.g. `2024-25` |
| source_files | Names of the raw files read |
| source_layout | Which layout the adapter recognised (e.g. `docwise-merged`, `full-monthly-with-readme`) |
| adapter_version | Version of the adapter + mapping table |
| created_on | Timestamp |
| periods_present | List, e.g. `Apr-24 ... Mar-25` |
| periods_with_extra_blocks | Periods that came as more than one download (e.g. `Oct-24: 2 blocks`) |
| summary_available | Y / N - was the `ITC Available` control sheet in the source |
| b2ba_available, cdnra_available | Y / N - sheet present AND readable |
| unread_sheets | Raw sheets not used (e.g. ISD, IMPG, ECO) - recorded, never an error |
| status | `OK`, `OK_WITH_WARNINGS` or `FAILED` |

### 3.2 B2B (one row per source row)

*(as built)* "Req" is enforced from `gstr2b_mapping.json` (`"required": true`). Only these B2B fields raise an **ERROR** when their column cannot be found: `gstin, invno, taxable, igst, cgst, sgst, cess, itc_avail` (for CDNR: `gstin, note, taxable, igst, cgst, sgst, cess`). Every other field marked "yes" gives a WARNING and is left blank. Additionally, either the supplier filing-period column or `Rate(%)` must exist (it is how column-shifted rows are detected).

| # | Column | Type | Req | Meaning / what the engine uses it for |
|---|---|---|---|---|
| 1 | period | T | yes | **Statement period** the row sits under: `Apr-24` (monthly) or `Apr-24..Jun-24` (quarterly). This - not the invoice date - decides the month. |
| 2 | months_covered | T | yes | `Apr-24` or `Apr-24;May-24;Jun-24`. The reader returns the row for each covered month. |
| 3 | block_no | I | yes | 1 normally; 2, 3... when the same period was downloaded in parts (portal 1,000-row cap). |
| 4 | gstin | T | yes | Supplier GSTIN - match key for EWB-in, 2A-vs-2B, ITC 3B vs 2B, counterparty sheets |
| 5 | supplier | T | yes | Supplier name - detail sheets, blocked-credit keyword screen |
| 6 | invno | T | yes | Invoice number - match key, B2BA supersession |
| 7 | invtype | T | yes | Regular / SEZ / ... - part of the 2A-vs-2B join key |
| 8 | date | D | yes | Invoice date - month assignment for EWB-in matching |
| 9 | invval | N | yes | Invoice value |
| 10 | pos | T | no | Place of supply (display only) |
| 11 | rcm | T | yes | Reverse-charge flag (Y/N) - RCM comparison |
| 12 | rate | N | no | GST rate; **blank when the source row had no rate** |
| 13 | taxable | N | yes | Taxable value |
| 14 | igst | N | yes | IGST ITC |
| 15 | cgst | N | yes | CGST ITC |
| 16 | sgst | N | yes | SGST/UTGST ITC |
| 17 | cess | N | yes | Cess ITC |
| 18 | itc_avail | T | yes | Yes / No / Unconfirmed - decides what counts as available ITC |
| 19 | itc_avail_reason | T | no | Portal reason text shown with ITC=No rows |
| 20 | einv_source | T | no | `E-Invoice` etc. - analysis #28 |
| 21 | irn | T | no | IRN - analysis #28 |
| 22 | filed_period | T | no | Supplier's own return period tag (`Apr-24`), reference only; blank if unreadable |
| 23 | src_sheet | T | yes | Raw sheet the row came from |
| 24 | src_row | I | yes | Raw row number |
| 25 | row_shift | I | yes | 0 normally; -1 if the source row was column-shifted (blank Rate cell) and was corrected |
| 26 | arith_check | T | yes | `OK` / `MISMATCH` / `NA` - see 4.4 |
| 27 | extra | T | no | Unmapped raw columns as `header=value; header=value` |

### 3.3 CDNR (credit / debit notes)

Same as B2B for columns 1-3 and 22-27, and: `gstin, supplier, note, ntype, supplytype, date, noteval, pos, rate, taxable, igst, cgst, sgst, cess, itc_avail, itc_avail_reason, einv_source, irn` (the engine's existing keys). `note` replaces `invno`, `ntype` = Credit/Debit note, `noteval` replaces `invval`.

### 3.4 B2BA (amendments to invoices)

`period` (the amendment's **own** filing period - it lands in that month), `orig_invno`, `orig_date`, `gstin`, `supplier`, `invno` (revised), `invtype`, `date` (revised), `invval`, `pos`, `rcm`, `rate`, `taxable`, `igst`, `cgst`, `sgst`, `cess`, `itc_avail`, `itc_avail_reason`, `src_sheet`, `src_row`, `extra`.
The engine needs `(gstin, orig_invno)` to drop the stale original and the revised figures to add. If the raw sheet has no readable header, the sheet is written empty, `b2ba_available = N`, and an ISSUE says so (today this only prints a console warning).

### 3.5 CDNRA

`period, orig_note, gstin, supplier, note (revised), ntype, supplytype, date, noteval, pos, rate, taxable, igst, cgst, sgst, cess, src_sheet, src_row, extra`.

### 3.6 ITC_SUMMARY (the `ITC Available` control totals)

*(as built)* One row per **month** (a quarterly file gets three rows, each already reading its own column group): `period, months_covered, all_other_igst/cgst/sgst/cess, rcm_igst/cgst/sgst/cess, cn_igst/cgst/sgst/cess, qtr_total_mismatch, summary_source`.
`summary_source` = `ITC Available sheet` or `not in source`. When not in source the numbers are **blank** and the reader keeps today's fallback (compute from B2B/CDNR rows, marked `_summary_available = False`). The adapter does not compute it, so the canonical file only ever states what the portal stated.

### 3.7 MAPPING_REPORT (one row per canonical field per sheet)

`canonical_sheet, canonical_field, raw_sheet, raw_header_found, raw_column, matched_by, confidence, status, rows_filled, rows_blank, note`

- `matched_by`: `exact` / `alias` / `contains` / `content-inferred`.
- `confidence`: HIGH (exact or alias), MEDIUM (contains), LOW (content-inferred).
- `status`: `OK`, `MISSING`, `GUESSED` (any LOW confidence match - shown in orange, never silent).

This is the "which column am I actually using" answer on every run.

### 3.8 ISSUES

`id, severity, sheet, field, periods_affected, message, what_is_skipped, suggested_action`. Severity: ERROR (canonical file not produced, 2B reported as "not supplied" with the reason), WARNING (produced, something degraded), INFO.

## 4. Adapter rules

4.1 **Period resolution, in order:** the period-marker row the data sits under -> `Read me` -> portal filename -> **ERROR** (never guessed from invoice dates: an invoice can appear in a later month's 2B than its own date).

4.2 **Multi-block periods:** all blocks for a period are kept and numbered `block_no`. A row byte-identical to a row in an earlier block of the same period is skipped as a double download (count recorded), identical rows inside one block are kept.

4.3 **Row-shift correction (from the existing parser):** if a row's Rate cell does not hold a standard GST rate, every column from Rate onward is read one position left and `rate` is left blank; `row_shift = -1`. Applies to B2B and CDNR (CDNR also needs the date fallback for an unreadable period tag).

4.4 **Arithmetic check (proposal, to be validated on real data in step 4):** CDNR: `noteval ~ taxable + igst + cgst + sgst + cess` within Rs 1. B2B: rows can repeat the full invoice value per rate line, so the check is only `invval >= taxable + taxes - 1`. Used to confirm the row-shift detector and to warn on mismatches; not used by any check.

4.5 **Header detection order:** exact header text -> alias list from the mapping table -> "contains" fragment (e.g. `taxablevalue`). *(as built)* Content-based inference ("the column of 15-character GSTINs") is **not in v1**: for amendments it would mean applying half-known columns, which is worse than not applying them. The header rows are the 2 nearest non-empty rows above the first period banner / first data row.

4.6 **Nothing is fatal by accident:** one unreadable optional sheet (B2BA, CDNRA, ITC Available) degrades that sheet only; B2B/CDNR are still produced.

## 5. Errors and warnings (messages the user will see)

Every message names the file, sheet, field, what was looked for, what was found, the periods affected and what is skipped.

| ID | Sev | Example message |
|---|---|---|
| E201 | ERROR | `GSTR-2B "<file>", sheet 'B2B': required field 'Taxable Value' not found. Looked for: 'Taxable Value (₹)', 'Taxable value'. Headers found: [...]. Periods affected: all. GSTR-2B is treated as NOT SUPPLIED; checks skipped: Comparison C/D, ITC 3B vs 2B, EWB-in vs 2B.` |
| E202 | ERROR | `GSTR-2B "<file>": cannot determine the tax period. No period banner in the sheet, no 'Read me', filename does not match '<ts>_<MMYYYY>_<GSTIN>_GSTR2B_<DDMMYYYY>_<n>.xlsx'. Rename the file to the portal name or give the period.` |
| E203 | ERROR | `GSTR-2B "<file>": no GSTR-2B sheet recognised (expected 'B2B' or 'B2B-CDNR'; found: [...]).` |
| E204 | ERROR | `GSTR-2B "<file>": cannot open (corrupt, password-protected or not an Excel file).` |
| E210 | ERROR | `GSTR-2B: several files cover the same period (<periods>) - give one file per period, or one merged file.` |
| W301 | WARNING | `'B2BA' has no readable header (looked for 'GSTIN of supplier', 'Invoice number', period column). Amendments are NOT applied: N amended invoices are taken as-is. Periods affected: ...` |
| W302 | WARNING | `'ITC Available' sheet not in the file. Control totals recomputed from B2B/B2B-CDNR rows; checks F7a and R14's 2B total are skipped.` |
| W303 | WARNING | `N of M B2B rows had a blank Rate and were column-shifted; corrected (row_shift=-1). Rate is blank for those rows.` |
| W304 | WARNING | `Period Oct-24 came in 2 downloads (1,000 + 326 rows); all kept. / N duplicate rows skipped.` |
| W305 | WARNING | `N B2B rows are internally inconsistent as reported by the supplier (document value does not cover taxable value + taxes). Values are kept exactly as reported; the columns are correctly aligned.` *(Measured on sample taxpayer A 24-25: 52 of 15,863 B2B rows, all genuine supplier-side inconsistencies - e.g. invoice value 1,611 against taxable value 1,088,621.91; 0 of 881 CDNR rows.)* |
| W306 | WARNING | `N CDNR rows had an unreadable period tag; recovered from the note date.` |
| W307 | WARNING | `Field 'Rate(%)' not found on sheet 'B2B'. Rate is left blank (the engine reads it as 0.0); the supplier filing-period column is used to detect misaligned rows instead.` *(Correction to the draft: 'Applicable % of Tax Rate' is NOT the GST rate - it is the percentage of the rate applicable, e.g. 100% - so it is never mapped to `rate`.)* |
| I401 | INFO | `Sheets not used: ISD, IMPG, ECO.` |
| I402 | INFO | `Sheet B2BA is not in the file; nothing to net for that amendment type.` |
| W309 | WARNING | `Recipient GSTIN not in the file (no 'Read me') and not supplied by the caller; META left blank.` (the engine supplies it in real runs) |

## 6. Parity and test plan (step 4)

1. Run sample taxpayer A FY 24-25 twice - today's reader vs the canonical reader - and require identical `parse_2b_excel` output for all 12 months (row counts, every key, the summary), then an identical master workbook.
2. Same for sample taxpayer A FY 23-24 (different layout), sample taxpayer B and sample taxpayer C 2B files.
3. Negative tests: rename a required header, delete `Read me`, empty first period's B2BA, corrupt file, two files for one period - each must give exactly the ID in section 5, and nothing else may change.
4. The arithmetic check (4.4) is measured on all these files before it is trusted.
5. If any parity difference cannot be explained as an intended fix, the reader is **not** switched.

## 7. Open decisions

1. **Empty amount cells.** Today an empty tax cell reads as 0. Proposal: keep that when the *column exists* (the portal leaves nil cells empty), count them in `MAPPING_REPORT.rows_blank`; a *missing column* stays blank and raises E201/W307. OK?
2. **Dates.** Store `date` as a real Excel date (recommended; the reader hands the engine the `dd/mm/yyyy` text it gets today) or keep the text?
3. **`extra` column.** Keep unmapped raw columns as text in one column (recommended), or a separate `B2B_EXTRA` sheet?
4. **File name / location.** `Canonical_GSTR2B_<GSTIN>_<FY>.xlsx` in the run's working folder, downloadable from the UI. OK?
5. **Mapping table format.** JSON next to the code (default), or an Excel sheet you can edit?

## 8. Implementation status (as built, 2026-09-20)

| Piece | File |
|---|---|
| Adapter, writer, reader, drop-in `parse_month` | `gstr2b_adapter.py` |
| Column-spelling table (add a spelling here for a renamed portal column) | `gstr2b_mapping.json` |
| Switch | `gst_config.GSTR2B_USE_CANONICAL` (env override `GST_2B_CANONICAL=1/0`); a canonical file supplied as input is always read through the adapter |
| Hook | `gst_parsers_returns.parse_2b_excel` dispatches to the adapter when the switch is on |
| Classification | `gst_core.classify_folder` recognises a canonical file by its `META` sheet; `_gstr2b_months` reads its months |
| Output location | `<working folder>/_canonical/Canonical_GSTR2B_<GSTIN>_<FY>.xlsx` (a sub-folder, so it is never mistaken for an input) |
| UI | results card offers the canonical file and lists its warnings/errors; each run starts with fresh parser caches |
| Tests | `test_gstr2b_adapter.py` (synthetic, no taxpayer data): 20 checks, all pass |

**Parity proof (sample taxpayer A FY 2024-25, 12 months):** `parse_2b_excel` old route vs canonical route: 16,744 rows (B2B + CDNR), 0 differences, and a mutation test confirms the comparison catches a 1-rupee change. Full engine runs, both routes, fixed hash seed: see the section below.

**Known limits of v1**
- A `B2BA` sheet with no header block is still not applied (W301, loudly). Guessing its columns is not safe; the merge tool now keeps the header so a re-merge fixes it.
- Cross-block dedupe applies within one period label only.
- Only GSTR-2B is converted. GSTR-1, 3B, 2A, E-Invoice, e-way bills and ledgers still read raw files.
