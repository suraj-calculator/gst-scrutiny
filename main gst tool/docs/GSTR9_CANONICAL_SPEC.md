# GSTR-9 / GSTR-9C canonical format - Spec v1 (as built, 2026-09-24)

Tenth and eleventh source (one adapter, `gstr9_adapter.py`, two forms). Shared rules in GSTR2B_CANONICAL_SPEC.md, shared code in `canonical_common.py`. Switch:
`gst_config.GSTR9_USE_CANONICAL` (ON; `GST_R9_CANONICAL=0` forces the original readers).

| Return | Original reader | Canonical file |
|---|---|---|
| GSTR-9 annual return | `gst_parsers_dept.parse_gstr9(path)` | `Canonical_GSTR9_<GSTIN>_<FY>.xlsx` |
| GSTR-9C reconciliation statement | `gst_parsers_dept.parse_gstr9c(path)` | `Canonical_GSTR9C_<GSTIN>_<FY>.xlsx` |

Both are called once per run from `master_build`; the results feed the four-way ITC check, the turnover-gap rule R13, the BS/PL rules, the ITC annual summary
(Table 13 - previous-FY ITC) and the GSTR-9 vs 3B / 1 comparisons.

## What the files look like
The government's Excel export: **one sheet per Item / Part** (`Part I - Basic Details`, `Item 4 - Advances & Outward+Inw...`, ... names truncated to 31
characters), each with a title, a heading block and rows whose **label is in column B** (column A holds the row letter). GSTR-9 Part I is label / value pairs in
columns A / B; GSTR-9C Part I has the label in B and the value in C. The original found a sheet by its leading words, a row by the words in its label
(within the first 60 rows), and then took the numbers from **fixed columns of that row** (Taxable / Central / State / Integrated / Cess for Items 4-8,
Payable / Paid for Item 9, ...). The converter keeps exactly that method; the sheet prefixes, label fragments and occurrence numbers are in `gstr9_mapping.json`.

## The canonical workbook
| Sheet | Content |
|---|---|
| META | schema_version `R9-1.0` / `R9C-1.0`, `form`, GSTIN / legal name / FY from Part I, and which sheet was found for every Item / Part (`sheet_item4` ...) |
| R9_KV | the Part I label / value rows the engine uses: `key` (fy, gstin, legal_name, arn, date_of_filing / arn_date), source row, label, value |
| R9_FACTS | one row per figure the engine uses (a "fact", e.g. `t4_b2b`, `c12_unrec`): the matched row's label, source sheet and row number, and the numbers of its first eight columns `c0..c7` (blank = none, `-` = 0, text that is not a number = blank) |
| R9_LISTS | Item 19 late-fee rows (letters A / B) and the auditor's typed reason lines of GSTR-9C Items 6 / 8 / 13 |
| MAPPING_REPORT | per fact: the sheet, the label matched (or the words looked for), the row, OK / MISSING / SHEET_MISSING |
| ISSUES | problems in plain words |

The file is loaded whole (not in Excel's fast read-only mode): government exports can carry a wrong sheet dimension, which the fast mode trusts and silently
truncates. The readers keep the original's logic: which columns each fact takes and how they are summed, the "system draft" test (no ARN), Item 19 superseding Item 9's
late-fee row, and every `notes` line (byte-for-byte, in the same order).

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E1301 | ERROR | the file cannot be opened - `available = False`, reason "Could not read GSTR-9 workbook: [E1301] ..." (the original's reason, now in plain words) |
| E1304 | ERROR | none of the expected sheets was found: this is not a GSTR-9 / 9C export |
| W1302 | WARNING | an Item / Part sheet the original also noted as missing (Part I, Items 4-9, Part V for GSTR-9; Part I, Items 5 / 7 / 9 / 12 for 9C) - the figures read from it are unavailable |
| W1303 | WARNING | a key row was not found although its sheet is (Table 4B, 4I, 6A, 13 of GSTR-9) - the message names the words looked for. **Seen on a real file:** the FY 2024-25 GSTR-9 words Part V rows 12 / 13 differently ("ITC of the financial year reversed / availed in the *next* financial year") - a different meaning from the older "ITC availed for the previous financial year", so Table 13 is deliberately not read from it (the original does not either, and recorded only a `notes` line); W1303 now says so. Whether the new rows should feed the ITC annual summary is a design decision, not a parsing one, and is left open |
| I1302 | INFO | an optional sheet is absent (Item 19, GSTR-9C reason sheets) |

## Deliberate differences from the original readers
None in results. Visibility only: numbered issues (with the words looked for) in the canonical file, a fact-by-fact MAPPING_REPORT with the source row of every figure,
and a plain message for an unreadable file.

## Known limit (unchanged from the original)
Row labels are matched by words, but the numbers are then taken from fixed columns of the matched row. A new column inserted into a return would still be read at the
old position; the canonical file at least shows the row and all eight column values (`c0..c7`) next to each other, so it is visible.

## Proof (2026-09-24)
* Reader level, every GSTR-9 / GSTR-9C workbook on disk - **10 distinct real files** (5 GSTR-9, 5 GSTR-9C; 4 taxpayer-years, FY 2022-23 to 2024-25, filed and system-draft
  variants, with and without Table 13, Item 19, reason lines): result identical to the original for every one, including every `notes` line.
* Full engine, old route vs canonical route, fixed hash seed: 1,576,926 cells (sample taxpayer A, FY 23-24, the run that has both returns): 0 differing cells; and 598,063 cells (sample taxpayer B, FY 25-26, which has neither, so the "not supplied" path): 0 differing.
* `test_gstr9_adapter.py`: 26 checks - both forms in the real layout, text amounts, no-ARN draft, missing Item 19 / Part V / Item 12 sheets, a missing Table 4B row,
  Table 7B zero note, reason lines, unreadable / not-supplied files, canonical files as input, classification of canonical 9 and 9C into their own lists.
