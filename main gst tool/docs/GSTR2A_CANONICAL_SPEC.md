# GSTR-2A canonical format - Spec v1 (as built, 2026-09-23)

Sixth source converted (after GSTR-2B, GSTR-3B, GSTR-1, E-Invoice and E-Way Bill; shared rules in
GSTR2B_CANONICAL_SPEC.md, shared code in `canonical_common.py`). Switch: `gst_config.R2A_USE_CANONICAL` (ON;
`GST_R2A_CANONICAL=0` forces the original reader).

## What the engine reads
One function: `gst_parsers_dept.parse_r2a_excel(path)` (called once per run from `master_build._safe_parse_r2a`),
returning `available, reason, months_present, b2b / b2ba / cdnr / cdnra / isd = {month: [row dicts]},
total_row_missing, malformed_gstin`. Every G-series check (G1-G10: 2A vs 2B, 2A vs 3B and the credit ledger, amendments,
data quality) reads that dict.

## Why 2A is different: the original read every column BY POSITION
GSTR-2A sheets have a wrapped two-row heading block (group headings such as `Invoice details` above `Invoice number`),
a period banner per month, and rate-lines plus a `<docno>-Total` rollup row per document. The original took every
column by fixed index ("column 9 is taxable value") and trusted the government template never to change - so a column
GSTN adds or moves would silently make every figure wrong.

The converter reads each column by its **heading first** and falls back to the original position:

| Situation | What happens | Reported as |
|---|---|---|
| the heading at the usual position matches | read from there (this is every real file so far) | MAPPING_REPORT `matched_by = position`, `HIGH` |
| the heading moved and exactly one unused column matches | read from where the heading is now | **W703** "'taxable' is not at its usual column J; found by heading at K" |
| no column matches the heading | read by position exactly as the original did | **W704** "cannot confirm the column for 'taxable' from its heading (column J reads '...')" |
| a required column does not exist at all (sheet cut short) | that sheet reads as empty | **E701** |

Headings are compared after lower-casing and stripping non-alphanumerics; a group heading directly above the heading row
that spans several columns (`Original details` / `Revised details` on the amendment sheets) is applied to the columns it
spans, so original and revised `Invoice number` are told apart. The amendment sheets repeat a sub-heading row right after
every banner; the converter reads it for the labels and the reader then drops such rows by GSTIN shape, as the original did.
Field names, legacy positions and heading fragments are all in `gstr2a_mapping.json`.

## The canonical workbook `Canonical_GSTR2A_<GSTIN>_<FY>.xlsx`
| Sheet | Content |
|---|---|
| META | schema_version `2A-1.0`, GSTIN, legal name and FY (from the `Read me` sheet), months, sheets found, status |
| RETURNS | one row per month: fy, tax period, banner text |
| COVERAGE | per (sheet, month) the number of source rows; a sheet that cannot be used carries one row with month `*` and the reason |
| R2A_B2B, R2A_CDNR, R2A_B2BA, R2A_CDNRA, R2A_ISD | source rows: `months_covered, src_row` + canonical fields (numbers converted, everything else exactly as in the source) + `extra` (unmapped columns, e.g. Rate, as `heading=value`) |
| MAPPING_REPORT | per table and field: heading found, column, how matched, confidence, rows filled/blank |
| ISSUES | problems in plain words |

All source rows under a banner are kept, including rate-lines, `-Total` rows and repeated sub-heading rows; the reader applies
the original's filters (GSTIN-shaped rows only, then only the `-Total` rollup row per document, with any document lacking one
reported in `total_row_missing`). A quarterly banner fans out to its three months, as before.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E701 | ERROR | a REQUIRED column of a sheet could not be located: that sheet reads as empty for the run (other sheets unaffected) |
| E702 | ERROR | the file cannot be opened: `available = False`, reason "Could not read GSTR-2A workbook: [E702] ..." |
| E703 | ERROR | a period banner is not understood (the original raised an unexplained exception here) |
| E704 | ERROR | none of B2B / CDNR / B2BA / CDNRA / ISD found (same reason text as the original) |
| W701 / W702 | WARNING | a sheet has no period banners / its heading row could not be found (the original skipped it silently) |
| W703 / W704 | WARNING | a column found by heading at a new position / a column that cannot be confirmed from its heading (see table above) |
| I702 | INFO | sheets not in the workbook (read as empty) |

## Deliberate differences from the original reader
1. **Columns follow their headings** (W703) - see above. Identical on an unchanged layout.
2. **`B2B-CDNR` / `B2B-CDNRA` sheet names** - which the classifier has always accepted for 2A (`_has_cdnr_pair`) - are now read; the original
   looked for the literal `CDNR` / `CDNRA` and silently read no credit notes from such a file.
3. Unparsable banner / unreadable file are reported in plain words instead of raising.
4. Cost: the first read converts the workbook (about 16 s for a 41,000-row B2B sheet versus about 11 s for the original's single read), and the result is cached, so nothing is read twice.

## Proof (2026-09-23)
* Reader level, 2 real 2A workbooks (sample taxpayers A FY 23-24: 13,565 B2B + 1,824 CDNR rows; B FY 25-26: 13,778 B2B + 66 CDNR + 12 B2BA rows,
  12 months each): `parse_r2a_excel` result identical to the original, including `total_row_missing` and `malformed_gstin`; every field
  matched by position AND heading, no warnings.
* Full engine, old route vs canonical route, fixed hash seed: 598,063 cells (sample taxpayer B, FY 25-26) and 1,576,926 cells (sample taxpayer A, FY 23-24): 0 differing cells.
* `test_gstr2a_adapter.py`: 23 checks on a synthetic all-five-sheets workbook in the real layout - identical results, `-Total` logic,
  B2BA/CDNRA/ISD (no real data in the samples), quarterly ISD banner, moved column (W703 + the original's wrong figures documented), unmatched heading
  (W704, identical), `B2B-CDNR` alias, sheet cut short (E701), canonical file as input, unreadable file, no 2A sheet, not supplied.
