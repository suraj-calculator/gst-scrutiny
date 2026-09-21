# GSTR-3B canonical format - Spec v1 (as built, 2026-09-21)

Second source converted after GSTR-2B (see GSTR2B_CANONICAL_SPEC.md for the shared rules: one fixed Excel
file per source, every value traceable, nothing guessed silently, mapping in a table not in code).
Switch: `gst_config.GSTR3B_USE_CANONICAL` (now ON; `GST_3B_CANONICAL=0` forces the original readers).

## What GSTR-3B is, and why it is easier than 2B
A FORM: label on the left, figures to the right, one sheet per month. The month is stated inside the sheet
(`Year` + `Tax Period` rows), so there is no "period problem". Lines are found by their label text.

## The four things the engine reads (all served by the adapter now)
| Legacy reader | Adapter function | Used for |
|---|---|---|
| `gst_parsers_returns.parse_gstr3b(path, month)` | `parse_gstr3b` | the 11 form lines below, every monthly comparison |
| `gst_checks_hsn_fraud._gstr3b_month_fields` | `month_fields` | ARN/filing date, 4A5 IGST, 4B reversals (checks #18, #25, cash timing) |
| `gst_checks_forensic.gstr3b_arn_dates_by_month` | `arn_dates_by_month` | late fee / interest per month |
| `gst_core._gstr3b_months`, `_gstr3b_gstin_and_name` | `months`, `gstin_and_name` | which months exist; GSTIN/name fallback |

## The canonical workbook `Canonical_GSTR3B_<GSTIN>_<FY>.xlsx`
| Sheet | One row per | Columns |
|---|---|---|
| META | fact | schema_version `3B-1.0`, recipient_gstin, legal_name, fy, source_files, months_present, sheets_not_used, status |
| RETURNS | month sheet | month, months_covered, fy, tax_period, gstin, legal_name, trade_name, arn, arn_date, generated_on, src_sheet, sheet_no, used |
| LINES | recognised form line per month | month, code, description, raw_label, **taxable, igst, cgst, sgst, cess** (aligned to the form's own column headings, blank where the form leaves the cell blank), n_values, values_in_row_order, src_sheet, src_row |
| MAPPING_REPORT | line code | code, description, rule, required, months_found, months_missing, status (OK / PARTIAL / MISSING) |
| ISSUES | problem | id, severity, sheet, field, periods_affected, message, what_is_skipped, suggested_action |

## Line codes (rules live in `gstr3b_mapping.json`)
`3.1a` outward taxable (required) - `3.1b` zero rated - `3.1c` nil/exempt - `3.1d` reverse-charge inward - `3.1e` non-GST -
`4A3` RCM ITC - `4A4` ISD ITC - `4A5` all other ITC (required) - `4B1` ITC reversed (1) - `4B2` ITC reversed (2) others -
`4C` net ITC (required) - `4D175` 17(5) line of the pre-Aug-2022 form.
`4B1`/`4B2` are read only between the "B. ITC Reversed" and "C. Net ITC available" rows, because "(2) Others" also
appears under D.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E301 | ERROR | no sheet with readable Year + Tax Period - nothing to identify a month from |
| E302 | ERROR | a REQUIRED line (3.1a, 4A5, 4C) missing for a month: that month is refused (the run reports "ERROR processing <month> ... SKIPPED") instead of silently running on zeros, which the original reader did |
| E304 | ERROR | file cannot be opened |
| W301 | WARNING | second sheet for the same month: the first is used, the second ignored |
| W302 | WARNING | a line matched several rows: the last is used |
| W303 | WARNING | an optional line missing for some months (4B1/4B2 fall back to zeros as before, others are absent) |
| W305 | WARNING | Year/Tax Period not recognised on a sheet: the sheet is not used |
| W306 | WARNING | `Date of ARN` is not a recognised date: filing-date checks for that month cannot use it |
| I401 | INFO | sheets that are not a 3B return (e.g. a help sheet) |
| I402 | INFO | a line that only older forms have (currently 4D175, the pre-Aug-2022 17(5) line) is absent, as expected |

## Deliberate differences from the original readers
1. **4B(1) is read from where it sits, whatever its wording.** The fraud-check reader looked for a label starting
   "(1) As per rules 42", but since Aug-2022 the form says "(1) As per rules 38,42 & 43 of CGST Rules and section 17(5)",
   so it always read 0. Found on real data (sample taxpayer B Dec-25: 11,277; another taxpayer Oct-23: 1,764 and Dec-23: 4,596.48).
   Pre-Aug-2022 wording gives identical values.
2. A required line missing for a month is now an explicit error for that month, not silent zeros.
3. Duplicate sheets for one month: the first is used everywhere (the ARN-date reader used to take the last).
Kept as-is on purpose: `rcm_liability_igst` still holds the 3.1(d) first figure (the taxable value) although its name
says IGST; it is not used downstream. A quarterly 3B still parses only its first month (`parse_gstr3b`), months 2-3 are
refused as before.

## Proof (2026-09-21)
* Reader level, 5 real workbooks (sample taxpayer A 23-24 and 24-25, sample taxpayer B 25-26, another taxpayer 23-24): `parse_gstr3b`,
  ARN dates and months identical in every month; the only differences are the intended 4B(1) values above.
* Full engine, old route vs canonical route, fixed hash seed: sample taxpayer A 24-25 (850,444 cells) and sample taxpayer B 25-26
  (598,063 cells): 0 differing cells.
* `test_gstr3b_adapter.py`: 26 checks (both form wordings, missing required/optional lines, duplicate sheet, quarterly,
  corrupt file, raw per-period files = merged workbook, editing the canonical file changes the engine's input).
