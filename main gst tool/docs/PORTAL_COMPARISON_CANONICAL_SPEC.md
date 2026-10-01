# Portal "Tax liability and ITC comparison" canonical format - Spec v1 (as built, 2026-09-24)

Last of the portal downloads converted (shared rules in GSTR2B_CANONICAL_SPEC.md, shared code in `canonical_common.py`). Switch:
`gst_config.PORTAL_USE_CANONICAL` (ON; `GST_PORTAL_CANONICAL=0` forces the original reader). Adapter: `portal_adapter.py` + `portal_mapping.json`.

## What the engine reads
`gst_parsers_dept.parse_portal_comparison(path)` (called once per run via `master_build._safe_parse_portal`) returns `{ 'Apr-25': {...}, ... }`: per tax period the
liability per GSTR-1 and GSTR-3B with the shortfall and cumulative shortfall, the ITC per GSTR-3B and GSTR-2B excluding reversal with shortfall and cumulative shortfall,
and (newer reports) the ITC per GSTR-3B considering reversal with its shortfall and cumulative shortfall (`None` when the cell is blank). Only the sheet
`Comparison Summary` is read; the report's other sheets (Tax liability, Reverse charge, Export and SEZ, ITC ...) are not used by any check.

## What the file looks like
A title block (GSTIN / Legal name / Financial Year), then a heading of two rows (newer reports) or **three** rows (older reports, which name the "excluding ITC
reversal" block on its own middle row), then one row per tax period (`Apr-25`) and a `Total` row. The original took every figure from a fixed column.

## Columns are found by heading, with the original position as the fallback
A group heading spans its block (tax liability / ITC excluding reversal / ITC considering reversal); the last heading row holds the sub-headings (As per GSTR-1, As per
GSTR-3B, Shortfall, Cumulative shortfall ...). Each field states the group and sub-heading words it needs (`portal_mapping.json`). A field is read from its original position if the
headings there fit, else from the one other unused column that fits (**W1402**), else from the original position anyway (**W1403**); a required column that does not exist at all
rejects the file (**E1403**). Both the two-row and the three-row layouts confirm every column with no warning.

## The canonical workbook `Canonical_PORTALCOMPARISON_<GSTIN>_<FY>.xlsx`
| Sheet | Content |
|---|---|
| META | schema_version `PC-1.0`, GSTIN / legal name / FY read from the report's own title block, row count, status |
| COMPARISON_ROWS | one row per tax period: `src_row`, `period` and the 11 figures (numbers; the three "considering reversal" figures blank when the source cell is blank) |
| MAPPING_REPORT | per field: heading found, column, how matched, confidence, rows filled / blank |
| ISSUES | problems in plain words |

The workbook is loaded whole (not in read-only mode) because portal files can carry a wrong sheet dimension. The classifier recognises a canonical portal-comparison workbook by its META.

## Errors and warnings
| ID | Sev | Meaning |
|---|---|---|
| E1401 | ERROR | the file cannot be opened - the reader prints the plain message and returns `{}` (the shape for "not supplied"); the original crashed the run with a Python traceback |
| E1403 | ERROR | a required column could not be located - same degrade |
| W1401 | WARNING | no `Comparison Summary` sheet - reads as `{}` (as the original did, silently) |
| W1402 / W1403 | WARNING | a column found by heading at a new position / a column that cannot be confirmed from its heading |
| W1404 / W1405 | WARNING | no heading row starting `Tax Period` (every figure read by position) / no tax-period rows under it |
| I1401 | INFO | sheets no check uses |

## Deliberate differences from the original reader
A moved column is followed and reported instead of silently misread (W1402); an unreadable file degrades to "not supplied" with a printed reason instead of crashing the run.
Identical results on every real report.

## Proof (2026-09-24)
* Reader level, every "Tax liability and ITC comparison" report on disk - **4 distinct real files** (3 taxpayer-years in the newer two-row layout, FY 2023-24 to 2025-26, and
  one FY 2022-23 report in the older three-row layout with blank cells in the "considering reversal" block): identical for every one, every heading confirmed, no warnings.
* Full engine, old route vs canonical route, fixed hash seed: 598,063 cells (sample taxpayer B, FY 25-26) and 1,576,926 cells (sample taxpayer A, FY 23-24): 0 differing cells.
* `test_portal_adapter.py`: 18 checks - both layouts, the Total row, text amounts, blank vs numeric "considering reversal" cells, a moved column (W1402, and the original's wrong figures
  documented), an unmatched heading (W1403, identical), no sheet, unreadable file, a report cut short (E1403), canonical file as input, classification.
