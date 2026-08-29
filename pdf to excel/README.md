# PDF to Excel (GST Portal Reports)

## Overview

This folder converts specific GST-portal PDF reports into structured, formatted Excel
workbooks. It is a **preparation/helper tool**: it does not perform any GST analysis
itself, it just turns raw PDF exports (BO Profile / risk-intelligence report, GSTR-9,
GSTR-9C, E-Way Bill Analytics) downloaded from the GST portal into clean `.xlsx` files,
presumably to be fed as input into the "main gst tool" elsewhere in this repository.

It currently supports four document types, auto-detected from each PDF's own text
(see `gst_pdf_to_excel.py`):

- **BO Profile** (GST risk-intelligence / taxpayer profile report) — `parse_bo_profile.py`
- **GSTR-9** (Annual Return) — `parse_gstr9.py`
- **GSTR-9C** (Reconciliation Statement) — `parse_gstr9c.py`
- **E-Way Bill (EWB) Analytics report** — `parse_ewb_analytics.py`

Design principle used throughout (stated in the code comments): a worksheet must
contain the **complete** data for one logical heading/section of the source PDF. A new
heading always starts a brand-new worksheet, and a section is never split mid-table
across two worksheets. Section boundaries are detected from the PDF's own structural
markers (item numbers, "Pt." labels, known field/column-header names) rather than page
breaks or fixed row counts.

An example input/output pair is included in this folder:
`05ASQPB9012R1ZA_BO_Profile_08_08_2026.pdf` -> `05ASQPB9012R1ZA_BO_Profile_08_08_2026.xlsx`
(filename pattern: `<GSTIN>_BO_Profile_<DD_MM_YYYY>.pdf`).

## Core module — `pdf_to_excel_core.py`

Shared utilities imported by all four parsers and by `gst_pdf_to_excel.py`:

- **Cell cleaning** (`clean_cell`, `clean_row`): strips watermark artifacts that bleed
  into extracted table cells (a diagonal "FINAL" watermark on the GSTR forms, a
  report-generator stamp on the BO Profile report — e.g. a stray leading letter +
  newline, or a lone single-letter cell), collapses newlines/whitespace, and returns
  `None` for empty cells.
- **Letterhead/footer detection** (`is_letterhead_or_footer`): matches the repeating
  `GSTIN: ... TradeName` letterhead and the "This is system generated report" footer
  that appear on every BO Profile page, so they can be dropped from extracted data.
- **PDF extraction**:
  - `extract_flat_rows(pdf_path, drop_letterhead=False)` — walks every page with
    `pdfplumber`, extracts every bordered table via `page.find_tables()`, cleans each
    row, merges short orphan fragments left over from a table cut across a page break,
    and de-duplicates repeated header/table blocks (while still preserving genuinely
    distinct side-by-side tables on the same page). Returns one flat, ordered list of
    cleaned rows for the whole document.
  - `extract_verification_rows(pdf_path)` — pulls the free-floating (borderless)
    "Verification of registered person" declaration block at the end of GSTR-9 /
    GSTR-9C directly from page text (date, name of authorised signatory,
    designation/status, declaration paragraph), since it never appears as a table.
- **Numeric formatting** (`to_cell_value`): classifies a cleaned text cell as a
  percentage, an Indian-formatted number (`#,##,##0.00` / `#,##,##0` style, negative in
  red), or plain text/identifier (e.g. long bare-digit strings such as phone numbers,
  TINs, account numbers are kept as text so leading zeros/exact digits survive).
- **Excel writing** (`write_section`, `safe_sheet_name`): writes one logical section as
  one worksheet, with a merged/styled title banner row, optional note row, a blank
  spacer row, then the data rows; auto-detects and bolds header rows (rows of
  label-like text with no numeric values), applies number formats, auto-sizes columns
  (capped 10–55/60 chars), freezes panes below the header, and hides gridlines. Sheet
  names are sanitized to Excel's 31-character/invalid-character rules and de-duplicated.

## Entry point — `gst_pdf_to_excel.py`

Usage:
```
python3 gst_pdf_to_excel.py <input.pdf> [output.xlsx]
python3 gst_pdf_to_excel.py <folder_of_pdfs> [output_folder]
```

- If given a single PDF, converts it to `output.xlsx` (default: same name as the PDF
  with a `.xlsx` extension, in the same location).
- If given a folder, converts every `*.pdf` in it into `output_folder` (default: the
  same folder), writing `<pdf_basename>.xlsx` for each, and prints a per-file
  `[doc_type] file.pdf -> file.xlsx (N sheets)` line (or `[FAILED] file.pdf: <error>`
  on error) without stopping the batch.

**Dispatch (`detect_doc_type`)** — the document type is auto-detected from the PDF's
own extracted text (not from the filename), by scanning the first page's text (first
2000 chars), then falling back to a scan of pages 1–3 if needed:

| Doc type | Detected by |
|---|---|
| `gstr9c` | `'GSTR-9C'` or `'GSTR-9c'` present |
| `gstr9` | `'Form GSTR-9'` or `'Annual Return'` present |
| `ewb_analytics` | `'OUTWARD SUPPLIES'` and `'EWB'` (case-insensitive) present, or later `'HSNWISE'` + `'EWB'` |
| `bo_profile` | `'Demographic Details'` present, or both `'GSTIN'` and `'Overall Risk Score'` present, or later `'Overall Risk Score'`/`'BIFA'` on pages 1-3 |

If none match, it raises `ValueError: Could not auto-detect document type for <path>`.
The resolved type is mapped via a `CONVERTERS` dict to the matching parser module's
`convert(pdf_path, out_path)` function, which does the actual extraction and writes the
workbook.

## Parsers

### `parse_bo_profile.py` — BO Profile (GST risk-intelligence / taxpayer profile report)

- **Input filename pattern seen in this folder**: `<GSTIN>_BO_Profile_<DD_MM_YYYY>.pdf`.
- Calls `extract_flat_rows(pdf_path, drop_letterhead=True)` then splits the flat row
  stream into sections using an ordered list of `BOUNDARIES` — each boundary is a
  `(canonical_title, matcher(row), keep_trigger_row, note_col)` tuple matched against
  each row in sequence, e.g. matching on `row[0] == 'Lead Type'`, `'Bank Name'`,
  `'Member Name'`, a leading `'BIFA...'`, or `_cell_has(row, 'EWB Related Information')`.
  This produces one worksheet per section, including (in the order matched):
  `Demographic Details` (implicit starting section), `Flagged in Lead Based Dashboard`,
  `Pre-GST Details`, `Places of Business`, `Bank Account Details`, `Member Details`,
  `Shared Entity`, `Shared Members`, `Financial Information`,
  `BIFA Specific Information`, `ITC Passed On`, `ITC Received`,
  `Top 10 Beneficiaries based on ITC Passed (last 12 months)`,
  `Top 10 Suppliers based on ITC Received (last 12 months)`,
  `EWB Related Information`, `E-Invoice Related Information`, `Refund Details`,
  `ITC Received from Related / Cancelled Party`,
  `ITC Passed On to Related / Cancelled Party`, `HSN as per REG01`,
  `Top 10 HSN as per GSTR-1 (last 12 months)`, `Top 10 HSN as per EWB (last 12 months)`,
  `Top 10 HSN as per E-Invoice (last 12 months)`, `Export / ICEGATE Information`,
  `Appeal Information`, `Case Information`, `DRC Payment Information`,
  `Transfer Information`.
- It also specially splits the repeated "Return Filing Details" table (same
  `['Date of Filing', 'Return Period', 'IP Address']` header used twice) into two
  distinct sheets: `Return Filing Details - GSTR-3B` (first occurrence) and
  `Return Filing Details - GSTR-1` (second occurrence).
- **Output**: one `.xlsx` workbook with one worksheet per detected section above (via
  `write_section`), default filename `<pdf_basename>.xlsx`.

### `parse_gstr9.py` — GSTR-9 (Annual Return)

- Calls `extract_flat_rows(pdf_path)` (no letterhead-dropping) and groups rows into
  sections keyed by the item number found in column 0, using an `ITEM_GROUPS` map:
  Items 1–3 stay together as `Part I - Basic Details (Items 1-3)`; item 4 ->
  `Item 4 - Advances & Outward+Inward Supplies (Tax Payable)`; item 5 ->
  `Item 5 - Outward Supplies (Tax Not Payable)`; item 6 -> `Item 6 - ITC Availed During
  the Year`; item 7 -> `Item 7 - ITC Reversed & Ineligible ITC`; item 8 -> `Item 8 -
  Other ITC Related Information`; item 9 -> `Item 9 - Tax Paid as Declared in Returns`;
  items 10–14 all collapse into one sheet, `Part V - Transactions Declared in Next FY
  (Items 10-14)`; item 15 -> `Item 15 - Particulars of Demands and Refunds`; item 16 ->
  `Item 16 - Supplies from Composition Taxpayers etc.`; item 19 -> `Item 19 - Late Fee
  Payable and Paid`.
- Handles the PDF's repeated "Pt. X" banner + column-header reprint at the top of every
  page: it uses a `carry` buffer to tell apart a genuine new-item header (kept) from a
  redundant mid-table repeat (dropped).
- `_extract_hsn_note`: items 17 & 18 (HSN-wise summary of outward/inward supplies) have
  no table in this filing, just a one-line note in the page text (mentioning "HSN Wise
  Summary" / "download GSTR 9"), which becomes its own sheet, `Items 17-18 - HSN Wise
  Summary (Note)`.
- Appends a `Verification` sheet from `extract_verification_rows` (declaration, date,
  authorised signatory name, designation/status) when present.
- **Output**: one `.xlsx` with one sheet per item-group above, plus the HSN note sheet
  and Verification sheet.

### `parse_gstr9c.py` — GSTR-9C (Reconciliation Statement)

- Calls `extract_flat_rows(pdf_path)` and groups rows via a similar item-number-driven
  `ITEM_GROUPS` map, starting from `Part I - Basic Details (Items 1-4)`: item 5 ->
  `Item 5 - Reconciliation of Gross Turnover`; item 6 -> `Item 6 - Reasons for Turnover
  Difference`; item 7 -> `Item 7 - Reconciliation of Taxable Turnover`; item 8 -> `Item
  8 - Reasons for Taxable Turnover Difference`; item 9 -> `Item 9 - Rate-wise Tax
  Liability & Amount Payable`; item 10 -> `Item 10 - Reasons for Un-reconciled
  Payment`; item 11 -> `Item 11 - Additional Amount Payable but Not Paid`; item 12 ->
  `Item 12 - Reconciliation of Net ITC`; item 13 -> `Item 13 - Reasons for
  Un-reconciled ITC Difference`; item 14 -> `Item 14 - ITC Reconciliation with
  Books/Expenses`; item 15 -> `Item 15 - Reasons for Un-reconciled ITC (Expense
  Basis)`; item 16 -> `Item 16 - Tax Payable on Un-reconciled ITC Difference`; item 17
  -> `Item 17 - Late Fee Payable and Paid`.
- A headless "Pt.V" banner (with no numbered item following it directly) is mapped via
  `HEADLESS_PT` to `Part V - Additional Liability due to Non-reconciliation`.
- Also appends a `Verification` sheet via `extract_verification_rows`, same as GSTR-9.
- **Output**: one `.xlsx` with one sheet per item-group above plus Verification.

### `parse_ewb_analytics.py` — E-Way Bill (EWB) Analytics report

- Unlike the other three parsers, this one does not use `extract_flat_rows`/table
  extraction at all — the source PDF has no bordered tables, so it reads raw page text
  with `pdfplumber` (`extract_text(use_text_flow=True)`), strips the diagonal "EWB
  ANALYTICS" watermark and known section-title occurrences, and parses the flowing text
  positionally, section by section, using a fixed, numbered `SECTION_HEADERS` list
  (1 through 20), each with a "shape" that selects a parsing function:
  1. `OUTWARD SUPPLIES` (shape `supplies`) — per-month table of EWB count/assessable
     value/tax value across Supplies, Non-Supplies, and Exports.
  2. `INWARD SUPPLIES` (shape `supplies`) — same structure, third group is Imports.
  3. `HSNWISE OUTWARD SUPPLIES` (shape `hsn`) — per-HSN-code table: S.No, HSN Code, HSN
     Description, EWB Count, Assessable Value, Tax Value.
  4. `HSNWISE INWARD SUPPLIES` (shape `hsn`) — same structure for inward supplies.
  5. `EWB VERIFICATION` (shape `month`) — Month-Year, Total Verification EWB, Ok Count,
     Not Ok Count.
  6. `Top 50 Suppliers` (shape `month`, no extra numeric columns beyond Month-Year).
  7. `Top 50 Recipients` (shape `month`).
  8. `CANCELLATIONS & REJECTIONS` (shape `month`) — Cancel EwayBill Count/Assessable
     Value, Rejection EwayBill Count/Assessable Value.
  9. `EWB BLOCKING` (shape `month`).
  10. `EXTENSION OF EWB` (shape `month`) — No. of EWBs, Assessable Value.
  11. `EXTENSION OF TRANSPORTERS` (shape `month`).
  12. `EWB OF CRTICAL COMMODITIES` (shape `month`) — No. of EWBs, Assessable Value, Tax
      Value.
  13. `BILL TO SHIP TO TRANSACTIONS` (shape `month`) — No. of EWBs, Assessable Value.
  14. `OUTWARD B2C TRANSACTIONS` (shape `month`) — No. of EWBs, Assessable Value.
  15. `INWARD B2C TRANSACTIONS` (shape `month`).
  16. `OUTWARD SUPPLIES TO WORK CONTRACTORS` (shape `nomonth`) — S.No-indexed rows: No.
      of EWBs, Assessable Value, Tax Value (no month column).
  17. `ODC EWBs` (shape `month`).
  18. `PART A EWAYBILLS` (shape `month`) — No. of EWBs, Assessable Value.
  19. `EWB-03 Details` (shape `month`).
  20. `RISK INVOLVED` (shape `risk`) — free-form closing narrative: abnormal EWB growth
      by month (>75% of previous month, with count and assessable value) and a
      "Sales to URP (Individual invoice value > Rs 10 Lakh)" entry.
  - `month`/`nomonth` sections whose body says "Details Not Found !!" are recorded as a
    single "Details Not Found" data row instead of being parsed as a table.
- **Output**: one `.xlsx` with one worksheet per numbered section, titled
  `"<number>. <SECTION NAME>"` (e.g. `1. OUTWARD SUPPLIES`), each ending with a
  computed `Total` row (except the `risk` section).

## How to run

**Windows, via `run_converter.bat`**: double-click the batch file. It:
1. Hardcodes `TOOL_DIR` to `C:\Users\admin\Documents\merger\merger\pdf to excel` and
   `cd`s into it.
2. Locates a Python interpreter (`python`, falling back to `py`), and prompts to
   install Python if neither is found.
3. Runs `pip install --quiet pdfplumber openpyxl` to ensure dependencies are present.
4. Runs `python gst_pdf_to_excel.py "%TOOL_DIR%"` — i.e. batch-converts every PDF found
   directly in that folder, writing each `.xlsx` next to its source PDF.
   (Note: `TOOL_DIR` is a hardcoded path specific to the machine/user it was written
   for — it must be edited to match wherever this folder actually lives before the
   `.bat` file will work elsewhere.)

**Directly, via Python** (requires `pdfplumber` and `openpyxl`):
```
python3 gst_pdf_to_excel.py <input.pdf> [output.xlsx]
python3 gst_pdf_to_excel.py <folder_of_pdfs> [output_folder]
```
A single parser module can also be run standalone for testing, e.g.:
```
python3 parse_bo_profile.py <input.pdf> <output.xlsx>
```
(same pattern for `parse_gstr9.py`, `parse_gstr9c.py`, `parse_ewb_analytics.py`, each
exposing a `convert(pdf_path, out_path)` function and a `python3 <file>.py in.pdf
out.xlsx` CLI.)
