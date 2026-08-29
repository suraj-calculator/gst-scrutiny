# GST 360° Scrutiny Tool

## 1. Overview

This is the **core/main tool** of the repository: a Python-based GST (Indian tax)
scrutiny and compliance-analysis engine. Given one taxpayer's GST data for a
financial year (or several), it cross-checks the taxpayer's own filed returns
(GSTR-1, GSTR-3B, GSTR-2B, E-Invoice), the department's own records (GSTR-2A,
Table 8A, the BO/360° risk-intelligence Profile, Electronic Cash/Credit/Liability
ledgers, the portal's own Tax Liability & ITC Comparison report), E-Way Bill
movements (both directions), the annual return (GSTR-9/9C), and optionally a
hand-typed Balance Sheet/P&L — and produces one large multi-sheet Excel workbook
of findings: mismatches, red flags, and reviewable anomalies, each tied where
possible to the specific section of GST law it tests (16(2), 17(5), 18(6),
36(4), 49(4), 86A, Rule 138, etc.) and backed by invoice-level detail, not just
an aggregate number.

**This tool does not fetch or convert anything itself.** Every input it reads is
prepared by the sibling folders in this repository, whose scripts merge/convert
raw GST-portal downloads into the workbook formats this tool consumes:

- **`forms merger/`** — merges per-month/per-quarter GSTR-1, GSTR-2A, GSTR-2B,
  GSTR-3B, and E-Invoice downloads into one `*_Merged.xlsx` workbook per return
  type, covering a whole FY.
- **`e_way bill/`** — converts and merges raw "EWB MIS Report" downloads into
  `inward_eway_bill_merged.xlsx` / `outward_eway_bill_merged.xlsx`.
- **`pdf to excel/`** — converts the GST portal's PDF exports (BO Profile,
  GSTR-9, GSTR-9C, E-Way Bill Analytics) into structured `.xlsx` workbooks,
  since this tool never parses PDF directly.

Every input this tool reads is classified by its **actual content** (sheet
names, header rows, banner text) — never by filename — via `gst_core.classify_folder()`.
This lets the merged/converted outputs of the sibling folders simply be dropped
into one working folder alongside this tool's `.py` files; the tool figures out
what each file is on its own.

The methodology and the "why" behind each check family is documented in
[`docs/GST_360_SCRUTINY_DIRECTIVE.md`](docs/GST_360_SCRUTINY_DIRECTIVE.md) — see
Section 8 below.

## 2. Entry point — `master_build.py`

**Run it as:**
```
python master_build.py [folder]
```
`folder` is an optional positional argument (defaults to `.`, the current
directory). There is no `argparse`/flags — put every input file for one
taxpayer/FY in one folder next to the tool's `.py` files and run this. The
`if __name__ == "__main__":` block just calls `main(folder)`.

**Orchestration sequence inside `main()`:**

1. `gst_core.classify_folder(folder)` — content-based classification of every
   input file into the dict described in Section 7 below. Raises `RuntimeError`
   if the self-GSTIN can't be resolved, or if no GSTR-1/GSTR-3B months are
   found — every other source is optional and degrades gracefully.
2. Parses annual EWB files (`gst_parsers_returns.parse_annual_ewb`), builds the
   set of months covered (intersection of GSTR-1 and GSTR-3B months present),
   and pulls ARN/filing dates (`gst_checks_forensic.gstr1_arn_dates_by_month` /
   `gstr3b_arn_dates_by_month`).
3. Pre-passes GSTR-3B and GSTR-2B for every month to build lookup indexes used
   later by the flow/forensic checks.
4. **Per-month loop** (`run_month()` for each month in scope): parses GSTR-1,
   GSTR-3B, E-Invoice, GSTR-2B for that month, runs the "14 checks" and the
   "27-check EWB matrix" (`gst_checks_monthly`), and the GSTR-1 amendment/
   document-series checks (`gst_parsers_returns.parse_b2ba/parse_cdnra/parse_docs/doc_series_gap_check`).
   A failing month is skipped (logged, not fatal) — the run continues.
5. Cross-month amendment pairing (`build_rectification_pairs`).
6. Annual-source parsing: Cash/Credit/Liability ledgers, Portal Comparison, BO
   Profile (`gst_parsers_dept`), GSTR-2A (`parse_r2a_excel`), GSTR-9/9C/Table
   8A (`parse_gstr9`, `parse_gstr9c`, `parse_table_8a`).
7. `gst_checks_hsn_fraud.run_all()` — the HSN & fraud-pattern check suite.
8. `gst_checks_forensic.check_turnover_gap()` (R13), `check_four_way_itc()`
   (R14), and — if a taxpayer-specific `bs_pl_input.py` is present and its
   `_gstin` tag matches the GSTIN being processed — `check_bs_pl_rules()`
   (R0-R12). A mismatched or absent GSTIN tag causes the BS/PL data to be
   **refused** with a printed warning, never silently applied to the wrong
   taxpayer.
9. `gst_checks_flow.build_context()` / `build_all()` — the annual flow/stock/
   ITC-roll-forward/payment/counterparty check suite (F1–F12, G1–G10),
   wrapped so a failure degrades to one INFO finding instead of aborting.
10. `gst_blocked_credit.build_and_write()` and `gst_machinery_scan.build_and_write()`
    — each individually fault-tolerant, degrading to a `"SKIPPED -- ..."` sheet
    status on error rather than aborting the whole run.
11. Workbook assembly (`openpyxl`): a Master Dashboard, per-month Comparison/
    Analysis/EWB sheets, and every annual/HSN/forensic/flow/blocked-credit/
    machinery sheet, finishing with a 4-sheet QA review layer
    (`QA Summary`, `Action Required`, `EWB Full-Year Reconciliation`,
    `Reviewed Master Dashboard`).

**Output:** `GST_MASTER_<GSTIN>_FY<fy_tag>.xlsx` (e.g.
`GST_MASTER_05AAECM6380J1ZA_FY2022-23.xlsx`, or `..._FY2022-23_to_2023-24.xlsx`
for a multi-year run), written to the current working directory. A run summary
(months covered, gaps, finding counts by pipeline/severity) is printed to
stdout.

## 3. Configuration — `gst_config.py`

A small, purely-additive settings file — described in its own docstring as
holding only thresholds for the *newest* features, never overriding or
duplicating any pre-existing per-file constant (e.g. `MATERIAL`/`TOL` in
`gst_checks_flow.py`, which stay local to that file):

- `BLOCKED_ITC_MASTER_FALLBACK_PATH` — fallback path for the blocked-credit
  keyword master, used only if `classify_folder()` finds no matching file by
  content in the run folder. `None` by default.
- `ITC_ROLLFORWARD_MOM_THRESHOLD = 2.0` — a month-over-month change of more
  than 2x (or less than half) in an ITC roll-forward column is flagged for
  review.
- `ITC_ROLLFORWARD_MOM_FLOOR = 1000.0` — a prior/current pair both under Rs
  1,000 is skipped from that ratio check (statistically meaningless on
  rounding-sized figures).

## 4. Core modules

### `gst_core.py` — period-marker utilities + folder classification

Two consolidated sub-modules:

- **Merged-period utilities** — every sub-sheet of a merged GSTR-1/E-Invoice/
  GSTR-2B workbook carries period-marker rows (e.g.
  `"Financial Year: 2022-23 | Tax Period: January | ARN: ..."`) ahead of each
  month's/quarter's data block; GSTR-3B instead uses one sheet per month
  identified by its own in-sheet `Year`/`Tax Period` rows (sheet *names* like
  `Jan_2022-23` are never trusted). Key functions: `parse_marker_text`,
  `split_rows_by_month`, `rows_for_month`, `find_block_for_month`,
  `find_block_and_index_for_month` (resolves GSTR-2B's quarterly summary,
  which lays 3 months' figures side by side in one row block, to the correct
  month's column group), `months_present`. Raises `PeriodParseError` rather
  than silently guessing when a marker or requested month can't be found.
- **`classify_folder(folder=".")`** — scans every `.xlsx`/`.xlsm`/`.csv` in the
  folder and classifies each by content signature (sheet-name sets, banner
  text, or header-row text) into: merged GSTR-1 / GSTR-3B / E-Invoice /
  GSTR-2B files (with per-month file maps for multi-year folders), annual
  outward/inward E-Way Bill workbooks (direction inferred from which GSTIN
  dominates the `From`/`To` columns), Cash/Credit Ledger and Liability
  Register (Part I)/Ledger (Part II, DRC) CSVs (by first-line banner text),
  Portal Comparison, BO Profile, GSTR-9, GSTR-9C, Table 8A, GSTR-2A, an
  HSN/SAC code master, a blocked-ITC keyword master, and a machinery-HSN
  master. Returns one dict consumed directly by `master_build.py` (see
  Section 7 for the exact keys/signatures).

### `gst_parsers_dept.py` — department/government-side sources

Consolidated from `annual_sources.py`, `annual_return_parser.py`,
`bo_profile_parser.py`, plus GSTR-2A parsing:

| Function | Parses |
|---|---|
| `parse_cash_or_liability_ledger(path, kind)` | Cash Ledger (`kind='cash'`), Liability Register Part I (`'liability'`), Liability Ledger Part II/DRC (`'liability_demand'`) — CSV, distinct column layouts confirmed per kind |
| `parse_credit_ledger(path)` | Electronic Credit Ledger — CSV |
| `parse_portal_comparison(path)` | Portal's own "Comparison Summary" sheet (Tax Liability & ITC Comparison report) |
| `parse_gstr9(path)` | GSTR-9 Annual Return — Excel export, item-numbered sheets (`"Item 4 - Advances"`, `"Item 6 - ITC Availed"`, etc.) |
| `parse_gstr9c(path)` | GSTR-9C Reconciliation Statement — Excel export |
| `parse_table_8a(path)` | Table 8A — Excel export, `B2B`/`CDNR` (or `B2B-CDNR`) sheets |
| `parse_bo_profile(path)` | BO/360° Profile — Excel export, ~16 named sheets (Demographic, Financial, BIFA, ITC Passed/Received, EWB/E-Invoice info, top counterparties, related-party ITC, DRC payments, appeals, cases) |
| `parse_r2a_excel(path)` | GSTR-2A — Excel export, `B2B`/`CDNR`/`B2BA`/`CDNRA`/`ISD` sheets |

All parsers return `dict(available=..., reason=...)` plus the parsed fields —
never raise on a missing/unreadable file. Financial figures in the BO Profile
are in **lakhs**, not rupees (callers scale by 1e5).

### `gst_parsers_returns.py` — taxpayer/return-side sources

Consolidated from `gst_scrutiny_tool.py`, `gstr2b_parser.py`,
`ewb_annual_parser.py`, `amendments.py`:

| Function | Parses |
|---|---|
| `parse_gstr1(path, month)` | GSTR-1 — sub-sheets `b2b, sez, de_inv`, `b2cl`, `exp`, `b2cs`, `cdnr`, `cdnur`, `exemp`, plus HSN summary (`hsn` or split `hsn(b2b)`/`hsn(b2c)`) |
| `parse_gstr3b(path, month)` | GSTR-3B — sheet located by its own `Year`/`Tax Period` rows; extracts 3.1(a)-(e), 4A(3)/(4)/(5), 4B(1)/(2), 4C by label-anchored scanning |
| `parse_einv(path, month)` | E-Invoice — sheet `b2b, sez, de`; separates cancelled IRNs before totalling (cancelled invoices correctly absent from GSTR-1) |
| `parse_2b_excel(path, month)` / `summary_for_month(path, month)` | GSTR-2B — `ITC Available` (Table 3 summary, quarter-aware), `B2B`, `B2B-CDNR`, plus `B2BA`/`B2B-CDNRA` amendments spliced in by content-based header matching |
| `parse_annual_ewb(path)` / `filter_by_month(rows, month)` | Annual E-Way Bill (either direction) — one whole-FY workbook, sheet located by `EWB No.`/`From GSTIN & Name`/`To GSTIN & Name` header |
| `parse_b2ba(path, month)`, `parse_cdnra(path, month)`, `parse_docs(path, month)`, `doc_series_gap_check(...)` | GSTR-1 amendment sheets (`b2ba`, `cdnra`) and Table 13 document-series (`docs`), with 3-tier fuzzy invoice-prefix matching against the actual invoice numbers used |

All month-scoped parsers rely on `gst_core`'s period-marker functions; ledger
and EWB parsers use their own FY-wide date logic instead (these sources are
annual, not marker-blocked).

### `bs_pl_input.py` — Balance Sheet / P&L structured input

Not a parser of any file — a **hand-typed, per-taxpayer Python dict**
(`BS_PL_DATA`), since the tool never OCRs a scanned Balance Sheet. Keys (all
optional; anything omitted still produces an explicit "not tested" finding,
never a silent skip): `total_assets`, `total_equity_liab`, `share_capital`,
`reserves_and_surplus`, `trade_payables`, `short_term_provisions`,
`fixed_assets_tangible`, `non_current_investments`, `inventories`,
`trade_receivables`, `revenue_from_operations`, `other_income`,
`other_expenses`, `finance_costs`, `depreciation`, `net_profit_after_tax` —
each a `{"fy_prior": ..., "fy_current": ...}` pair. A mandatory `"_gstin"` tag
is checked by `master_build.py` before use: if it doesn't match the GSTIN
being processed this run, the data is refused (never applied to the wrong
taxpayer). To supply BS/PL data for a new taxpayer, edit this file's values in
place and keep the variable name `BS_PL_DATA` — `master_build.py` imports that
exact name.

## 5. Checks

### `gst_checks_flow.py` — annual flow, ITC roll-forward, ledger & counterparty forensics (F1–F12, G1–G10)

The largest module in the tool. Recomputes GSTR-2B figures from invoice-level
B2B/B2B-CDNR data rather than trusting the quarterly summary sheet (documented
bug: that sheet's reader used to take the first of three side-by-side monthly
column groups regardless of which month was requested, understating a real
taxpayer's FY ITC total by ~Rs 21 lakh). Constants: `TOL = 1.0` (rupee
tolerance), `MATERIAL = 100000.0` (Rs 1 lakh materiality floor).

- **F1/F1b** — purchase (2B) vs sales (GSTR-1) value-flow, a money-terms stock
  proxy, cross-checked against audited Inventories movement from `bs_pl_input.py`.
- **F2–F7 (ITC roll-forward, "ITC Annual Summary" sheet)** — 4B(1) vs 4B(2)
  reversal split; exempt turnover with zero Rule 42/43 reversal (F2b); Credit
  Ledger tie-out (F3, F5); blocked-credit scan cross-linked to 4D(1)/Sec 17(5)
  (F4); month-over-month outlier detection across every 4A–4D column (F7, uses
  `gst_config`'s threshold/floor).
- **F4/F5 (three-way)** — GSTR-1 vs GSTR-3B(3.1a) vs deduplicated outward EWB,
  narrative branches on gap direction, downgraded to `EXPLAINED` for
  services-dominant HSN profiles (Rule 138 doesn't apply to services).
- **F7/F7a** — ITC claimed (4A(5)) vs invoice-level GSTR-2B available.
- **F8/F8a (RCM)** — liability (3.1d) vs ITC (4A3) vs actual cash-ledger RCM
  debits (Sec 49(4)); flags both underpayment and (documented bug fix)
  large overpayment, previously invisible.
- **F9/F9a/F9b/F9c** — DRC-03/voluntary payments; refund debits vs BO Profile;
  Rule 86A credit blocks; Liability Ledger Part II demand entries matched to
  BO Profile by exact Demand/Source ID.
- **F10** — turnover growth vs cash-tax-paid share, flags persistent low cash
  share.
- **F11/F11a (counterparty)** — same-day repeat transactions (no value floor);
  reciprocal (both-buyer-and-seller) counterparties — circular-trading signal.
- **F12/F12a** — top-10 supplier/beneficiary cross-check against the BO
  Profile's own top-10 lists.
- **F6 (B2B→B2C shift)** — flags a ≥10 percentage-point month-over-month shift.
- **G1–G10 (GSTR-2A cross-checks)** — duplicate/reused 2A invoice numbers
  (G1); malformed GSTIN/missing-total-row data-quality notes; invoice-level
  2A-vs-2B existence/value comparison (G2) and a monthly running-total
  wash-out view (G2a); 2A vs 3B 4A(5) and vs Credit Ledger (G3/G4, Sec
  16(2)(aa)/Rule 36(4)); 2A amendment linkage (G5); RCM cross-check (G6,
  informational — F8/F8a is authoritative) and IGST-vs-CGST+SGST state-code
  validation (G7); ISD credit vs 4A(4) (G8); supplier-filing-date vs Sec
  16(4) deadline (G9); ITC from GSTINs later cancelled (G10).

### `gst_checks_monthly.py` — per-month analysis (14 checks) + E-Way Bill matrix (27 checks)

Consolidated from `gst_analysis_checks.py` and `gst_eway_recon.py`.

**"14 checks" (`run_checks`)**, `TOL = 1.0`: GSTR-1 B2B reconciliation (#0);
Table 8 nil/exempt/non-GST vs 3.1(c)/(e) (#1); credit-note effect on liability
vs 3.1(a) (#2); ITC arithmetic 4C = 4A5+4A3−4B1−4B2 (#3, documented bug fix —
4B1 was silently omitted from the formula); effective-rate comparison, ±0.10pp
tolerance (#4); orphan-invoice re-linking to e-invoice by value/rate/tax key
(#5); duplicate invoice numbers (#6); e-invoice auto-population errors (#7);
IRN-to-filing lag >30 days (#8); rate-wise HSN cross-check, informational (#9);
GSTR-1-vs-3B filing gap >20 days (#10); POS-vs-tax-head mismatch (#11); RCM
routing vs 3.1(d)/4A3 (#12); HSN-summary vs named-invoice IGST gap (#13);
ITC/output-liability ratio >95% "near-zero cash payout" flag (#14).

**27-check EWB matrix (`run`)**: `TOL=1.0`, `VALUE_TOL_PCT=0.01` (1%),
`EWB_THRESHOLD=50000.0` (Rule 138). Covers EWB-Out↔GSTR-1↔E-Invoice
triangulation (#1–9, #17), EWB-In↔GSTR-2B matching (#10–13), same-doc
self-transaction overlap (#15), EWB-vs-document date lag (#7/#16), HSN
cross-check (#18), multiple EWBs per invoice (#23), the EWB/turnover ratio
with a services-HSN carve-out (#25, `EXPLAINED` not `FLAG`), an inward-value
ratio (#26), and repeated-vehicle same-GSTIN-pair trips (#27, circular-trading
signal). Checks #14, #19–22, #24 are explicitly out-of-scope (no
purchase-register/validity/cancellation-status data available from the
exports seen) and render as documented `SKIPPED`/`INFO` findings, never a
silent gap. A documented fix ensures a wholly-missing EWB direction produces
an honest SKIP rather than a misleading false-PASS/false-REVIEW.

### `gst_checks_forensic.py` — filing compliance, R0–R14 rule engine, cancelled-e-invoice cross-checks

Consolidated from `forensic_checks.py` + `filing_compliance.py`.
`TOL_RS=200.0`; `MATERIALITY_PCT=0.01`/`MATERIALITY_FLOOR=100000.0`;
`LATE_FEE_PER_DAY_NORMAL=25.0`/`_NIL=10.0` (Sec 47); `INTEREST_RATE_ANNUAL=0.18`
(Sec 50(1)).

- **R13** (`check_turnover_gap`) — GSTR-9C Table 7B exempt/nil adjustment must
  trace to actual GSTR-1 Table 8 rows.
- **R14** (`check_four_way_itc`) — GSTR-9 6A vs recomputed GSTR-2B FY total vs
  Table 8A vs GSTR-9C 12A (books); flags return-side sources agreeing with
  each other but diverging from books.
- **R0–R12** (`check_bs_pl_rules`, on `bs_pl_input.BS_PL_DATA`) — R0 is a
  pre-flight Assets=Equity+Liabilities gate that halts R1–R12 on failure;
  R1 Revenue vs GSTR-9C turnover; R2 other-income taxability screen; R3 trade
  payables vs Rule 37 (180-day reversal); R4 bad-debt GST-relief screen; R5
  inventories vs Sec 17(5)(h)/Rule 56; R6 fixed assets vs Sec 18(6)/16(3);
  R7 investments vs Schedule I deemed supply; R8 provisions vs undisclosed
  tax exposure; R9 other expenses vs Sec 17(5) blocked categories; R10
  finance costs (exempt interest vs ITC-eligible bank charges); R11 share
  capital non-cash-consideration screen; R12 reserves roll-forward
  (Opening+NPAT=Closing).
- **Filing compliance** (`due_date_gstr1/3b`, `gstr1/gstr3b_arn_dates_by_month`,
  `compute_late_fee`, `compute_interest`, `month_filing_compliance`) — Sec 47
  late fee (capped per Notification 07/2023-CT by turnover slab) and Sec
  50(1) interest, per month.
- **Cancelled e-invoice cross-checks** — a still-live GSTR-1 B2B row for a
  cancelled IRN (D2a); a still-live outward EWB against a cancelled IRN
  (D2b); document-series-gap explanations sourced from B2CS or cancelled
  e-invoices.

### `gst_checks_hsn_fraud.py` — HSN rate/master checks + numbered fraud-pattern library

Consolidated from `hsn_fraud_checks.py`. Documents up front that HSN exists at
invoice level only in EWB data (GSTR-1's `hsn` sheet is a monthly aggregate,
not invoice-linked; GSTR-2B's B2B sheet carries no HSN at all).

- **A1/A1-EXT** wrong GST rate vs a curated, date-versioned `HSN_RATE_HISTORY`
  table (A1) or a supplementary third-party HSN-code CSV (A1-EXT, always
  REVIEW, GST-2.0-era only). **A2** GST charged on an exempt HSN. **A3**
  missing Compensation Cess. **A4** same HSN taxed at two rates in one month.
  **A5** blocked-ITC-by-HSN — always INFO, documented as not computable (no
  HSN column on GSTR-2B's purchase side; see also `gst_blocked_credit.py`
  below for the trade-name-based alternative for this same gap). **A6** HSN
  reported at <6 digits despite turnover >Rs 5Cr (Rule 46). **A7** HSN/SAC
  code doesn't exist in the official code master.
- **B1–B4** POS-vs-tax-head validation (judged against the *supplier's* state,
  not the recipient's — a documented fix to avoid false-flagging normal
  inter-state sales); B2C-Large invoice with no matching outward EWB.
- **C1–C5** RCM HSN cross-check (not computable, INFO); branch-transfer
  detection (same PAN, different state — Rule 28 valuation flag); export/LUT
  and EWB-distance checks (not computable, INFO); intra-state-vs-EWB
  interstate-implied contradiction.
- **~30 numbered fraud-pattern checks** (of a documented 57-item catalogue),
  each independently callable: round-number invoices, below-EWB-threshold
  pricing, reciprocal (BO Profile) trading, credit-note timing gaps, HSN
  mix-share drift and new-code timelines, ITC/liability volatility spikes,
  year-end (last-15-days-of-March) sales dumping, zero-cash-ledger months,
  ghost-supplier PAN clusters, credit notes with no matching inward EWB,
  cross-FY invoice/EWB date splits, post-filing IRN generation, credit
  hoarding vs cash tax paid, negative ITC reversal, midnight/Sunday/holiday
  EWB generation, EWB-generation bursts, PAN-level related-party
  superclusters, EWB-to-invoice date gaps, recurring late-cash-deposit
  patterns, 3-way invoice/IRN/EWB date spread, GSTR-1-vs-E-Invoice matching
  accuracy, BO-Profile high-risk-supplier ITC concentration, head-wise
  GSTR-1-vs-3B mismatches, and invoice-rate outliers. Checks that this tool's
  current inputs genuinely cannot compute (invoice splitting, EWB
  cancellation rate, price-per-unit variance, import BoE mismatch, etc.) are
  emitted as explicit `not_feasible_notes()` INFO findings, never silently
  dropped. Orchestrated by `run_all()`, every check individually
  fault-isolated.

### `gst_blocked_credit.py` — Potential Blocked Credits (Sec 17(5), trade-name screen)

Scans merged GSTR-2B invoice-level data for suppliers whose Trade/Legal name
matches a **taxpayer-supplied** keyword/HSN master list (content-detected:
single sheet, header `Category / Search keyword / Indicative HSN/SAC`) against
Section 17(5) blocked-credit categories (accommodation, motor vehicles, club
membership, insurance, etc.). This is the trade-name-based counterpart to
`gst_checks_hsn_fraud.check_blocked_itc_by_hsn()`'s finding A5 — the same
underlying gap (GSTR-2B carries no HSN/SAC or line-item description), attacked
from the one signal that *is* available from that source. `Match_Confidence`
is capped at `"Medium"` for the same reason. Screening aid only — never
adjusts, removes, or blocks any ITC figure elsewhere in the workbook; degrades
to a `SKIPPED` status (never a crash) if no master file or no GSTR-2B data is
available. Writes one sheet, **"Potential Blocked Credits"**: a per-category
summary table, the flagged invoices grouped by category with full invoice
detail, and (appended) the complete GSTR-2B invoice register for the FY with
every row tagged Yes/No against the same flag.

(**`gst_machinery_scan.py`**, present in this folder but not separately named
in the brief, is the module that closes the parallel "capital goods purchase/
sale" gap using the taxpayer-supplied Machinery HSN master — see Section 6,
which covers it alongside the reporting module since it writes its own
standalone sheet the same way `gst_blocked_credit.py` does.)

## 6. Reporting

### `gst_report.py`

Historically two standalone scripts (`gst_unified_scrutiny.py` and
`build_annual_workbook.py`) kept in this file **only for their reusable
`write_*` sheet-writer functions** — its own two `main()`s
(`main_unified_legacy()`, `main_annual_standalone()`, writing
`GST_Scrutiny_Unified.xlsx` and `GST_Annual_Reconciliation_FY2022-23.xlsx`
respectively) are legacy/standalone and are **not** part of the production
pipeline. `master_build.py` is the actual entry point; it imports this module
(aliased `uni` and `annualwb`) purely for its sheet writers —
`write_comparison`, `write_analysis14`, `write_eway`, `write_monthly`,
`write_fy_total_vs_bifa`, `write_related_party`, `write_top_counterparties` —
and calls the check modules (`gst_checks_flow`, `gst_checks_forensic`,
`gst_checks_hsn_fraud`, `gst_blocked_credit`, `gst_machinery_scan`) itself,
passing their `Finding` objects into these writers. `gst_report.py` does not
import any of those five check modules.

Shared styling: header fill `1F3864` (dark navy) with white bold text, a
severity palette (`FLAG`/`MISMATCH`=red `FFC7CE`, `REVIEW`=amber `FFEB9C`,
`PASS`/`MATCH`=green `C6EFCE`, `INFO`=blue, `SKIPPED`=grey), plus a separate
rupee-magnitude "Severity" band (`severity_band()`: Critical/High/Medium/Low/
Resolved) layered onto some sheets in addition to the Result column. Most
sheets use frozen header panes and an autofilter.

The actual production filename/location is set in `master_build.py` (Section
2): `GST_MASTER_<GSTIN>_FY<fy_tag>.xlsx`, saved to the working directory.

### `gst_machinery_scan.py`

Purely additive, same fault-tolerant convention as `gst_blocked_credit.py`.
Uses the content-detected **machinery HSN master** (a workbook with several
sheets — per-chapter references, a Read Me, and one match-ready sheet whose
header row reads `S.No / HSN Heading (4-digit) / Chapter / Description /
Category / Flag as Machine Purchase? (Y/N/Review) / Match Rule...`). Detects
two independent, cross-checked signals via 4-digit HSN-prefix matching:

- **Sale (Outward)** — GSTR-1's own HSN Summary and outward EWB, matched by
  real HSN code — a possible Sec 18(6) disposal-of-capital-goods signal.
- **Purchase (Inward)** — inward EWB matched by real HSN code (chosen because
  GSTR-2B carries no HSN on the purchase side — the same gap A5/blocked-credit
  document), plus a lower-confidence trade-name keyword screen over GSTR-2B
  supplier names (a curated ~51-term list, deliberately pruned of two terms —
  "crusher", "motor" — after real-data testing showed they were dominated by
  false positives).

Self-to-self (same-GSTIN-both-ends, e.g. branch-transfer) EWB movements are
excluded from both directions. Writes one sheet, **"Machinery HSN Scan"**: a
4-row summary table, full "Sale (Outward)" and "Purchase (Inward)" detail
tables (HSN, description, category, flag, counterparty, value, tax,
reference), and footnotes on source-coverage limitations.

## 7. Expected inputs summary

`gst_core.classify_folder(folder)` scans `*.xlsx`/`*.xlsm`/`*.csv` in one
folder and returns paths for the following — content-classified, never by
filename, so the exact filenames below (from the sibling helper tools) are
illustrative, not required:

| Input | Typically produced by | Detected by |
|---|---|---|
| GSTR-1 (merged, whole FY) | `forms merger/merger-tool/gstr1/` → `GSTR1_Merged.xlsx` | sheets `b2b, sez, de_inv` + `hsn` |
| GSTR-3B (merged) | `forms merger/merger-tool/gstr3b/` → `GSTR3B_Merged.xlsx` | any sheet containing `"Form GSTR-3B"` banner text |
| E-Invoice (merged) | `forms merger/merger-tool/e invoice/` → `EINV_Merged.xlsx` | sheet `b2b, sez, de` (without `b2b, sez, de_inv`) |
| GSTR-2B (merged) | `forms merger/merger-tool/gstr2b/` → `GSTR2B_Merged.xlsx` | sheets `ITC Available` + `B2B` |
| GSTR-2A (whole FY) | `forms merger/merger-tool/gstr2a/` → `R2A_Merged.xlsx` | sheets `B2B`/`B2BA`/`CDNR`(A) + `Read me`, plus a `GSTR 2A`/`GSTR-2A` banner |
| Outward E-Way Bill (annual) | `e_way bill/merge_ewb/outward_eway_bill_merged.xlsx` | header `EWB No.` + `From/To GSTIN & Name`; direction inferred (self-GSTIN mostly in `From`) |
| Inward E-Way Bill (annual) | `e_way bill/merge_ewb/inward_eway_bill_merged.xlsx` | same header signature; self-GSTIN mostly in `To` |
| Cash Ledger | GST portal CSV export | first line contains `"cash ledger"` |
| Credit Ledger | GST portal CSV export | first line contains `"credit ledger"` |
| Liability Register (Part I) | GST portal CSV export | first line contains `"liability register"` |
| Liability Ledger (Part II, DRC) | GST portal CSV export | first line contains `"liability ledger"` |
| Portal Tax Liability & ITC Comparison | GST portal Excel export | sheet `Comparison Summary` |
| BO / 360° Profile | `pdf to excel/parse_bo_profile.py` → `<GSTIN>_BO_Profile_<date>.xlsx` | sheets `Demographic Details` + `Financial Information` + `BIFA Specific Information` |
| GSTR-9 | `pdf to excel/parse_gstr9.py` | sheets starting `Item 4 - Advances` + `Items 17-18` |
| GSTR-9C | `pdf to excel/parse_gstr9c.py` | sheets starting `Item 5 - Reconciliation` + `Item 12 - Reconciliation` |
| Table 8A | GST portal Excel export | sheets `B2B`/`CDNR`(A) + `Read me` (same shape as GSTR-2A, checked second) |
| HSN/SAC code master | e.g. NIC e-Invoice system's `HSN_SAC.xlsx` | sheets `HSN_MSTR` + `SAC_MSTR` |
| Blocked-ITC keyword master | taxpayer-supplied, e.g. `GSTR_2B_Blocked_ITC_Master.xlsx` | single sheet, header `Category / Search keyword / Indicative HSN/SAC` |
| Machinery HSN master | taxpayer-supplied, e.g. `Machinery_HSN_Master.xlsx` | a sheet with header `S.No / HSN Heading (4-digit) / Chapter / Description / Category / Flag as Machine Purchase?... / Match Rule...` |
| Balance Sheet / P&L | hand-typed into `bs_pl_input.py`, GSTIN-tagged | not file-based — Python dict `BS_PL_DATA`, opt-in per taxpayer |

Only the GSTR-1 and GSTR-3B merged workbooks (with the self-GSTIN resolvable,
usually from GSTR-1's own `Read me` sheet, or from the EWB files) are
mandatory; every other source is optional and each check/module degrades to a
documented `SKIPPED`/`INFO` finding rather than failing the run when its
source is absent.

## 8. Further reading

[`docs/GST_360_SCRUTINY_DIRECTIVE.md`](docs/GST_360_SCRUTINY_DIRECTIVE.md) is
the standing methodology brief for this tool: it states the non-negotiable
rules every check must follow (content-based classification, never fabricate,
explicit gaps over silent skips, consistent severity meaning, aggregate +
invoice-level detail together, cite the section of law, cross-check every
pair of sources that can be cross-checked), a full coverage checklist of what
is/isn't built across outward supply, ITC, E-Way Bill, ledgers, annual
returns, capital-goods signals, and counterparty risk, and a list of known
structural gaps (e.g. GSTR-2A/2B carrying no HSN/line-item detail, GSTR-1
credit notes carrying no original-invoice-number field, GSTR-9/9C only being
parsable from Excel not PDF) that explain why certain checks in this codebase
are permanently INFO/SKIPPED rather than built.
