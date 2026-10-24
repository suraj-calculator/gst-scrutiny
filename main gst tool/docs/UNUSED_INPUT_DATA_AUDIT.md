# Unused Input Data Audit - what the GST tool reads but does not (fully) use

Date: 2026-10-11. Scope: the real taxpayer data in the repo root (Bajrang Motors FY23-24 and FY24-25, Jagnath FY23-24, Tirupati Cement FY25-26) against the engine in `main gst tool/` (identical to `docs/py/core`).

## How to read this

Every raw sheet, table row and column was traced through three layers: the **canonical layer** (adapter plus mapping JSON), the **parsers**, and the **checks**. Each item gets a status:

| Status | Meaning |
|---|---|
| **A** | Never read. The sheet or column is dropped, or sits only in an `extra` blob. |
| **B** | Read or canonicalised, but no check uses it. |
| **C** | Used for display only, not for any decision or comparison. |
| **D** | Used in a check (shown only for contrast). |

Each row says what the data could tell us and a priority (High / Med / Low). Examples come from the real files. Some counts come from merged workbooks, because several raw monthly zips are trimmed to about 29 rows per month. Each section's own caveats are kept inside it.

## Cross-cutting headline findings

1. **Whole sheets or files never read:** GSTR-1 `at`, `atadj`, `ata`, `atadja` (advances), `eco*`, `b2cla`, `expa`, `cdnura`, `b2csa`. GSTR-3B Tables 3.1.1, 3.2, 5, 5.1 and 6.1. GSTR-2B ITC Rejected and ITC not available, the IMS "(Rejected)" sheets, ISD, IMPG and ECO. GSTR-2A TDS, TCS and IMPG. E-invoice `cdnr`. The Electronic Credit Reversal and Re-claimed Statement CSV. Seven of the eight sheets in the portal comparison report. Most BO Profile sheets and GSTR-9/9C items.
2. **Fields read with wrong meaning (these can produce false findings):**
   - Debit notes are subtracted as credit notes in GSTR-1, because Note Type is ignored. This causes the Bajrang Feb-24 CRITICAL mismatch of -1,118,080.18, which is exactly 2 x the 559,040 of debit notes that month.
   - 3B 4(D)(1) is labelled "17(5)" but is now a reclaim of earlier 4(B)(2) reversals.
   - For 3B 4B(1) and 4B(2) only the IGST value is taken, so CGST/SGST reversals count as zero.
   - PMT-09 transfers are counted as cash tax paid in the cash ledger.
   - The "B2B without IRN" check can never fire (`str(None)` is truthy).
   - Table 8A `rate` holds the IGST amount in the FY24-25 export.
   - Bajrang's BO Profile has 12 of its 16 tables stacked in the "ITC Passed On" sheet, so they read as empty.
3. **Cross-checks that become possible only with unused data:**
   - Portal shortfall vs Liability Ledger DRC-03 (Tirupati Jun-25: 14,76,38,341 vs 14,76,38,342).
   - 3B 5.1/6.1 vs the engine's own interest and late-fee estimates.
   - 2A TDS/TCS vs cash-ledger TDS credit (Bajrang 24-25: 2A shows 22,923.95 each head, ledger shows 17,689).
   - GSTR-1 advances (`at`/`atadj`) vs 3B 3.1(a) (Tirupati Jun-25: 34.06 cr advances).
   - E-invoice credit notes vs GSTR-1 CDNR (30 of 60 not found for Bajrang).
   - Credit Reversal statement vs 3B 4(B)(2)/4(D)(1) (Bajrang Aug-24 reclaim 1.42 cr).
   - EWB PINs vs the BO Profile places of business.
   - 2B IMS rejected invoices and credit notes.

## Sections

| # | Source | Section |
|---|---|---|
| 1 | GSTR-1 | all sheets |
| 2 | GSTR-3B | all tables |
| 3 | GSTR-2A and GSTR-2B | all sheets |
| 4 | E-Invoice and E-Way Bill | both directions |
| 5 | Ledgers, Credit Reversal statement, Portal comparison, masters, Balance Sheet/P&L | |
| 6 | BO Profile, GSTR-9, GSTR-9C, Table 8A | |

Caveat: an audit like this has not been re-run to confirm every claim. Verify an item in the code before building a check on it. Counts marked "matched on note number only" and the "impossible travel" style indicators are leads, not proof.

---

# GSTR-1: unused input data audit

Scope: `main gst tool/` (identical to `docs/py/core`), canonical layer `gstr1_adapter.py` + `gstr1_mapping.json` + `docs/GSTR1_CANONICAL_SPEC.md`, consumers in `gst_parsers_returns.py`, `gst_checks_monthly.py`, `gst_checks_flow.py`, `gst_checks_hsn_fraud.py`, `gst_checks_forensic.py`, `gst_machinery_scan.py`, `master_build.py`.

Status key: A = never read; B = read/canonicalised but no check uses it; C = display only; D = used in a check (shown briefly for contrast).

## Data actually inspected

| Taxpayer / layout | What exists |
|---|---|
| Bajrang FY23-24 raw zips (12 monthly `*_Inv_1.xlsx`) | 24 tabs each. Only 29 b2b rows per month, no hsn, no docs: these look like a trimmed first chunk. Full-detail evidence is `output/GSTR1_Merged.xlsx` (same taxpayer, 15,644 b2b invoice header rows, hsn, docs). |
| Bajrang FY24-25 `GSTR1_Merged (7).xlsx` | Merged, banner-per-month. 16,701 b2b invoices, 189 b2cl, 388 b2cs rows, 444 cdnr, hsn 212 rows, docs 561 rows. |
| Jagnath FY23-24 raw zips | Only `b2b` (29 rows per month) and `b2cs` (1 row per month) carry data. All else empty. |
| Tirupati FY25-26 monthly xlsx | `b2b`, `b2cs` (1 row), `at` (2 rows), `atadj` (5 rows). `hsn`, `hsn(b2b)`, `hsn(b2c)`, `docs` all empty. |

Tabs that exist in every layout: Read me, b2b/sez/de_inv, b2cl, exp, b2cs, exemp, b2ba, b2cla, expa, cdnr, cdnur, cdnra, cdnura, b2csa, at, atadj, ata, atadja, hsn (or hsn(b2b)/hsn(b2c)), docs, eco, ecoa, eco_9(5), ecoa_9(5).

Tabs the adapter reads (`raw_names` in the mapping): b2b, b2cl, exp, b2cs, exemp, cdnr, cdnur, b2ba, cdnra, hsn(+b2b/b2c), docs. Every column of those tabs is mapped (nothing falls into the `extra` blob on the real files). So the gaps are (1) 12 tabs never read, (2) mapped columns that no check consumes, (3) wrong logic over a column that is consumed.

Whole-sheet "A" list, which the adapter itself reports as `I401: Tabs not used by any check: ['b2cla','expa','cdnura','b2csa','at','atadj','ata','atadja','eco','eco_9(5)','ecoa','ecoa_9(5)']`.

---

## Sheet `b2b, sez, de_inv` (all layouts, 17 columns)

Contrast (D): Taxable/Integrated/Central/State/Cess (comparison with 3B), GSTIN (PAN match C2, POS vs GSTIN state B1/B3), Invoice Number (E-inv/EWB matching, dups, doc series), Rate (rate outliers), Invoice date (year-end dumping, EWB date gap), Invoice Value and Receiver Name (zero-tax scan, same-day repeats, top buyers), Reverse Charge (RCM line check).

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Invoice Type | `R` on all 15,644 (BM 23-24), 16,701 (BM 24-25), 50 (Tirupati), 348 (Jagnath) | B | Carried into `read_gstr1_b2b_ff` (`invtype`) but no consumer. `check_sez_misclassification` (B4, `gst_checks_hsn_fraud.py` ~1082) literally sets `itype = ""` ("not retained by read_gstr1_lines") and always emits a hard-coded INFO saying "all are R". Real use: SEZWP/SEZWOP/DE/CBW rows vs recipient GSTIN (SEZ units start `..Z`), IGST-on-SEZWOP, Rule 96/LUT, Sec 16 IGST Act. Latent for these four taxpayers (all `R`), but the text is not computed from data. | Med |
| Source | blank on every row, all taxpayers | B | Would split e-invoice-sourced vs manually keyed B2B (Sec 12 GSTR-1 auto-population). Dead column in all four layouts. | Low |
| IRN | blank on every row, all taxpayers | B (check exists but cannot fire) | `parse_gstr1` counts "B2B invoices WITHOUT IRN (should be 0)" with `str(r["irn"]).strip()`. A blank cell is `None`, `str(None)` is `"None"`, which is truthy, so the count is always 0 even though 100% of rows have no IRN (output master Sep-23 etc. shows `Flag 0 / MATCH`). Both the legacy reader (`gst_parsers_returns.py` ~289) and the adapter (`gstr1_adapter.py` parse_gstr1) have this bug. Fix: test `r["irn"] in (None,"")`. Because the portal download does not carry IRN, the real source for this check must remain the E-Invoice file (already done in section B). | Med |
| IRN date | blank everywhere | B | Feeds IRN-lag checks only as a fallback behind e-invoice lines (`einv_lines or g1_lines`), which is empty. Dead. | Low |
| Receiver Name | blank on 33 header rows (BM 23-24, e.g. GSTIN `05MRTE01767D1DR`, 7 invoices), 22 rows (BM 24-25) | C | Used for display and the GSTIN-to-name lookup. Blank names are all government/TDS-style IDs (`05MRT...D..`). Not used to flag name drift: same GSTIN with different trade names across invoices, or recipient name vs 2A/BO profile name. | Low |
| Place Of Supply | `Maharashtra` (8,947 rows of 15,644 for BM) | D | Used in B1/B3. Contrast only. Not used for: destination-state concentration vs EWB destination for those 57% Maharashtra rows (C5 does cover EWB) or vs state-wise registrations in BO profile. | Low |
| Invoice date (vs return period) | `08-05-2023` reported in the **June-23** return; `18-10-2023` (inv 8200853284) reported in **March-24** | D but wrong question | Invoice date is consumed, but nothing compares it to the GSTR-1 tax period. 61 of 15,644 BM FY23-24 invoice header rows (0.4%) sit in a return whose month differs from the invoice date (1 of 16,701 in FY24-25, `DN/23-24/01` dated Feb-24 reported Aug-24). Use: delayed reporting = late time-of-supply declaration, interest u/s 50(1), Sec 37(3) time limit, and the 3B/GSTR-1 monthly mismatch explanation. Grep for "later period / belated / time of supply" in the engine finds no check. | High |

Layout notes: the Tirupati b2b rows show three invoices to the same GSTIN with identical invoice value and taxable (`TCP/UK/25-26/005`, `/007`, `/008`, value 870,815 / taxable 737,979, dated 11-04 and 16-04-2025). Existing same-day/repeat logic uses invoice date and counterparty, but no check looks at identical invoice value + taxable across different serial numbers (duplicate / split billing). That uses already-consumed columns, so it is an idea, not an unused column.

---

## Sheet `b2cl` (BM only; 8 columns)

Contrast (D): Invoice Number to EWB match (B2), taxable/IGST/Cess into 3B comparison, Invoice Value (display in B2 finding).

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Invoice date | `28-01-2024` | B | `_b2cl_rows_by_month` keeps it, but B2 never uses it. Could test EWB date vs invoice date for large inter-state B2C. | Low |
| Place Of Supply | `Uttar Pradesh` (91 of 128 FY23-24 rows) | C | Shown in the B2 message only. Table 5 must be inter-state: POS equal to the taxpayer state (Uttarakhand) is an error. Real data: 0 such rows, so check would PASS, but it is not coded. Also state concentration (UP 71%) vs EWB destination. | Low |
| Invoice Value | min 110,000 in FY24-25 (3 invoices below 2.5 lakh) | C | The Table 5 threshold test (value more than the statutory limit, which changed during 2024) is not applied; 3 FY24-25 rows are below Rs 2.5 lakh. Worth a rule-dated threshold check. | Low |
| Cess Amount | 23,296.49 on `INV24A000017` | D (3B cess row) | Contrast. Compensation cess on 28% vehicles in B2CL but zero in all 15,644 B2B rows: cess presence by rate and HSN is not cross-checked (hsn check ~893 does a cess-prefix test only for HSN rows). | Med |

---

## Sheet `b2cs` (BM, Jagnath, Tirupati; 7 columns)

Contrast (D): taxable/tax into 3B comparison and B2C share (`read_gstr1_b2c`, "B2B to B2C Shift" sheet).

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Place Of Supply | `Uttarakhand` | B | POS-vs-tax-head (B1/B3) runs only on B2B lines. On B2CS: POS = own state must carry CGST+SGST, other state IGST. FY23-24 BM has 0 violations and FY24-25 0 (so low yield), but nothing tests it. | Low |
| Rate x Taxable x Tax | Aug-23 `Haryana`, rate 18, taxable 38,010.93, tax 9,809.00 (25.8%); rate 28, taxable 38,728.14, tax 9,360.00 (24.2%); Mar-25 Haryana 18%, 45,488.29, tax 7,929.89 | B | Tax is not equal to rate x taxable on 6 of 370 rows (FY23-24) and 1 of 388 (FY24-25); the Haryana rows look like rate/tax swapped between the 18% and 28% lines. A row-level rate-arithmetic check exists for nothing in GSTR-1 (b2b rows tie out perfectly: 0 of 21,480). Cheap integrity check, understated liability evidence. | Med |
| Whole sheet as share of turnover | Jagnath: B2CS = 79.9% of GSTR-1 taxable (Rs 5.06 cr of 6.33 cr) with no HSN/docs filed | D | Contrast. Already surfaced by B2B-to-B2C shift; Table 13 gap (below) is the cross-check that would size implied B2C invoices. | Med |

---

## Sheet `exp` (empty in every file; 14 columns)

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Export Type, Port Code, Shipping Bill Number, Shipping Bill Date | no data | B | Mapped (`EXP_LINES`) and only taxable/IGST/Cess are summed into outward totals. WPAY vs WOPAY is ignored: a WOPAY export (3B 3.1(b)) would be wrongly added to the 3.1(a) comparison. SB number/date/port vs ICEGATE shipping-bill data, LUT (C3 is hard-coded "not computable"), refund (RFD-01) cross-check. No data on these four taxpayers, so latent. | Low |

---

## Sheet `exemp` (4 columns, 4 rows per month)

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Description (4 supply classes: intra/inter x registered/unregistered) | `Inter-State supplies to unregistered persons` | B | Rows are summed (nil + exempt + non-GST) into one number. The class split is dropped; it would let you test inter-state vs intra-state nil/exempt against POS and against 3B 3.1(c) split table 3.2. All values are 0 on every real file, so no finding today. Sec 17(2)/Rule 42-43 (#25 uses BO/3B not GSTR-1 exempt). | Low |

---

## Sheet `cdnr` (BM; 18 columns)

Contrast (D): taxable/tax summed into `cn_*` and netted against 3B; GSTIN, Note Number, Note Date in #4 (CN timing) and #11 (CN vs inward EWB).

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| **Note Type** (`C`/`D`) | `D`: 145 distinct debit notes in BM FY23-24 (taxable Rs 5.68 lakh), 9 in FY24-25 (Rs 5.76 lakh) | B (kept by `_cdnr_rows_by_month`, never branched on) | `parse_gstr1` adds every cdnr/cdnur row into `cn_taxable`/`cn_IGST`/... and the comparison does `GSTR-1 gross minus cn_*`, so every debit note is **subtracted** as if it were a credit note. Hard evidence: the output master shows **Feb-24 outward taxable MISMATCH Critical, diff -1,118,080.18**; Feb-24 cdnr debit notes = 559,040 and 2 x 559,040 = 1,118,080 (the 0.18 is paise). Mar-24 debit notes 8,475 -> 2x = 16,950 vs reported diff -17,562 (small remainder, same direction). Sec 34(3) CN reduces, Sec 34(3) DN increases liability. Also #4 and #11 treat debit notes as credit notes (wrongly raise "paper CN risk"). Fix: sign by Note Type in `parse_gstr1`, `cdnr_rows_by_month`, `read_gstr1_lines`. | **High** |
| Note Value | 246,388 (IFFCO-Tokio, `VOU24A000517`) | B | Mapped as `noteval`, dropped by all readers; #11 recomputes value as taxable+tax. Note value vs taxable+tax gap shows missing component (cess/round off). Also Note Value 246,388 against taxable 28.56 on the same line shows a multi-rate note (continuation rows), so it is the only place the whole-note value is recorded. | Low |
| Place Of Supply, Reverse Charge, Note Supply Type | `Uttarakhand`, `N`, `R` | B | Not returned by `_cdnr_rows_by_month`. POS-vs-tax-head for notes, RCM on notes, SEZ/DE note supply types (all `R`). | Low |
| Receiver Name | `IFFCO-TOKIO GENERAL INSURANCE COMPANY LTD.` | B | Not carried into CN readers; CN-heavy counterparties (government canteens, insurers) not named in #4/#11 findings. | Low |
| Note Date vs return month | 0 of 1,136 notes dated outside the return month | D-able | Contrast: notes are all same-month. The tax-point of CN versus original invoice is not computed (cdnr has no invoice reference in the portal file, which #4 acknowledges), but Sec 34(2) CN deadline (30 Nov following FY) could use Note Date vs FY end. | Low |
| Source / IRN / IRN date | blank everywhere | B | Same as b2b. | Low |

---

## Sheet `cdnur` (BM; 13 columns)

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Note Number + Note Value + Note Date | `CNINV24A001220`, 28-01-2024, 866,501, `C`, POS Uttar Pradesh | B (only taxable/tax summed) | `UR Type = B2CL`: the note number embeds the original invoice number. B2CL `INV24A001220` (28-01-2024, Rs 866,501) is fully reversed by a credit note of exactly the same value on the **same day**. FY24-25: `CNINV25A002227` (11-03-2025, Rs 368,837) reverses `INV25A002227` (28-02-2025, Rs 368,837) in full, 11 days later. This is a paper-sale / invoice-then-cancel signature (outward inflation, or cancelled sale that never moved, cross-check with EWB B2). CDNUR is **excluded** from `_cdnr_rows_by_month`, so #4 and #11 never look at it. | Med |
| UR Type | `B2CL` | B | Not used. Distinguishes B2CL notes (inter-state, tied to Table 5 invoices), EXPWP/EXPWOP (export notes) and unregistered import notes. | Low |
| Note Type | `D` once in FY24-25 (`CNINV...`, taxable 19,531) | B | Same debit-note sign issue as cdnr. | Med |
| Place Of Supply, Source/IRN/IRN date | UP; blank | B | Minor. | Low |

---

## Sheet `b2ba` (BM; 16 columns)

Contrast (D): `master_build.build_rectification_pairs` uses gstin, recipient, original/revised invoice number, taxable and tax heads to pair with the original month and show delta.

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Original Invoice date / Revised Invoice date | `02-03-2023` amended in the **Apr-23** return (FY22-23 invoice) | B (parsed as `orig_date`, `revised_date`; grep finds no consumer) | Amendment lag in months and prior-FY originals: such originals are always "NOT FOUND in any earlier month provided". Sec 37(3) cut-off (30 Nov after FY end) and interest/penalty. Date change on revision (not just number) is not shown. 145 amendment rows in FY23-24, 191 in FY24-25. | Med |
| Place Of Supply, Invoice Value, Invoice Type, Reverse Charge, Rate | `Uttarakhand` vs original `Maharashtra`, 708.00 | B | Not shown on the rectification sheet. On the Mahindra rows the amendment moves tax from IGST to CGST+SGST (recipient GSTIN changes `27AAACM...` to `05AAACM...`): that is a place-of-supply correction, which has a state-wise tax-head consequence not reported. Rate change (original rate not in this tab; only inferable from the paired B2B line). | Low |

## Sheet `cdnra` (14 data rows)

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Original/Revised Note Date, Note Value, POS, Rate, Cess, Receiver Name | `31-01-2024` | B | Parsed partly; rectification sheet states that note amendments are "not month-matched". The original note is in `cdnr` of an earlier month, so a pairing like the invoice one is feasible via Original Note Number. | Low |

---

## Sheet `hsn` / `hsn(b2b)` / `hsn(b2c)` (BM both years: data; Tirupati, Jagnath: empty)

Contrast (D): HSN, Description, Rate, Taxable, tax heads feed rate-master checks A1-A7, HSN drift #6/#27, hsn vs invoice reconciliation, machinery scan, blocked-credit HSN.

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| **Total Quantity** | `87033299` Vehicles: 700 units, taxable Rs 84.85 cr (FY24-25); BM FY23-24: 406 units, Rs 49.5 cr | C (only printed on the "HSN Timeline" sheet, `master_build.py` ~2414) | The engine says the opposite: `not_feasible_notes()` (`gst_checks_hsn_fraud.py` ~2046) and forensic R5 (~377) state the HSN quantity "is 0 on every row inspected / a known structural zero", so #16/#44/#53 (average price variance, price band, MoM price drift by HSN) and #22 are published as INFO "not computable". Real data: **123 of 212 HSN rows (FY24-25) and 124 of 196 (FY23-24) have quantity > 0, covering 97.0% and 96.5% of taxable value**. Only service rows (`996111 Commission Finance Co`, `997161 Distribution Fees`, `998599 Dealer Incentive`) have qty 0. Per-unit price is computable: `87033299` runs Rs 1.15-1.26 lakh per unit across months; `87042190` Rs 7.16-7.75 lakh. Uses: price-band/drift (undervaluation, Sec 15 / Rule 27-28), units sold vs outward EWB count and vs RTO/vehicle stock, input-output quantity ratio with purchase side (2B has no qty, but stock register demanded by R5), cross-check against `Quantity` in E-invoice item data, and units reconciliation to Form 22A/dealer invoices. | **High** |
| UQC | `NOS-NUMBERS`, `OTH-OTHERS`; blank for 52 of 196 rows | C | Display only. Unit consistency per HSN across months (a vehicle HSN in OTH, mixed units breaking the per-unit price), and missing UQC on goods rows (Rule 46(h)/GSTR-1 Table 12 validation). Blank only on service rows here. | Low |
| Cess Amount | 0 on all HSN rows with no Cess | D (check ~893 cess-prefix) / B for `hsn_CESS` | Parsed totals `hsn_CESS` never used; the cess-prefix check covers rows. Cross-check of HSN cess against B2B cess (b2b cess is 0 in all 15,644 rows, while b2cl cess is nonzero): cess on 87033299 should appear in B2B. | Med |
| Rate | 28, 18, 12, 5 | D | Contrast. Rate mix of the HSN table vs rate buckets on B2B/B2CS (tax-rate coverage) not compared. | Low |
| Description | `Spare Parts`, `Vehicles` | D (rate master only) | Free-text description vs HSN code mismatch (e.g. a vehicle description on an accessory HSN) only used for fuzzy master matching, not as a mismatch flag. | Low |
| Whole tab when empty | Tirupati: 4 `hsn*` tabs empty for 12 months; Jagnath: empty for 12 months | B (a `month_marker_gaps` "hsn" entry exists) | A taxpayer with Rs 181 cr (Tirupati) / Rs 6.3 cr (Jagnath) outward and no Table 12 is a compliance finding in itself (Notification 78/2020 mandatory Table 12 reporting; 4/6-digit rule by turnover). Currently silent gap, not a flagged finding. | Med |

---

## Sheet `docs` / Table 13 (BM merged only; 5 columns)

Contrast (D): "Invoices for outward supply" rows go through `doc_series_gap_check` (series gaps, excess count, cancelled count explaining gaps).

| Raw column / row group | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Cancelled (rate) on outward invoices | FY24-25: **4,998 cancelled of 38,579 issued (13.0%)**; FY23-24: 523 of 30,340 (1.7%) | B/C (used only to "explain" gaps) | No cancelled-rate flag. A jump from 1.7% to 13.0% year on year is an anomaly (backdated cancellations, sale then cancel around quarter ends, gaming turnover threshold). Cross-check: cancelled e-invoices (IRN cancelled) in the E-Invoice file, cancelled EWBs. | High |
| Total Number minus Cancelled vs reported invoices | FY24-25: 38,579 - 4,998 = 33,581 issued; B2B invoice headers 16,701 + B2CL 189 = 16,890; implied B2C-small count 16,691 | A (no cross-count) | B2CS is an unnumbered summary, so Table 13 is the only way to size B2C volume. Implied average B2CS invoice = Rs 1.726 bn / 16,691 = about Rs 1.03 lakh (FY24-25) and Rs 1.27 bn / 14,045 = Rs 0.91 lakh (FY23-24). Flags B2C invoice-splitting (#13 is currently marked "not detectable") and revenue completeness. | High |
| Nature of Document = `Credit Note` | FY23-24: 1,814 issued, 186 cancelled -> 1,628 live vs 991 cdnr + 4 cdnur notes reported (about 633 unreported). FY24-25: 828 - 34 = 794 vs 435 + 3 | B (parsed, not tested; `doc_series_gap_check` skips every nature but outward invoices) | Credit-note serial coverage: credit notes issued but absent from Table 9B (could legitimately be B2C notes netted in B2CS, but then B2CS should be net of them). Sec 34 CN not declared = ITC wrongly retained by recipient / output overstated. | High |
| Nature = `Debit Note` | FY23-24: 79 issued vs 145 distinct D notes in cdnr; FY24-25: 24 (7 cancelled) vs 9+1 | B | Reverse inconsistency (reported D notes exceed issued Table 13 count in FY23-24). Either Table 13 under-declared or reported notes mis-typed. | Med |
| Nature = `Invoices for inward supply from unregistered person` | FY23-24: 24 documents; FY24-25: 11 | B | RCM self-invoices count (Sec 31(3)(f)). Tie to 3B 3.1(d) RCM liability and to the RCM triangulation sheet; zero count with RCM tax in 3B (or vice versa) is a flag. | Med |
| Nature = other (delivery challan, receipt/refund voucher, revised invoice) | not present on these files | A | Latent. | Low |

---

## Sheet `at` / `atadj` / `ata` / `atadja` (Tirupati has data; whole sheets A)

No code reads these (grep for `Gross Advance`, `atadj`, `11A` in the engine returns nothing). They are not in the mapping.

| Raw column | Example value (Tirupati FY25-26) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| `at`: Gross Advance Received, tax heads | Jun-25: Rs 34.06 cr advance at 18% (CGST 3.065 cr + SGST 3.065 cr = Rs 6.13 cr tax); Aug-25: Rs 11.94 cr | A | Table 11A(1)/(2) output tax on advances. In Jun-25 GSTR-3B 3.1(a) taxable = 45,339,111, **exactly equal to the GSTR-1 b2b taxable** (Rs 4.53 cr, CGST 40.8 lakh), so the Rs 6.13 cr advance tax is not in the 3B outward row the engine compares (the engine compares only invoice tabs against 3B and reports MATCH). Either 3B omits the liability, or advance tax is paid elsewhere (DRC-03/cash ledger); this has to be reconciled against liability register and cash ledger. Sec 13/Notification 66/2017 (no tax on goods advances) is also a question if the supply is goods. | **High** |
| `atadj`: Gross Advance Adjusted | Mar-26: Rs 14.23 cr (CGST 1.281 cr each); Aug-25 6.45 cr; Sep-25 3.19 cr; Oct-25 8.80 cr; Dec-25 5.96 cr; total Rs 38.63 cr adjusted | A | Table 11B: advance adjusted against invoices shown in Tables 4-7. Liability for those invoices is already paid on advance, so output tax for invoices that month is reduced by this adjustment. Net adjusted 38.63 cr vs received 46.00 cr leaves Rs 7.38 cr advance unadjusted (open at year end). Adjustment above cumulative advance received would be a flag. Reconciliation: GSTR-1 net liability = invoice tax + at tax - atadj tax, compare to 3B 3.1(a) and to ledger debits per month. Currently the Mar-26 b2b Rs 60.9 cr is compared to 3B without the Rs 14.2 cr adjustment. | **High** |
| `ata`, `atadja`: Financial Year, Original Month, Original POS, Rate, Gross Advance | empty | A | Amendments of earlier advances; same pairing as b2ba. Latent. | Low |

## Sheets `b2cla`, `expa`, `cdnura`, `b2csa` (all empty in every file)

| Sheet | Status | What it could tell us | Priority |
|---|---|---|---|
| b2cla (Revised POS, Rate, Taxable, IGST), expa (SB no, port), cdnura, b2csa (Original Month, Financial Year) | A (not in the mapping; the amendments module header even says it covers `b2csa/expa`, but only b2ba and cdnra are parsed) | Amendments to B2C large, exports, unregistered notes and B2CS. All empty here, so no finding today. b2csa matters because B2CS is the only B2C reporting: a later upward revision of B2CS months is an under-reporting admission (Sec 37, 39). | Low |

## Sheets `eco`, `ecoa`, `eco_9(5)`, `ecoa_9(5)` (empty; banner rows only)

| Raw column | Status | What it could tell us | Priority |
|---|---|---|---|
| GSTIN of E-Commerce Operator, Net value of supplies, tax heads | A | Table 14: supplies through e-commerce operators (Sec 52 TCS). Cross-check to TCS credit received in the electronic cash/credit ledger (`ElectronicCreditLedger` TCS entries) and GSTR-2A TCS section; Table 15 (9(5)): ECO-liable supplies (restaurant / transport / accommodation) for recipient/supplier GSTINs. Empty for all four taxpayers, so latent. Note the 23-24 merged file has only 3 banner rows for `eco` (Jan-24) and 2 for `ecoa` (Feb-24): months absent from banners also mean month coverage cannot be confirmed. | Low |

## Sheet `Read me`

| Item | Status | Note |
|---|---|---|
| GSTIN, Legal Name, ARN, ARN date, Tax Period | D | Used for self-GSTIN and filing-date/late-fee (`arn_dates_by_month`). Per-month banner ARN dates are canonicalised in `RETURNS`. |
| Trade Name, Date and Time of Generation, Financial Year | A | Trade vs legal name could be checked against BO profile name (identity mismatch); generation timestamp (`31/08/2026`) shows how stale the pull is (relevant to "amendment since" risk). Low. |
| 269-row data-entry instruction table | A | Static text; no data. |

---

## Top 5 opportunities

1. **Sign debit notes correctly (cdnr/cdnur `Note Type`).** Feb-24 BM FY23-24 critical outward mismatch of Rs 11.18 lakh is exactly 2 x the Rs 5.59 lakh debit notes; fixes false CRITICALs in `Comparison A` and in #4/#11. Code: `parse_gstr1` (`cn_*`), `_cdnr_rows_by_month`, `gstr1_adapter.cdnr_rows_by_month`.
2. **Use HSN Total Quantity and UQC.** The "known structural zero" statement is wrong for Bajrang (97% of taxable value has quantity); unlock #16/#44/#53 (unit price, band, drift) and units sold vs EWB/stock; remove the hard-coded INFO and R5 sentence.
3. **Read `at` / `atadj` (and `ata`/`atadja`).** Tirupati Jun-25 shows Rs 34.06 cr advances and Rs 6.13 cr tax in Table 11A that the engine ignores while 3B 3.1(a) equals invoice-only taxable; Rs 38.63 cr adjusted across the year. Needs a GSTR-1-total-liability (invoices + at - atadj) vs 3B vs ledger reconciliation.
4. **Table 13 beyond outward invoices.** Cancelled-rate anomaly (1.7% to 13.0%), credit-note serial coverage (1,628 live notes vs 995 reported, FY23-24), implied B2C invoice count and size, RCM self-invoice count vs 3B 3.1(d).
5. **Invoice/amendment date vs reporting period.** 61 BM FY23-24 invoices reported 1 to 5 months after their invoice date, plus `b2ba` Original Invoice date (prior-FY originals) for Sec 37(3)/interest u/s 50. Also: CDNUR to B2CL link (same-day full reversal of Rs 8.67 lakh B2CL invoice), and fixing the `str(None)` bug so the "B2B without IRN" check can actually fire.


---

# GSTR-3B - unused input data audit

Scope: raw 3B of Bajrang Motors (BM) FY23-24 (12 per-month zips -> xlsx) and FY24-25 (`GSTR3B_Merged (6).xlsx`), Jagnath (JG) FY23-24 (12 zips), Tirupati Cement (TC) FY25-26 (12 xlsx, newer layout). No `Canonical_GSTR3B_*.xlsx` exists in the repo or in `output/`; I read the adapter/mapping instead.

Status key: A = never read; B = read/canonicalised but no check uses it; C = display only; D = used in a check (listed only for contrast).

## 0. How 3B is consumed today (for contrast, status D)
Canonical layer (`gstr3b_adapter.py` + `gstr3b_mapping.json`) recognises 12 lines only: 3.1a, 3.1b, 3.1c, 3.1d, 3.1e, 4A3, 4A4, 4A5, 4B1, 4B2, 4C, 4D175 (the pre-Aug-22 line, which never exists in these files), plus the header block (GSTIN, year, period, ARN, ARN date, legal name, trade name, generated-on). Nothing else is in `LINES`; the spec says nothing for the rest of the form, and the META sheet has no "sections not read" notice for them (only `sheets_not_used`).

Bypass of the canonical layer: `gst_checks_flow.read_gstr3b_extra()` opens the RAW 3B workbook again (`pr.load_xlsx` + `_gstr3b_sheet_month`) to fetch 4A(1), 4A(2), 4A(4) and 4(D)(1)/(D)(2). So these five lines never pass through the adapter and are not in the canonical file at all (4A4 is in both). Side effects: (i) not covered by the E302/W303 issue log; (ii) for a quarterly 3B `_gstr3b_sheet_month` returns `labels[0]` (first month) while the adapter anchors on the LAST month, so extra would be keyed to a different month; (iii) a user-supplied canonical 3B file would yield all-zero extras with no warning.

D (used in a check): 3.1a (vs GSTR-1, e-way bill #8, ITC/liability ratio #14, effective rate #4); 3.1c/3.1e (check #1 vs GSTR-1 Table 8); 3.1d + 4A3 (RCM triangulation, #12, F8); 4A5 IGST/CGST/SGST (vs 2B, D-section); 4B1/4B2 (ITC roll-forward, D2 outlier, F2/F3); 4C (check #3 arithmetic); 4A1/4A2/4A4/4D1 (roll-forward, 4D(1) vs blocked-credit scan F4, G8 ISD vs 2A); ARN + ARN date (late fee, interest, #8/#10/#15); `month_fields` 4A5-IGST, 4B1, 4B2 (#18, #25).

## 1. Header block (rows 5-12)
| Raw row | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| GSTIN | 05AAAFB8782C1ZO | D (fallback self-GSTIN only) | Fine. Could assert every sheet's GSTIN equals the taxpayer GSTIN: a 3B of another registration (or an ISD/second-registration sheet) pasted into the merged file would currently be accepted silently (E-codes check only Year/Tax Period). | Low |
| Legal name / Trade name | BAJRANG MOTORS / M/s.BAJRANG MOTORS | B (RETURNS sheet; legal name is only a name fallback) | Compare with BO Profile legal/trade name and e-invoice seller name: name change or different trade name across years/sources is a profile-integrity flag. | Low |
| Date and Time of Generation | 31/08/2026, 16:02:00 | B (stored in RETURNS.generated_on, never read) | Only a data-freshness stamp: all BM FY23-24 sheets generated 31/08/2026, FY24-25 on 14/09/2026. Useful to warn "3B downloaded after later amendments" or mix of vintages in one run. | Low |
| ARN (number itself) | AA050423149295B | D for existence; the CONTENT is unused | Format AA + state(05) + MMYY + serial + check char. In all 48 sheets MMYY equals the tax period. Check "ARN MMYY = tax period" and "serial/ARN dates monotonic" to catch a sheet relabelled/copied from another month (the tool identifies the month only from the Tax Period text). Also ARNs of the Credit Reversal and Re-claimed Statement (BM 24-25 csv, e.g. Aug-24 AA050824183714Z dated 20/09/2024) are an independent source for ARN+date: cross-check against 3B header. | Med |
| Date of ARN | 21/02/2024 (Jan-24 3B, due 20th) | D | Used for late fee/interest. Missed use: compare with the ARN dates in the Credit Reversal statement and the ledgers (see 5.1). | Low |

## 2. Table 3.1 outward + RCM (rows 17-21)
| Raw row/column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| 3.1(b) zero-rated, all 5 columns | 0 in all 48 sheets | C (only quoted in a check #1 text) | No exports in these files. For any exporter: 3.1(b) taxable + IGST vs GSTR-1 Table 6A/6B (exports, SEZ), e-invoice export docs, shipping-bill/EGM; IGST on exports with payment (refund Sec 54) and Rule 96 mismatch. Zero here but the compare is not built. | Med (for exporters) |
| 3.1(a) IGST/CGST/SGST/Cess columns as separate heads | BM Apr-23 cess 7,868,187.99 on 93.5m taxable | D for taxable + all heads in A-section vs GSTR-1; but the 3.1(a) CESS is not used in check #4 (rate), #14 (ratio) | BM carries heavy compensation cess (cess = ~8-10% of taxable in BM). Cess is only compared to GSTR-1 cess; a cess RATE check (cess/taxable by HSN: autos 1-22%) vs GSTR-1 HSN summary and e-invoice is not done. | Low |
| 3.1(d) first column (taxable value of RCM inward) | BM Nov-24 71,000; TC Jan-26 1,063,092 | D (tax portion), the TAXABLE VALUE is B | `rcm_liability_igst` in `month_fields` stores the taxable value under an IGST name and no check uses it. Compare 3.1(d) taxable value to RCM-flagged lines in 2B/2A (GTA, legal, security, unregistered purchases Sec 9(4)) and to ledger expense heads (freight / director remuneration). TC 12/12 months nonzero RCM with ~11% effective rate - consistent with GTA 5%+... a freight vs RCM sanity check is possible. | Med |
| 3.1(e) Non-GST outward | 0 | D | ok | - |
| 3.1(c) Nil/exempt | 0 | D (#1) | ok | - |

## 3. Table 3.1.1 (Sec 9(5) supplies via e-commerce operators)
Whole sub-table unused. Rows (i) "ECO pays tax under 9(5)" and (ii) "supplies made through ECO on which ECO pays tax" - 5 columns each; all 0 in the four taxpayers (A, canonical LINES has no code for them, no check reads them, mapping.json has no entry).
| Raw row | Example | Status | Use | Priority |
|---|---|---|---|---|
| 3.1.1 (i), (ii) | 0 | A | (ii) must tie to GSTR-1 Table 15 (supplies through ECO u/s 9(5)) and to TCS received in GSTR-2A/2B (Table 9 of 2A: TCS credit) and cash ledger credit; a non-zero (ii) with no Table 15 = outward-undeclared; (i) is only for ECOs. Also tells us the taxpayer sells via a marketplace (e.g. passenger transport, restaurant services). No real data triggers it, so impact is low now. | Low |
| 6.1(B) heading in newer layout "Reverse charge **and supplies made u/s 9(5)**" | TC Dec-25 | A | Layout variant: the 6.1(B) row now also holds 9(5) liability; any 6.1 reader must handle it. | Low |

## 4. Table 3.2 inter-state supplies to unregistered / composition / UIN (rows 33-35)
Whole table unused: A (no canonical code, not in LINES, no engine grep hit for "unregistered persons"/3.2/UIN).
| Raw row/column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Supplies to Unregistered Persons: Total taxable value + IGST | BM Jan-24 15,859,064.71 / 4,410,932.16; BM Apr-23 4,707,236.49 / 1,299,765.35; TC Apr-25 1,016,949.15 / 183,050.85 | A | The IGST must be <= 3.1(a) IGST and the taxable value must be matched to GSTR-1 B2CL (invoice >Rs 1 lakh inter-state B2C) + B2CS rows with POS != supplier state. In BM the 3.2 B2C inter-state share is the only place 3B declares the place-of-supply split; a gap vs GSTR-1 is a PoS/IGST-vs-CGST mis-declaration (wrong head -> revenue/refund issue) or understated B2C. Note BM 3.1(a) IGST (Apr-23 1.99m) is larger than 3.2 IGST (1.30m): the remainder (~0.69m) should be B2B inter-state - verifiable against GSTR-1 B2B IGST. | High |
| Supplies to Composition persons | BM Feb-24 5,932.33 / 1,068.15 | A | A cross-check of GSTR-1 B2B rows whose recipient GSTIN is a composition dealer (BO profile / 2A not applicable). Also: composition customers cannot take ITC - confirm no ITC-eligible tax invoice reasoning. | Low |
| UIN holders | 0 | A | Cross-check with GSTR-1 table 4 UIN recipients / refund claims by embassies. | Low |
| 3.2 row sum vs 3.1(a) | - | A | Arithmetic integrity: sum(3.2 IGST) <= 3.1(a)+3.1.1(i) IGST; violations show a manual edit/ wrong table fill. | Low |

## 5. Table 4 ITC
| Raw row/column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| 4A(1) Import of goods | 0 | D via `read_gstr3b_extra` (bypasses adapter) | Zero in all four. For importers, compare with Bill of Entry IGST (ICEGATE) / 8A. No BoE source in the repo; flag only. Also not in canonical LINES. | Low |
| 4A(2) Import of services | 0 | D (as above) | Compare with RCM on import of services (3.1(d) is for domestic RCM, import-of-service tax shows in 3.1(d) too) - the tool never checks that 4A(2) matches 3.1(d). | Low |
| 4A(3) / 4A(4) | TC Jan-26 4A3 66,353; 4A4 0 | D | ok. NOTE check #3 (4C arithmetic) uses only 4A5+4A3-4B1-4B2 and ignores 4A1/4A2/4A4: a month with ISD credit or imports will fake a "4C mismatch". Latent bug, zero in this data. | Low |
| 4A(5) **Cess column** | BM Apr-24 cess ITC 17,144,569.31 vs 3.1(a) cess liability 17,452,667.41 | B (value is parsed; compared only for IGST/CGST/SGST in `gst_parsers_returns` D-section and ratio #14) | Compensation-cess ITC is ~98% of cess liability in BM and 4B(2) cess reversal 966,216 (Sep-24); cess is paid through the cess ITC 7.87m/9.58m (6.1) with partial cash 620,632 (Feb-24). Compare 4A5 cess to 2B cess (2B ITC columns exist) for each month. Only 3 heads are compared today, so the biggest-rupee head (Rs ~17m/month ITC) in BM is un-reconciled. Also cess credit may be used only for cess liability - a check that cess ITC <= cess liability+openings. | High |
| 4B(1) CGST/SGST/Cess columns | TC Oct-25 [IGST 0, CGST 25,928, SGST 25,928, Cess 0]; TC Dec-25 [11,277; 8,150; 8,150] | D for IGST only (`month_fields` takes `vals[0]`) | #25 "exempt turnover vs Rule 42 reversal" counts a month as "zero reversal" when the IGST value is 0, so TC Oct-25 (25,928+25,928 reversed) is falsely counted as zero. Same for #18 (compares IGST reversals to IGST ITC only). Fix: sum all heads. Useful result: the Rule 42/43 ratio check could be done on the actual total. | High |
| 4B(2) all heads | BM Sep-24 [0, 676,351, 676,351, 966,216]; BM Dec-23 [199,555.44 IGST]; TC 11/12 months nonzero | D for IGST/CGST/SGST (outlier), Cess = B | Reversal under 4B(2) (non-payment to supplier in 180 days Rule 37, Rule 39 ISD etc.) is "reclaimable" -> must show later in 4D(1). BM Sep-24 reversal 676k is reclaimed in the portal statement? The statement csv for 24-25 shows closing 0/676,351/676,351/966,216 still carried; the tool says "reclaim not trackable" (F2a text). Use the Credit Reversal and Re-claimed Statement csv (exists for BM 24-25) with 4B2/4D1 to close the loop. | Med |
| 4C Net ITC | BM Apr-23 [29,441,622.17; 2,358,448.93; ...] | D (arithmetic) | Also compare to Electronic Credit Ledger credits: done in F3 (D). | - |
| 4(D)(1) ITC reclaimed (earlier 4B2) | BM Aug-24 [12,099,914.44; 148,790; 148,790; 1,806,103] | D (4D1 vs blocked scan F4) **but semantically misused**: 4(D)(1) is "ITC reclaimed which was reversed under 4(B)(2)", not sec 17(5). F4 compares the blocked-credit scan with 4D(1), and the code label says "4D(1) sec 17(5)" (that wording was the pre-Aug-22 form). | In the new form 4D(1) is a RECLAIM (an additional credit, not an ineligible block). The BM Aug-24 12.1m IGST reclaim is the single biggest 4D item in the data and ties to the opening balance of the Credit Reversal statement (12,099,914.44 opening, reclaimed on Aug-24 filing). It needs: (a) the matching earlier 4B(2) reversal (FY23-24 Dec-23 4B2 IGST 199,555 does not explain 12.1m, so the original reversal sits in an earlier FY: Rule 37/ Sec 16(2) 180-day timing: check payment evidence in ledger/9C), (b) FY24-25 reclaim-by-FY-end rules. It also inflates `claimed` in ITC roll-forward only if included; today D1 is not in claimed/avail but is used for the F4 blocked comparison. F4 gap is therefore reading the wrong semantic. | High |
| 4(D)(2) Ineligible ITC u/s 16(4) & PoS restricted | BM May-23 [0; 132,584.92; 132,584.92; 189,407.03]; Apr-23 [0; 510.07; 510.07; 0]; TC 6/12 months (84.39 etc.) | B for display; D only in F7 month-over-month outlier | This is a self-declared ineligible ITC (time-barred under 16(4) or PoS = wrong state / Sec 17(5) type). Cross-check against 2B ITC-availability = No rows and reason codes (the code at gst_checks_flow ~275 already has the "POS differs" reason breakdown from 2B/8A) - D2 total should explain the "2B ITC not available" amount; mismatches mean taxpayer under-reports ineligible ITC or claims ITC outside time limit. Also 16(4) amounts show invoices older than the 30-Nov deadline are being claimed. | High |
| 4(A)(5) etc. - per-sheet "Net ITC" cess value | - | B | (see cess row) | - |

## 6. Table 5 inward exempt / nil / non-GST (rows 57-58)
Whole table unused: A. No canonical code, no engine reference.
| Raw row/column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| "From a supplier under composition scheme, Exempt, Nil rated supply" - Inter-State, Intra-State | BM Apr-23 41,575 / 998,019; BM Mar-24 intra 27,036; BM Jul-24 29,769, Nov-24 27,432; JG/TC 0 | A | Cross-check against inward e-way bills (exempt/nil HSN or composition suppliers) and 2B nil/exempt supply tables (2B "Nil/exempt, Non-GST" ISD/ B2B rate 0) - supplier type from the Trade/GSTIN master. A taxpayer that buys from composition/exempt suppliers has exempt inward input: part of the Rule 42 context (not a reversal trigger), and it tests whether purchases from composition dealers are properly excluded from ITC (ITC claimed on bills of supply). Under-declared Table 5 vs EWB/GSTR-9 Table 8 (the inward-exempt line in GSTR-9 Table 5/6 area) shows the return is not reflecting actual purchase mix. | Med |
| Non GST supply | 0 | A | Similar for fuel (petroleum), liquor purchases. | Low |

## 7. Table 5.1 interest and late fee (rows 63-65; TC adds "Late fee (GSTR-3B/3BQ)" and "Break-up of late fee levied for GSTR-1")
Whole table unused: A (canonical adapter has no code; `gst_checks_forensic.compute_interest` and `late_fee` compute their own values from ARN date and cash-paid tax instead).
| Raw row/column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Interest Paid (IGST/CGST/SGST/Cess) | BM Jan-24 [0,0,1,569.39,0]; BM Feb-24 [0,0.07,0.07,0] | A | THE portal-recorded interest actually paid on the previous period's late payment: BM Dec-23 3B filed 21/01/2024 (1 day late) -> 1,569.39 interest and 25+25 late fee in the Jan-24 sheet. The tool's "Sec 50 Interest" (18% x cash-paid tax x days) and "late fee" in `Filing Compliance & Late Fee` are approximations; the portal's actual figures validate/falsify the due-date logic. Also shows asymmetry (interest booked all under SGST - 1,569.39 CGST 0) = wrong head payment, plus interest on delayed payment where ARN is on time but the cash payment was deferred. | High |
| System computed Interest | "-" (dash) in all 48 sheets | A | Would be the portal's own computed interest vs paid: dash everywhere; with Interest Paid = 0 and a late filing it indicates interest not recomputed (under-payment). Cannot be tested with this data but column should be recorded. | Low |
| Late fee (CGST/SGST) | BM Jan, Feb, Mar-24, Nov-23: 25/25 | A | 25+25 = Rs 50 late fee for a 1-day delay - validates the filing dates (Dec-23, Jan-24, Feb-24, Oct-23 returns filed 21st vs due 20th). It lets the tool prove which return was late without relying on the due-date table, and catch the CBIC extension notifications the table may not know. In older (BM) layout one cell mixes GSTR-3B and GSTR-1 fee; newer TC layout splits "Late fee (GSTR-3B/3BQ)" and "Break-up of late fee levied for GSTR-1" (period + Central + State columns, empty in TC). | High |
| Break-up of late fee levied for GSTR-1 (TC new layout) | present in Dec-25..Mar-26 sheets, empty | A | Directly names the GSTR-1 period for which the late fee was levied and the amount: independent evidence of a late GSTR-1 without GSTR-1 ARN dates. | Med |

## 8. Table 6.1 payment of tax (rows 70-81 old; 74-85 new layout)
Whole table unused: A. Columns: Total tax payable | Tax paid through ITC (IGST, CGST, SGST, Cess) | Tax paid in cash | Interest paid in cash | Late fee paid in cash (+ in newer layout "Adjustment of negative liability of previous tax period" and "Net Tax payable"), rows Integrated/Central/State/Cess, (A) Other than reverse charge and (B) Reverse charge.
| Raw row/column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| ITC-vs-cash split per head | TC Dec-25 CGST 16,842,932 = ITC IGST 3,739,686 + ITC CGST 10,076,888 + cash 3,026,358; BM Feb-24 Cess 19,449,742 = ITC 18,829,110 + cash 620,632 | A | (1) Rule 88A utilisation order (IGST ITC must be exhausted first, against IGST then CGST then SGST; CGST ITC cannot pay SGST): validate in every month from the 3 ITC columns. (2) Cash share of tax: the BO Profile `tax_paid_in_cash` is only annual (F10); month-wise it detects the "pay nothing in cash" pattern behind #14 (ITC>95%). (3) Compare 6.1 cash with Electronic Cash Ledger debits for that period (ledger debits by return period exist in the ledgers read by the tool) - mismatch = payment made against another liability or wrong period. (4) 6.1(B) RCM must be paid wholly in cash (Sec 49(4)); BM/TC (B) rows are cash only, consistent; any ITC there would be a clear violation. | High |
| Total tax payable vs 3.1(a)+3.1(d) tax | BM Apr-23 CGST payable 11,641,302 = 3.1(a) 11,641,302.28; (B) 1,980 | A | Internal consistency (6.1 must equal 3.1 tax); a user edit or negative-liability adjustment shows here. The "Adjustment of negative liability of previous tax period" column in newer sheets (Dec-25 to Mar-26 sheets) = credit/debit note netting (negative liability carried forward) - shows when a month's outward is net-negative. | Med |
| Interest paid in cash / Late fee paid in cash columns | BM Jan-24 Central tax row [13,880,730; 13,880,730; 0; 0; 0; 0; 1,569; 25] | A | Actual cash interest/late fee at payment time - duplicates 5.1 and is the true cash leakage; sum per FY = cost of non-compliance, and "interest paid with no late return" = delayed payment. | Med |
| Tax paid in cash = 0 for entire FY | BM 24-25 cash ~0 in most heads; TC Jul-25 total payable only 378,622 | A | Months where 3.1(a) liability is large but cash paid is 0 and ITC is exactly equal: "ITC covers 100%" red flag already used (#14) at a different granularity; 6.1 gives it per head. | Med |

## 9. "Breakup of tax liability declared (for interest computation)" (bottom of sheet)
| Raw row | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Period / Integrated / Central / State / Cess per liability period | BM Apr-23: April 2023 [1,989,211; 11,643,282; 11,643,282; 7,868,188] | A | Equals 3.1(a)+(d) liability of the period (checked on BM Apr-23: 11,641,302.28 + 1,980). If a row carries an EARLIER period it is liability declared late (amendments / missed returns) - drives Sec 50 interest per period. All 48 sheets contain a single row for the same period, so no late-declared liability in this data. Check "row period = return period" and "amount = 3.1+6.1" per sheet. | Low |

## 10. Whole-sheet / structure observations
- Sheet-level: every month sheet is read; sheets of other types would be `I401`. No unread sheet types in these files (merged BM FY24-25: 12 month sheets; TC 12 per-month workbooks). **Rows below row 21 of the month sheet are NOT read at all** except 4A/4B/4C/4D175: that is the totality of Sections 3.1.1, 3.2, 4A1/4A2/4D (via raw re-read), 5, 5.1, 6.1 and the breakup table.
- Layout drift the mapping would need: TC sheets Dec-25..Mar-26 have 92 rows with new 6.1 columns ("Adjustment of negative liability", "Net Tax payable") and the late-fee break-up block; Apr-Nov have 87 rows. Row positions differ - good that 3B rules are label-based, but any future 6.1/5.1 reader must be too.
- The canonical `LINES` sheet stores values by column name for 12 codes. Nothing in the `ISSUES` log tells the user which sections of the form were ignored.

## Top 5 opportunities
1. **Use 5.1 + 6.1 interest/late-fee as ground truth for Sec 47/50** (BM Jan-24: interest 1,569.39 and fee 25+25 for the Dec-23 return filed 21/01/2024): validate the tool's due-date table and late-fee/interest estimates; flag where the portal levied nothing although the ARN is late (system "computed interest" is "-" everywhere).
2. **Fix and widen the Rule 42/43 and 4B checks**: #18 and #25 look only at the IGST column (`vals[0]`), so TC Oct-25 4B(1) CGST/SGST 25,928 is counted as "zero reversal"; add CGST/SGST/Cess to #18/#25 and do the cess compare.
3. **Reconcile 4(A)(5) CESS and 4(B)(2) CESS to 2B** - BM has Rs 17m/month of cess ITC against Rs 17.5m cess liability (98%) and nothing checks it; and cap cess ITC use against cess liability in 6.1.
4. **4(D)(1) reclaim and 4(D)(2) ineligible**: stop treating 4D(1) as "sec 17(5)" (BM Aug-24 12.1m IGST reclaim) - tie it to the Credit Reversal and Re-claimed Statement csv and to prior 4B(2); tie 4D(2) (BM May-23 132.6k+189k cess, TC 6 months) to 2B "ITC not available" / 16(4) invoices.
5. **Table 3.2 and 6.1 payment mechanics**: 3.2 inter-state B2C vs GSTR-1 B2CL/B2CS PoS split (BM Jan-24 15.86m), and 6.1 Rule 88A utilisation order + cash vs Electronic Cash Ledger per period; plus move `read_gstr3b_extra` into the canonical adapter so 4A1/4A2/4D and all of the above are in `LINES` with issue logging and quarterly anchoring.


---

# Unused-input audit: GSTR-2B and GSTR-2A

Scope: every sheet and column of the GSTR-2B and GSTR-2A files in `input_files_bajrang_motors` (FY23-24 monthly 2B, 2A zips), `input_files_bajrang_motors_24_25` (merged 2B, merged 2A), `input_files_jagnath` (2BQ quarterly), `input_files_tirupati_cement` (full monthly 2B with IMS sheets, monthly 2A), plus `output/GSTR2B_Merged.xlsx`, `output/R2A_Merged.xlsx`.
Engine reviewed: `main gst tool/` (gstr2b_adapter.py, gstr2a_adapter.py, mappings, gst_parsers_returns.py, gst_parsers_dept.py, gst_checks_*.py, gst_blocked_credit.py, gst_machinery_scan.py, master_build.py).

Status key: A = never read (sheet unread or column only in `extra`/dropped); B = canonicalised/parsed but no check reads it; C = display-only (copied into a detail table, no logic); D = used in a check (mentioned briefly only).

## 0. How the data is consumed (evidence)

**GSTR-2B**
- `gstr2b_mapping.json` knows only five sheets: B2B, B2B-CDNR, B2BA, B2B-CDNRA, ITC Available. Everything else is written to META `unread_sheets` and logged as INFO I401 (confirmed by a live run on Tirupati Apr-25: "Sheets not used: B2B(Rejected), B2B-CDNR(Rejected), B2B-CDNRA(Rejected), B2BA(Rejected), ECO, ECO(Rejected), ECOA, ECOA(Rejected), IMPG, IMPGA, IMPGSEZ, IMPGSEZA, ISD, ISDA, ITC Rejected, ITC not available").
- Row dict handed to the engine (`gstr2b_adapter._b2b_dict/_cdnr_dict`) has: gstin, supplier, invno/note, invtype, date, invval/noteval, pos, rcm, rate, taxable, igst, cgst, sgst, cess, itc_avail, itc_avail_reason, einv_source, irn, filed_period, via_amendment (+ original_invno). `extra` (e.g. GSTR-1 Filing Date, IRN Date, Applicable % of Tax Rate, "Whether ITC to be reduced", "Amount declared by taxpayer", Remarks) is stored in the canonical workbook only and is not in the dict at all.
- Fields actually driving logic (D): gstin, invno/note, date, taxable, igst/cgst/sgst/cess, itc_avail (YES/NO/other split in `read_2b_invoice_level`, gst_checks_flow.py:303-335), ntype (credit vs debit sign), rcm (rcm_itc, flow:311), einv_source+irn (only check #28, gst_checks_monthly.py:1592), supplier (blocked-credit keyword screen, machinery name scan, top-10 supplier F12), ITC Available "All other", "RCM" and "Part B Others" rows IGST/CGST/SGST (D-section comparisons in gst_parsers_returns.py:970-1016, 909-930; F7a flow:1564; R14 master_build:2718).
- `invtype`, `pos`, `rate`, `invval`, `noteval`, `supplytype`, `filed_period`, `via_amendment`, `itc_avail_reason` (outside B3 / "ITC=No" tables): display columns or unread (see per-sheet tables). `pos` is used only as display in 2B-side tables (monthly:1043, 1063) and in blocked-credit output; the POS-vs-state logic in hsn_fraud (1029, 1168) runs on GSTR-1, not 2B.
- ITC Available sheet: only the "All other ITC" (4A5), "Inward supplies liable for reverse charge" (3.1(d)/4A3) and "Part B Others" (CN) rows are extracted (`_summary_from_block`, gst_parsers_returns.py:1479-1512, label matching). All other lines of that sheet are not read.

**GSTR-2A**
- `parse_r2a_excel` / gstr2a_adapter read B2B, B2BA, CDNR, CDNRA, ISD only. ECO, ECOA, ISDA, TDS, TDSA, TCS, IMPG, IMPG SEZ are explicitly "OUT OF SCOPE ... deliberately not parsed" (gst_parsers_dept.py:1234-1238). Unlike 2B, the 2A adapter does NOT record these as `unread_sheets`; it only logs I702 for sheets absent from the file, so unread-but-populated sheets (TDS) are silent.
- Checks read only `r2a["b2b"]`, `["b2ba"]`, `["isd"]` (flow.py 2599, 2660, 2878, 2948-2967, 3017, 3052, 3137, 3170, 3208, 3414). `r2a["cdnr"]` and `r2a["cdnra"]` are parsed and never consumed by any check (grep: no `r2a.get("cdnr"` / `["cdnra"]` anywhere in gst_checks_*, master_build).
- B2B row fields used (D): gstin, supplier, invno, invdate, invtype (join key only), invval, taxable, igst/cgst/sgst/cess, rcm, cancel_date (G10), g1_filing_date (G9 deadline test). Parsed but unread: g1_status, g3b_status, g1_filing_period, amend_type/amend_period (B2B), source, irn, irn_date, pos (G7 uses the supplier GSTIN state code instead), and Rate (not even parsed; only `extra` in the canonical file, and only the "-Total" rows are kept so the per-rate lines are discarded).

## 1. GSTR-2B sheet by sheet

### 1.1 Whole-sheet items (status A unless noted)

| Sheet | Present in (real data) | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|---|
| ITC Rejected (summary) | Tirupati (all 12 months; non-zero in 2) | Jul-25 IGST 19,57,145.40 (B2B invoices IMS), and Part B credit notes 19,57,145.40; May-25 56,711.60 | A | IMS rejection is the recipient's own act: ITC deliberately not taken. Control total for the B2B(Rejected) rows; tells the scrutiny officer which ITC the taxpayer refused and when; cross-check 3B 4A5 does not include it (if 4A5 > 2B "ITC Available", rejected ITC may have been claimed anyway). Also IMS acceptance behaviour per supplier. | High |
| B2B(Rejected) | Tirupati May-25 (1 row), Jul-25 (4 rows) | 07BLSPS2793A1ZP OJAS INFRA, 25-26/2, 01/07/2025, taxable 27,88,590, IGST 5,01,946.20, Remarks "No remarks available" | A | Invoices the taxpayer rejected in IMS. 4 rows total 32.9L+29.0L+38.6L+27.7L invoice value. Each has an identical-value rejected credit note in B2B-CDNR(Rejected) (CN/25-26/1..4: 32,90,536 / 29,04,133 / 38,61,657 / 27,73,848, same dates). Invoice then equal credit note, both rejected, is a classic round-trip/fake-billing or supplier-error footprint; Sec 16(2)(aa)/Rule 36(4) and Sec 34. Never surfaced anywhere in the engine. | High |
| B2B-CDNR(Rejected) | Tirupati Aug-25 (4 rows) | CN/25-26/1 .. /4, Credit Note, filing date 20/08/2025 | A | Pair with B2B(Rejected) to detect invoice/credit-note wash trades; credit-note rejection means the taxpayer refuses ITC reduction, so tie to 3B 4(B)(2) (no reversal expected). | High |
| B2BA(Rejected), B2B-CDNRA(Rejected), ECO(Rejected), ECOA(Rejected) | Tirupati (empty in all files) | header only | A | Rejected amendments / ECO documents. Empty here, but any non-empty row is IMS evidence. | Low |
| ITC not available (summary) | Tirupati, Jagnath (and Bajrang monthly when full export) | Tirupati: B2B invoices CGST+SGST 144+144 in Apr-25; 6 months non-zero, total 2,329.56 | A | Official portal total of ITC the supplier-side rules say is blocked (4(D)(2)): independent control total for the row-level "ITC=No" test B3; also the only source of the Part A II/III/Part B ISD & RCM not-available lines. Cross-check to 3B Table 4(D)(1)/(2) (not parsed by the engine at all: no `4D` in gst_parsers_returns) and to 4(A)(5) overclaim. | Med |
| ECO | Tirupati (empty), Jagnath (empty) | header "Documents reported by ECO on which ECO is liable to pay tax u/s 9(5)" | A | Supplies on which ECO (Zomato/Amazon-style, 9(5): restaurants, passenger transport, housekeeping) pays tax: recipient ITC rules differ; also cross-check to GSTR-1 table 14/15 on the supplier side. Empty in all samples, so no immediate finding, but a populated sheet would carry ITC that sits outside 4A5. | Low |
| ECOA | same | header only | A | amendments of ECO docs | Low |
| ISD | Jagnath, Tirupati (empty in all samples) | header: GSTIN of ISD, ISD Document type, Document number/date, Original invoice number/date, Input tax distribution, GSTR-6 period/filing date, Eligibility of ITC | A | Cross-check to 3B 4A(4) (ISD credit). Note the engine ALREADY does ISD vs 4A(4) using 2A's ISD sheet (G8), so 2B ISD would be the legally binding side (2B is the Sec 38 statement). Eligibility of ITC "No" lines = cross-charge-ineligible. | Med |
| ISDA | same (empty) | header only | A | ISD amendments | Low |
| IMPG | Tirupati, Jagnath (empty) | header: Icegate Ref Date, Port Code, Bill of Entry no./date, taxable value, IGST, Cess | A | Import ITC (3B 4A(1)) cross-check against BoE; the only 2B source of import ITC. Also should tie to e-way bills with import flags and to Customs. Empty in all four samples; for an importer it is the whole import ITC. | Med |
| IMPGA | same | header only | A | BoE amendments | Low |
| IMPGSEZ | same | header: GSTIN of supplier, trade name, BoE | A | SEZ import ITC (3B 4A(1)/4A(2)): supplier GSTIN of an SEZ unit. | Low |
| IMPGSEZA | same | header only | A | | Low |
| Read me | all full exports | Legal name, GSTIN, FY, "Date of generation" | B | GSTIN/FY/period used by the adapter (`_period_from_readme_or_filename`, recipient GSTIN). Legal name and generation date are not carried to META. Generation date matters: 2B is a static snapshot; a download made months later (Bajrang files generated 31/08/2026, Tirupati 04/08/2026) can contain amendments (B2BA) not present at claim time. | Low |
| ITC Available (all lines other than 4A5 / RCM / CN totals) | all full exports | see 1.3 | A | see 1.3 | Med |

### 1.2 B2B sheet (used heavily; columns not or only partly used)

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| GSTR-1/1A/IFF/GSTR-5 Filing Date | `13/05/2024` (Bajrang 24-25 row 8); Tirupati `10/02/2026` | A (no mapping entry; goes to `extra` only) | The supplier's GSTR-1 filing date. 2B already tells us the ITC is available; the date tells us supplier filing discipline: in Bajrang 24-25, 263 of 15,863 rows (tax Rs 77.5L) were filed more than 2 days after the 11th of the following month, max lag 57 days; Tirupati 994 of 13,798 rows (Rs 87.0L), max lag 109 days. Use: (1) supplier-reliability score per vendor, (2) Sec 16(2)(c)/16(4) and Rule 36(4) tracing (the engine's G9 does this only from 2A, which most FYs do not have: the 2B file also carries it), (3) lag between invoice date and filing date to find invoices pushed into later periods. No other source gives this when 2A is absent. | High |
| GSTR-1/1A/IFF/GSTR-5 Period (`filed_period`) | `Apr'24`, `Jul'24` | B (canonicalised, reference-only; `filed_period` in the dict, never read; code comment "reference data") | Compare to invoice date month: invoices reported by the supplier two or more periods after the supply (delayed reporting) and invoices dated before the FY but arriving in this FY (Bajrang 24-25: 41 rows, Rs 3.33L tax; Tirupati: 105 rows, Rs 16.36L tax, of which 82 rows Rs 13.18L in Apr-25). A Sec 16(4) deadline test (invoice-date FY + 30 Nov) on 2B itself does not exist; no breach found in these four samples (0 rows in statement months after 30-Nov), but the control is absent. | Med |
| IRN (`irn`) | `b9dee16cded...` (11,557 of 15,863 rows have one) | C/D (only used in #28 and only for invoices matched to an inward EWB, monthly:1587-1599) | The all-register view is missing: 43 Bajrang suppliers issue some invoices with IRN and some without (2,098 no-IRN rows). M&M (27AAACM3025E1ZZ) alone has 1,206 no-IRN rows with tax Rs 8.85 crore; all 43 together Rs 12.36 crore. A supplier above the e-invoice threshold issuing invoices without IRN means the document is not a valid tax invoice (Rule 48(5)): ITC risk. Tirupati: 37 mixed suppliers, 269 rows, Rs 41.2L. #28 reports only INFO and only on EWB-matched documents. | High |
| IRN Date | `15/04/2024` | A (`extra`) | IRN date vs invoice date: lag in Bajrang 24-25 up to 35 days, 3 rows over 30 days (Rule 48(4)/GSTN 30-day reporting limit for AATO >= 100 Cr). Also IRN date vs supplier filing period (backdating). Zero rows with IRN date before invoice date. Weak in these samples but absent from the engine. | Low |
| Source | `E-Invoice` (11,557 rows), blank (4,306) | C/D (same #28 use) | Blank "Source" = supplier reported without e-invoice. Use as a per-supplier e-invoice adoption flag in the vendor risk table. | Med |
| Invoice type (`invtype`) | `Regular` in 100% of rows (also SEZ/DE possible) | B for 2B (only used as the 2A-vs-2B join key, flow:2615, and printed) | SEZ-with/without-payment and deemed-export invoices change ITC eligibility; all-Regular in these samples, so no signal yet, but a non-Regular row is never flagged. | Low |
| Place of supply (`pos`) | `Uttarakhand` (15,855), `Uttar Pradesh`/`Maharashtra`/`Himachal Pradesh`/`Tamil Nadu` (8 rows) | C | The 8 non-home POS rows are exactly the 8 ITC=No rows ("POS and supplier state are same but recipient state is different"); engine prints them in B3. Could add: recipient-state vs POS vs tax head (IGST on same-state supply) the way G7 does for 2A; B2B 2B never gets that test. | Low |
| Supply Attract Reverse Charge (`rcm`) | `Yes` (Tirupati 95 rows, Rs 66,752 tax; Bajrang 1 row Rs 6) | D (rcm_itc) but only inside the RCM triangulation | OK as is; but the RCM list itself (supplier, invoice, value) is not shown as a table in 2B terms (G6 shows RCM only for 2A). | Low |
| Applicable % of Tax Rate | `100%` in every row | A (`extra`) | Always 100% in these samples. Non-100% means a concessional/partly-exempt rate (e.g. notified old-rate goods); signals reduced-rate purchases. Zero signal here. | Low |
| Rate(%) | absent in Bajrang 24-25 and Tirupati monthly (portal dropped it); present in Jagnath `18`, `28`, `5` and Bajrang 23-24 | C (printed in EWB-in detail; never tested). In the new layout `rate` is blank/0.0 for every row | Derive rate = (IGST+CGST+SGST)/taxable when the column is missing. Rate mix per supplier: purchase at 28% / sales at 18% (inverted duty), or 5% vs 18% mismatch vs GSTR-1 HSN rate. Nothing in the engine reads a 2B rate. | Med |
| Invoice Value(₹) (`invval`) | `180009` | C (used in "g2b_zero" listing and display, plus agg in key of 2A-vs-2B in flow:3436, where it is overwritten per rate-line) | Invoice value vs taxable+tax: 52 of 15,863 Bajrang rows are internally inconsistent per the adapter's own `arith_check` (W305), but `arith_check` is "not used by any check" per spec 4.4. Use as a supplier data-quality indicator. | Low |
| Reason | `POS and supplier state are same but recipient state is different` (8 / 11 rows) | C (B3 text, flow tables) | Fine; only one reason value ever appears in these files. | Low |
| Trade/Legal name (`supplier`) | `M/S M S ENTERPRISES` | D (name screen) | Ok. Could be matched with BO Profile risk lists; not done. | Low |
| GSTIN of supplier characters 3-12 (PAN) | `05ABOPK4335R1Z2` | D (as key only) | PAN 4th char P = individual/proprietor. In Bajrang 24-25 suppliers like RAKHI CHOUHAN (05AGGPC4967Q1ZW, P) supply Rs 4.16L tax in 2A with GSTR-3B not filed; not exploited in 2B either. | Med |

### 1.3 ITC Available (summary) sheet

| Raw row / column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Part A I "All other ITC" 4A5 (IGST/CGST/SGST/Cess) | Tirupati Apr-25 IGST 2,41,11,710.06 | D | used (F7a, R14, D section) | - |
| "Details" lines inside Part A I: `B2B - Invoices (IMS)`, `B2B - Debit notes (IMS)`, `ECO - Documents (IMS)`, `B2B - Invoices (Amendment) (IMS)` | Tirupati Apr-25: Invoices IGST 2,40,33,736.09; Amendment 77,973.97 IGST | A | The amendment line separates ITC coming from B2BA (supplier amended an earlier invoice) from ordinary invoices: 7 of 12 Tirupati months carry amendment ITC, FY total Rs 13.42L; the Apr-25 amount is the one B2BA row where the amendment type was "GSTIN Amended" (invoice 9, RAHUL TOMER, 25-03-2025, Rs 5.11L, original period Mar-25, amended GSTIN to the taxpayer). Amendments that re-point an invoice to the taxpayer's GSTIN are a classic ITC-shifting red flag. Not shown anywhere. | High |
| Part A II "Inward Supplies from ISD" 4A4 | zero in every sample | A | cross-check 3B 4A(4) (see ISD sheet). | Med |
| Part A III RCM (3.1(d)) incl. `Details` lines | Tirupati Apr-25: IGST 128.75, CGST 600, SGST 600 | D (ITC_rcm_* used in RCM triangulation); cess not used | OK. | - |
| Part A IV "Import of Goods" 4A1 (IMPG, IMPGSEZ and amendments) | zero | A | Import ITC control total for 3B 4A(1). | Med |
| Part B "Others" credit notes (IMS) 4(A) (IGST/CGST/SGST) | Tirupati Apr-25: IGST 57,785.40 | D (CN_* in D2) | Cess (`CN_CESS`) and debit notes are never compared separately: engine only nets CN in D2. | Low |
| Part B lines: `B2B - Credit notes (Reverse charge)`, ISD CNs | zero | A | RCM and ISD credit notes net-off. | Low |
| "Advisory" column | "Net input tax credit may be availed under Table 4(A)(5) of FORM GSTR-3B" | A | Portal guidance; constant text, no information. | Low |
| "GSTR-3B table" column | `4(A)(5)`, `3.1(d) / 4(A)(3)` | A | Confirms mapping; constant. | Low |
| Quarterly file only (Jagnath): the 4th column group "Total (Jan-Mar)" | e.g. 4,536,353.88 CGST | D partly (sum-check `_qtr_total_mismatch`, set but never surfaced) | The mismatch flag is set when m1+m2+m3 != quarter-total but grep finds no consumer (`_qtr_total_mismatch` appears only in the parser/adapter and one comment at flow:1635), so a self-inconsistent government summary would be silent. | Low |

### 1.4 B2B-CDNR sheet (credit / debit notes)

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Whether ITC to be reduced (Taxpayer's Input) | `NA` on all 108 Bajrang credit notes, 57 Tirupati | A (`extra`) | The recipient's own IMS declaration that ITC is to be reduced for a credit note, plus the next column; in this layout it is NA everywhere so no signal now, but when populated it is direct evidence that the taxpayer reversed ITC (supports 4(B)(2) cross-check, D2). | Med |
| Amount declared by taxpayer for ITC reduction (IGST/CGST/SGST/Cess) | `NA` | A (`extra`) | Same: declared reversal amount vs 3B 4(B)(2). | Med |
| Remarks | `No remarks available` | A (`extra`) | IMS remark text on a rejected/pending note. | Low |
| GSTR-1 Filing Date | `11/01/2025` (Bajrang 24-25, 262 rows) | A | Credit notes issued after the supplier's 30-Sep/30-Nov Sec 34(2) cutoff; delay between note date and filing. | Med |
| Note Supply type (`supplytype`) | `Regular` | B (dict, never read) | Same as invoice type. | Low |
| Note value (`noteval`) | `17497.03` | C | Used only for arithmetic check. | Low |
| Note type | 773 Debit, 108 Credit in Bajrang 24-25 (Mahindra) | D (ntype) | M&M debit notes (Rs 2.69L tax in 2A) = additional ITC; ratio of debit to credit notes and the Dec/Sep 2024 bunching (26/12/2024: 474 notes; 18/09/2024: 180) is an unusual year-end settlement pattern not examined. Could be used as a supplier-level credit/debit note intensity indicator. | Med |
| Source / IRN / IRN Date | `E-Invoice` 712 rows of 881 | C/A (not used in #28 because that is invoice-only) | Credit notes without IRN from e-invoicing suppliers (169 of 881 here): credit note not e-invoiced = ITC reduction cannot be enforced; flip side of the invoice test. | Low |

### 1.5 B2BA, B2B-CDNRA

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Original invoice number / date, revised invoice | Tirupati Feb-26: original `02`, 14/01/2026, revised `1`, GSTIN 09GKHPK2685C1ZY | B/D (orig_invno used only to drop the stale B2B row; spliced as `via_amendment=True`) | `via_amendment` is never read: no list of "ITC claimed through amendments", no amendment count per supplier, no check that the revised invoice was not also claimed under B2B. The Feb-26 sheet shows a Rs 11.8 lakh Giscon Engineering invoice and a Rs 9.08 lakh F.S. Enterprises invoice amended 1-2 months after the supply. | Med |
| Invoice Value / Taxable Value / Tax (revised) | `1180000` | D | used when spliced. | - |
| Original amounts (not in file) | n/a | - | The 2B B2BA sheet gives only revised amounts, so "amount changed from X to Y" cannot be derived from 2B; 2A's B2BA gives orig vs revised taxable (G5). | - |
| Whether ITC to be reduced / Amount declared / Remarks | `NA` | A | same as CDNR | Low |
| Filing Date, Period | `10/03/2026` | B/A | Amendment filed very late (after the original period's 16(4) cutoff) = ITC on amended invoice may fall outside Sec 16(4). | Med |

## 2. Layout notes (quarterly / summary-only / merged)

- **GSTR2BQ (Jagnath)**: B2B sheet has 22 columns and a real `Rate(%)` column (mapped to `rate`, never tested). Sheets ISD, ISDA, IMPG, IMPGSEZ, ECO, ECOA, ITC not available are all present and empty. The Jun-23 file's B2B sheet reports max_row = 1,048,576 (dimension record is corrupt, not data; read-only openpyxl iterates past real data).
- **"_summary" files (Tirupati May-25, Jan/Feb/Mar-26 etc.)**: ITC Available, ITC not available, ITC Rejected plus B2B, B2BA, B2B-CDNR only; the other IMS sheets (ECO, ISD, IMPG, *Rejected) are missing. Adapter correctly handles this; but "B2B(Rejected)" and "B2B-CDNR(Rejected)" exist only in the non-summary downloads (Jul-25, Aug-25), so rejected rows can be absent from "summary" months. The unread-sheet INFO does not warn that rejected-row detail is missing for those months while "ITC Rejected" total shows a non-zero number. Example: May-25 ITC Rejected shows Rs 56,711.60 in the summary sheet, yet the file is a `_summary` variant whose `B2B(Rejected)` sheet has 1 row; and the Jul-25 non-summary download has 4 rows.
- **Merged files (Bajrang 24-25 `GSTR2B_Merged (1) (1).xlsx`, `output/GSTR2B_Merged.xlsx`)**: only B2B, B2BA, B2B-CDNR are present. The merge tool supports every sheet (merge_gstr2b.py lists ISD, ECO, IMPG, *(Rejected), ITC *), but the input was a "document-wise" export so there is no ITC Available at all for FY24-25 Bajrang: summary-driven checks (F7a, R14's 2B total, D-section) fall back to recomputing from rows. Nothing in the UI warns that ISD/IMPG/ECO would be invisible.

## 3. GSTR-2A sheet by sheet

### 3.1 Whole-sheet items

| Sheet | Present in (real data) | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|---|
| TDS | Bajrang 24-25 merged: 39 rows (non-empty); Tirupati monthly: empty | 05MRTO00357A1DK OFFICE OF THE SUPERINTENDENT OF POLICE ALMORA, GSTR-7 period 042024, taxable 1,99,800, CGST 1,998, SGST 1,998 | A | TDS (Sec 51) credit credited to the taxpayer by government deductors. FY24-25 totals from 2A: taxable Rs 23.04L, CGST 22,923.95 + SGST 22,923.95. Electronic Cash Ledger "TDS Credit" entries for the same year: CGST 17,689 + SGST 17,689 (9 entries). So ~Rs 5,235 CGST + 5,235 SGST (Rs 10.5k) of 2A-reported TDS credit has not been credited to the cash ledger (unaccepted in GSTR-2, or timing). Also: deductors 05MRT... (govt bodies) must appear in GSTR-1 B2B as customers with matching taxable value; the engine already trips over a blank name of such a deductor (flow:2100) but never reads the TDS sheet. Cross-check 1) TDS sheet vs ledger TDS credit; 2) TDS taxable vs GSTR-1 supplies to that deductor (deductor under-/over-reporting); 3) TDS credit utilised in 3B 6.1. Zero use today (grep: no TDS/GSTR-7 logic in any check). | High |
| TDSA | same (empty) | | A | amendments of TDS | Low |
| TCS | Bajrang 24-25 (empty), Tirupati (empty) | header: GSTIN of E-com operator, GSTR-8 period, gross value, returns, net value, TCS | A | TCS credit (Sec 52): taxpayer selling via Amazon/Flipkart etc. Cross-check to cash ledger "TCS Credit" and to GSTR-1 table 14/15. Empty here, but for an e-commerce seller it is the only source of ECO-collected tax. | Med |
| ECO / ECOA | empty in all samples | | A | ECO-supplied documents (9(5)), inward side | Low |
| ISDA | empty | | A | ISD amendments (ISD itself IS read, G8) | Low |
| IMPG | empty | header: Reference date, Bill of entry (port, number, date, taxable, IGST, cess), amended flag | A | Import ITC credit from BoE; cross-check to 3B 4A(1) and to Customs. | Med |
| IMPG SEZ | empty | | A | SEZ imports | Low |
| CDNR | Bajrang 24-25: 881 notes (773 debit, 108 credit), credit-note tax Rs 27.60L, debit-note tax Rs 2.69L | MAHINDRA AND MAHINDRA ..., Debit note | B | Parsed into `r2a["cdnr"]` and the canonical sheet but no check reads it (no `r2a.get("cdnr")` anywhere). 2A is the supplier-filed view (earlier than 2B, includes GSTR-3B status of the supplier on notes). Use: 2A CDNR vs 2B B2B-CDNR (G2-style existence check for notes, missing today: G2 is B2B only), plus 4(B)(2) cross-check using the 2A note list when the 2B ITC Available sheet is absent. | Med |
| CDNRA | empty | | B | parsed, never used | Low |
| Read me | all | field list | B | GSTIN/legal name/FY | Low |

### 3.2 B2B sheet (24 columns)

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| GSTR-3B Filing Status (`g3b_status`) | `Y` 34,123 rows, `N` 259 rows (Bajrang 24-25, per rate-line); on the -Total rows 129 documents, tax Rs 8,06,568, 18 suppliers; top: RAKHI CHOUHAN 05AGGPC4967Q1ZW Rs 4,15,546 tax; NARINDER KUMAR KHANEJA 52 docs Rs 1,64,140; DOLLY KHANEJA 22 docs Rs 1,25,160 | B (parsed in both readers, no check reads it) | Supplier has filed GSTR-1 but not GSTR-3B: tax collected but not paid to government. Section 16(2)(c): ITC is available only if the tax has been actually paid by the supplier; this is exactly the 2A "red flag" officers use. In 2B the ITC still shows because 2B only needs GSTR-1/3B for auto-population in 2B generation; 2A's `N` identifies suppliers who will fail Sec 16(2)(c) in scrutiny. Directly feeds a "ITC at risk" table by supplier and a Sec 16(2)(c) / Rule 37A finding. Also tie to cancelled suppliers (below). | High |
| GSTR-1 Filing Status (`g1_status`) | `Y` for all 34,382 rows | B | Y everywhere here; any `N` row means invoice in 2A via IFF/QRMP only. | Low |
| GSTR-1 Filing Period (`g1_filing_period`) | `Apr-24`, `Jul-24` | B | Same lag analysis as 2B's `filed_period` (supplier reporting later than invoice month). | Med |
| GSTR-1 Filing Date (`g1_filing_date`) | `13-May-24` | D (G9 16(4) test) | Used only for the 30-Nov deadline, not for the late-filer lag/regularity score per supplier. | Low |
| Effective date of cancellation (`cancel_date`) | `31-May-26` in 365 rows, 31-Oct-25 (10 rows), 01-Apr-25 (6), 25-Dec-24 (4); 387 rows, 6 suppliers, tax Rs 3.21L | D (G10 counts and lists) | Used; but G10 does not compare the invoice date to the cancellation date (finding text says "verify"). The code has both dates (`invdate`, `cancel_date`): invoices dated AFTER cancellation are ITC blocked (Sec 16(2)(a)/Rule 21/ retrospective cancellation) and could be flagged FLAG instead of REVIEW, and cancelled-with-retrospective-effect (01-Apr-25 and 25-Dec-24 inside FY24-25 months) should be priority. | Med |
| Amendment made, if any (`amend_type`) | 24 rows, e.g. `Invoice Details` (20), `Invoice Number` (4) | B (B2B; only the B2BA amend_type is displayed in G5) | Supplier amended a B2B row after filing. Amended-after-claim pattern per supplier: amendments of "Invoice Details" within the same period in which the taxpayer claimed. | Low |
| Tax Period in which Amended (`amend_period`) | `May-24` (20), `Mar-25`, `Oct-25` | B | Amendment in a later FY (Oct-25 for a FY24-25 invoice) = invoice value changed after the 16(4) cut-off. | Med |
| Source (`source`) | `E-Invoice` 24,220 rows, blank rest | B | Same IRN mix test as 2B. The mixed-IRN supplier list (M&M etc.) can also be built from 2A. | Low |
| IRN / IRN date | present on 24,220 rows | B | IRN date vs invoice date; as above. | Low |
| Rate (%) | `28`, `18`, `5`, '-' on the "-Total" row | A (not parsed; only in canonical `extra` as `Rate (%)=...`; and the engine keeps only the "-Total" row) | 2A is the only place with a genuine per-line rate on the Bajrang FY24-25 side (2B lacks it). Rate per supplier gives: 28% vehicle/parts purchase mix, rate vs GSTR-1 outward rate (HSN rate check A-series), Section 17(5) vehicle ITC (28% motor vehicle lines), cess presence on 28% lines. | Med |
| Place of supply (`pos`) | `Uttarakhand` 34,364; others 18 rows | B (G7 uses supplier GSTIN state) | POS differs from recipient state = charge is IGST or POS-rule issue. | Low |
| Invoice type | `R` (all rows) | D (join key) | no signal | Low |

### 3.3 B2BA

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Amendment made, if any | `GSTIN Amended`, `Invoice Details`, `Invoice Number` (Bajrang 24-25: 14 docs; Tirupati Apr-25: 1 doc `GSTIN Amended` RAHUL TOMER, taxable 4,33,188.70, IGST 77,973.97) | C (counted in G5) | GSTIN-amended invoices: original recipient GSTIN was someone else; moved to the taxpayer after the supply. Flag per rule: link to the 2B "B2B - Invoices (Amendment)" ITC line (same Rs 77,973.97 in Tirupati 2B Apr-25). Cross-source proof: 2A amendment type + 2B amendment ITC total = same figure. | High |
| Original tax period in which reported | `Mar-25` | C (G5 table) | time gap original-period to amended-period (Mar-25 original, Apr-25 revised, FY boundary crossing). | Med |
| GSTR-3B status / G1 status / filing date / cancel date | | B | same as B2B. | Med |
| Original invoice number / date | | D | used to link to B2B. | - |

### 3.4 CDNR / CDNRA columns (all B because the sheet is never consumed)

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Note type, Note number, Note supply type, Note date, Note value | `Debit note` 773, `Credit note` 108 | B | Note-level existence check vs 2B CDNR; missing from G2. | Med |
| Taxable value, IGST/CGST/SGST/Cess | credit-note tax Rs 27.60L (Bajrang 24-25) | B | Total credit-note tax available to compare with 3B 4(B)(2) and the engine's 2B CN total. | Med |
| GSTR-3B Filing Status | 0 `N` among 881 | B | same Sec 16(2)(c) test on notes; credit notes from suppliers not filing 3B are non-enforceable. | Low |
| Amendment made / Tax period in which amended | blank in all 881 | B | none now | Low |
| Source/IRN/IRN date | 712 of 881 `E-Invoice` | B | credit notes without IRN from e-invoicing suppliers (169 rows). | Low |

### 3.5 ISD (parsed, used only for total tax in G8)

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Eligibility of ITC, ISD doc type, ISD invoice number/date, credit note number/date, original invoice number/date, ISD GSTR-6 filing status, amendment, tax period of amendment | sheet empty in all real data | B (parsed, only igst+cgst+sgst+cess are summed in G8: flow:3137-3141) | `eligibility` ("Eligible"/"Ineligible") is never read: G8 sums ineligible ISD credit as if eligible. Document type (invoice vs credit note) is also ignored, so ISD credit notes are added instead of subtracted. Cannot be exercised with these samples (zero rows) but is a latent logic gap. | Med |

## 4. Cross-source opportunities that need no new sheet beyond what is already present

1. **2A TDS vs cash-ledger TDS credit** (Section 3.1): Rs 10.5k CGST+SGST in FY24-25 Bajrang and a direct unaccepted-credit finding.
2. **2A B2B GSTR-3B Filing Status = N vs 3B 4A(5) claim** (Section 3.2): 129 documents, Rs 8.07L, 18 suppliers.
3. **2B Rejected (IMS) sheets** (Section 1.1): 4 invoices Rs 27.9L-38.6L each and 4 identical credit notes, Tirupati Jul/Aug-25, Rs 19.57L IGST rejected.
4. **2B IRN/no-IRN mix by supplier** (Section 1.2): M&M 1,206 no-IRN rows, Rs 8.85 cr tax.
5. **2B / 2A amendment ITC with `GSTIN Amended`** (Sections 1.3, 3.3): Rs 77,973.97 IGST in Tirupati Apr-25, Rs 13.42L amendment ITC over the year.

## 5. Top 5 opportunities (ranked)

1. **GSTR-2A TDS (and TCS) sheet reconciliation** - whole sheet unread; real data present (39 TDS rows, Rs 23.04L taxable, CGST+SGST Rs 22,923.95 each) vs Rs 17,689 each credited in the Electronic Cash Ledger. Add a TDS vs ledger vs GSTR-1-deductor table and a "TDS credit not yet taken" finding. Small change: new 2A table in `gstr2a_mapping.json` + one check.
2. **GSTR-2A `g3b_status`/`g1_status` supplier non-filing** - already parsed in `gst_parsers_dept.parse_r2a_excel` (B), zero checks. Group by supplier, show ITC at risk (129 docs / Rs 8.07L / 18 suppliers in Bajrang FY24-25), Sec 16(2)(c)/Rule 37A. Pure engine-side change.
3. **GSTR-2B IMS sheets: ITC Rejected, B2B(Rejected), B2B-CDNR(Rejected) (and ITC not available)** - whole sheets unread; Tirupati shows invoice/credit-note wash pairs (Rs 27.7L to Rs 38.6L each, identical values/dates, Jul/Aug-25) with Rs 19.57L ITC rejected. New adapter sheets plus a check "rejected invoice with equal-value rejected credit note" and a control total vs `ITC Available`.
4. **GSTR-2B whole-register IRN / e-invoice discipline** - `irn` and `einv_source` exist in every row but are consulted only inside #28 for EWB-matched invoices. A supplier-level "mixed IRN" table (M&M 1,206 no-IRN rows Rs 8.85 cr tax; 43 suppliers Rs 12.36 cr) and IRN-date lag (`IRN Date` is only in `extra`) would be a Rule 48 / Sec 16(2)(a) risk list at no new data cost.
5. **GSTR-2B/2A supplier filing date and amendment signals** - 2B `GSTR-1 Filing Date` is dropped to `extra`; `filed_period`, `via_amendment`, 2A `amend_type`/`amend_period`/`orig_tax_period` (B2B) are canonicalised but unused. Combine into a supplier reliability score (late filing: 263 rows Rs 77.5L Bajrang, 994 rows Rs 87.0L Tirupati; amended-after-claim; GSTIN-amended invoices; prior-FY-dated invoices Rs 16.4L Tirupati) and a 2B-based Sec 16(4) test that works when 2A is not supplied.

## 6. Caveats and things not verified

- Unreadable: Jagnath 2BQ Jun-23 file reports 1,048,576 rows for B2B (dimension bug); counts above exclude it.
- The monthly 2A files in Tirupati and Bajrang 23-24 zips are tiny (about 30 rows per sheet, B2B total 192-201 rows for the whole year) versus 800-1,800 rows per month in 2B: they look truncated, so the FY23-24 / Tirupati 2A figures above come only from the merged Bajrang FY24-25 2A.
- `ITC not available` sheet could not be cross-checked to 3B 4(D) because the engine does not parse 4(D) at all; this is stated as an opportunity, not a verified mismatch.
- Totals computed with ad-hoc scripts over the raw workbooks (kept outside the repo); numbers are for -Total rows (2A) or one row per rate-line (2B), as stated.


---

# Unused-input audit: E-Invoice and E-Way Bill (both directions)

Status legend: A = never read (dropped / only in `extra` blob / sheet ignored); B = read or canonicalised but no check uses it;
C = used only for display (detail-row column, no test on it); D = used in a check (mentioned briefly only).
Engine files read: `einv_adapter.py`, `ewb_adapter.py`, `einv_mapping.json`, `ewb_mapping.json`, `gst_parsers_returns.py`
(`parse_einv`, `parse_annual_ewb`), `gst_checks_monthly.py` (`read_einv_lines`, `read_einv_invoices`, run_checks #1-#12, EWB matrix #1-#27),
`gst_checks_hsn_fraud.py` (#15, #23, #24, #49, #51 and the EWB pattern checks), `gst_checks_flow.py`, `gst_machinery_scan.py`, `master_build.py`.
Method: the spec said "the engine only reads the `b2b, sez, de` tab"; I confirmed that in code and then read every consumer.
Python was run without `-I` for openpyxl only because `-I` hides the user site-packages where openpyxl lives; scripts were kept in the scratchpad, input folders untouched.

## 0. What the real files look like

### E-Invoice (EINV_<GSTIN>_<FY>.xlsx, one file per tax period, merged by `forms merger`)
| Taxpayer | Files | Sheets in each file | Rows (b2b) |
|---|---|---|---|
| Bajrang Motors FY23-24 (`input_files_bajrang_motors/e inv/*.zip`, 10 zips x 2 xlsx) | 20 xlsx | `Read me`, **`b2b,sez,de` (no space)**, `cdnr`, `cdnur`, `exp`, **`hsn`** | 9,000 rows in the 20 files (6,285 distinct invoice numbers, Rs 22.9 cr invoice value), 42 Cancelled, 60 cdnr rows |
| Jagnath Enterprises FY23-24 (`input_files_jagnath/e invoice/`) | 11 xlsx + 1 zip | `Read me`, `b2b, sez, de`, `cdnr`, `cdnur`, `exp` | 1,898 rows; 5 "Auto-population failed" with an error text |
| Tirupati Cement FY25-26 (`input_files_tirupati_cement/e invoi/`) | 11 xlsx | same as Jagnath | 50 rows |
Bajrang FY24-25 has no E-Invoice file at all.

b2b tab columns (22): GSTIN/UIN of Recipient, Receiver Name (spelled `Reciever Name` in 25-26 files), Invoice Number, Invoice date, Invoice Value,
Place Of Supply, Reverse Charge, **Applicable % of Tax Rate**, Invoice Type, **E-Commerce GSTIN**, Rate, Taxable Value, Integrated Tax, Central Tax, State/UT Tax,
Cess Amount, IRN, IRN date, E-invoice status, GSTR-1 auto-population/ deletion upon cancellation date, GSTR-1 auto-population/ deletion status, Error in auto-population/ deletion.
There is **no HSN, no item description, no quantity, no UQC, no unit price, no item-level line** in this download (the portal export is rate-wise per invoice, not item-wise).
The `hsn` tab exists only in Bajrang's files and has a header row and zero data rows in all 20 files. `Read me` carries GSTIN, legal name, trade name, `Date Updated till`.

### E-Way Bill
All real inputs are the portal "EWB MIS report" (`.xls` that is really an HTML table: Jagnath raw, `docs/sample_data/inward|outward`) or the `*_eway_bill_merged.xlsx`
made from them (Bajrang x2 years, Tirupati). Always ONE sheet, 12 columns, no period banners:
`EWB No. | From GSTIN & Name | To GSTIN & Name | From Place & Pin | To Place & Pin | EWB No. & Dt. | Doc No. & Dt. | Assess Val. | Tax Val. | HSN Code | HSN Desc. | Latest Vehicle No.`
Row counts (inward/outward): Bajrang 23-24 9,425/117; Bajrang 24-25 10,478/247; Tirupati 25-26 7,029/132; Jagnath 964/72.
The MIS report has NO transporter name/ID, distance, transport mode, doc type, status (active/cancelled/rejected), validity, quantity, unit, consignment value split, supply type, or Part-A/Part-B flag.
The engine itself says so (`gst_checks_monthly.py` #22 "EWB export has no validity/expiry column"; `gst_checks_hsn_fraud.py` #40 "no vehicle-capacity reference... no product-weight field").
`ewb_mapping.json` has an extra `TAX RATE` field that no real file contains (always 0).
No E-Way Bill Analytics output (converted PDF) exists anywhere in the repo or input folders; only the converter `pdf to excel/parse_ewb_analytics.py` (20 sections) does, and **no engine file reads it** (grep for the section titles in `main gst tool/*.py` returns nothing).

---
## 1. E-INVOICE

### 1.1 Sheet `b2b, sez, de` (the only sheet read)
Columns that ARE used (D, brief): GSTIN, Invoice no, Invoice date, Rate, Taxable, IGST/CGST/SGST (totals, E-Inv vs GSTR-1 line compare, #6 EWB-vs-E-Inv taxable), IRN, IRN date (#8 lag to filing, #15 IRN after GSTR-1 filing, #23 IRN delay, #49), POS and RCM (#12 RCM uses `einv_lines`; POS only displayed), Status (cancelled split), Error column (#7).

| Raw column | Example value (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| GSTR-1 auto-population/ deletion upon cancellation date (mapped as `autopop_date`, deliberately NOT `cancel_date`) | Cancelled invoice RBC24A001122: IRN 01-Jun-2023, this column 03-Jun-2023 (42 cancelled rows in Bajrang: all deleted 1-2 days after IRN) | B (mapped, `cancel_date` stays None, so the "Cancelled E-Invoices" sheet in `master_build.py:2206` prints a blank cancel date) | For rows with status Cancelled/Deleted it IS the deletion date. Use as proxy to test: cancellation after the GSTR-1/3B for that month was filed (liability reduced in a later 3B, Sec 34/Rule 53 credit note is the correct route, not IRN cancel after 24h - Rule 9(?) of e-invoice cancellation window, Notification 13/2020); cancelled invoice that still has a live EWB (check exists but without the date); cancellation cluster at month end (window dressing). For Valid rows, autopop date minus IRN date: 373 of 8,958 valid Bajrang rows are 3-7 days late (portal-side lag, low value). | Med |
| GSTR-1 auto-population/ deletion status | `Deleted` x42, `Auto-population failed` x2 (Bajrang); `Auto-population failed` x5 (Jagnath) | B (in canonical `autopop`, no check reads it; only the Error text column is counted in #7) | `Auto-population failed` / `Deletion failed` mean the e-invoice never reached GSTR-1 (or a cancelled one stayed there): turnover not in GSTR-1 unless taxpayer re-keyed; a failed row with a BLANK error text is currently invisible. Cross-check each failed IRN against GSTR-1 invoice list (Sec 37, Rule 59(6)/ GSTR-1 auto-draft). Today #7 only fires if the error TEXT is non-blank. | High |
| Error in auto-population/ deletion (text) | `Document already present with same document number in same period`; `Zero items added or added with zero value` (Bajrang), `Error in auto-population` (Jagnath) | D (#7 prints it) but the error TEXT is not classified | Duplicate-document-number errors = same invoice number reused (Rule 46(b) uniqueness); zero-value item = invoice with a zero-value line (possible free supply/ sample not taxed under Sch I). Classify and escalate. | Low |
| Invoice Value | 107,017,982.00 vs taxable 90,693,205 + tax 16,324,777 (Tirupati TCP/UK/25-26/009) | C (kept in `read_einv_invoices.invval`, summed, only shown; `_val_mismatch` uses taxable only) | Invoice value minus (taxable+tax+cess) = round-off / freight / non-GST charges loaded in the invoice (value > tax base implies unreported other charges, Sec 15(2)); invoice value >= Rs 50,000 => EWB mandatory (Rule 138) - join to EWB by invoice no. In Bajrang 327 distinct invoices >= Rs 50k, today only checked via taxable. | Med |
| Cess Amount | 0.00 on every real row | C/B (summed into `parse_einv` CESS, not compared to anything) | Vehicles/ tobacco/ aerated drinks: cess on e-invoice vs GSTR-1 cess vs 3B cess (3.1(a)); Bajrang sells HSN 8703 (EWB implied 48% = 28% + cess) yet the e-invoice cess is 0 - needs a look that e-invoiced cess = EWB tax value gap. | Med |
| Receiver Name (`Reciever Name` in 25-26) | `SURENDRA SINGH`, blank in 17/50 Tirupati rows | B (canonical `receiver`; not read by any reader or check) | Name vs GSTIN registry/2B supplier name; name vs EWB `To name` for the same invoice (bill-to/ship-to); blank name = recipient GSTIN not resolved; consistency of one GSTIN under several names (I tested: none in Bajrang). | Low |
| Invoice Type | `Regular B2B` on 100% of 11,000+ rows | B (canonical `invtype`; no check reads it; GSTR-1 `invtype` is used elsewhere) | SEZWP/SEZWOP/DE/CBW types change the tax-head expectation and LUT/ refund claims; all real rows are Regular B2B so low yield today but zero-cost guard (SEZ w/o payment with IGST > 0, CBW with CGST). | Low |
| Place Of Supply | `27 - Maharashtra` 4,432 of 9,000 (Bajrang) | C (einv POS is only shown in #8 detail; #11 runs on GSTR-1 lines, hsn_fraud #11 etc. on GSTR-1 invoices) | Run the POS-vs-recipient-GSTIN-state head test (#11) on E-Invoice lines too: e-invoice is the primary document; GSTR-1 can be edited after auto-population. Compare POS e-invoice vs POS GSTR-1 for same invoice (differing = post-edit manipulation). | Med |
| Reverse Charge | `N` on every row | D (#12) | all N; fine | - |
| Rate | 18/28/12/5/0 | D | Rate 0.00 with positive taxable (4 Bajrang rows) is not tested against HSN; no HSN here so cannot. | Low |
| **Applicable % of Tax Rate** (not in `einv_mapping.json`) | blank on all 11,000+ rows | A (lands in `extra`) | Non-blank = 65% concessional rate (Notification 1/2017-CT(R) eligible supplies, old schemes). Blank everywhere so no yield now; add to mapping so a future non-blank value is flagged. | Low |
| **E-Commerce GSTIN** (not mapped) | blank on all rows | A (`extra`) | Supplies through an ECO attract Sec 52 TCS and 9(5): TCS credited in the electronic cash ledger (`Electronic_Cash_Ledger`) vs ECO-routed sales; blank here so nil. | Low |
| IRN (64-char hash) | `50f7091bc3ce14f4...` | D (only to show) | Duplicate IRN across invoices; same IRN on a Cancelled and Valid row; IRN vs invoice (hash of supplier GSTIN+FY+doc no) can be recomputed to prove the IRN belongs to that invoice number (SHA-256 of GSTIN|doc type|doc no|FY). Not done. | Low |
| Invoice date vs period banner | n/a | D (placement of quarterly banners only) | Invoice date older than IRN date by > 30 days (Rule 48(4), reporting limit for turnover > Rs 100 cr, 7 days from 1-Nov-2023 for >= Rs 10 cr) - `check_irn_delay` fires at >= 3 days only; add statutory threshold. | Med |

### 1.2 Sheet `cdnr` (e-invoiced credit/debit notes to registered persons) - WHOLE SHEET UNUSED (A)
Spec §"What the engine reads": "cdnr, cdnur, exp tabs are not used by any check and are listed as 'not used' in ISSUES" (I502). Confirmed: `einv_adapter.py:14`.
Real data: Bajrang 60 rows (all `C` credit notes, 33 at 28%, 18 at 18%, Rs 29.9 lakh taxable, tax Rs 7.8 lakh), Jagnath/Tirupati 0.
Columns: GSTIN/UIN of Recipient, Receiver Name, Note Number, Note Date, Note Type, Place Of Supply, Reverse Charge, Note Supply Type, Note Value, Applicable % of Tax Rate, Rate, Taxable Value, Integrated/Central/State Tax, Cess, IRN, IRN date, E-invoice status, autopop date/status/error.

| Raw column | Example | Status | Use | Priority |
|---|---|---|---|---|
| Note Number / Note Date / Taxable / tax (all cdnr columns) | `VOU24A000064`, 02-Jun-2023, taxable 10,135.14, CGST 1,418.92, recipient 05EJNPM3606N1ZS | A | Reconcile E-Invoice credit notes to GSTR-1 CDNR (Sec 34): in Bajrang 30 of the 60 distinct e-invoice note numbers (about Rs 23.6 lakh taxable) do not appear in the 12 supplied GSTR-1 `cdnr` tabs (GSTR-1 has 72 note lines). First-pass match on note number only, needs verifying, but if real this is credit notes issued/IRN-registered yet missing from the return or the reverse. Also feeds the Sec 34(2) time-limit test (note raised after 30-Nov following FY end / link to original invoice's ITC reversal at recipient), the GSTR-3B 4(B)(2)/ 3.1(a) net-off and the CN-vs-sales-return EWB logic (`gst_checks_hsn_fraud.py:1483`, which today matches EWB to GSTR-1 CDNR only). Credit notes are 3.3 percent of invoice taxable; large post-year-end notes are a classic turnover-reduction device. | High |
| E-invoice status (Cancelled notes) | all `Valid` | A | cancelled credit notes excluded | Low |
| Note Value (note value vs taxable+tax) | 15,000.00 vs 10,135.14 + 2,837.84 | A | note value above taxable+tax = amount refunded without GST reversal (e.g. note 15,000 for taxable 10,135 at 28%) | Med |

### 1.3 Sheets `cdnur` and `exp` - WHOLE SHEETS UNUSED (A), no rows in any real file
Header-only in all 31 real files. `exp` carries Port Code, Shipping Bill Number/Date, Export Type (WPAY/WOPAY): would feed export-turnover vs ICEGATE (`BO Profile` "Export - ICEGATE Information" sheet) and LUT/ refund checks (Sec 16 IGST Act). `cdnur` has UR Type (B2CL/EXPWP/EXPWOP). No yield for these three taxpayers; priority Low until an exporter is loaded. Important: the adapter should at least raise a count when these tabs contain data rows (today I502 says "not used" even when populated).

### 1.4 Sheet `hsn` (Bajrang only) - WHOLE SHEET UNUSED (A) and EMPTY
Header: HSN, Description, UQC, Total Quantity, Total taxable value, Rate (%), Integrated/Central/State tax, Cess. Zero data rows in all 20 Bajrang files, so there is no item/quantity data from E-Invoice at all today. If a taxpayer's file does have rows this is the ONLY place a UQC/quantity appears (HSN-wise outward summary auto-drafted from IRP): it would give a second HSN/rate source independent of GSTR-1 table 12 (compare to `check_b2b_matching_accuracy`, HSN rate tests in `gst_checks_hsn_fraud.py` 845/930). Also the `forms merger` (`gst_merge_common.py`) does not carry this tab into `EINV_Merged.xlsx` (output has only b2b/cdnr/cdnur/exp). Priority Med (Low today, because empty).

### 1.5 Sheet `Read me`
Used only for GSTIN and legal name. `Date Updated till` (e.g. 31-Aug-2026 15:08:28) unused (B/A): it tells how fresh the IRP data is; a file "updated till" before the GSTR-1 due date can be missing late IRNs. Trade name unused. Low.

---
## 2. E-WAY BILL (MIS report; inward and outward use the same 12 columns)

| Raw column | Example (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| **Tax Val. / Assess Val. as an implied tax rate** | Bajrang outward Thar 502,600.22 / 1,047,083.79 = 48%; Bajrang inward 28% bucket 4,861 rows, 20-26% buckets ~1,050 rows, 4% bucket 28 rows, 0% 26 rows; Tirupati inward 4% on 821 rows (HSN 251710 at 5%: 37 rows) | D for sums; **the ratio itself is B** (the `rate` field comes from the `TAX RATE` column that no real file has, so `rate` is always 0.0; `gst_checks_flow.py:1065` only tests tax = 0) | Per-EWB effective rate vs the statutory rate for that HSN (HSN master already loaded in `gst_checks_hsn_fraud.py`): inward EWB with lower rate than the rate the supplier should charge = ITC risk / supplier under-charging; outward EWB with rate lower than the HSN rate = under-declaration of tax value. Check on Tirupati: HSN 25232930 (cement) is 18% (28% before 22-Sep-2025) - the EWB rate by date can prove the rate change applied. Bajrang vehicles 87038040 show 5% (EV) correctly vs 48%/31% on petrol/diesel; mismatch of rate to HSN is directly testable. | High |
| **Zero Tax Val. (inward and outward)** | Tirupati outward: 123 of 132 EWBs tax 0 (Rs 1.65 cr of Rs 2.27 cr); Jagnath inward 248 of 964; Bajrang inward 26/30 | D (flow 1065/1070 list them) | already covered; cross-check with ITC table 4 is the open part | - |
| **From Place & Pin / To Place & Pin: the 6-digit PIN** | Tirupati inward `Malli Madli Champawat / 262523` x 85+, `KICHHA / 263148` x1,141, `TANAKPUR / 262309` x 90 outward; BO Profile (Places of Business) declares only Haldwani (principal) and Rishikesh 249204 (additional) | **A/C**: `parse_annual_ewb` keeps `from_place`/`to_place` strings, but only `to_place` is echoed in the #11-type detail (`gst_checks_hsn_fraud.py:1227`); the PIN is never parsed or compared. All state tests use the GSTIN prefix (`gstin_state`) | (1) **Undeclared places of business**: distinct own-side PINs of EWBs vs BO Profile principal + additional places (Sec 25(1)/(4), Sec 2(85), Rule 19 amendment of registration). Tirupati: 85 distinct own to-places inward and 17 outward from-places, main ones Tanakpur 262309, Champawat 262523, Kichha 263148, Udham Singh Nagar 263153 - none in the BO Profile. 1,141+118+97 inward EWBs to Kichha alone. (2) 117 Tirupati outward EWBs are from the GSTIN to itself (Tanakpur to Champawat, Rs 1.67 cr) = movement between own sites that are not registered. (3) PIN-derived state vs To GSTIN state vs POS (bill-to/ship-to, Sec 10(1)(b)): Bajrang 24-25 has 10 outward EWBs where from GSTIN = to GSTIN = Bajrang, Haldwani to Haldwani, HSN 87038040, tax 5%, with INV25A series numbers - a sale-series invoice raised on one's own GSTIN (demo/ own-use vehicles, Sec 17(5)(a)(b) and Sch I self-supply). (4) Bajrang inward: To PIN 263620 once (delivery elsewhere than Haldwani). (5) PIN to PIN distance = a proxy for Rule 138(10) validity (1 day per 200 km): not computable without a pincode DB, but state-level PIN distance is possible. | High |
| **Latest Vehicle No.: format, blank, mode strings, state** | Bajrang 23-24 inward: 837 blank (8.9% = Rs 1.28 cr), strings `Air`(20), `Ship`(6), `Rail`; 58 rows with `TM` prefix (`TMDL08TC0296`, `TMMH14TC305176`); 762 NL (Nagaland) plates, 3 `TMHR26TC/2019/1`; Tirupati `UK063CC9558` malformed; 5 blank | **C/B**: vehicle is shown on detail rows and used only in #27 (same vehicle, same GSTIN pair). No format or presence test | (a) **Blank vehicle** after the EWB is generated = Part-B not filled: Rule 138(3) says no movement without Part B; 837 inward EWBs in Bajrang and 644 in 24-25 show goods without a vehicle (or via transporter ID only) - cross-check to GSTR-2B date and purchase booking; (b) vehicle field holding `Air/Ship/Rail` = mode text instead of a vehicle (Rule 138(1)? Part B alternative for rail/air/ship needs RR/AWB no.) - genuine for rail but check the consignor/ value; (c) invalid/ non-standard registration string (not matching `^[A-Z]{2}\d{1,2}[A-Z]{0,3}\d{4}$`) = fake vehicle number; (d) **vehicle state plate vs route**: NL/ TM plated trucks on a Haridwar to Haldwani leg are a common carrier pattern, but a plate registered in one state used for hundreds of unrelated legs, or the SAME vehicle number appearing in EWBs ~15-45 minutes apart in far-apart towns, is an impossible-trip signal (my first-pass count: 533 pairs in Bajrang 23-24, 687 in 24-25, 4 in Tirupati, 0 in Jagnath; NOTE: the column is the LATEST vehicle, which after a Part-B update replaces the original, so treat as indicator needing the full EWB history, not proof); (e) a vehicle appearing both on taxpayer's inward and outward legs (14 / 23 / 42 / 7 vehicles) = owned or hired fleet - relevant to transport-cost and ITC claims. | Med (blank/ mode/ format = quick win; impossible-travel Med) |
| **HSN Desc.** | Bajrang inward 114 blank + vehicle model names (`MAHINDRA BOLERO B6 O BSVI`), Tirupati outward 128 of 132 blank, Jagnath outward `721420` (desc = HSN code, 59 rows) | C (only `gst_checks_flow.py:519/553` commodity list) | Description vs HSN consistency (e.g. `BISON WALL PUTTY` filed as HSN 38244090 - putty is 3214, rate 18% either way but different heading); description naming an item that belongs to Sec 17(5)/ blocked list or capital goods (machinery scan reads HSN+name, not this description for EWB); blank description in 220 Tirupati inward EWBs hides what moved. Also vehicle EWBs carry the chassis/engine number in the description (`MA1NM2EW1R2B71341`, `RJK41523`): a chassis number appearing on two different outward EWBs / invoices = same vehicle sold twice; matching chassis from inward EWB to outward EWB = stock-in/ stock-out traceability. | Med |
| **HSN Code (digit length, vs GSTR-1/ E-Invoice HSN)** | Tirupati: 907 rows 6-digit, 1 blank; Bajrang 1 outward row of length 1; Jagnath 64 six-digit | D at 4-digit heading (flow YoY comparison, #18 info only) | Digit-count vs turnover limit (Rule 46(g): 4 digits up to Rs 5 cr, 6 digits above); EWB HSN of 4 digits against 8-digit in GSTR-1 for the same doc; EWB HSN vs GSTR-1 HSN for the same document number (not compared at document level, only #18 info text) - classification shift between EWB and return. | Med |
| **EWB No.** (12-digit) | `371592191376` | D (dedupe key) | EWB-number prefix/ check: none | Low |
| **EWB No. & Dt. : time part** | `26/04/2023 19:53:00` | D (#21 midnight, burst cluster) | Time of EWB vs invoice date: EWB generated the evening before the invoice date, i.e. before the document existed (not tested; only the date gap is); night generation 23:00-05:00 is 584/880 inward EWBs for Bajrang, 621 for Tirupati (normal for trucking but unusual for outward: 0-1). | Low |
| **Doc No. & Dt. : date part** | `INV24A000152 - 12/05/2023` | D (#7, #16 date gaps) | EWB date minus doc date up to 10+ days in 47 (23-24) and 68 (24-25) inward Bajrang rows: EWB generated long after the invoice (Rule 138(1): before movement; delayed Part-B / goods not moving, late receipt = Sec 16(2)(b) receipt date vs ITC month). Only outward/in-doc gap is flagged, not tied to the 2B ITC month. | Med |
| **From GSTIN & Name / To GSTIN & Name: name part** | `Geeken Design Concepts Pvt.Ltd 2023-2024 - (from 1-Apr-23)` (name carries a ledger-year suffix = accounting-software masters), `URP` as a GSTIN (Tirupati, 1 row) | C (only the name learner at `gst_checks_flow.py:3758` and machinery scan) | Name suffix `(from 1-Apr-xx)` shows the recipient name is typed from the books of account, not from the GST portal: nothing to test legally, but EWB `To` GSTIN vs the e-invoice/GSTR-1 recipient GSTIN of the same invoice number is NOT compared anywhere (grep: to_gstin used only for state tests and zero-tax/ self tests): a mismatch = bill-to/ship-to or an EWB generated against another GSTIN to move goods. `URP` = unregistered recipient, EWB for goods to an unregistered person, check against B2C turnover. | High |
| Self-to-self GSTIN legs | Tirupati inward 126 / outward 117 rows with from = to = own GSTIN; Bajrang 24-25 10 | C/D (excluded from third-party checks in flow) | See PIN row: they are not "nothing" - they are the proof of unregistered sites and unrecorded stock movement (Sec 25(4), Sch I item 2, Rule 55 delivery challan for job work/ supplies on approval; Sec 143 job-work ITC-04). ITC-04 not a loaded source. | High |
| Same-pin legs | from PIN = to PIN: Jagnath inward 422 of 964, Tirupati inward 311, Bajrang 20 / 37 | B | Local delivery inside one PIN: EWB for goods that did not travel; Rule 138(14)? (distance under 10 km is exempt for transporter-to-consignee only, not for the consignor); also reveals the taxpayer's real godown vs registered address. | Low |

### 2.1 Fields the EWB workflow could provide but the real MIS files do NOT contain (cannot be tested from today's input)
Transporter name/ID, transporter doc no/date, distance (km), transport mode, document type (invoice/ challan/ bill of supply/ credit note), supply type (outward/inward, job work, SKD), EWB status (Active/Cancelled/Rejected), validity "Valid Upto", number of extensions, quantity and unit, consignee-state/ dispatch-state codes, tax split into CGST/SGST/IGST/cess and cess value, Part-A-only flag, multi-vehicle info.
The engine already prints the limitation (#22, #40). Recommendation: also accept the detailed EWB download (or EWB Analytics sections 5, 8, 10, 13, 17, 18, 19 below).

### 2.2 E-Way Bill Analytics (PDF converted) - WHOLE SOURCE UNUSED (A) / NOT PRESENT
No converted workbook is in `input_files_*`, `output/`, `docs/sample_data`; `pdf to excel/parse_ewb_analytics.py` defines 20 sections the engine never reads:
| Section | Content | What it could tell us | Priority |
|---|---|---|---|
| 1,2 Outward/Inward supplies (monthly) | EWB value by month incl. exports/imports | Whole-year EWB totals vs GSTR-1 / GSTR-3B 3.1 / 2B totals without the 12-column MIS (recon of the MIS file against GSTN's own totals - proves the merged MIS is complete) | High |
| 3,4 HSN-wise outward/inward | HSN totals | HSN mix vs GSTR-1 table 12 and E-Invoice HSN | Med |
| 5 EWB Verification (ok / not ok) | Verification outcomes of vehicle checks | Interception/ verification exceptions (Rule 138B) | Med |
| 6,7 Top 50 suppliers/ recipients | | Concentration, shell-supplier discovery without row data | Med |
| 8 Cancellations and rejections | counts and values | cancelled/ rejected EWB values with live GSTR-1 invoice, Rule 138(9) | High |
| 9 EWB blocking | blocked GSTINs (Rule 138E) | Blocking itself is a red flag (non-filing of 3B for 2 periods) | High |
| 10,11 Extension of EWB/ transporters | extensions | validity breaches; transit delays | Med |
| 12 Critical commodities, 13 Bill-to-ship-to, 14/15 B2C outward/inward, 16 Work contractors, 17 ODC, 18 Part-A EWBs, 19 EWB-03 | | each is a ready-made department risk indicator for the same GSTIN (bill-to/ship-to = Sec 10(1)(b) triangulation; EWB-03 = verification report; Part-A-only = goods never moved) | High |
| 20 Risk involved | narrative | Direct department risk text for the assessee | High |

---
## 3. Per-taxpayer observations (grounded)
* Bajrang FY23-24: 9,425 inward EWB (Rs 167 cr assessable); 8.9% have no vehicle; 4,861 at an implied 28% (cars/parts), ~1,050 at 20-26% (28% plus cess reflected unevenly), outward 48% on SUV HSN 87033xxx; E-invoice has 42 cancelled and 2 auto-population failures; 60 credit notes ignored.
* Bajrang FY24-25: 10 outward EWBs to its own GSTIN Haldwani to Haldwani (HSN 87038040, tax 5%, INV25A series, chassis numbers in HSN Desc.); 880 inward EWBs generated 23:00-05:00; no E-Invoice file.
* Tirupati FY25-26: 85 inward + 17 outward distinct own-site places, 117 self-to-self outward and 126 self-to-self inward EWBs (Rs 1.67 cr), none of Tanakpur/ Champawat/ Kichha in the BO Profile; outward is 94% zero tax value (Rs 1.65 cr of Rs 2.27 cr) incl. 4 EWBs to 07AACFT2702L1Z9 (same PAN, Delhi = distinct person, deemed supply Sch I(2)); desc blank on 128 of 132 outward; only 50 E-invoices.
* Jagnath FY23-24: 248 of 964 inward EWBs carry no tax value (Berger Paints 273,866 with blank tax); 5 e-invoices failed auto-population (error text present); outward HSN Desc. equals the HSN code on 59 of 72 rows.

## 4. TOP 5 OPPORTUNITIES
1. **Parse the PIN from From/To Place & Pin and compare own-side PINs to the BO Profile "Places of Business" (and registered addresses)** - finds unregistered places of business and self-to-self movements (Tirupati: Tanakpur, Champawat, Kichha, 1,300+ inward EWBs, Rs 1.67 cr self-moves). Also add PIN-state vs To-GSTIN state vs POS test. Source: EWB (A/C) x BO Profile.
2. **Per-EWB implied tax rate (Tax Val./Assess Val.) vs HSN-master statutory rate and vs the GSTR-1/2B rate for the same document** - `rate` is dead (always 0, no TAX RATE column in real files). Covers under-charged inward supplies (ITC), under-declared outward tax value, cess handling, and the 22-Sep-2025 cement rate change. Source: EWB (B).
3. **Process the E-Invoice `cdnr` sheet and reconcile it to GSTR-1 CDNR and 3B** (Bajrang: 30 of 60 distinct e-invoice credit notes, about Rs 23.6 lakh taxable, not found in the supplied GSTR-1 cdnr; verify). Whole sheet unused today. Source: EINV cdnr (A).
4. **Use E-Invoice `GSTR-1 auto-population status` and `autopop date`**: flag `Auto-population failed` / `Deletion failed` even when the error text is blank, and fill the missing cancellation date for Cancelled rows (deletion date) to test cancellation after the filing/3B date and with live EWBs. Source: EINV (B).
5. **Vehicle-field tests**: blank Part-B vehicle (837 + 644 inward EWBs in Bajrang), `Air/Ship/Rail/TM...` strings, malformed plates, same plate in distant places within an hour, vehicle common to inward and outward legs; plus compare EWB `To GSTIN` to e-invoice recipient GSTIN for the same invoice number (bill-to/ship-to) and EWB HSN to GSTR-1 HSN at document level. Source: EWB (B/C) x EINV x GSTR-1.

## 5. Caveats
* Vehicle number is the LATEST vehicle: impossible-travel counts are indicators only.
* `cdnr` mismatch was matched on note number only across 12 GSTR-1 files; confirm before reporting.
* No sample EWB Analytics workbook exists; section list taken from `parse_ewb_analytics.py`.
* Rule citations beyond 138, 46, 48, Sec 34, 25, 17(5) are for orientation and should be confirmed against the current notification text before use in a notice response.


---

# Unused-input audit: Ledgers (cash / credit / liability register / liability ledger / credit-reversal statement), Portal comparison, BS-PL and masters

Scope: `main gst tool/` (identical copy in docs/py/core). Real data: Bajrang Motors FY23-24 and FY24-25, Tirupati Cement FY25-26 (Jagnath has no ledgers/portal file).
Status key: A = never read; B = read / canonicalised but no check uses it; C = display only; D = used in a check (mentioned briefly).

## 0. How the engine really consumes the four ledger CSVs

Both readers (`gst_parsers_dept.parse_cash_or_liability_ledger` / `parse_credit_ledger`, and the canonical `ledger_adapter.py` clone) reduce each row to:
`date, ref, tax_period, description, ttype, heads{IGST/CGST/SGST/CESS: {tax,total,balance}}, total (= sum of the four head TOTALS), balance_total, ledger_used, demand_id, stay_status`.
The canonical workbook (`LEDGER_ROWS`) keeps all 6 parts x 4 heads x (amount + balance) and `time`/`reporting_date`, but `parse_*` drops them. Nothing is left in the `extra` blob: every column is mapped.

Grep result (gst_checks_*.py, master_build.py, gst_report.py): the ONLY ledger fields any check reads are `total`, `balance_total` (credit: `bal_total`), `date`, `ref`, `tax_period`, `description`, `ttype`, `demand_id`, `stay_status`, and `monthly_by_tax_period` / `monthly_by_txn_date` roll-ups. `heads[...]`, `ledger_used`, per-head credit-ledger `igst/cgst/sgst/cess` and `bal_igst/...`, `accrued_desc`, and every Interest/Penalty/Fee/Others part are read by NO check.
Consumers: F4 (credit tie-out), F5 (cash tie-out), F3/ITC roll-forward (credit by tax period), F8 (RCM cash vs liability, `gst_checks_flow.py:1650-1800`), F9/F9a/F9b/F9c (`:1814-1960`, keyword scan on description), `#9 #17 #47/#55` (`gst_checks_hsn_fraud.py:1412,1569,1837`), Annual Ledger Walkthrough (`gst_report.build_monthly_rows`).

Engine caveats found while reading (affect how far the used fields can be trusted):
* Cash ledger `AA...` debit+credit pairs on the same date are PMT-09 inter-head transfers (e.g. Tirupati 25/02/2026 `AA050226024524M` debit CGST tax 34,01,361 / credit CGST interest 34,01,361; Bajrang 21/01/2025 `AA050125013688D` SGST tax 50,00,000 debit / SGST "Others" credit 50,00,000). The debit leg carries a Tax Period so it lands in `monthly_by_tax_period[...].debited` and is counted as "cash tax paid" in #9, #17 and `cash_ratio_pct`; the credit leg has no period. Because Interest/Others parts are never separated, the engine cannot tell a transfer from a payment.
* `check_cash_timing` (#47/#55) tests the day-of-month of EVERY cash-ledger row (debits, TDS credits, PMT-09 legs), not just deposits (`Amount deposited`), so it fires on the 19th-21st of nearly every month purely because DC/AA debits are booked on filing day.
* F9 `VOL = ("voluntary","drc","demand","scrutiny","annual return")` is applied to the credit ledger too, so Bajrang's `Payment of Demand under Appeal` (appeal pre-deposit, 05/07/2024, ITC debit Rs 71,87,858 -> `DI0507240001391`) is labelled "DRC-03 / voluntary".
* `compute_interest` docstring (`gst_checks_forensic.py:1000`) says "verify against the Liability Register's own Interest-head entry" - that column is never read, so the verification it promises cannot happen.

---

## 1. Electronic Cash Ledger CSV (Bajrang 23-24 / 24-25, Tirupati 25-26)

File is read. Columns:

| Raw column | Example value (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Interest (x4 heads, Debited/Credited block) | Tirupati 25/02/2026 `DC0502260055083` Voluntary payment, CGST Interest 1,09,01,361 + SGST Interest 1,00,40,587; 23/09/2025 IGST Interest 2,325 | B (canonical keeps, `parse_*` drops; only Total used) | Separates s.50 interest from tax. Tirupati paid Rs 2.09 cr interest vs Rs 14.76 cr DRC-03 tax = 14.2%, i.e. ~290 days at 18% p.a., yet the tax was paid 17/07-30/07/2025 (<=10 days after the 20/07 due date). Either the base is older periods or interest was over/under-computed - an explicit rule-50 recomputation (interest only on cash portion, s.50(1) proviso) should be run against the actual interest-head debit. Also gives "interest paid = admission of delay" per month for Bajrang (SGST interest 1,569 Feb-24). | High |
| Fee (x4 heads) | Bajrang 23-24: CGST 25 + SGST 25 on 20/12/2023, 21/02/2024 (x2 filings), 21/03/2024; 24-25: 25+25 on 20/04/2024 | B | Actual late fee levied by the portal (s.47) per return. `gst_checks_forensic` computes a *theoretical* late fee from ARN dates; the ledger Fee column is the government's own figure - tie them out (also exposes late GSTR-1 fee vs 3B fee split, CGST/SGST 50/50, nil-return Rs 20/day cap). Red flag if ledger fee > or < computed. | High |
| Penalty (x4 heads) | none in these cash ledgers (0.00) | B | Penalty debits (s.73/74/122, DRC-07 recovery) in cash. Zero here but a non-zero entry would be invisible to every check. | Med |
| Others (x4 heads) | Bajrang 24-25 `AA050125013688D` 21/01/2025: SGST Others credited 50,00,000 (balance col shows SGST Others 5000000 parked) | B | Cash parked in the "Others" minor head (pre-deposit under s.107(6)/112(8), deposit under protest, DRC-03 for non-tax). Rs 50 lakh sat unused to FY-end beside a stayed demand of Rs 2.01 cr; links to the Liability Ledger demand and BS contingent liability (R8). | High |
| Tax part per head (`heads[h].tax`) | Tirupati 20/12/2025 PMT-09 `AA051225021944G` IGST 192 + SGST 97,000 | B (`heads` dict built, never read) | Head-wise cash utilisation: IGST cash paid while CGST/SGST ITC idle, SGST paid in cash while CGST ITC balance stays high (s.49(5) cross-utilisation order; Bajrang 24-25 Oct: SGST cash Rs 1.14 cr with CGST ITC exhausted). Needed for a correct "credit hoarding" test (#17 currently nets heads). | Med |
| Balance block per head and per part (Tax/Interest/Penalty/Fee/Others, x4) | Bajrang 24-25 closing: SGST 2512104, SGST Others 5000000 (total 7512335) | B (only `balance_total` summed) | Idle cash by head and by purpose at year-end; Others/Interest/Fee balances cannot fund tax; unrefunded excess cash (s.54(1) refund of cash balance) and opening carry-forward by minor head. | Med |
| Time of deposit | `18:11:35` (Tirupati 19/08/2025), `23:xx` possible | B | Deposit after bank cut-off / midnight vs filing time (same-day-payment timing for s.50 day count); weak on its own. | Low |
| Reporting date (by bank) | `19/08/2025` (blank for non-deposits) | B | Lag between bank debit date and ledger credit (when != Date). Interest runs to the ledger credit date; lag > 0 explains small interest days. | Low |
| Reference No. | `CK00BXCAY5`, `E2521165748463` (deposits); `AA...` (TDS credit or PMT-09 ARN); `DC...` (liability discharge) | C (F9 display, F9c ID match) | Prefix classification: bank-CIN deposits vs `AA` ARN (GSTR-7/8 TDS-TCS credit or PMT-09) vs `DC`/`DI` (DRC-03/return payment). Can drive PMT-09 netting (see caveats) and the deposit-vs-TDS split. | Med |
| Description = `TDS Credit` | Tirupati 16/04/2025 CGST 22,10,990 + SGST 22,10,990; total TDS credit FY25-26 ~ Rs 3.04 cr (CGST+SGST) | D for keyword scan (not classified) - effectively A for "TDS Credit" meaning | s.51 TDS credited to cash ledger by govt deductors (GSTR-7). Cross-check against GSTR-2A TDS table and against GSTR-1 B2B supplies to govt/PSU GSTINs: TDS credit with no matching customer in GSTR-1 = undeclared supply (s.37/s.51) or mis-credited TDS. Also share of liability met by TDS (not cash) for #9 "zero cash months". | High |
| Description = `Payment of Demand under Appeal` / `Voluntary payment` | Bajrang 05/07/2024 `DC0507240003288` Mar-19 SGST 1,08,592; Tirupati 4 voluntary debits Jun-25 | C (F9 counts them, no amounts per head/component) | Ties to Liability Ledger by Demand ID (`IP0507240000456`); 10% pre-deposit test (1,08,592 = 10% of SGST tax 10,85,916 -> s.107(6) compliant). | Med |
| Tax Period if applicable (on PMT-09 and DRC-03 rows) | `Mar-19` on appeal payment; `Mar-22` on Tirupati 23/09/2025 voluntary payment (tax period tagged "Apr 2021 to Mar 2022" in Liability Ledger) | D (monthly roll-up) | Lets DRC-03 payments be bucketed by period they relate to rather than month paid; Tirupati's Mar-22 voluntary payment (IGST 67,341 + interest 2,325) is a prior-period tax admission (s.73(5)/74(5)). | Med |
| Sr.No | `1..65` | D (row filter only) | Gap check of statement completeness; nil value. | Low |

Whole file used? Yes (D) for date, description, ttype, total, balance_total, tax_period. Everything else above is under-used.

---

## 2. Electronic Credit Ledger CSV

File read. Layout: one Credit/Debit block of 5 columns (IGST, CGST, SGST, CESS, Total) + Balance block of 5.

| Raw column | Example value (real) | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Credit/Debit - IGST / CGST / SGST / CESS (individual heads) | Bajrang 24-25 `DI0510240054335` 20/10/2024 debit IGST 34,625,759 CGST 0 SGST 0 CESS 10,620,665; Tirupati `DI0507250024878` Voluntary payment debit IGST 8,501,475, CGST 13,539,628, SGST 8,891,260 | B (`rec` has igst/cgst/sgst/cess; only `total` read) | Head-wise ITC accrual vs GSTR-3B 4C by head (engine ties only grand totals, F3/F4); head-wise utilisation order check (s.49A/49B, rule 88A: IGST first for IGST then CGST/SGST; CGST ITC cannot pay SGST). Bajrang 24-25 shows ITC exhausted in the wrong order only if heads are tested. Cess ITC of Rs 1.9-2.8 cr/month at Bajrang can only offset cess - a head check detects cess ITC used elsewhere / cess cash paid (Oct-24 Rs 38.95 lakh) while cess ITC idle. | High |
| Balance Available - per head | Tirupati opening IGST 19,935,139 / CGST 11,835,215 / SGST 7,186,848 / CESS 22,223; closing CGST 10,994,245 | B | Year-end idle ITC by head (Rs 10.99 cr CGST + Rs 10.99 cr SGST + Rs 2.49 cr IGST at Tirupati) vs cash paid same year (Rs 14-15 cr cash in Jul-25..Feb-26): shows ITC trapped in the wrong head while cash is paid in another (s.49(5), s.54(3) refund of unutilised ITC for inverted duty). Also persistent CESS 22,223 stuck balance. Only `bal_total` is used (F4). | Med |
| Description = `Payment of Demand under Appeal` (Debit) | Bajrang 05/07/2024 Mar-19 CGST 394,258 + CESS 132,518 | C (mis-classified by F9 as voluntary) | Pre-deposit paid out of ITC. Verify mode permitted for pre-deposit (s.107(6)/112(8) read with CBIC circular on pre-deposit via credit ledger) and that the Liability Ledger shows the matching `ITC` credit for the same Demand ID (it does: row 3 `DI0507240001391`). Reduces usable ITC; ITC reversal interplay. | Med |
| Description = `Voluntary payment` (Debit, ITC) | Tirupati 17/07/2025 `DI0507250024878` total Rs 3,09,32,363 | C (F9 counts only) | Rs 3.09 cr of ITC consumed to pay a Jun-25 DRC-03 liability. If the liability was an under-reporting of output tax, the ITC used to settle it is only valid if ITC was genuinely available at that time and interest must be cash (interest cannot be paid from ITC, only cash (s.49(3)/Rule 85)). Check that interest was paid in cash (it was: Cash ledger 25/02/2026). | Med |
| Transaction type + Description `ITC accrued through - Inputs` | 12 monthly credits, all "ITC accrued through - Inputs" | D (monthly credited). `accrued_desc` set built but unused | `accrued_desc` would reveal other accrual types (ISD, import, reclaim of reversal, RCM `... - Reverse charge`, refund re-credit). In these files only "Inputs" appears, so reclaimed 4D(1) is not visible as a separate ledger credit (cross-check the Credit Reversal statement, section 5). | Low |
| Date | `20/04/2024` (credit) / `20/04/2024` (debit) | D (tie-outs) | The credit entry date is the actual GSTR-3B filing date and the debit date is the liability-discharge date. Two independent sources for filing-delay (s.47 late fee / s.50 interest) beyond the 3B ARN date; engine never uses them for compliance dates. | Med |
| Reference No. | credit `AA050424144472I` (ARN of 3B), debit `DI...` | C | `AA`-ARN encodes tax period (AA+state+MMYY) - cross-links to the Liability Register ARN and Reversal statement; lets 3B revisions/duplicate ARNs be spotted. | Low |

---

## 3. Electronic Liability Register (Part I)

File read; only `total`, `ttype`, `description` ("reverse charge" keyword) and `date` are used (`liab_register_booked` in the walkthrough and F8). Tax Period is NOT a column (parser sets `col_taxperiod=None`).

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Interest (x4 heads), Debit rows | Bajrang 23-24 Jan-24 return: SGST Interest 1,569 (debit and matching credit by cash `DC0502240040757`) | B | Interest levied for late payment of the return (s.50); proves the filer knew of delay. Drives interest-vs-delay recompute (`compute_interest`). | High |
| Fee (x4 heads) | Bajrang 23-24: CGST 100 + SGST 100 in total (4 returns x 25) | B | Late-fee debits per return (s.47) - tie to Cash Ledger Fee credits and to computed late fee. | High |
| Penalty / Others | all 0 | B | Would carry s.122/125 penalty if any. | Low |
| Tax part per head, Total per head, Balance per head | Bajrang 24-25 Oct-24 return (`AA051024188343G`): IGST 19,217,485 / CGST 45,652,763 / SGST 45,652,763 / CESS 45,261,101 | B | Head-wise liability vs GSTR-3B 6.1 head-wise (engine compares only one grand total with the portal's 3B total). Detects CGST vs SGST asymmetry (SGST cash in Sep/Oct/Apr: 9,47,211; 1.14 cr; 1.72 cr) = SGST output > CGST ITC usable, fuel for s.49(5) analysis. Balance block should be 0 at every return end; a non-zero residual = unpaid liability (never tested). | Med |
| Ledger Used For Discharging Liability | `Credit`, `Cash`, `-` | B (`ledger_used` carried, never read) | Per-return cash/credit split by tax and by RCM vs non-RCM (RCM must be cash - s.49(4)). F8 infers it from the cash ledger instead; this column gives the authoritative discharge mode per liability row. | Med |
| Description (`Reverse charge and supplies made u/s 9(5)`, `Reverse charge`, `Other than reverse charge`) | 24-25: new description `Reverse charge and supplies made u/s 9(5)` | D for "reverse charge" substring | 9(5) e-commerce operator liability is lumped into RCM; separate by description to avoid RCM-vs-3.1(d) false variance. | Low |
| Reference No. with date != filing month | Bajrang 24-25 rows 44-45: `AA050325474894V`, `AA050325475003L` dated 31/12/2025, all-zero amounts | C/A (dated outside FY walkthrough months, so silently dropped) | Two Mar-25-period returns/amendments filed on 31/12/2025 (9 months after due date, ARN prefix `0325`). Late filings with nil effect still attract s.47 late fee (Rs 20/day nil cap Rs 500 per return) and indicate GSTR-3B/1-style belated filing or post-year amendment; never surfaced. | Med |
| Sr.No / Date for ARN-derived period | `AA050424144472I` dated 18/05/2024 | D (date) | Period can be derived from ARN (Apr-24) giving filed-late days = date - 20th of next month for every month, independent of the 3B file. | Med |

---

## 4. Electronic Liability Ledger (Part II, DRC / demand) - read, but only display

Used: `date, ref, tax_period (display), description + [Stay], ttype, total` in F9c, exact-ID matched to BO Profile DRC. Not used: everything below.

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Tax / Interest / Penalty parts per head, Debit rows | Bajrang `ZD050824020999Y` (29/08/2024, DRC-07, FY19-20): CGST tax 52,61,956 + interest 3,43,881 + penalty 5,86,888; SGST tax 64,75,792 + interest 3,43,881 + penalty 5,86,888; FY18-19 `ZD050524001618N` IGST tax 28,56,660 + penalty 2,85,666 etc. | B | Demand quantum split tax/interest/penalty (s.73 10% penalty vs s.74 100% penalty -> which section was invoked; Bajrang penalty = 10% of tax -> s.73). Needed for contingent-liability disclosure (R8) and Rule 142. | High |
| Balance block (all heads, all parts) | Bajrang closing outstanding: IGST 30,10,000 + CGST 72,90,387 + SGST 85,04,222 + CESS 13,25,175 = Rs 2.01 cr, all `Stayed` | B (nothing sums or reports outstanding demand) | Outstanding demand at FY-end (Rs 2.01 cr at Bajrang) vs BS provisions / contingent liabilities (`bs_pl_input` R8), vs pre-deposit made (Rs 2.97 lakh SGST+IGST for FY18-19 only; FY19-20 demand Rs 1.35 cr has NO pre-deposit in the ledger). Stay needs 10% pre-deposit under s.107(6); an un-appealed stayed demand is a red flag. | High |
| Relevant Demand ID / Liability ID | `ZD050524001618N`, `IP0507250001871` | C (F9c key) | ID prefix classifies: `ZD`=DRC-07 order, `IP`=voluntary payment liability (DRC-03), `DC`/`DI`=discharge legs. Graph DRC-07 -> payment (`IP0507240000456`) linking appeal payments to demands. | Med |
| Ledger used for discharging liability | `Cash`, `ITC` | B | Tirupati's Rs 14.76 cr voluntary tax was paid Rs 7.4 cr cash and Rs 3.09 cr ITC (rest cash on 25/07, 30/07): interest cannot go through ITC. Check mode legality per component. | Med |
| Type of transaction (DR/CR/RD/RF) | `Debit`, `Credit` (RD = reduction, RF = refund never occurs here) | D (sign only) | RD (reduction after appeal order) and RF would show appeal outcomes; currently collapsed into "Credit". | Low |
| Stay status | `Stayed` (Bajrang), `Un-stayed` (Tirupati) | C (appended to description in F9c) | Stayed demand (appeal pending) = recovery barred: no interest/penalty accrual; must be reflected in contingent liability, not provision. Never flagged. | Med |
| Tax Period (text: `Apr 2018 to Mar 2019`, `Jun 2025`, `Apr 2021 to Mar 2022`) | Bajrang FY18-19, FY19-20; Tirupati `Apr 2021 to Mar 2022` | C (not parsed; F9c just prints it) | Demand period -> previous-year exposure; Tirupati Mar-22 voluntary admission (IGST 67,341 + interest 2,325 on 23/09/2025). Not compared to the FY under review. | Med |

Tirupati key reconciliation (not built): sum of "Voluntary payment" tax credits for Jun-25 = CGST 7,38,19,171 + SGST 7,38,19,171 = Rs 14,76,38,342 matches the portal Comparison Summary Jun-25 shortfall -14,76,38,341 to the rupee (see section 6), and the portal's own "Payment made through DRC-03" column is blank. The engine's F9c compares only to BO Profile, so the closure of the biggest finding of the year (Jun-25 under-declared output tax, 94.7% of liability shortfall) is never shown.

---

## 5. Electronic Credit Reversal and Re-claimed Statement CSV (Bajrang 24-25 only)

Whole file unused (A). `gst_core.classify_folder` (`:1183-1193`) matches the first line against `"cash ledger"`, `"credit ledger"`, `"liability register"`, `"liability ledger"`; "Electronic Credit Reversal and Re-claimed Statement" matches none, so the CSV is dropped without a message (the `csvs` list is only reused at `:1219` to find the GSTIN). No adapter, mapping, spec, or grep hit for "re-claimed" anywhere in the repo.

| Raw column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| ITC Claimed (Table 4A(5)) per head | 20/09/2024 Aug-24: IGST 45,396,432.6, CGST 3,389,036.65, SGST 3,389,036.65, CESS 20,114,805.15 | A | Government's own record of 3B 4A(5); matches the Credit Ledger accrual (IGST 45,396,433 Aug-24) - a three-way tie (3B, 2B, ledger) per head. | Med |
| ITC Reversal (Table 4B(2)) Eligible to Re-claim | Sep-24 `AA050924262377X`: CGST 676,351, SGST 676,351, CESS 966,216 | A | Exact pool of reversible-and-reclaimable ITC per head and period. `gst_checks_flow.py:720-740` (F2 / closing section) explicitly says the reclaim "isn't trackable in this tool... treat as an upper bound until reclaim tagging exists" and hands over `closing_pool = T["b2"]`; this statement IS that tag. | High |
| ITC Reclaimed (Table 4D(1)) | Aug-24: IGST 12,099,914.44, CGST 148,790, SGST 148,790, CESS 1,806,103 (Rs 1.42 cr) | A | Opening 4B(2) pool of Rs 1.42 cr reclaimed in the first return of the series; Rule 37/37A/(s.16(2)/(4)) compliance and 4D(1) vs 2B check; ITC reclaimed must appear in the Credit Ledger credit - ledger Aug-24 credit IGST 45,396,433 = 4A(5) 45,396,432.6 so the reclaim is inside 4A(5) (verify not double counted: net fresh ITC = 4A(5) - 4D(1)). | High |
| Closing Balance (Opening + 4B(2) - 4D(1)) | 31/03/2025: CGST 676,351, SGST 676,351, CESS 966,216 (Rs 23.19 lakh) un-reclaimed since Sep-24 | A | Reversed ITC still not reclaimed at year-end (7 months): either condition (payment to supplier within 180 days, s.16(2)(b) second proviso / Rule 37) not yet met or forgotten. Interest s.50(3) at 24% runs from reversal-to-reclaim only if reversal was due to s.16(2)(b)/Rule 37 failure, otherwise none - needs the reason. Opening carry-in pool at 24/08/2024 (Rs 1.42 cr) proves earlier reversals sit outside this FY's data. | High |
| Return Period, Description (`Filing of Form GSTR-3B`), Reference No. (ARN), Date | `AA050824183714Z`, 20/09/2024 | A | Another independent source of the 3B filing date and ARN per month (the file starts 24/08/2024 - first return after the facility). | Low |
| Legal Name line | `BAJRANG MOTORS` | A | Taxpayer identity check (GSTIN/name vs rest of the folder); the file would otherwise be an orphan in `GSTIN` discovery. | Low |

---

## 6. Portal "Tax liability and ITC comparison" xlsx (8 sheets)

Only `Comparison Summary` is read (`parse_portal_comparison`, `portal_adapter.py`; spec says so and emits INFO I1401 "Sheets not used by any check"). From it 11 columns are mapped; of these only `gstr3b_liability` (D, vs Liability Register), `itc_3b_adj/unadj` (D, vs Credit Ledger accruals and #7 volatility), `gstr1_liability` and `itc_2b` (C, shown in walkthrough) are consumed. The shortfall / cumulative shortfall fields `diff_liability, cum_diff_liability, diff_itc_unadj, cum_diff_itc_unadj, diff_itc_adj, cum_diff_itc_adj` are canonicalised and returned in the dict but not read by any check (B) - they are recomputed by the tool from its own sources.

### 6a. Comparison Summary (all 15 columns)

| Raw column | Example | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Shortfall (-)/Excess (+) in liability; Cumulative shortfall (liability) | Tirupati Jun-25 -14,76,38,341 (cum stays -14.76 cr for the rest of the year) ; Bajrang 23-24 Jul-23 +3,85,22,094 then Aug-23 -3,83,72,625 | B | Government's own view of 3B vs GSTR-1 gap; engine does not alert on it (no finding is generated from `diff_liability`). The Jul-23/Aug-23 Bajrang swap is a timing shift of Rs 3.85 cr output tax between months (possible s.50 interest on the under-reported month). | High |
| Cumulative Shortfall as percentage (%) (liability) | Tirupati Jun-25 -574, Mar-26 -53.08 | A (unmapped; column 5 is not in `portal_mapping.json`) | Cumulative exposure as % of GSTR-1 liability: threshold for scrutiny selection (dept. uses % bands). | Med |
| Shortfall in ITC (excl. reversal) and cumulative; Cumulative % | Bajrang 23-24 Aug-23 -10,83,211 cum; Tirupati Total +36,053 | B / A for the % columns (cols 10 and 14 unmapped) | ITC claimed in 3B above/below 2B (s.16(2)(aa), 38, Rule 36(4)): Bajrang excess 3B vs 2B stays at -1,216,334 (3B claimed LESS than 2B = under-claim; no leakage). | Med |
| ITC claimed in GSTR-3B (considering reversal) / shortfall / cum. | Tirupati Mar-26 -12,66,596 | B (diff), D (value) | The "considering reversal" series exposes Rule 37/42/43 reversals: Tirupati Mar-26 reversal of Rs 12.67 lakh is visible only here (3B 4B(1)+(2) reversal not separately tested). | Med |
| (blank-formatted) Total row | `Total` | A (period regex skips it) | Portal's FY total as a reconciliation anchor for the tool's own month sums (catches a missing/duplicated month). | Low |
| Title block: `Report generated at: 20/04/2026 08:40 PM`, Trade name | | A (only GSTIN / Legal name / FY read into META) | Data currency: report date vs due dates - a report generated before later 3B amendments / DRC-03 postings is stale (explains why Tirupati's DRC-03 column is still blank on 20/04/2026 although the ledger shows DRC-03 payments from July 2025). | Low |

### 6b. Sheets never read (whole sheet unused, A)

| Sheet | Columns / content | Real-data example | What it could tell us | Priority |
|---|---|---|---|---|
| Tax Liability Summary | GSTR-1/IFF/GSTR-1A,2B/2A liability (IGST/CGST/SGST/CESS/Total) vs 3B paid vs shortfall vs `Payment made, if any through DRC-03` vs shortfall including DRC-03 vs `Percentage of liability paid (%) including DRC-03` | Tirupati Jun-25: GSTR-1 liab 15.58 cr vs paid 0.82 cr, % paid 5; `Jan-26` % paid 1237; Bajrang Jul-23 359%, Aug-23 58% | Per-head liability gap (3B vs 1) and the **DRC-03 payment column** (blank for all three taxpayers although Tirupati's ledgers hold Rs 14.76 cr DRC-03 for Jun-25). The % paid column is the department's own risk indicator (e.g. 5% -> auto-scrutiny). F9 notes mention this column "is read on the Annual Ledger Walkthrough" - it is not. | High |
| Tax liability | 3B 3.1(a) per head vs GSTR-1 (4A,4C,5,6C,7,9A..) per head, shortfall, cum, % | Bajrang Jul-23 CGST 3B 1.75 cr vs GSTR-1 0.32 cr (+1.42 cr); Jun-23 CGST/SGST 3B 16,000,522.75 vs 16,000,772.75 | Head-wise (CGST/SGST/CESS) mismatch with no netting; the tool's own #56 (head-wise 1 vs 3B) can be validated against the portal's figures. | Med |
| Reverse charge | 3B 3.1(d) per head vs 2B RCM (B2B, B2BA, CDNR), shortfall, cum %, | Bajrang 23-24 cumulative 93.42% of RCM liability has no 2B RCM invoice (3B 8,517.50 per head vs 2B ~560); Tirupati cum % 97.4% | RCM-vs-2B gap: RCM on supplies (legal fees, GTA) that suppliers did not report; the tool's F8 uses cash ledger instead, never the portal figure. | Med |
| Export and SEZ | 3.1(b) vs GSTR-1 6A/6B/9..: IGST, CESS | all zero (all three) | Zero-rated mismatch; empty here but independent evidence for the refund test (F9a). | Low |
| ITC (Other than IMPG) | 3B 4A(4)+4A(5)-4D(1) vs 2B per head, shortfall, cumulative, % and "considering reversal" block | Bajrang Aug-23 IGST 3B 33,223,969 vs 2B 34,197,865 (-973,896) | Head-wise ITC excess/under-claim (the Comparison Summary only has the grand total). Head-wise detection of IGST over-claim hidden by CGST under-claim. | Med |
| ITC (IMPG) | 3B 4A(1) vs 2B IMPG/IMPG-SEZ | all zero | Import ITC vs BoE; empty here. | Low |
| RCM_LIABILITY_ITC | 3B 4A(2)+4A(3) ITC vs 3.1(d) liability per head | Tirupati Jan-26 ITC 4,775 = liability 4,775, shortfall nil | s.16(2)(c)/49(4): RCM ITC cannot exceed liability paid in cash - tie each month to Cash Ledger `Reverse charge` debits (F8 recomputes). Nil gap here = clean. | Low |

---

## 7. Hand-typed bs_pl_input.py and the three masters

`bs_pl_input.py` (`BS_PL_DATA`, tagged `_gstin`; used only if tag == run GSTIN, else silently ignored; the shipped file is a different taxpayer `05AAECM6380J1ZA`, so for Bajrang/Tirupati/Jagnath **the whole file is unused (A)** and R0-R12 print "not tested"). When it is applied:

| Key | Status | Use |
|---|---|---|
| `total_assets`, `total_equity_liab` | D | R0 gate |
| `revenue_from_operations` | D | R1 vs GSTR-9C turnover |
| `reserves_and_surplus` + `net_profit_after_tax` | D | R12 roll-forward |
| `inventories` | C | F1b display, R5 INFO |
| `other_income`, `trade_payables`, `trade_receivables`, `fixed_assets_tangible`, `depreciation`, `non_current_investments`, `short_term_provisions`, `other_expenses`, `finance_costs` | C (INFO prose only) | R2-R10 give no numeric test, e.g. `trade_payables` should drive a 180-day ITC (Rule 37) test, `finance_costs` an interest-income/exempt test, `short_term_provisions` should be compared with the Liability Ledger outstanding demand (Rs 2.01 cr Bajrang). |
| `share_capital` | B (R11 fixed text, key never read) | none |

Masters:

| File | Columns | Status |
|---|---|---|
| HSN/SAC master (`HSN_SAC_default.xlsx`, sheets `HSN_MSTR` 21,936 rows, `SAC_MSTR` 682 rows) | `HSN_CD`/`SAC_CD` (code) and description | Both columns read; code = D (A7 validity), description stored (B/C - not used for description-vs-HSN mismatch test, e.g. "cement" goods under a service SAC). No rate, no UQC, no effective-from/to date columns exist in the file, so no rate-by-date validation is possible from it. |
| `GSTR_2B_Blocked_ITC_Master_Updated.xlsx` (sheet `2B Keyword HSN Master`, 415 rows) | `Category` (D), `Search keyword` (D), `Indicative HSN/SAC` (`hsn_raw` loaded, never used - B) | HSN column unused; code itself states 2B has no HSN column, but GSTR-2 purchase-register / e-way bill inward HSN (available in EWB inward) could use it: `hsn_raw` x EWB inward HSN = s.17(5) screening on goods lines. Medium. |
| `Machinery_HSN_Master_v2.xlsx` (sheets `Ch.84`, `Ch.85`, `Master HSN Flag List`, `Read Me`) | Only the 7-column `Master HSN Flag List` sheet is loaded: `HSN Heading` (D), `Chapter`, `Description`, `Category`, `Flag` (D); `S.No` and `Match Rule` (col 7) unused (B); other 3 sheets never read (A) | `Review` flag (Parts & Accessories) is surfaced but no value threshold; `Category` ("Machine / Equipment" vs "Parts") not used to split capital goods (s.18(6) / Rule 43) from consumables. Low. |

---

## 8. Top 5 opportunities

1. **Read the Interest / Fee / Penalty / Others components of all four ledgers** (currently summed away into `total`). Immediately gives: late fee actually levied (cross-check `month_filing_compliance`), interest actually paid (Tirupati Rs 2.09 cr on Rs 14.76 cr DRC-03 tax = ~290-day equivalent against payments made 17-30 Jul 2025; Bajrang SGST interest Rs 1,569), "Others" parked cash (Bajrang Rs 50 lakh SGST), and removes the PMT-09 / TDS contamination in #9, #17, #47/#55. Rule: s.47, s.50(1)/(3), s.107(6).
2. **Tie Part II Liability Ledger voluntary/DRC payments to the portal Comparison Summary shortfall** and populate the (blank) "Payment made through DRC-03" column of the unread `Tax Liability Summary` sheet. Tirupati Jun-25: shortfall Rs 14,76,38,341 vs Part II Jun-25 voluntary tax credits Rs 14,76,38,342 - an exact closure that converts the year's biggest red flag into an explained under-reporting plus interest; same pattern for Mar-22 prior-period admission. Also compute Bajrang outstanding stayed demand Rs 2.01 cr (balance block) vs BS provisions (R8) and pre-deposit coverage (FY19-20 demand Rs 1.35 cr has none).
3. **Read the Electronic Credit Reversal and Re-claimed Statement** (currently dropped by `classify_folder`): it supplies the 4B(2) reclaim tagging that `gst_checks_flow.py:720-740` says does not exist; gives Rs 1.42 cr reclaimed Aug-24 and Rs 23.19 lakh still un-reclaimed at 31/03/2025 (Rule 37 / s.16(2)(b) / s.50(3)) and per-head 4A(5) tie to ledger accruals.
4. **Use head-wise (IGST / CGST / SGST / CESS) values in the credit, cash and liability ledgers** (already in `heads`/`igst..cess`, never read): s.49/49A/49B/49(5) cross-utilisation test, ITC idle in one head while cash is paid in another (Tirupati closing CGST+SGST ITC Rs 21.99 cr + IGST Rs 2.49 cr while Rs 14-15 cr cash was paid; Bajrang Oct-24 SGST cash Rs 1.14 cr, cess cash Rs 38.95 lakh), head-wise 3B 6.1 vs Register tie-out, and Cess ITC-only-for-cess rule.
5. **Classify the cash-ledger Description / Reference prefix** (`TDS Credit`, `Amount deposited`, `AA` PMT-09, `DC`/`DI` discharge, `Payment of Demand under Appeal`, `Voluntary payment`): TDS credit of ~Rs 3.04 cr at Tirupati cross-checked to GSTR-2A TDS table and to GSTR-1 govt customers (s.51), PMT-09 netting before "cash paid" is computed, appeal-payment not mislabelled DRC-03, and `31/12/2025` nil returns for Mar-25 in the Bajrang register flagged as late filings (s.47).


---

# Unused-input audit: BO Profile, GSTR-9, GSTR-9C, Table 8A (R9_8A)

Scope: `main gst tool/` engine (identical to `docs/py/core`). Real data read: `input_files_bajrang_motors{,_24_25}/`, `input_files_tirupati_cement/`, `pdf to excel/` (Jagnath BO), `docs/sample_data/sample_bo_profile.pdf`.
No input file was modified. Amounts in BO are lakh; GSTR-9/9C/8A are rupees.

Status key: **A** = never read (sheet/row/column not in the canonical layer or not parsed at all); **B** = read/canonicalised but no check/finding/report cell uses it; **C** = display only (listed on a sheet, no rule); **D** = used in a check (mentioned briefly).

How consumption was verified (grep of the whole engine, not assumed):
* `parse_bo_profile` output is consumed only by: `check_reciprocal_trading` (#3), `check_ghost_supplier_cluster` (#10/#19), `check_exempt_turnover_rule42` (#25), `check_related_party_supercluster` (#43), `check_bo_high_risk_supplier_itc` (#54) in `gst_checks_hsn_fraud.py`; `bo_financial_rows` -> F10, F9/F9a/F9c, F11/F12 name lookups in `gst_checks_flow.py`; DRC in `check_bs_pl_rules` (`gst_checks_forensic.py:428`); `write_fy_total_vs_bifa`, related-party and proceedings listings in `gst_report.py`. The `demographic` dict has **no consumer at all** (`grep demographic` hits only the parser and an empty-shape default in master_build.py:2053).
* `parse_gstr9` keys consumed: `table6a_total` (R14, `gst_checks_forensic.py:154`) and `table13_*` / `table12_*` (ITC annual summary, `gst_checks_flow.py:718-784`). Nothing else (`grep table4_|table5_|table7_|table8_|table9_` finds no consumer outside parser/adapter/tests).
* `parse_gstr9c` keys consumed: `turnover_declared_gstr9` (BS/PL R1 + materiality in R14), `exempt_nil_nongst_adjustment` (R13, F2b), `itc_per_books` (R14). Everything else unused.
* `parse_table_8a` output consumed only as `totals.total` / `totals.no_reason_breakdown` (R14, `gst_checks_flow.py:795-823`). No check reads a single invoice row (`b2b[]`) or note row (`cdnr[]`).

---------------------------------------------------------------------------------------------------

## 0. Structural findings that change the picture (read these first)

| # | Finding | Evidence | Status | Why it matters | Priority |
|---|---|---|---|---|---|
| S1 | **The main taxpayer's BO Profile (Bajrang, both 31_08_2026 and 14_09_2026 files) has 12 of its 16 tables stacked inside the single sheet "ITC Passed On" (rows 15-140)**. The reader looks tables up by sheet name, so ITC Received, Top-10 Beneficiaries, Top-10 Suppliers, EWB, E-Invoice, Refund, Related-party (received/passed), Appeals, Cases, DRC, Transfers all come back EMPTY. The adapter only prints `I1202 sheets ... not in this file`. | `parse_bo_profile(bajrang 14_09_2026)` returns `top_suppliers=0, top_beneficiaries=0, related_*=0, drc_payments=0, appeals=0, ewb_by_fy={}, einv_by_fy={}`. Tirupati/Jagnath files (30 sheets) fill all of them. | A (whole stacked block) | Every BO-driven check (#3, #10/#19, #43, #54, F9c DRC tie-out, F12, proceedings listing) silently runs on nothing for the largest taxpayer. Real content being dropped: 7 of the top-10 beneficiaries are ALSO top-10 suppliers (Indraprastha Automobiles 07AABCI6357A1ZV: passed 131.11, received 165.28 lakh, risk HIGH; 09AACFA8668A2ZH; 05AAACD8573A1ZZ; 09AALCA8926F1ZC; 05AAKCD9198F1Z8; 09ADYFS1339P1ZZ; 09AABCV1426F1ZK) - #3 would flag all seven; DRC demands `ZD050824020999Y` (29-08-2024: CGST 52.62 / SGST 64.76 / other 18.62 lakh) and `ZD050524001618N`; 3 appeals; related supplier Kaveosoft (05AAJCK6991Q1ZK) cancelled 2022-08-31 yet 5.64 lakh ITC received after cancellation. The same DRC ids exist in `Electronic_Liability_Ledger.csv` / `Electronic_Cash_Ledger (1).csv`, so F9c will say "NOT found in BO Profile DRC" for ids that ARE in the BO file. | **High** |
| S2 | Same stacking corrupts what IS read: for Bajrang `itc_passed_by_fy` is **0.0 beneficiaries / 0.0 ITC for all 10 FYs** (true FY24-25: 1006 beneficiaries, 2077.21 lakh) because later FY-first rows (ITC Received block, EWB, e-invoice, refund = 0 rows) overwrite earlier ones ("later duplicate FYs overwrite"). Also a junk FY `2008-9` row is carried in `financial_by_fy`/`bifa_by_fy`. | adapter run above | B (wrong data, unused) | Detect "FY seen twice" / headings found in the wrong sheet and split the stacked sheet by its section titles (the titles "Top 10 Beneficiaries based o...", "EWB Related Information", ... are present as rows). | **High** |
| S3 | BO **E-Invoice sheet columns are swapped relative to their headings**: under "Active Assessable Value" sit TAX figures and under "Active Taxes" sit TAXABLE values. Proof: Tirupati FY25-26 values 3259.6 / 18108.91 (3259.6 = 18% of 18108.91 = ITC passed total); FY24-25 1717.44 / 9541.34 (= tax liability / turnover); Bajrang FY24-25 2003.63 / 5648.23 (tax is the smaller). | `einv_by_fy` | B | Any future use must swap; as read today `active_taxes` > `active_assessable_value`. | Med |
| S4 | BIFA column **"Liability as per EWB" repeats the value of "Difference (Supply value to TDS/TCS - 3B)" in alternate years** (Bajrang FY25-26, 23-24, 21-22, 19-20, 17-18: both -30367.01, -17947.03, -8462.39, -9134.62, -7760.64; in the other years it is a genuine outward-EWB tax e.g. 908.84 = EWB sheet outward tax). `write_fy_total_vs_bifa` compares it blindly with the recomputed outward-EWB tax and would print a false REVIEW. | BIFA sheet, `gst_report.py:975` | C | Guard: if `liability_ewb == diff_supply_tds` treat as not available. | Med |
| S5 | Related-party tables have the FY printed only on the first row of each FY group; the adapter keeps `fy=None` for the continuation rows (Tirupati: 12 of 22 "ITC Received from Related" rows have `fy None`). | adapter output | B | Forward-fill FY; today the ghost/supercluster checks say "SUMMED ACROSS ALL YEARS" because FY is lost. | Med |

---------------------------------------------------------------------------------------------------

## 1. BO Profile / 360 profile (`boprofile_mapping.json`, `boprofile_adapter.py`)

Sheets in the real files: Bajrang 13 sheets (rest stacked, S1); Tirupati and Jagnath 30 sheets.

### 1.1 Sheet-level map

| Sheet | Read? | Mapped to canonical? | Used by a check? |
|---|---|---|---|
| Demographic Details | yes | `BO_DEMOGRAPHIC` (whole label/value dict) | only GSTIN/Legal/Trade name (D); all other 20 labels B |
| Flagged in Lead Based Dashboard | **no** (I1203) | no | A (whole sheet) |
| Pre-GST Details | **no** | no | A (whole sheet) |
| Places of Business | **no** | no | A (whole sheet) |
| Bank Account Details | **no** | no | A (whole sheet) |
| Member Details | **no** | no | A (whole sheet) |
| Shared Entity / Shared Members | **no** | no | A (whole sheets; header-only in all 3 files) |
| Financial Information | yes | 10 of 14 columns | D (#25: turnover, taxable) / C (F10) / B (rest) |
| BIFA Specific Information | yes | 8 of 12 columns | C (sheet "FY Total vs BIFA", no Finding) / B / A |
| ITC Passed On / ITC Received | yes | TOTAL pair only | B (never consumed) |
| Top 10 Beneficiaries / Suppliers | yes | 6 of 7 cols | D (#3 #10 #43 #54) / C / B |
| EWB Related Information | yes | all | **B** (never consumed) |
| E-Invoice Related Information | yes | all | **B** (never consumed) |
| Refund Details | yes | all | claimed D (F9a only), rejected/sanctioned B |
| ITC Received from / Passed On to Related - Cancelled | yes | 8 of 11 cols | D (#10 #43), C (F12 text) |
| HSN as per REG01; Top 10 HSN (GSTR-1 / EWB / E-Invoice) | **no** | no | A (header-only in all files) |
| Export - ICEGATE | **no** | no | A (header-only) |
| Appeal / Case / Transfer Information | yes | partial | C (listed only) |
| DRC Payment Information | yes | all | D (BS/PL, F9c) |
| Return Filing Details GSTR-3B / GSTR-1 | **no** | no | A (header-only: Date of Filing, Return Period, IP Address) |
| (PDF page header) Data Refresh Date / Report generated on / by | not in xlsx | no | A |

### 1.2 Demographic Details (status B for everything except GSTIN / names)

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Overall Risk Score | Tirupati `7.50 HIGH`; Bajrang `4.50 MEDIUM`; Jagnath `1.08 LOW` | B | The department's own composite risk for the taxpayer. Put on the cover page and use it to set the scrutiny depth / to explain why a taxpayer is likely already on the officer's list. Trend vs prior BO files (Bajrang 31_08 vs 14_09 files exist). | **High** |
| Nature of Business | `Services` for BOTH Bajrang Motors (vehicle dealer, 28% goods + cess) and Tirupati Cement Products (cement goods) | B | Registered nature of business inconsistent with supplies (goods). Compare to GSTR-9 Table 4 (rate mix) / HSN (goods vs SAC); Sec 25 registration particulars; if a taxpayer is registered only for services but sells goods, ask for amendment history. | Med |
| Constitution of Business | `Partnership` | B | Partners = Member Details; ties to the signatory name in GSTR-9/9C Verification (SANJAY AGRAWAL, Partner). | Low |
| Income Tax Paid Range (FY) + GTI Range | Bajrang `2L-5L 2021-22` with `GTI >50 lakh`; FY21-22 BO turnover 8485 lakh | B | Income-tax paid of 2-5 lakh on 84.9 Cr turnover is a ~0.04% tax/turnover; and GTI ">50 lakh" is incompatible with only 2-5 lakh tax for a firm (30%+). Use for profit-margin plausibility vs GSTR-9C audited P&L, BS/PL rules (R1..), and as a Sec 73/74 lead (suppressed profits / inflated purchases = inflated ITC). Tirupati `>20Lakhs` is consistent. | Med |
| Last GSTR3B / GSTR1 Filed (YYYYMM) | `202607` (Bajrang), `202606` (Tirupati) | B | Completeness check: the engine's last supplied period vs the department's last filed period (are we missing a month; was a return filed after the extract?). | Low |
| No.of GSTR3B / GSTR1 defaulted | `0` | B | Cross-check with the engine's own late-filing/late-fee sheet (derived from ARN dates). A non-zero here with no late fee in the files = unfiled/NIL gaps. | Med |
| QRMP | `No` | B | Compare with the engine's inferred QRMP (from 3B sheet count); a mismatch means wrong period unit and wrong due dates. | Med |
| E-invoice Enablement Status | Bajrang `Y`; Tirupati **blank** although BO e-invoice table shows e-invoices since 2020-21 | B | Blank/N with e-invoices issued, or `Y` with B2B turnover outside e-invoice, is an e-invoicing-compliance lead (Rule 48(4), turnover > threshold). | Med |
| Registration Date / End Date / GSTIN status / Taxpayer type / Category / Jurisdiction / Level | `2017-07-01`, `NA`, `Active`, `CENTER`, `Uttarakhand \| MEERUT \| DEHRADUN \| DIVISION HALDWANI` | B | Registration age vs ITC from newly registered suppliers; jurisdiction (Centre) for who issues SCN. Low individually. | Low |
| Mobile Number / Email | `9999404546`, `accounts@tcp.in.net` | B | The portal's "Related Parameter" (Mobile/Email) in the related-party tables is based on exactly these; allows own-side confirmation when a supplier's contact (from 2A/e-invoice) matches. | Low |

### 1.3 Sheets never read (status A)

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Flagged in Lead Based Dashboard: Spike in ITC Passed / Spike in ITC Utilized / Missing Taxpayer / High ITC Utilization / Suspicious Transactions / Suo Moto Cancelled Related PAN / Related Member PAN (count of appearances in last 6 months) | Tirupati: `Spike in ITC Utilized = 1`, rest 0; Bajrang all 0 | A | These are the department's analytics-lead hits. A non-zero count is a direct "why was this taxpayer selected" signal; "Suo Moto Cancelled Related PAN/Member PAN" = promoter linked to cancelled GSTINs (Rule 21 / Sec 29(2)). Cheap, high-signal: show on cover and raise an INFO/REVIEW Finding when > 0; correlate "Spike in ITC Utilized" with the month-wise ITC-utilisation checks. | **High** |
| Member Details (name, PAN, type) | Bajrang: Sanjay Kumar Agrawal ADCPA0656P (Primary Signatory, Promotor), Dilip Kumar, Inder Lal; Tirupati 4 promoters | A | Characters 3-12 of every supplier/customer GSTIN is the PAN: match promoter/partner PANs (and their relatives' surnames) against top suppliers, 8A suppliers, GSTR-1 buyers = related-party transactions (Sec 15(2)(b), Sch I; Rule 28 valuation) and circular-trading rings. Tested on the real files: 0 PAN hits today, so this is a ready rule that would currently PASS - but nothing performs the test. Also check the GSTR-9/9C signatory against Member list. | Med |
| Shared Entity / Shared Members (shared mobile / email / PAN / address with other GSTINs, with their risk score and status) | header-only in all 3 files | A | When populated: other GSTINs under the same mobile/email = ghost network; feed them into #10/#19/#43 PAN clusters and to the 2A/2B supplier list (are any of them suppliers?). | Med |
| Pre-GST Details (Central Excise no., CST no., CIN, **IEC**, Service Tax no., TIN under VAT) | Tirupati IEC `0514024607`, ST `AACFT2702LSD001`, CST `05016429885`; Bajrang TIN `05001453764` | A | IEC with export turnover 0 / no import ITC in 3B 4(A)(1) / GSTR-9 6E-6F=0: dormant IEC; IEC + 6E nonzero but no Bill of Entry data = check. Service-tax/excise numbers: continuity of pre-GST credit (TRAN-1, GSTR-9 6K-6L = 0 in all files). | Low |
| Places of Business | Tirupati principal Haldwani; **additional** `E-165 Gumani Wala ... Rishikesh` | A | Compare additional places with EWB from/to places and with the same-PAN GSTINs in the related-party tables (Tirupati's own `07AACFT2702L1Z9` Delhi and `02AACFT2702L1ZJ` Himachal GSTINs show inter-GSTIN ITC flows). Place of supply / ISD / branch-transfer review (Sec 25(4), Sch I(2)). | Med |
| Bank Account Details | HDFC `01298710000055`, New Delhi Janakpuri branch for a Haldwani firm; Validation "Valid & PAN Matches with NPCI PAN" | A | Bank in a different state; validation status (a mismatch is a red flag); account numbers to match refund/DRC payment accounts and any payee in ledgers. | Low |
| Return Filing Details GSTR-3B / GSTR-1 (Date of Filing, Return Period, IP Address) | header-only in all 3 files | A | When populated: filing dates vs due dates (late fee Sec 47, interest Sec 50), **same IP address filing for unrelated taxpayers** (shared operator/ghost cluster), burst filing. | Med |
| Top 10 HSN as per GSTR-1 / EWB / E-Invoice; HSN as per REG-01; Export - ICEGATE (shipping bill, scroll amount) | header-only in all files | A | HSN mix vs GSTR-1 Table 12 and registered HSN (REG-01); ICEGATE scroll vs GSTR-1 Table 6A / GSTR-9 5A (zero-rated claims) and IEC. Empty here, build mapping so a populated file is used. | Low |
| PDF page header: Data Refresh Date, Report generated on, Report generated by | `Data Refresh Date: 2026-08-01`, `Report generated on 2026-08-08` (sample PDF; not in the portal xlsx) | A | `gst_report.py:1023` says the tool "doesn't capture the BO Profile's own generation date" - staleness of BIFA vs late-filed returns. The sample pdf has it; the Table 8A "Read me" sheet also has a "Date of generation" (see section 4). | Low |

### 1.4 Financial Information (per FY)

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Export Turnover | `0` all years | B | Cross-check with GSTR-9 Table 5A/4C, GSTR-1 Table 6A/6B, IEC (Pre-GST), refund claims. | Low |
| ITC Utilization % | `0.0098` (Bajrang), `0.0072` (Tirupati) | B | ITC utilization 97-99% = near-zero cash tax: a ratio the department's lead engine ("High ITC Utilization") uses. Use as a time-series (Tirupati FY25-26 0.72% cash 785.89 lakh is the only year with real cash) and as a trigger for deeper ITC checks. | Med |
| Cash Utilization % | `0.0003` | **A** (no mapping) | Same, complementary. | Low |
| Effective Tax Rate % | Bajrang `0.3133`..`0.3684`; Tirupati `0.18` | B | Effective rate vs statutory mix: Bajrang 31-37% includes cess (28%+cess on vehicles) -> sanity check on GSTR-9 rate-wise table (9C Item 9); Tirupati exactly 18% every year -> single-rate. A dip vs prior year with rising turnover = under-reporting lead. | Med |
| ITC Reversal | Bajrang FY24-25 `29.93`; Tirupati FY25-26 `16.41` | B | Equals GSTR-9 Table 7 total (29.93 lakh = 2,992,859 in FY24-25) and should be compared with GSTR-3B 4(B) and the Electronic Credit Reversal and Re-claimed Statement. Year with reversal 0 but exempt turnover = Rule 42 lead. | Med |
| **RCM "Liability" and "Cash"** (unlabelled duplicate headings; the sample PDF shows them under group "RCM") | Bajrang FY24-25 `0.22` / `0.22` = GSTR-9 4G RCM tax 22,320; Tirupati FY25-26 `21.06` / `21.06` | **A** (not mapped; headings "Liability"/"Cash" match nothing) | Department's RCM liability and cash discharge per FY. Tie out to GSTR-3B 3.1(d)/6.1 RCM and GSTR-9 4G / 6C-6D; Tirupati 21.06 lakh RCM in FY25-26 (vs 4.27 in FY23-24) is a jump to inspect (services: GTA, legal, import of services). | **High** |
| Turnover / Taxable Turnover / Tax paid in cash / ITC availed / Total tax liability / Tax paid by ITC | e.g. Bajrang FY24-25 25165.85, 9270.19 | D (#25) / C (F10 table) | Not compared to GSTR-9 / 9C / 3B. Concrete gaps: BO turnover 25165.85 lakh vs GSTR-9 5N 25159.32 vs 9C audited 25159.81 (FY24-25, 6.5 lakh); FY23-24 17954.18 vs 17942.63 (11.55 lakh); BO ITC availed 8844.20 = 9C 12A (net of reversals) vs GSTR-9 6A 8874.13. A 3-way tie (BO vs 3B vs 9/9C) is not done anywhere. | **High** |
| 10-year history (FY 2017-18 ..) | Bajrang ITC availed 2663 -> 9277 lakh | C/B | Only rows for the scrutinised FY / a trend table are used; year-on-year jump detection (Tirupati turnover 4877 -> 9541 -> 15451) is not a rule. | Low |

### 1.5 BIFA Specific Information (department's own analytics)

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Liability as per GSTR1 / Difference Liability ("Supply Tax Deficit") | **Tirupati FY25-26: 4257.55 vs 2781.17 = 1476.38 lakh**; FY24-25 0; Bajrang FY20-21 510.19 vs 2237.34 (-1727.15), FY19-20 -1227.52 | B (mapped; `liability_gstr1`, `diff_liability` never read) | The department already sees GSTR-1 liability exceeding GSTR-3B by 14.76 Cr for Tirupati FY25-26. It is exactly equal to the sum of the three July-2025 voluntary DRC-03 payments in the same BO file (613.06 + 213.40 + 649.92 = 1476.38 lakh, ids IP0507250001871, DC0507250079030, DC0507250088643) - i.e. under-payment in 3B later cured with DRC-03, plus a 209.42 lakh "other than tax" (interest/penalty) DRC on 2026-02-25 and a pending **Section 74** transfer (see 1.9). Sec 73/74, Sec 50 interest, Sec 122 penalty. The engine's own GSTR-1 vs 3B comparison should be reconciled to this department number (are our files complete?). Only `liability_gstr3b` is used. | **High** |
| Difference ITC ("Excess ITC claimed (other than RC)") | Tirupati FY24-25 `27.87` lakh (availed 1540.98 vs accrued in 2B/2A 1513.11); Bajrang FY20-21 `76.50` | B | Excess of 3B ITC over 2A/2B: Sec 16(2)(aa)/Rule 36(4)/Sec 42 reversal with interest. Compare to the engine's own month-wise 3B vs 2B excess (A5 note says both exist, nothing is flagged). | **High** |
| Supply value to TDS/TCS taxpayers + "Difference (supply to TDS/TCS taxpayers - supply per GSTR3B)" (group "TDS/TCS & GSTR3B mismatch") | **Tirupati FY23-24: 11142.00 vs 3B 4876.75 = +6265.25 lakh**; FY24-25 +2621.40; Bajrang tiny (7.15) | **A** (columns 8 and 10 of the BIFA table are not in the mapping) | Deductors under Sec 51 (govt. bodies - Tirupati sells to UUSDA etc.) reported more supplies to the taxpayer than the taxpayer's 3B turnover: possible turnover suppression. Cross-check with GSTR-1 B2B to govt GSTINs (05MRTU..., 05AAALU...), GSTR-7 credits in ledger/2A (TDS credit), Table 4 of GSTR-9. A strong department lead the engine ignores. | **High** |
| Liability as per EWB / Difference of Liability EWB vs 3B | Bajrang FY24-25 `908.84` / `-8361.35` | C (compared to recomputed outward-EWB tax, sheet only) | Data-quality bug S4. Use only when `ewb != tds-diff`. | Med |
| Supply value as per GSTR3B | = BO turnover | B | Redundant with Financial.turnover; confirms BO turnover is the 3B outward figure, which is why it differs from GSTR-9 5N by the amendments/credit-note lag. | Low |
| ITC Availed in R3B / Accrued in R2B-R2A; Liability as per GSTR3B | | D/C | used in the BIFA sheet only. | - |

### 1.6 ITC Passed On / ITC Received (INTRA / INTER / TOTAL groups)

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| TOTAL No. of beneficiaries / ITC passed; TOTAL No. of suppliers / ITC received | Tirupati FY24-25: 5 / 1717.44; 419 / 1535.34 | B (`itc_passed_by_fy`, `itc_received_by_fy` read, never used; Bajrang's corrupted, S2) | ITC received per department (2A/2B based) vs the engine's 2B, Table 8A (Bajrang FY24-25: BO 290 suppliers / 8760.12 lakh vs 8A 289 distinct suppliers / 8782.86 lakh) and BO ITC availed (8844.2): completeness test for the supplied 2B/8A, and "ITC claimed above ITC received" lead. ITC passed total = output tax (Tirupati 1717.44 = tax liability). | Med |
| INTRA / INTER-state split (counts and ITC) | Tirupati FY25-26: intra 1196.45 (346 suppliers), inter 1154.00 (192) | **A** (explicitly dropped; parser says TOTAL only) | Inter-state ITC share and supplier count vs GSTR-3B 3.2/IGST ledger; large inter-state supplier base with small EWB trail = bill-trading lead. | Low |

### 1.7 Top-10 Beneficiaries / Suppliers

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| `status`, `reg_start`, Registration End Date (not mapped) | Tirupati supplier 05AABCI8790B1ZO reg `2022-12-28`, risk HIGH, **name blank**, 276.92 lakh; Kalyan Associates reg 2022-06-16 HIGH 65.76; Bajrang supplier 05AAKCD9198F1Z8 reg 2024-07-22, 128.21 lakh | B (`reg_start`); A (end date) | Supplier registered shortly before large ITC / registered after the taxpayer started buying; Sec 16(2)(c) genuineness, Sec 29 retrospective cancellation risk. Compute months between `reg_start` and the first 2B invoice from that GSTIN. | Med |
| `risk` for Top-10 BENEFICIARIES (customers) | Bajrang: Indraprastha `10.00 HIGH`, Kamlesh Autowheels `7.90 HIGH`; Tirupati govt entities `0` | B (only suppliers' `risk` is tested, #54) | High-risk customers to whom ITC/tax is passed: outward-side risk (fake sales/ fictitious customers); show customers' risk in GSTR-1 buyer concentration (F11). | Med |
| `amount` (last-12-months ITC) vs the FY totals | | C (printed next to F11 tables) | Fine; but the BO list is "last 12 months" - not time-aligned to the scrutinised FY (the code itself says "different period basis"). | Low |
| GSTIN-shape anomalies | `05MRTU01195F1DS` (govt body, non-PAN-form GSTIN) | B | `gstin_pan()` on such ids gives nonsense PAN in clusters #10/#43; exclude. | Low |

### 1.8 EWB Related Information and E-Invoice Related Information (status B, whole tables never consumed)

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Inward Supply / Inward Supply Tax / Number of Inward EWB | Bajrang FY24-25: 21029.12 / **7665.00 lakh** / 10675; Tirupati FY25-26 5719.72 / 985.48 / 7067 | B | Department's EWB-based view of goods inward: compare inward EWB tax with ITC availed. Bajrang: 8844.2 ITC vs 7665 EWB tax = 1179 lakh (13%) of goods-type ITC with no EWB trail; Tirupati FY25-26: 2361.78 ITC vs 985.48 (1376 lakh). Also tie to the engine's own inward EWB file (completeness of the supplied EWB). | **High** |
| Outward Supply / Outward Supply Tax / Number of Outward EWB / Ratio Inward:Outward | Tirupati FY25-26: 58.84 lakh, 8 EWBs, ratio 97.2 vs turnover 15450.95 lakh; FY24-25 ratio 40.13 | B | Goods inward 100x goods outward: either services/works-contract supplier (Nature "Services"), or ITC without corresponding outward goods movement (inflated ITC, fake inward). Rule 138/Sec 16 genuineness; compare ratio trend and with HSN. | **High** |
| E-Invoice: Count of Active / Cancelled E-Invoices; Cancelled Assessable Value / Cancelled Taxes | Bajrang FY25-26 active 17490, cancelled 85; Tirupati FY22-23 3 cancelled, 83.46 lakh | B | The engine has a "Cancelled E-Inv" pipeline that does not use these; cancelled e-invoices after 24h (Rule 48(...)) and cancelled-vs-GSTR-1 credit notes; active e-invoice **tax** (swapped column, S3) vs GSTR-1: Tirupati FY25-26 e-invoiced tax 3259.60 lakh already exceeds 3B liability 2781.17 (confirms the GSTR-1 side of the 1476.38 lakh BIFA gap); e-invoiced value 18108.91 > 3B turnover 15450.95. | **High** |

### 1.9 Proceedings, related parties, DRC

| Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Transfer Information: **Section Informati on**, Initial Module Name, Recommendation Remarks, Target Case ID | Tirupati: `Section 74`, Enforcement -> Adjudicating Authority, `Pending for action by tax officer`, source case AD051125004285L | **A** (mapping keeps date, ids, source/target module, status only) | **Section 74 (fraud/wilful misstatement) proceedings at the adjudicating authority.** Elevate to a top-of-report Finding; the section decides the penalty (100%) and the limitation (5 years). | **High** |
| Case Information: SCN number/date, tax period from/to, CGST/SGST/IGST/Cess, Other than tax | header-only in the 3 files | B (amounts, SCN fields mapped, never shown); `Tax Period To` and `Other Than Tax` A | When populated: demand quantum vs the engine's own exposure, limitation check on SCN date. | Med |
| Appeal Information: Tax Officer Name, Order Number, Selected Declaration | Tirupati `2023-24, AD0509230044658, Adjournment Granted`; Bajrang 3 appeals (stacked, not read) | C (listed); officer/order B; declaration A | Appeal vs DRC "Payment of Demand under Appeal" (pre-deposit 10%, Sec 107(6)) rows - Bajrang DRC `DC0507240003288`/`DI0507240001391` on 2024-07-05 match appeal AD050724001087F. | Med |
| DRC Payment: payment method "Paid by ITC" for voluntary payment; `other` (Other-than-tax = interest/penalty) | Tirupati 2025-07-17 "Paid by ITC" CGST 135.4 lakh (+SGST) inside a 613.06 lakh voluntary payment; 2026-02-25 `Other than Tax` 209.42 lakh in cash | D (BS/PL, F9c) but `method`/`other` B | ITC used to discharge a voluntary payment for short-paid output tax: Sec 49/86B interplay and Rule 142(2)/(2A) - reduce eligibility; the 209.42 lakh "other than tax" = interest/penalty paid, size of the underlying default. Also demand rows with CGST != SGST (Bajrang ZD050824020999Y: 52.62 vs 64.76) - impossible for a pure intra-state demand, flags a mixed or mis-allocated demand. | Med |
| Related party: `related_parameter` (Mobile / Email / PAN,Mobile,Email) | Tirupati: **Pratham Metallics 07AAZFP7501N1Z** (Mobile-linked, **Active**) supplies 532.54 lakh (FY25-26), 532.33 (FY24-25) AND receives 63.44 lakh (FY23-24) and is #2 top supplier (460.91); own other-state GSTINs `07AACFT2702L1Z9`, `02AACFT2702L1ZJ` (PAN,Mobile,Email) | C (printed in F12 only) | #10/#19 flag only **Cancelled** parties and #43 needs 3 GSTINs, so an ACTIVE party sharing the taxpayer's mobile that is simultaneously a major supplier and customer produces no finding: related-party / circular trading (Sec 15(2)(b), Sec 16(2)(c)). Add rule: related_parameter != None AND total_itc >= threshold AND both lists. Same-PAN own GSTINs: ITC from own branch = taxable supply between distinct persons (Sec 25(4), Sch I). | **High** |
| Related party: **ITC Received After Cancellation / ITC Passed After Cancellation** (col 9), Order Passed On (col 7), Over All Risk Score (col 11) | Bajrang (stacked) Kaveosoft cancelled 2022-08-31, order 2025-11-17, `5.64` lakh ITC received after cancellation; Tirupati all 0 | **A** (not in mapping; only `cancellation_date`, `reason`, `total_itc` kept) | ITC after the date of (retrospective) cancellation is directly ineligible (Sec 16(2)(c) / 29(2)(c)); the rule is trivial once the column is mapped. `cancellation_date` itself is B (not compared to invoice dates in 2A/2B). | **High** |
| Related party: `reason` (Cancellation Reason) | `Others\|returns furnished by ...`, `Discontinuation/Closure of Business` | B | "Non-filing of returns for 6 months" vs "closure" separates defaulting suppliers (Sec 16(2)(c) tax not paid) from closures. | Low |
| Refund Details: Total Rejected / Sanctioned | all 0 | B (claimed D in F9a) | When populated, rejected refunds = department already disputes ITC; ties to Sec 54 and to the credit-ledger "refund" debits. | Low |
| Financial `2008-9` junk FY row | Bajrang only | B | Filter rows with FY earlier than registration date. | Low |

---------------------------------------------------------------------------------------------------

## 2. GSTR-9 annual return (`gstr9_mapping.json`: 23 facts + Part I + Item 19)

Real files: Bajrang FY23-24 (ARN AA050324442575B, filed 31-12-2024) and FY24-25 (AA050325474894V, filed 31-12-2025); Jagnath/Tirupati: none supplied. Sheets in file: Part I, Item 4, 5, 6, 7, 8, 9, Part V (10-14), Item 15, 16, 19, Items 17-18 (note only), Verification. **Sheets never read: Item 15, Item 16, Items 17-18, Verification.**

| Table / item | Raw item | Example value (Bajrang) | Status | What it could tell us | Priority |
|---|---|---|---|---|---|
| Part I | 3(d) Date of Filing, 3(c) ARN | `31-12-2025`, `AA050325474894V` | B | Late fee Sec 47(2) if after due date; Item 9/19 late fee zero consistent? Filed on the last due day both years, 9C ARN date the same day (31-12). `is_system_draft` computed but never surfaced/used (a draft would silently feed R14). | Low |
| Part I | 3(b) Trade name | `M/s.BAJRANG MOTORS` | A | cosmetic | Low |
| Verification | Authorised signatory, designation, date | `SANJAY AGRAWAL, Partner, 31-12-2025` | A | Match to BO Member Details (primary signatory) and to GSTR-3B signatory; mismatch = unauthorised filer. | Low |
| Item 4 (tax payable) | 4A B2C supplies (taxable, tax by head) | `1,927,851,165.47` taxable; CGST/SGST 232.9M each, IGST 57.8M, cess 195.5M | A | 77% of turnover is B2C. Compare 4A+4B with GSTR-1 annual (b2cs+b2cl+b2b, sums of 12 months) and 3B 3.1(a); a mismatch = Table 4 not reconciled to returns (Sec 44, Rule 80). | **High** |
| Item 4 | 4B B2B taxable and tax (read) | `600,236,289.02` | B (`table4_b2b_*`) | Tie to GSTR-1 B2B annual and BO e-invoice assessable (BO 5648 lakh, swapped column S3) - 4B 6002 lakh vs e-invoice 5648: ~354 lakh B2B without e-invoice. | **High** |
| Item 4 | 4I credit notes (read), 4J debit notes, 4K/4L amendments, 4M subtotal | CN taxable `12,817,634.23` | B (4I) / A (4J-4M) | Credit notes after year-end/ amendments (Sec 34 timing) and tie to GSTR-1 CDNR (engine has check #4 cn timing, not tied to Table 4). | Med |
| Item 4 | 4G inward RCM (taxable read, tax not) | `124,000` taxable, tax 11,160+11,160 | B (taxable only) | RCM tax 22,320 = BO RCM Liability 0.22 lakh and 9C Item 9F; vs 3B 3.1(d) sum and 6C. Supports BO RCM mapping (1.4). | Med |
| Item 4 | 4C/4D/4E/4F/4G1 (exports, SEZ, deemed, advances, ECO) | all `0` | A | Zero here with 5A=0; cross-check against ICEGATE/BO export and GSTR-1 6A. | Low |
| Item 4 | 4N net supplies (only the IGST cell read) | taxable `2,516,055,647.39`, tax: CGST 277.5M + SGST 277.5M + IGST 123.6M + cess 248.2M = 926.83M (9268.31 lakh) | B (IGST only) / A (CGST/SGST/cess/taxable) | Annual liability to compare with BO total tax liability 9270.19 (1.88 lakh diff), 3B and ledger. | **High** |
| Item 5 | 5A zero-rated, 5B SEZ, 5D exempted, 5E nil, 5F non-GST (read) | all `0` | B (`table5_*`, `table5_all_zero`) | `table5_all_zero` is computed but R13 uses GSTR-1 instead; combine: 9C 7B > 0 while 9 Table 5 zero is the exact contradiction R13 text describes but it is never tested here. | Med |
| Item 5 | 5C RCM by recipient, 5C1, 5G..5M subtotals, **5N Total turnover** | `2,515,931,647.39` | A | 5N = 9C 5Q = 9C 5P; vs BO turnover 25165.85 lakh (+6.53 lakh) and 3B annual 3.1: the main turnover three-way tie. | **High** |
| Item 6 | 6A (read, 4 heads summed) | `887,413,076.49` | D (R14) | note: R14 compares 6A (gross) with 9C 12A, while 6O - 7I = 884,420,215.98 = 9C 12D = BO ITC availed 8844.20. See Table 7. | - |
| Item 6 | 6A1 (preceding-FY ITC), 6A2 | `0` | A | ITC of earlier FY taken now: Sec 16(4) time-bar test. | Med |
| Item 6 | 6B inputs / capital goods / input services split | inputs 45.17M+45.17M+547.7M+236.9M; CG 0; services 0 | A | Capital goods 0 and input services 0 for a dealer with insurance (IFFCO-Tokio notes in 8A), rent and freight: everything booked as "Inputs" (also 9C Item 14 puts 100% in 14A "Purchases"). Category-wise mapping to 8A supplier names (hotels, insurance, transport = 17(5)). | Med |
| Item 6 | 6C/6D ITC on RCM (unregistered / registered) | CGST `9,270` | A | 6C 9,270 x2 = 18,540 vs 4G tax 22,320 -> ITC on RCM is 3,780 less than the RCM tax paid: mismatch, or the RCM invoice ITC not taken; match with 3B 4(A)(3). | Med |
| Item 6 | 6E/6F imports, 6G ISD | `0` | A | vs IEC (BO Pre-GST) / BOE; fine. | Low |
| Item 6 | **6H ITC reclaimed** | CGST `148,790`; SGST same; **IGST `12,099,914.44`**; cess 0 | A | ITC reversed earlier and reclaimed: Rule 37 (180-day payment) / Rule 37A / Sec 41. Should equal the `Electronic Credit Reversal and Re-claimed Statement.csv` (supplied for 24-25) re-claim rows and 3B 4(A)(5)/(B)(2); not tied. 1.2 Cr IGST re-claim is material. | **High** |
| Item 6 | 6J difference (I - A2), 6K-6N transition | FY23-24 `26,460` | A | FY23-24 6J = 26,460 (RCM input service ITC outside 6A). 6K-6M TRAN/ITC-01 zero. | Low |
| Item 6 | 6O total ITC availed (read in total only) | `887.4M` | B (`table6_total_itc_availed`) | with 7I -> 6O-7I = 9C 12D; see R14 note. | Med |
| Item 7 | 7I total reversed (read) | CGST 677,050.38 x2, IGST 400,051.59, cess 1,238,708.16 = 2,992,859 | B (`table7_total_reversed`) | = BO ITC Reversal (29.93 lakh). R14 should compare 9C books ITC with 6O-7I, not with 6A, otherwise a real reversal shows up as a "books vs returns gap" (FY24-25: 2.99M). | **High** |
| Item 7 | 7A (Rule 37), 7A1 (37A), 7A2 (38), 7B (39), 7C (42), 7D (43), 7E (17(5)), 7F/7G TRAN | all `0` | A | The whole reversal sits in **7H1 "Reversed"** (the catch-all) in both years, with 7C/7D (Rule 42/43) and 7E (17(5)) nil. For a taxpayer whose BIFA/Table 5 show no exempt turnover that is plausible, but 17(5)(a) blocked motor-vehicle credits (a vehicle dealer's own use vehicles) and 7C/7D can only be judged with the categories; ties to 3B 4B(1)/(2) split and `gst_blocked_credit.py`. | Med |
| Item 8 | 8A ITC per GSTR-2B (read, 4 heads summed), 8B, 8D (read, signed sum) | 8A `875,970,714.53`; 8B `874,997,042.05`; 8D CGST -29,326.12, SGST -29,326.12, **IGST +1,408,506.59**, cess **-376,181.87** | B (`table8_*`) | Per-head sign matters (summing nets over- and under-claims): cess over-claimed by 376k vs 2B while IGST under-claimed by 1.4M; Sec 16(2)(aa) / Rule 36(4). Table 8A in the return is the portal's 2B total - compare with the engine's recomputed 2B (R14 uses only recomputed) and with Table 8A workbook total 878,285,609.61 (gap 2.3M). | Med |
| Item 8 | 8C ITC on inward supplies of prior FY | FY23-24 `115,081.63` (= Part V 13) | A | duplicate of Table 13 (consistency). | Low |
| Item 8 | **8E ITC available but not availed, 8F ITC available but ineligible, 8G-8K (import IGST, lapsed ITC)** | all `0` | A | 8F = ITC the taxpayer itself declares ineligible (17(5)). 0 for a dealer whose 8A shows hotel/insurance invoices (8A "No" rows: 17,378 in FY24-25; **462,899 in FY23-24**) - ineligible ITC not disclosed? 8E/8K = lapse. | **High** |
| Item 9 | Tax payable (IGST, CGST, SGST) read | `277,505,905.14` | B (`table9_liability_*`) | tie to BO liability and ledger Part I. | Med |
| Item 9 | **Tax paid through cash / through ITC (cols 3-8) and col 9 difference** | CGST cash `11,160`, ITC `45,124,974` ...; difference CGST **-144,133.86**, IGST +80,742.23, cess -3,716.53 | **A** | Payable vs paid per head - negative = paid more than payable (annual excess payment); cash share per head vs BO `tax_paid_in_cash` 386.92 lakh and the cash ledger; the 2 cr of IGST paid by ITC. Today only "payable" is read, so the cash/ITC split declared in the annual return is never compared with BO or ledgers. | **High** |
| Item 9 | Cess payable row | `248,203,327.47` | A (`cess` row not matched) | cess liability tie to 9C Item 9 and 3B. | Med |
| Item 9 | Interest (payable/paid), late fee (read), penalty, other | FY23-24 Interest `1,569`, late fee `200` | B | Declared interest/late fee = evidence of late payment/filing in the year; compare to the interest in the electronic ledger and the Filing Compliance sheet; nonzero when 3B were all on time = inconsistent. | Med |
| Part V (10-14) | 12 (reversal of prior-FY ITC), 13 (ITC availed for prior FY) | FY23-24 12: 676,351/676,351/966,216; 13: 115,081.63 | D (ITC annual summary) / C | FY24-25 file uses new wording ("ITC of the financial year reversed/availed in the **next** financial year") -> not read (W1303, known design question). | Med |
| Part V | 10 and 11 (supplies declared/reduced in next FY), 14 differential tax | FY23-24 10: taxable 17,926.74; 11: 60,051.86; 14: IGST -4,709.27, CGST 4,441.27 | **A** | Amendments after year-end (Sec 39(9)): feeds the 9C turnover plug; compare with GSTR-1/3B of the following FY (Apr-Sep amendments to FY rows). | Low |
| Item 15 | A-D refund claimed/sanctioned/rejected/pending; E-G demand raised/paid/pending | all `0` (FY24-25) | **A (sheet never read)** | Bajrang BO shows DRC-07 demands of 29-08-2024 (52.62/64.76/18.62 lakh) and 03-05-2024 (10.86/10.86/28.57/13.25/6.35 lakh) and 3 admitted appeals in FY24-25 - Table 15E-G shows **0** -> demands not disclosed in the annual return. Also refund 15A vs credit-ledger refund debits (F9a). | **High** |
| Item 16 | composition supplies, deemed supply (143), goods on approval | all `0` | A | vs 8A composition suppliers (invoice type) - low here. | Low |
| Item 19 | late fee payable/paid (read, A/B) | `0` | B | supersedes Item 9 late fee; use as above. | Low |
| Items 17-18 | HSN summary | note only (data not in the Excel export) | A | HSN table not available in this export; need JSON/PDF to compare HSN outward/inward with GSTR-1 Table 12 and BO HSN. | Low |

---------------------------------------------------------------------------------------------------

## 3. GSTR-9C (`gstr9c` facts)

Real files: Bajrang FY23-24 (ARN AA050324442606E) and FY24-25 (AA050325475003L). Sheets: Part I, Items 5-17, Part V, Verification. **Never read: Item 10, 11, 14, 15, 16, Part V, Item 17, Verification.**

| Part / item | Raw item | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|---|
| Part I | 3(d) ARN date; Item 4 "Name of Act under which liable to audit" | `31-12-2025`; **blank** | arn_date B; Item 4 A | Self-certified 9C (no auditor details); the blank "Act" field means the statutory-audit basis is not stated. 9C filed same day as 9. | Low |
| Item 5 | 5A audited turnover (read) | FY24-25 `2,515,981,440.32` | B (`turnover_audited_bs`) | The engine's BS/PL R1 compares P&L revenue with `turnover_declared_gstr9` (the GSTR-9 figure, i.e. 5Q) - not with 5A audited. | Med |
| Item 5 | 5B-5N adjustments (unbilled revenue, advances, deemed supply, credit notes, discounts, SEZ, composition, s.15, forex) | all `0` | A | Each is a legal reconciling item; non-zero ones must be matched to GSTR-1/3B. | Low |
| Item 5 | **5O "Adjustments in turnover due to reasons not listed above"** | FY24-25 **`-49,792.93`** | A | The only bridge between audited turnover 2,515,981,440.32 and GSTR-9 5N 2,515,931,647.39. A residual "other" plug = unexplained difference hidden in the "reconciled" statement (5R=0). Flag any 5O != 0 (and the size vs materiality). | **High** |
| Item 5 | 5P after adjustments, 5R un-reconciled (read) | `0` | B | trivially 0 because 5O plugs; test (5A - 5Q) not 5R. | Med |
| Item 6 | Reason text for turnover difference (read, `turnover_diff_reasons`) | FY24-25: `Value of turnover omitted to be shown however the tax has been declared correctly` (while 5R = 0) | B | A typed explanation although unreconciled = 0 shows a real difference existed (the 5O -49,792.93). The reason itself says turnover was **omitted to be shown** (Sec 39(9)/ Rule 80). Search reason text for amounts and keywords ("omitted", "not reflected", "timing"). `itc_diff_reasons` (Item 13) likewise. | Med |
| Item 7 | 7B exempt/nil/non-GST (read) | `0` | D (R13, F2b) | - | - |
| Item 7 | 7C zero-rated w/o tax, 7D RCM, 7D1 ECO | `0` | A | vs 9 Table 5A/5C/4G (Bajrang 4G has 124,000 RCM taxable but 7D = 0 - **inconsistent: 9C 7D RCM supplies 0 while GSTR-9 4G shows 124,000 inward RCM**; note 7D is outward-RCM so the check is 9 Table 5C). | Low |
| Item 7 | 7E, 7F (read), 7G unreconciled | equal | B | tie-outs. | Low |
| Item 9 | **Rate-wise taxable value and tax rows A..O** (5%, 12%, 18%, 28%, RC rows, interest, late fee, penalty) | FY24-25 28%: taxable `2,318,758,819.15`, CGST/SGST 265.7M, IGST 117.8M, cess 248.2M; 18%: 149.3M; 12%: 1.25M; 5%: 46.6M; 18% RC 124,000 | **A** (only the totals P/Q are read, as sums, B) | Rate-wise declared turnover vs GSTR-1 Table 12/ HSN and rate-wise tax in 3B; the 28% bucket is 92% of turnover - a drift of goods into 5%/12% buckets (Mahindra spares, services) shows as a mix change. Effective-rate test vs BO `effective_tax_rate_pct` (31.3%). | **High** |
| Item 9 | P/Q totals (read) | `277,505,905.14` | B (`tax_payable_total`, `tax_paid_declared`) | vs GSTR-9 Item 9 and 4N (all equal in FY24-25). | Low |
| Item 10, 11 | Reasons for un-reconciled payment / additional amount payable not paid | blank / `0` | A (Item 10 not even a reason-list) | When non-blank: auditor-admitted tax short payment (Sec 73/74 self-assessment). | Med |
| Item 12 | 12A books ITC (read) | FY24-25 `884,420,215.98` | D (R14) | R14 uses the gross books figure 12A, not 12D (after B and C). FY23-24: 12A 639,799,501.75, 12C `1,494,696.39` (ITC booked in the year, claimed later) -> 12D 638,304,805.36 = 6O - 7I. R14 gap 12A vs 6A is 1.35M, almost entirely explained by 12C; the engine reads 12C (`itc_booked_now_claimed_later`) but never uses it. | **High** |
| Item 12 | 12B (earlier FY claimed now), 12C, 12E, 12F | FY23-24 12C `1,494,696.39` | B | explains the books/returns gap above; 12B>0 = Sec 16(4) time-bar test. | **High** |
| Item 13 | Reasons for un-reconciled ITC | blank | B | see Item 6. | Low |
| Item 14 | Expense-wise ITC reconciliation (A Purchases ... Q capital goods, services, freight, power, rent, bank charges, etc.), "Value" and "Amount of ITC" per head | FY24-25: 14A Purchases value `0`, ITC `884,420,215.98`; every other head `0` | **A** (whole sheet) | The entire ITC is parked under "Purchases" (and "Value" = 0, so ITC/value ratio cannot be tested): no ITC on freight, power, rent, insurance, bank charges - while the 8A shows insurance (IFFCO-Tokio), hotel and transport suppliers. Evidence that 9C Item 14 is not a real reconciliation; rate test: ITC/purchase value by head. | Med |
| Item 15, 16 | Reasons un-reconciled (expense basis); tax payable on un-reconciled ITC (CGST/SGST/IGST/cess/interest/penalty) | `Reason number 1` blank, `0` | A | Any non-zero 16 is an auditor-admitted ITC reversal liability. | Med |
| Part V | Additional liability due to non-reconciliation (rate-wise, ITC, interest, late fee, penalty, erroneous refund, outstanding demands) | `0` | A | non-zero = tax payable in addition; "Outstanding demands to be settled" should reconcile with the BO DRC/Liability ledger (Bajrang has real demands: see Item 15 of GSTR-9). | Med |
| Item 17 | Late fee payable/paid | `0` | A | duplicate of 9 Item 19. | Low |
| Verification | Signatory, designation | `SANJAY AGRAWAL, Partner` | A | as for GSTR-9. | Low |

---------------------------------------------------------------------------------------------------

## 4. Table 8A (R9_8A) (`table8a_mapping.json`)

Real files: Bajrang FY23-24 (generated **30/11/2024**), FY24-25 (generated **14/09/2026**): 15,711 / 15,867 B2B rows, 1,827 / 881 note rows, 217 / 289 distinct suppliers. Sheets: Notes, Read me, B2B, B2BA, B2B-CDNR, B2B-CDNRA, ECO, ECOA.

| Sheet / column | Example value | Status | What it could tell us | Priority |
|---|---|---|---|---|
| Notes, Read me instructions | text | A | none | - |
| Read me: **Date of generation**, Trade name | `30/11/2024` vs `14/09/2026` | A | The FY23-24 8A is a 30-Nov-2024 snapshot (Sec 16(4) cut-off date), the FY24-25 file a 2026 snapshot: different cut-offs; supplier filings after 30 Nov are absent from the first. Compare which invoices disappear/appear. | Low |
| **B2BA** (amended invoices: original and revised details, "Whether ITC to be reduced", "Amount declared by taxpayer for ITC reduction") | 12 rows FY24-25, e.g. Dehradun Premier Motors CSR25A000103 `2650.42`, tax 251.32x2 | **A (sheet never read)** | Supplier amended an invoice after the recipient claimed; if the revision reduces tax and the taxpayer has not reversed ("Whether ITC to be reduced = Y"), ITC is over-claimed (Sec 34(2)/Rule 37). Also the B2B sheet may still carry the original line, so totals double count. | Med |
| B2B-CDNRA, ECO, ECOA | 0 rows | A | e-commerce operator supplies (ECO) belong with Sec 52 TCS, absent here. | Low |
| B2B `period` | maps to the supplier's **GSTR-1 period** (`APR-24`), not the 2B period | B | wrong-meaning field name. | Low |
| B2B **GSTR-2B Period** (col "GSTR-2B Period") | `APR-24` / `MAY-24` | **A** (not mapped; the first "period" match is the GSTR-1 period) | Compare with the month the ITC was claimed in 3B: ITC claimed before it appears in 2B, or claimed 3+ months after (Sec 16(4) clock). | Med |
| B2B `supplier_filing_date` (GSTR-1 filing date) | `13/05/2024` | B | Supplier's filing lag vs invoice date: FY24-25 45 invoices with lag > 60 days (Rs 1.62 lakh ITC), none after the 30-Nov cut-off. Late-filing suppliers = unreliable suppliers, and invoices re-appearing after 30 Nov are time-barred ITC (Sec 16(4)). Computable now. | Med |
| B2B `rate` | FY24-25 export: `0.0 ... 223051.36` (the IGST amount; this export has **no Rate column** and "integrated tax" contains the substring "rate") | B (wrong column in FY24-25) | Real rate = tax / taxable. Rate-wise ITC vs Table 4 / 9C Item 9; 28% supplier (Mahindra) share. Fix match to exact `rate`. | Low |
| B2B `invno`, `invdate`, `invval`, `pos`, `rcm`, `invtype`, `taxable`, per-head tax, `supplier`, `gstin` (invoice rows) | whole table | B (rows never read by any check; only the Yes-sum is) | Row-level: (1) invoice dates outside FY (FY23-24: 55); (2) duplicates by (gstin, invno) (FY23-24 2112 = line-level splits - verify dedupe); (3) RCM = Y rows (FY23-24 3 Harbans Transport invoices Rs 1,200 tax; FY24-25 1) vs 3B 3.1(d); (4) POS != Uttarakhand (hotels in Pune/Agra: IGST/17(5)(b)); (5) **concentration**: Mahindra 27AAACM3025E1ZZ = 60.4% of FY24-25 and 71.7% of FY23-24 8A ITC, top-3 Mahindra GSTINs = 77% (the engine's F12 uses 2B, not 8A); (6) supplier-by-supplier match against GSTR-2B/2A to find invoices in 8A but absent from the 2B the ITC was claimed on. | **High** |
| B2B `itc_available`=No + `reason_not_available` | FY24-25 9 rows Rs 17,378 ("POS lies in supplier's State", "Rev charge available"); **FY23-24 5 rows Rs 462,899** ("POS and supplier state are same but recipient state is different", incl. Dewan Fourwheels INV24C000300 Rs 265,170) | C (reason breakdown printed in flow row 800; amounts and invoices not) | These invoices are ITC-ineligible: confirm they were NOT claimed in 3B 4(A) (cross-check by invoice no. against 3B ITC detail / ledger); tie to GSTR-9 8F "ITC available but ineligible" (0). | **High** |
| B2B-CDNR (credit/debit notes) all columns | FY24-25 CN 108 (ITC 2.76M), DN 773 (0.27M); FY23-24 CN 102, DN 1,725 (DN ITC 4.20M) | B (`cdnr` read but `totals.total` is B2B only) | The 8A figure used in R14 (878,285,609.61) **excludes** notes. Net ITC per 8A = B2B + DN - CN: FY24-25 -2.49 lakh, FY23-24 +9.31 lakh. Also 8C/13 prior-year and the Sec 34(2) reversal on credit notes (supplier CN: recipient must reverse ITC, GSTR-9 Table 7). | Med |
| B2B-CDNR `note_type`, `supply_type` | CN/DN, `R` | B | see above | Low |

---------------------------------------------------------------------------------------------------

## 5. Top 5 opportunities (ranked by signal per unit of work)

1. **Fix the stacked-sheet BO Profile (Bajrang) and then actually run the BO-driven checks.** (S1/S2) Split the "ITC Passed On" sheet at its section titles. Without it #3 (7 of top-10 beneficiaries are also top-10 suppliers), #54 (Indraprastha HIGH), F9c (DRC-07 demands 52.62/64.76/18.62 lakh), proceedings (3 appeals), related-party ITC after cancellation (5.64 lakh) and `itc_passed_by_fy` (currently 0.0) are all silently void for the largest taxpayer.
2. **Turn BIFA into findings.** Map the TDS/TCS supply columns (A) and consume `diff_liability`, `diff_itc`, `liability_gstr1`: Tirupati FY25-26 1476.38 lakh GSTR-1-vs-3B gap is exactly the July-2025 DRC-03 payments (613.06 + 213.40 + 649.92), with a further 209.42 lakh interest/penalty DRC and a pending **Section 74** transfer (map `Section Information`); Tirupati FY23-24 deductor-reported supplies exceed 3B turnover by 6265.25 lakh; Tirupati FY24-25 excess ITC 27.87 lakh. Also guard BIFA's "Liability as per EWB" duplicate-column bug (S4).
3. **Wire GSTR-9/9C into three-way tie-outs (9 vs 3B vs 1 vs BO) and fix R14.** Currently only 6A, 13 and 9C 5Q/7B/12A are used. Add: Table 4 (4A/4B/4N) and 5N vs GSTR-1/3B/BO turnover (BO 25165.85 vs 5N 25159.32 vs 9C 5A 25159.81); Item 9 cash/ITC split per head vs ledgers and BO; Table 7 by rule and **R14 against 6O-7I / 9C 12D (not 6A / 12A)** and 12C; 9C 5O plug (-49,792.93) and reason text; 6H reclaimed ITC (IGST 1.21 Cr) vs the Credit Reversal and Re-claimed statement; 8E/8F/8K ineligible ITC vs 8A "No" rows.
4. **Use Table 8A at invoice level.** Today only a sum is used (and it omits credit/debit notes, wrong "rate", wrong "period"). Add: supplier concentration (Mahindra 60-72%), supplier filing lag / post-30-Nov invoices (Sec 16(4)), ITC-ineligible invoices checked against 3B claim, B2BA amendments with "ITC to be reduced", RCM=Y rows vs 3B RCM, 2B-period vs claim month, invoice-by-invoice match with 2B.
5. **Department-signal and related-party columns that are free to read.** Map and surface: `Flagged in Lead Based Dashboard` (Tirupati: Spike in ITC Utilized = 1), `Overall Risk Score` (Tirupati 7.50 HIGH), `Suo Moto Cancelled Related PAN`, `ITC Received/Passed After Cancellation` (direct Sec 16(2)(c) violation), `related_parameter` for ACTIVE related parties (Tirupati: Mobile-linked Pratham Metallics supplies 532 lakh/yr and buys 63 lakh; own other-state GSTINs), BO EWB inward/outward ratio (Tirupati 97.2) and BO inward-EWB tax vs ITC availed (Bajrang 13% of ITC has no EWB trail), RCM Liability/Cash columns (Tirupati 21.06 lakh FY25-26), GSTR-9 Item 15 (demands not disclosed although DRC-07 demands exist in the BO file).


---


---

## Addendum: follow-up investigation results

**Credit Reversal and Re-claimed Statement.** Now parsed (`gst_credit_reversal.py`) and reconciled against
GSTR-3B 4(B)(2) and 4(D)(1) by tax period and head (sheet "Credit Reversal Recon", findings CR1-CR4). On
Bajrang FY24-25 every compared line agrees to the rupee (Aug-24 reclaim 12,099,914.44 IGST / 148,790 CGST /
148,790 SGST / 1,806,103 Cess; Sep-24 reversal 676,351 / 676,351 / 966,216). Mar-25's 3B reversal
(400,051.59 / 699.38 / 699.38 / 272,492.16) is filed after the statement's To date (31/03/2025) and is shown
as outside the window, not as a mismatch. Still reclaimable per the statement's closing balance: 676,351
CGST, 676,351 SGST, 966,216 Cess. The statement's "ITC Claimed (Table 4A(5))" columns are not yet compared
with 3B 4A(5).

**PMT-09.** None of the three taxpayers' cash, credit or liability ledgers contains a PMT-09 or any
major/minor-head transfer row. The only cash-ledger descriptions are "Amount deposited", "TDS Credit",
"Voluntary payment", "Interest" and "Payment of Demand under Appeal". No PMT-09 handling was added; the
#47/#55 exclusions stay as a guard only. Revisit when a real ledger with such rows is supplied.

**R14 (6O - 7I).** Verified. The audit's figure was right and the tool's extraction was wrong: row 6O of the
real GSTR-9 export has no blank "Type" cell, so its four heads are in columns 2-5, and the fixed 3-6 read
dropped CGST (842,088,760.34 instead of 887,413,076.49). Fixed in `gstr9_adapter.py` and
`gst_parsers_dept.py`. With the fix, 6O - 7I = 884,420,215.98, equal to GSTR-9C 12D and 12E to the paisa, so
9C books figures are net of reversals while 6A is gross. R14 keeps 12A + 12B - 12C as the books figure and
now adds 6O - 7I as a separate source; a gross-vs-books gap equal to Table 7 is reported as "the reversals".
The remaining Bajrang R14 gap is 8A (878,285,609.61) vs books, -6,134,606.37.
