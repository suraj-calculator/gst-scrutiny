# GST 360° SCRUTINY — STANDING DIRECTIVE

**Purpose:** this is the prompt/spec to hand to an AI (or a developer) working on this tool,
any time the goal is "make sure nothing is missed" rather than "add one specific feature."
It captures the hard rule the tool owner has stated explicitly:

> Complete scrutiny of the GST returns is non-negotiable. Every actionable data point that
> a GST officer could interpret must be surfaced — not summarised away, not left as an
> aggregate when invoice-level detail is available, not silently skipped when a source is
> present. This holds regardless of how much time or how many tokens it costs to get there.

---

## 1. The prompt itself (paste this to start a scrutiny-expansion session)

> You are extending a Python-based GST return-scrutiny tool. The tool already ingests
> GSTR-1, GSTR-3B, GSTR-2A, GSTR-2B, E-Invoice, both E-Way Bill directions, Electronic
> Cash/Credit/Liability Ledgers, the BO/360° Profile, GSTR-9/9C, Table 8A, and optionally a
> structured Balance Sheet/P&L. Your job is not to add one feature — it is to find every
> place this tool's OWN inputs contain a comparison, a cross-check, or an anomaly signal
> that isn't yet being computed, and add it.
>
> Work through Section 3 of this document (the coverage checklist) one row at a time.
> For each row: (a) confirm in the actual code whether it's already covered — read the
> function, don't assume from a sheet name; (b) if it's covered, move on; (c) if it's
> partially covered, note exactly what's missing; (d) if it's not covered at all and the
> data to compute it already exists in a parsed source, build it as a new finding/sheet
> following Section 2's non-negotiables; (e) if it's not computable from any currently
> supplied source, say so explicitly (as a SKIPPED/INFO finding, never a silent absence)
# and name what NEW input would make it possible.
>
> Test every claim against real taxpayer data before calling it done — reading the code is
> not verification, running it is. A full pipeline run takes several minutes; budget for
> it. If a check produces zero hits on your test taxpayer, that is not proof it works —
> construct a synthetic case (clearly labelled as synthetic, never mixed into real output)
> to prove the logic fires correctly, the way this tool's own §1–5 quarterly-2B fix was
> verified.
>
> Do not stop at "the sheet exists." A sheet that shows a number with no comparator, no
> severity, no red-flag, and no invoice-level backing is not scrutiny — it's a data dump.
> Every finding in this tool must answer: compared against what, is it within tolerance,
> and if not, here is the exact invoice/document behind the gap.

---

## 2. Non-negotiables — apply to every new check, every time

1. **Content-based, never filename-based.** Every input is classified by its actual sheet
   names / header row / banner text, never by what the file happens to be called.
2. **Never fabricate.** No invented invoice, no guessed value, no assumed classification
   presented as fact. If a comparator can't be built from real data, say why, don't
   approximate it silently.
3. **Explicit gap over silent skip.** A missing source, an empty result, or a check that
   doesn't apply this month must render as a stated SKIPPED/INFO finding with a reason —
   never a blank cell or a vanished row that looks like "nothing to report."
4. **Severity must mean something and must be consistent.** FLAG = a genuine issue past a
   materiality floor; REVIEW = worth a look but not yet material or not fully proven;
   PASS = actually checked and clean; INFO = observational, not itself a finding; SKIPPED
   = the source wasn't there. Two checks describing the same underlying fact (e.g. an
   invoice-level check and its FY-aggregate sibling) must never disagree on severity.
5. **Aggregate + invoice-level detail, not aggregate alone.** Any finding that says "N
   invoices differ" must have a table of those N invoices somewhere on the same sheet —
   not just a count in a sentence.
6. **Every table gets a total row** where the columns are summable, styled consistently
   with the rest of the sheet.
7. **Red/amber highlighting on the cell that IS the problem**, not only on a generic
   Result column three columns away. If column G is the number that's wrong, G is red.
8. **Cite the section of law** the check is actually testing (16(2), 17(5), 18(6), 36(4),
   37, 38, 42, 43, 49(4), 86A, Rule 138, etc.) in the finding text — a number without the
   legal basis for why it matters is not scrutiny, it's arithmetic.
9. **Cross-check every pair of sources that CAN be cross-checked**, not just the obvious
   ones. If two independent sources both carry a figure for the same thing (declared
   liability vs cash debited; 2B-flagged RCM vs 3.1(d); GSTR-1 HSN vs EWB HSN), they
   should be compared, even if the obvious primary check already exists elsewhere.
10. **Test on real data, not just compile.** A change that compiles but was never run
    against an actual taxpayer's files is not verified. Batch changes, then run once —
    don't skip the run to save time.

---

## 3. Coverage checklist — work through this, not from memory

Use this as the working list. Columns: **what to check**, **which source(s) it needs**,
**status as of this session** (fill in honestly when you touch a row), **law basis**.

### A. Outward supply (GSTR-1) integrity
| Check | Sources | Status | Law |
|---|---|---|---|
| GSTR-1 vs GSTR-3B 3.1(a) taxable/tax | G1, G3B | ✅ built | — |
| GSTR-1 vs E-Invoice completeness | G1, EINV | ✅ built | — |
| GSTR-1 vs Outward EWB (both directions of "present in") | G1, EWB-out | ✅ built, fixed this session (§1 bug) | Rule 138 |
| B2B→B2C shift, month-on-month | G1 | ✅ built | — |
| Document-series integrity (Table 13 vs actual) | G1 | ✅ built | — |
| Zero-tax invoices with real value | G1 | ✅ built this session | — |
| Same-day repeat invoices to one counterparty | G1 | ✅ built, enhanced this session | invoice-splitting risk |
| Credit-note original-invoice linkage | G1 | ⚠️ approximate only (no invoice-no field in GSTR-1's own CDNR sheet) — documented limitation |
| HSN-wise rate review / mix-shift / new-code timeline | G1 | ✅ built, made actionable this session |

### B. Inward supply / ITC (GSTR-2A, GSTR-2B) integrity
| Check | Sources | Status | Law |
|---|---|---|---|
| ITC claimed (3B 4A) vs available (2B invoice-level) | G3B, 2B | ✅ built, §7 bug fixed this session (B2BA/CDNRA netting) | 16(2)(aa), Rule 36(4) |
| ITC claimed vs GSTR-2A (completeness, filing-status aware) | G3B, 2A | ✅ built | 16(2)(c) |
| 2B quarterly-summary column misread | 2B | ✅ fixed this session (§1-5) | — |
| Duplicate/reused invoice number in 2A | 2A | ✅ built | — |
| RCM flag (2A) vs 3.1(d) liability vs 4A(3) ITC | 2A, G3B | ✅ built, complete-detail added this session | 9(3), 9(4) |
| RCM: liability declared vs cash actually debited | G3B, cash ledger | ✅ built, overpayment case added this session | 49(4) |
| State-code vs tax-head (IGST vs CGST+SGST) | 2A | ✅ built, complete-detail added (mismatch-only, per instruction) | POS rules |
| Blocked credits (Sec 17(5)) — HSN-based | 2B | ❌ not computable — 2B has no HSN column (documented, A5) |
| Blocked credits — trade-name keyword screen | 2B, master list | ✅ built this session | 17(5) |
| Zero-tax inward invoices with real value | 2B | ✅ built this session | — |
| Reciprocal counterparties (both buy and sell) | G1, 2B | ✅ built, complete-detail added this session | circular-trading risk |

### C. E-Way Bill cross-checks
| Check | Sources | Status | Law |
|---|---|---|---|
| EWB-out present in GSTR-1 / vice versa | EWB-out, G1 | ✅ built, #1/#3 duplication bug fixed this session | Rule 138 |
| EWB-in present in GSTR-2B / vice versa | EWB-in, 2B | ✅ built | Rule 138 |
| EWB value vs invoice value (both directions) | EWB, G1/2B | ✅ built | — |
| Same vehicle, repeated trips | EWB | ✅ built, annual pattern view added |
| Machinery/capital-goods HSN in EWB (purchase and sale) | EWB both, master list | ✅ built this session — the only source with real inward HSN | 18(6) |
| EWB validity/expiry vs supply date | EWB | ❌ not computable — export has no validity column |
| EWB cancellation status vs return filing | EWB | ❌ not computable — export has no cancellation-status column |
| Zero-tax EWB movements with real assessable value | EWB both | ✅ built this session |

### D. Ledger / cash / payment integrity
| Check | Sources | Status | Law |
|---|---|---|---|
| Credit ledger running balance tie-out (month-over-month) | Credit ledger | ✅ built this session, prior-FY-carryover bug fixed |
| Cash ledger sanity check | Cash ledger | ✅ built this session |
| DRC-03 / voluntary payments, Rule 86A blocks | Cash, Credit, Liability ledgers, BO Profile | ✅ built, made actionable this session |
| Refund claimed out of credit ledger vs BO Profile | Credit ledger, BO Profile | ✅ built |
| FY-total recomputed vs department's own BIFA figures | Ledgers, BO Profile | ✅ built, Diff% + fake-mismatch bug fixed this session |
| Liability Ledger (Part II) vs BO Profile DRC — exact ID match | Liability ledger, BO Profile | ✅ built |

### E. Annual-return cross-checks (GSTR-9/9C, Table 8A)
| Check | Sources | Status | Law |
|---|---|---|---|
| Exempt/nil turnover declared but zero Rule 42/43 reversal | GSTR-9C, G3B | ✅ built | Rule 42, 43 |
| GSTR-9 Table 8A vs 2B/3B ITC | Table 8A, 2B, G3B | needs GSTR-9/9C as Excel — PDF not parsed (documented) |
| BS/PL rule engine (R0-R12) | structured BS/PL | needs a per-taxpayer hand-typed input file — documented, not automatic |

### F. Capital-goods / business-nature signals
| Check | Sources | Status | Law |
|---|---|---|---|
| Machinery purchase (manufacturing signal) | EWB-in, 2B names, master list | ✅ built this session |
| Machinery sale (Sec 18(6) signal) | GSTR-1 HSN, EWB-out | ✅ built this session |
| Registration "nature of business" vs actual transaction pattern | BO Profile demographic, all sources | ❌ not built — would need a structured comparison of declared business type vs the HSN/keyword signals above |

### G. Counterparty risk
| Check | Sources | Status | Law |
|---|---|---|---|
| Related/cancelled-party ITC exchange | BO Profile | ✅ built, totals + departmental-proceedings merge this session |
| Top counterparties vs department's own top-10 | BO Profile, G1, 2B | ✅ built |
| Reciprocal (both-direction) counterparties | G1, 2B | ✅ built, complete detail this session |
| Same-day repeat transactions | G1, 2B | ✅ built, complete detail this session |
| GSTIN name resolution across sources | all sources | ✅ built this session (cross-source lookup fixing blank-name cases) |

---

## 4. What "done" means for a row in Section 3

A row only gets ✅ once ALL of these are true:
- The check exists in code and is wired into the actual output workbook (not just a
  standalone script).
- It was run against real taxpayer data and its output was read, not just "it didn't
  crash."
- Its severity logic was checked against at least one genuine positive case (something
  that SHOULD flag) and one genuine negative case (something that shouldn't).
- Invoice-level detail backs any aggregate claim.
- A totals row exists if the table has summable columns.
- The finding text names the section of law, where one applies.

A row that fails any of these is ⚠️ (partial) or ❌ (not done) — say so plainly rather than
marking it ✅ on the strength of "the code looks right."

---

## 5. Known structural gaps (not fixable without a new input source)

Listed here so nobody re-discovers these from scratch and burns time re-confirming them:

- **GSTR-2A/2B carry no HSN/line-item description** — invoice-level, not line-item.
  Blocked-credit and machinery detection on the INWARD side is therefore HSN-based only
  where EWB exists (EWB does carry HSN, confirmed), and trade-name-based (lower
  confidence) everywhere else. A purchase register would close this gap completely.
- **GSTR-1 credit notes carry no original-invoice-number field** — any CN-to-original
  link in this tool is by GSTIN+value approximation, never a proven document link.
- **GSTR-9/9C only parse from Excel, not PDF** — most taxpayers' annual returns are only
  ever downloaded as PDF from the portal; R13/R14 and the Table 8A cross-checks will show
  INFO/not-computable for those taxpayers until Excel is supplied instead.
- **BS/PL requires a hand-typed, GSTIN-tagged input file** — there is no automatic OCR of
  a scanned Balance Sheet; R0-R12 is opt-in per taxpayer by design (deliberately, to avoid
  ever silently applying one taxpayer's financials to another).
- **EWB exports used so far carry no validity/expiry or cancellation-status column** — if
  a taxpayer's actual EWB download includes those fields (a fuller export than seen so
  far), the corresponding checks (#22, #24 in the EWB matrix) would need their column
  detection extended to use them.

---

## 6. How to keep this document itself honest

Every time a row in Section 3 moves from ❌/⚠️ to ✅, or a new row is added, update this
file in the same commit/session as the code change — not as a separate cleanup pass later.
This document is only useful if it reflects the tool's ACTUAL current coverage, not an
aspirational one.
