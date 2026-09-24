#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TOOL-WIDE CONFIG -- settings for the NEW features added across this and the
next few sessions (Potential Blocked Credits, ITC Roll-Forward enhancements,
3-Way GSTR1/3B/EWB enhancements, ITC Annual Summary). Kept in one place per
explicit instruction, so future thresholds don't get scattered across files.

Purely additive: nothing here overrides, duplicates, or is read by any
PRE-EXISTING check. The tool's existing per-file constants (e.g. MATERIAL,
TOL in gst_checks_flow.py) are untouched and intentionally NOT duplicated
here -- only settings for the new code below exist in this file.
"""

# ---- Potential Blocked Credits sheet ----
# The master keyword/HSN file is normally auto-detected by CONTENT (single
# sheet, header row = Category / Search keyword / Indicative HSN/SAC) via
# gst_core.classify_folder() -- same principle as every other input in this
# tool, never by filename. This fallback path is used ONLY if no matching
# file is found in the run folder at all.
BLOCKED_ITC_MASTER_FALLBACK_PATH = None

# ---- ITC Roll-Forward 4A-4B-4C enhancements ----
# Month-over-month outlier flag: a column's value changing by more than this
# multiple (2.0 = 100%, i.e. more than double or less than half) vs. the
# prior month is flagged for review. Prior month = 0 is handled separately
# ("new activity this month"), not run through this ratio.
ITC_ROLLFORWARD_MOM_THRESHOLD = 2.0
# A ratio on two trivially small figures (a few hundred rupees of rounding/
# interest adjustment) is statistically meaningless -- skip flagging a
# month/column where BOTH the prior and current value are below this floor.
ITC_ROLLFORWARD_MOM_FLOOR = 1000.0

# GSTR-2B conversion layer (spec: docs/GSTR2B_CANONICAL_SPEC.md). False = the original direct
# reader (unchanged). True = raw 2B -> canonical workbook (gstr2b_adapter.py) -> engine. A canonical
# file supplied as input is always read through the adapter regardless of this switch. Env override
# for tests: GST_2B_CANONICAL=1 / 0.
GSTR2B_USE_CANONICAL = True

# GSTR-3B conversion layer (spec: docs/GSTR3B_CANONICAL_SPEC.md). False = the original direct readers
# (unchanged). True = raw 3B -> canonical workbook (gstr3b_adapter.py) -> engine. A canonical file
# supplied as input is always read through the adapter. Env override: GST_3B_CANONICAL=1 / 0.
GSTR3B_USE_CANONICAL = True

# GSTR-1 conversion layer (spec: docs/GSTR1_CANONICAL_SPEC.md). False = the original direct readers
# (unchanged). True = merged GSTR-1 -> canonical workbook (gstr1_adapter.py) -> engine. A canonical file
# supplied as input is always read through the adapter. Env override: GST_1_CANONICAL=1 / 0.
GSTR1_USE_CANONICAL = True

# E-Invoice conversion layer (spec: docs/EINV_CANONICAL_SPEC.md). False = the original direct readers
# (unchanged). True = merged E-Invoice -> canonical workbook (einv_adapter.py) -> engine. A canonical file
# supplied as input is always read through the adapter. Env override: GST_EINV_CANONICAL=1 / 0.
EINV_USE_CANONICAL = True

# E-Way Bill conversion layer (spec: docs/EWB_CANONICAL_SPEC.md). False = the original direct reader
# (unchanged). True = merged e-way bill (inward or outward) -> canonical workbook (ewb_adapter.py) ->
# engine. A canonical file supplied as input is always read through the adapter. Env override:
# GST_EWB_CANONICAL=1 / 0.
EWB_USE_CANONICAL = True

# GSTR-2A conversion layer (spec: docs/GSTR2A_CANONICAL_SPEC.md). False = the original direct reader
# (unchanged). True = merged GSTR-2A -> canonical workbook (gstr2a_adapter.py) -> engine. A canonical file
# supplied as input is always read through the adapter. Env override: GST_R2A_CANONICAL=1 / 0.
R2A_USE_CANONICAL = True

# Ledger conversion layer (spec: docs/LEDGER_CANONICAL_SPEC.md) - the four ledger CSVs (cash, credit, liability
# register, liability ledger/DRC). False = the original direct readers (unchanged). True = ledger CSV -> canonical
# workbook (ledger_adapter.py) -> engine. A canonical ledger workbook supplied as input is always read through the
# adapter. Env override: GST_LEDGER_CANONICAL=1 / 0.
LEDGER_USE_CANONICAL = True

# Table 8A and BO Profile conversion layers (specs: docs/TABLE8A_CANONICAL_SPEC.md, docs/BOPROFILE_CANONICAL_SPEC.md).
# False = the original direct readers (unchanged). True = government export -> canonical workbook
# (table8a_adapter.py / boprofile_adapter.py) -> engine. A canonical file supplied as input is always read through its adapter.
# Env overrides: GST_T8A_CANONICAL=1 / 0, GST_BO_CANONICAL=1 / 0.
TABLE8A_USE_CANONICAL = True
BOPROFILE_USE_CANONICAL = True

# GSTR-9 / GSTR-9C conversion layer (spec: docs/GSTR9_CANONICAL_SPEC.md). False = the original direct readers (unchanged).
# True = government Excel export -> canonical workbook (gstr9_adapter.py) -> engine. A canonical file supplied as input is
# always read through the adapter. Env override: GST_R9_CANONICAL=1 / 0.
GSTR9_USE_CANONICAL = True
