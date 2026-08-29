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
