#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BS/PL STRUCTURED INPUT  --  fill this in by hand from the real PDF/Excel,
then pass BS_PL_DATA (or your own copy of this dict) to
gst_checks_forensic.check_bs_pl_rules(). See docs/GST_360_SCRUTINY_DIRECTIVE.md
for why this is a typed-in dict and not an auto-OCR'd one.

Every key is OPTIONAL -- only fill in what you can read off the real
document. check_bs_pl_rules() emits an explicit "not tested -- <key> not
supplied" finding for anything left out, so the final report always shows
a complete checklist, never a silently-skipped row.

Values are {"fy_prior": <number>, "fy_current": <number>} pairs. Use the
SAME two financial years as the rest of the run.

GENERIC NAME (BS_PL_DATA, not a taxpayer-specific name): master_build.py
imports this exact variable from this exact file -- for a taxpayer, replace
the values below with their real figures (rename nothing, keep the variable
name BS_PL_DATA so the pipeline picks it up automatically), or leave the
dict empty to skip R0-R12 cleanly.

This repo ships this file as an empty template on purpose -- no taxpayer
data is committed here. Fill in your own figures locally; this file is
gitignored once you do (see .gitignore) so real financial data never gets
committed by accident.
"""

BS_PL_DATA = {
    # SAFETY TAG (checked by master_build.py before use): this dict is only applied if it
    # matches the GSTIN actually being processed this run. Prevents a stale/wrong-taxpayer's
    # BS/PL figures from silently being used against a different taxpayer's GST returns.
    "_gstin": None,
    "total_assets": {"fy_prior": None, "fy_current": None},
    "total_equity_liab": {"fy_prior": None, "fy_current": None},

    "share_capital": {"fy_prior": None, "fy_current": None},
    "reserves_and_surplus": {"fy_prior": None, "fy_current": None},

    "trade_payables": {"fy_prior": None, "fy_current": None},
    "short_term_provisions": {"fy_prior": None, "fy_current": None},

    "fixed_assets_tangible": {"fy_prior": None, "fy_current": None},
    "non_current_investments": {"fy_prior": None, "fy_current": None},
    "inventories": {"fy_prior": None, "fy_current": None},
    "trade_receivables": {"fy_prior": None, "fy_current": None},

    "revenue_from_operations": {"fy_prior": None, "fy_current": None},
    "other_income": {"fy_prior": None, "fy_current": None},
    "other_expenses": {"fy_prior": None, "fy_current": None},
    "finance_costs": {"fy_prior": None, "fy_current": None},
    "depreciation": {"fy_prior": None, "fy_current": None},
    "net_profit_after_tax": {"fy_prior": None, "fy_current": None},
}
