# =============================================================================
# debt_schedule.py — Single-year debt sweep helper
# =============================================================================
#
# Called once per projection year by forecast.py.
# Interest has already been computed from BEGINNING balances before this runs.
# Mandatory amort has already been deducted from fcf_before_optional.
#
# Sweep waterfall (fixed):
#   1. Revolver — draw if cash < MIN_CASH_BALANCE; repay if excess and drawn
#   2. Term Loan — sweep all excess cash above MIN_CASH_BALANCE
#   3. Mezzanine — only if MEZZ_CASH_SWEEP = True
#
# =============================================================================

import inputs as cfg


def sweep_debt(
    beg_cash: float,
    rev_beg: float,
    tl_beg: float,
    mezz_beg: float,
    fcf_before_optional: float,
    mand_amort: float,
    mezz_pik: float,
) -> dict:
    """
    Parameters
    ----------
    beg_cash              : cash balance at START of year (before FCF)
    rev_beg               : revolver beginning balance
    tl_beg                : term loan beginning balance
    mezz_beg              : mezzanine beginning balance (before PIK accretion)
    fcf_before_optional   : FCF after mandatory amort; drives optional paydown
    mand_amort            : mandatory amort already deducted from fcf and TL balance
    mezz_pik              : PIK accretion for the year (accretes to mezz balance)

    Returns
    -------
    dict with ending balances and repayment/draw amounts for each tranche.
    """

    # ── Adjust balances for items already booked ─────────────────────────────
    tl_after_mandatory  = max(0.0, tl_beg - mand_amort)
    mezz_after_pik      = mezz_beg + mezz_pik       # PIK accretes before any paydown

    # ── Cash available before optional sweep ─────────────────────────────────
    # beg_cash + operating FCF + (mandatory amort already reduced fcf, so just add fcf_before)
    cash_pool = beg_cash + fcf_before_optional

    # ── Step 1: Revolver ─────────────────────────────────────────────────────
    if cash_pool < cfg.MIN_CASH_BALANCE:
        # Cash shortfall — draw revolver up to available capacity
        shortfall    = cfg.MIN_CASH_BALANCE - cash_pool
        rev_draw     = min(shortfall, cfg.REVOLVER_CAPACITY - rev_beg)
        rev_repay    = 0.0
        cash_pool   += rev_draw
        rev_end      = rev_beg + rev_draw
    else:
        # Repay revolver before sweeping Term Loan
        rev_draw     = 0.0
        rev_repay    = min(cash_pool - cfg.MIN_CASH_BALANCE, rev_beg)
        cash_pool   -= rev_repay
        rev_end      = rev_beg - rev_repay

    # ── Step 2: Term Loan optional sweep ────────────────────────────────────
    excess_for_tl  = max(0.0, cash_pool - cfg.MIN_CASH_BALANCE)
    tl_repay       = min(excess_for_tl, tl_after_mandatory)
    cash_pool     -= tl_repay
    tl_end         = tl_after_mandatory - tl_repay

    # ── Step 3: Mezzanine optional sweep (controlled by flag) ────────────────
    if cfg.MEZZ_CASH_SWEEP:
        excess_for_mezz = max(0.0, cash_pool - cfg.MIN_CASH_BALANCE)
        mezz_repay      = min(excess_for_mezz, mezz_after_pik)
        cash_pool      -= mezz_repay
        mezz_end        = mezz_after_pik - mezz_repay
    else:
        mezz_repay = 0.0
        mezz_end   = mezz_after_pik

    return {
        "rev_draw":       rev_draw,
        "rev_repayment":  rev_repay,
        "rev_end":        rev_end,
        "tl_repayment":   tl_repay,
        "tl_end":         tl_end,
        "mezz_repayment": mezz_repay,
        "mezz_end":       mezz_end,
        "end_cash":       cash_pool,
    }
