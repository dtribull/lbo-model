# =============================================================================
# forecast.py — Year-by-year income statement, FCF, and debt schedule loop
# =============================================================================
#
# Uses beginning-of-period balances for interest (no circular reference).
# Calls debt_schedule.sweep_debt() each year to determine ending balances.
#
# Index convention: index 0 = entry/Year 0 reference row; 1..N = projections.
#
# =============================================================================

import inputs as cfg
from debt_schedule import sweep_debt


def run_model(txn: dict) -> dict:
    """
    Run the full year-by-year model.

    Parameters
    ----------
    txn : output of transaction.compute_transaction()

    Returns
    -------
    dict of lists, each of length (HOLD_YEARS + 1).
    Index 0 = entry reference; indices 1..N = projection years.
    """
    N = cfg.HOLD_YEARS

    def zeros():
        return [0.0] * (N + 1)

    # ── Revenue & Income Statement ────────────────────────────────────────────
    revenue       = zeros()
    gross_profit  = zeros()
    sga           = zeros()
    rd            = zeros()
    ebitda        = zeros()
    ebitda_margin = zeros()
    da            = zeros()
    ebit          = zeros()

    # ── Interest & Fees ──────────────────────────────────────────────────────
    rev_interest      = zeros()   # revolver drawn interest + commitment fee
    tl_interest       = zeros()
    mezz_cash_int     = zeros()
    mezz_pik          = zeros()   # PIK accretion — P&L AND non-cash FCF addback
    total_cash_int    = zeros()   # cash interest only (feeds EBT)
    fee_amort         = zeros()   # amortization of upfront financing fees

    # ── Below-the-line ───────────────────────────────────────────────────────
    ebt              = zeros()
    taxes            = zeros()
    net_income       = zeros()

    # ── FCF Bridge ───────────────────────────────────────────────────────────
    capex            = zeros()
    delta_nwc        = zeros()
    fcf              = zeros()    # operating FCF before debt service
    mand_amort_arr   = zeros()
    fcf_before_opt   = zeros()    # FCF after mandatory amort

    # ── Balance Sheet ─────────────────────────────────────────────────────────
    cash             = zeros()
    rev_bal          = zeros()
    tl_bal           = zeros()
    mezz_bal         = zeros()

    # ── Debt Schedule Detail ──────────────────────────────────────────────────
    tl_mand          = zeros()
    tl_opt           = zeros()
    mezz_opt         = zeros()
    rev_draw         = zeros()
    rev_repay        = zeros()

    # ── Year 0: entry reference row ──────────────────────────────────────────
    revenue[0]       = cfg.BASE_REVENUE
    ebitda[0]        = cfg.ENTRY_EBITDA
    ebitda_margin[0] = cfg.ENTRY_EBITDA / cfg.BASE_REVENUE   # sanity check

    # Post-close balance sheet (beginning of Year 1)
    cash[0]    = cfg.CASH_TO_BS
    rev_bal[0] = 0.0
    tl_bal[0]  = txn["tl_principal"]
    mezz_bal[0]= txn["mezz_principal"]

    # ── Derived constants ─────────────────────────────────────────────────────
    tl_rate            = cfg.BASE_RATE + cfg.TERM_LOAN_SPREAD
    rev_drawn_rate     = cfg.BASE_RATE + cfg.REVOLVER_SPREAD
    annual_fee_amort   = (txn["total_financing_fees"] / cfg.FINANCING_FEE_AMORT_YEARS
                          if cfg.FINANCING_FEE_AMORT_YEARS > 0 else 0.0)
    orig_tl_principal  = txn["tl_principal"]

    # ── Projection Loop ───────────────────────────────────────────────────────
    for y in range(1, N + 1):

        # Revenue & gross profit
        revenue[y]      = revenue[y - 1] * (1.0 + cfg.REVENUE_GROWTH_RATE)
        gross_profit[y] = revenue[y] * cfg.GROSS_MARGIN
        sga[y]          = revenue[y] * cfg.SGA_PCT
        rd[y]           = revenue[y] * cfg.RD_PCT
        ebitda[y]       = gross_profit[y] - sga[y] - rd[y]
        ebitda_margin[y]= ebitda[y] / revenue[y]
        da[y]           = revenue[y] * cfg.DA_PCT
        ebit[y]         = ebitda[y] - da[y]

        # Interest — all from BEGINNING-OF-PERIOD balances (no circularity)
        rev_drawn_int    = rev_bal[y - 1] * rev_drawn_rate
        rev_commit       = (cfg.REVOLVER_CAPACITY - rev_bal[y - 1]) * cfg.REVOLVER_COMMITMENT_FEE
        rev_interest[y]  = rev_drawn_int + rev_commit
        tl_interest[y]   = tl_bal[y - 1]   * tl_rate
        mezz_cash_int[y] = mezz_bal[y - 1]  * cfg.MEZZ_CASH_RATE
        mezz_pik[y]      = mezz_bal[y - 1]  * cfg.MEZZ_PIK_RATE
        total_cash_int[y]= rev_interest[y] + tl_interest[y] + mezz_cash_int[y]

        # Financing fee amortization (straight-line; stops after amort period)
        fee_amort[y] = annual_fee_amort if y <= cfg.FINANCING_FEE_AMORT_YEARS else 0.0

        # Income statement
        # PIK flows through EBT (reduces taxes) then added back in FCF bridge
        ebt[y]       = ebit[y] - total_cash_int[y] - mezz_pik[y] - fee_amort[y]
        taxes[y]     = max(0.0, ebt[y]) * cfg.TAX_RATE   # no tax benefit on losses
        net_income[y]= ebt[y] - taxes[y]

        # FCF bridge
        capex[y]      = revenue[y] * cfg.CAPEX_PCT
        delta_nwc[y]  = revenue[y] * cfg.NWC_PCT
        fcf[y]        = (net_income[y]
                         + da[y]
                         + mezz_pik[y]      # non-cash addback
                         + fee_amort[y]     # non-cash addback
                         - capex[y]
                         - delta_nwc[y])

        # Mandatory amort — bounded by current TL balance
        mand_amort_arr[y] = min(orig_tl_principal * cfg.TERM_LOAN_AMORT_RATE,
                                tl_bal[y - 1])
        fcf_before_opt[y] = fcf[y] - mand_amort_arr[y]

        # Debt sweep
        ds = sweep_debt(
            beg_cash            = cash[y - 1],
            rev_beg             = rev_bal[y - 1],
            tl_beg              = tl_bal[y - 1],
            mezz_beg            = mezz_bal[y - 1],
            fcf_before_optional = fcf_before_opt[y],
            mand_amort          = mand_amort_arr[y],
            mezz_pik            = mezz_pik[y],
        )

        # Store ending balances
        cash[y]    = ds["end_cash"]
        rev_bal[y] = ds["rev_end"]
        tl_bal[y]  = ds["tl_end"]
        mezz_bal[y]= ds["mezz_end"]

        tl_mand[y]  = mand_amort_arr[y]
        tl_opt[y]   = ds["tl_repayment"]
        mezz_opt[y] = ds["mezz_repayment"]
        rev_draw[y] = ds["rev_draw"]
        rev_repay[y]= ds["rev_repayment"]

    return {
        "N":                    N,
        "annual_fee_amort":     annual_fee_amort,
        "tl_rate":              tl_rate,
        "rev_drawn_rate":       rev_drawn_rate,
        # Income statement
        "revenue":              revenue,
        "gross_profit":         gross_profit,
        "sga":                  sga,
        "rd":                   rd,
        "ebitda":               ebitda,
        "ebitda_margin":        ebitda_margin,
        "da":                   da,
        "ebit":                 ebit,
        # Interest
        "rev_interest":         rev_interest,
        "tl_interest":          tl_interest,
        "mezz_cash_interest":   mezz_cash_int,
        "mezz_pik":             mezz_pik,
        "total_cash_interest":  total_cash_int,
        "fee_amort":            fee_amort,
        # Below the line
        "ebt":                  ebt,
        "taxes":                taxes,
        "net_income":           net_income,
        # FCF bridge
        "capex":                capex,
        "delta_nwc":            delta_nwc,
        "fcf":                  fcf,
        "mandatory_amort":      mand_amort_arr,
        "fcf_before_optional":  fcf_before_opt,
        # Debt schedule detail
        "tl_mandatory":         tl_mand,
        "tl_optional":          tl_opt,
        "mezz_optional":        mezz_opt,
        "rev_draw":             rev_draw,
        "rev_repay":            rev_repay,
        # Balance sheet
        "cash":                 cash,
        "rev_bal":              rev_bal,
        "tl_bal":               tl_bal,
        "mezz_bal":             mezz_bal,
    }
