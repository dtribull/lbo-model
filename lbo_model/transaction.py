# =============================================================================
# transaction.py — Entry valuation bridge and Sources & Uses
# =============================================================================

import inputs as cfg


def compute_transaction() -> dict:
    """
    Build the transaction drivers block:
      • Entry valuation bridge (EV → equity value)
      • Debt sizing (off LEVERAGEABLE_EBITDA)
      • Sources & Uses table
      • Implied PE ownership and equity investment

    Returns a flat dict used downstream by forecast, returns, and sensitivity.
    """

    # ── Entry Valuation ──────────────────────────────────────────────────────
    purchase_ev    = cfg.ENTRY_EBITDA * cfg.ENTRY_MULTIPLE
    equity_value   = purchase_ev - cfg.EXISTING_DEBT + cfg.EXISTING_CASH
    # "Refinancing Net Debt" — what we repay at close
    refinancing_nd = cfg.EXISTING_DEBT - cfg.EXISTING_CASH

    # ── Debt Sizing ──────────────────────────────────────────────────────────
    tl_principal   = cfg.LEVERAGEABLE_EBITDA * cfg.TERM_LOAN_LEVERAGE
    mezz_principal = cfg.LEVERAGEABLE_EBITDA * cfg.MEZZ_LEVERAGE
    total_debt     = tl_principal + mezz_principal

    # ── Financing Fees (upfront) ─────────────────────────────────────────────
    tl_fee         = tl_principal   * cfg.TERM_LOAN_FEE_PCT
    mezz_fee       = mezz_principal * cfg.MEZZ_FEE_PCT
    rev_fee        = cfg.REVOLVER_CAPACITY * cfg.REVOLVER_FEE_PCT
    total_fees     = tl_fee + mezz_fee + rev_fee

    # ── Sources & Uses ───────────────────────────────────────────────────────
    # Uses
    use_purchase_equity = equity_value        # buy out existing equity holders
    use_refinancing_nd  = refinancing_nd      # repay existing net debt
    use_cash_to_bs      = cfg.CASH_TO_BS      # operating cash post-close
    use_fees            = total_fees          # underwriting / OID fees
    use_tx_expenses     = cfg.TRANSACTION_EXPENSES
    total_uses          = (use_purchase_equity + use_refinancing_nd
                           + use_cash_to_bs + use_fees + use_tx_expenses)

    # Sources
    total_equity        = total_uses - total_debt
    rollover_equity     = total_equity * cfg.ROLLOVER_EQUITY_PCT
    pe_equity_invest    = total_equity * (1.0 - cfg.ROLLOVER_EQUITY_PCT)
    total_sources       = total_debt + total_equity

    # Sanity check
    sources_uses_check  = round(total_sources - total_uses, 6)

    pe_ownership        = 1.0 - cfg.ROLLOVER_EQUITY_PCT

    return {
        # Entry valuation
        "purchase_ev":          purchase_ev,
        "equity_value":         equity_value,
        "refinancing_nd":       refinancing_nd,
        # Debt tranches
        "tl_principal":         tl_principal,
        "mezz_principal":       mezz_principal,
        "total_debt":           total_debt,
        # Fees
        "tl_fee":               tl_fee,
        "mezz_fee":             mezz_fee,
        "rev_fee":              rev_fee,
        "total_financing_fees": total_fees,
        # Uses detail
        "use_purchase_equity":  use_purchase_equity,
        "use_refinancing_nd":   use_refinancing_nd,
        "use_cash_to_bs":       use_cash_to_bs,
        "use_fees":             use_fees,
        "use_tx_expenses":      use_tx_expenses,
        "total_uses":           total_uses,
        # Sources detail
        "rollover_equity":      rollover_equity,
        "pe_equity_invest":     pe_equity_invest,
        "total_equity":         total_equity,
        "total_sources":        total_sources,
        "sources_uses_check":   sources_uses_check,
        # Ownership
        "pe_ownership":         pe_ownership,
        # Leverage at entry (x entry EBITDA)
        "tl_leverage_x":        tl_principal   / cfg.LEVERAGEABLE_EBITDA,
        "mezz_leverage_x":      mezz_principal / cfg.LEVERAGEABLE_EBITDA,
        "total_leverage_x":     total_debt     / cfg.LEVERAGEABLE_EBITDA,
    }
