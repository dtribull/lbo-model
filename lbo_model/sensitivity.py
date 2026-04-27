# =============================================================================
# sensitivity.py — 5×5 IRR and MoM sensitivity tables
# =============================================================================
#
# Both tables show TERMINAL-YEAR (Year N) exit only.
# Rows: entry multiples    Columns: exit multiples
#
# Key insight: changing the ENTRY multiple only affects the PE equity investment
# (since debt is sized off LEVERAGEABLE_EBITDA, not entry multiple).
# The debt schedule, FCF, and EBITDA at exit are identical across all entry
# multiples — so we can reuse the base-case results and only recompute
# pe_equity_invest and value_to_pe for each cell.
#
# Changing the EXIT multiple only affects exit_ev and therefore value_to_pe.
# =============================================================================

import numpy as np
import numpy_financial as npf
import inputs as cfg


def _equity_invest_for_entry(entry_multiple: float, base_txn: dict) -> float:
    """Recompute pe_equity_invest for a given entry multiple, holding all else constant."""
    purchase_ev  = cfg.ENTRY_EBITDA * entry_multiple
    equity_value = purchase_ev - cfg.EXISTING_DEBT + cfg.EXISTING_CASH
    refinancing  = cfg.EXISTING_DEBT - cfg.EXISTING_CASH
    total_uses   = (equity_value + refinancing
                    + cfg.CASH_TO_BS
                    + base_txn["total_financing_fees"]
                    + cfg.TRANSACTION_EXPENSES)
    total_equity = total_uses - base_txn["total_debt"]
    return total_equity * (1.0 - cfg.ROLLOVER_EQUITY_PCT)


def compute_sensitivity(base_txn: dict, base_results: dict) -> dict:
    """
    Build IRR and MoM tables for the terminal exit year.

    Parameters
    ----------
    base_txn     : from transaction.compute_transaction()
    base_results : from forecast.run_model()

    Returns
    -------
    dict with entry_multiples, exit_multiples, irr_table, mom_table arrays.
    """
    N   = cfg.HOLD_YEARS
    own = base_txn["pe_ownership"]

    # ── Build axis arrays ────────────────────────────────────────────────────
    entry_mults = np.round(
        np.arange(cfg.SENS_ENTRY_MIN,
                  cfg.SENS_ENTRY_MAX + cfg.SENS_ENTRY_STEP / 2,
                  cfg.SENS_ENTRY_STEP), 4)
    exit_mults  = np.round(
        np.arange(cfg.SENS_EXIT_MIN,
                  cfg.SENS_EXIT_MAX  + cfg.SENS_EXIT_STEP  / 2,
                  cfg.SENS_EXIT_STEP), 4)

    n_entry = len(entry_mults)
    n_exit  = len(exit_mults)

    irr_table = np.full((n_entry, n_exit), np.nan)
    mom_table = np.full((n_entry, n_exit), np.nan)

    # Terminal-year values that are FIXED across all sensitivity runs
    ltm_ebitda  = base_results["ebitda"][N]
    total_debt  = (base_results["tl_bal"][N]
                   + base_results["mezz_bal"][N]
                   + base_results["rev_bal"][N])
    net_debt    = total_debt - base_results["cash"][N]

    for i, em in enumerate(entry_mults):
        pe_invest = _equity_invest_for_entry(float(em), base_txn)

        for j, xm in enumerate(exit_mults):
            exit_ev     = ltm_ebitda * float(xm)
            equity_val  = exit_ev - net_debt
            value_to_pe = equity_val * own

            cf = [-pe_invest] + [0.0] * (N - 1) + [value_to_pe]
            try:
                irr = float(npf.irr(cf))
            except Exception:
                irr = np.nan

            irr_table[i, j] = irr
            mom_table[i, j] = value_to_pe / pe_invest if pe_invest > 0 else np.nan

    return {
        "entry_multiples": entry_mults,
        "exit_multiples":  exit_mults,
        "irr_table":       irr_table,
        "mom_table":       mom_table,
    }
