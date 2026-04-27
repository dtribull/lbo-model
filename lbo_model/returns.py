# =============================================================================
# returns.py — IRR and MoM for each possible exit year
# =============================================================================

import numpy_financial as npf
import inputs as cfg


def compute_returns(txn: dict, results: dict) -> list[dict]:
    """
    Compute IRR and MoM for every exit year from Year 1 through HOLD_YEARS.

    Cash flow convention (PE firm's perspective):
      Year 0 : -pe_equity_invest  (equity check written at close)
      Year Y :  value_to_pe       (proceeds at exit)
      Intervening years: 0

    Returns
    -------
    List of dicts, one per exit year, sorted by exit year.
    """
    pe_invest  = txn["pe_equity_invest"]
    pe_own     = txn["pe_ownership"]
    N          = cfg.HOLD_YEARS

    rows = []
    for y in range(1, N + 1):
        ltm_ebitda = results["ebitda"][y]

        exit_ev      = ltm_ebitda * cfg.EXIT_MULTIPLE
        total_debt_y = results["tl_bal"][y] + results["mezz_bal"][y] + results["rev_bal"][y]
        net_debt_y   = total_debt_y - results["cash"][y]
        equity_val   = exit_ev - net_debt_y
        value_to_pe  = equity_val * pe_own

        # IRR cash flow array: [−invest, 0, ..., 0, proceeds]
        cf = [-pe_invest] + [0.0] * (y - 1) + [value_to_pe]
        try:
            irr = float(npf.irr(cf))
        except Exception:
            irr = float("nan")

        mom = value_to_pe / pe_invest if pe_invest > 0 else float("nan")

        rows.append({
            "exit_year":    y,
            "ltm_ebitda":   ltm_ebitda,
            "exit_ev":      exit_ev,
            "total_debt":   total_debt_y,
            "net_debt":     net_debt_y,
            "equity_value": equity_val,
            "value_to_pe":  value_to_pe,
            "irr":          irr,
            "mom":          mom,
        })

    return rows
