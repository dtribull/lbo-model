# =============================================================================
# main.py — Orchestration and terminal output
# =============================================================================
#
# Run:  python main.py
# Output: formatted terminal tables + lbo_output.xlsx in this directory.
#
# =============================================================================

import sys
import os

# Windows: force UTF-8 output so box-drawing characters render correctly
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Allow "python main.py" from any working directory
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import inputs as cfg
from cli          import collect_inputs
from transaction  import compute_transaction
from forecast     import run_model
from returns      import compute_returns
from sensitivity  import compute_sensitivity
from excel_export import write_excel


def _apply_overrides(overrides: dict):
    """
    Write every collected input into the live inputs module so all downstream
    modules see the updated values without any signature changes.
    Derived sensitivity bounds are recomputed from the new entry/exit multiples.
    """
    for key, val in overrides.items():
        setattr(cfg, key, val)
    # Re-derive sensitivity bounds if not explicitly overridden
    if "SENS_ENTRY_MIN" not in overrides:
        cfg.SENS_ENTRY_MIN = cfg.ENTRY_MULTIPLE - 1.0
    if "SENS_ENTRY_MAX" not in overrides:
        cfg.SENS_ENTRY_MAX = cfg.ENTRY_MULTIPLE + 1.0
    if "SENS_EXIT_MIN" not in overrides:
        cfg.SENS_EXIT_MIN  = cfg.EXIT_MULTIPLE  - 1.0
    if "SENS_EXIT_MAX" not in overrides:
        cfg.SENS_EXIT_MAX  = cfg.EXIT_MULTIPLE  + 1.0

# =============================================================================
# Terminal formatting helpers
# =============================================================================

W_LABEL = 38
W_COL   = 12

SEP  = "─" * (W_LABEL + (cfg.HOLD_YEARS + 1) * (W_COL + 1) + 2)
DSEP = "═" * (W_LABEL + (cfg.HOLD_YEARS + 1) * (W_COL + 1) + 2)


def _fmt_mm(v, dec=1) -> str:
    """$ in millions — parentheses for negatives, dash for zero."""
    if v is None:
        return " " * W_COL
    if abs(v) < 10 ** (-dec - 1):
        return f"{'  -':>{W_COL}}"
    if v < 0:
        return f"{'(' + f'{abs(v):,.{dec}f}' + ')':>{W_COL}}"
    return f"{v:>{W_COL},.{dec}f}"


def _fmt_pct(v) -> str:
    if v is None:
        return " " * W_COL
    if abs(v) < 0.00005:
        return f"{'  -':>{W_COL}}"
    return f"{v:>{W_COL - 1}.1%} "


def _fmt_mult(v) -> str:
    if v is None:
        return " " * W_COL
    return f"{v:>{W_COL - 1}.1f}x "


def _hdr(text: str):
    print()
    print(DSEP)
    print(f"  {text}")
    print(DSEP)


def _sub(text: str):
    print()
    print(f"  {text}")
    print(SEP)


def _yr_header(year_labels: list[str], txn_label="Txn"):
    cols = [f"{txn_label:>{W_COL}}"] + [f"{y:>{W_COL}}" for y in year_labels]
    print(f"  {'':>{W_LABEL - 2}}  " + "  ".join(cols))
    print(SEP)


def _row(label: str, values: list, fmt_fn, indent=2, bold_label=False):
    """Print one data row."""
    cols = "  ".join(fmt_fn(v) for v in values)
    lbl  = f"{'':>{indent}}{label}"
    print(f"  {lbl:<{W_LABEL - 2}}  {cols}")


def _total(label: str, values: list, fmt_fn):
    print(SEP)
    _row(label, values, fmt_fn, indent=2)
    print(SEP)


def _blank():
    print()


# =============================================================================
# Print sections
# =============================================================================

def print_transaction(txn: dict):
    _hdr("1.  TRANSACTION DRIVERS")

    _sub("Entry Valuation")
    _row("Entry EBITDA ($mm)",           [cfg.ENTRY_EBITDA],      _fmt_mm)
    _row("(x) LTM Entry Multiple",       [cfg.ENTRY_MULTIPLE],    _fmt_mult)
    _row("Purchase Enterprise Value",    [txn["purchase_ev"]],    _fmt_mm)
    _row("Less: Existing Debt",          [-cfg.EXISTING_DEBT],    _fmt_mm)
    _row("Plus: Existing Cash",          [cfg.EXISTING_CASH],     _fmt_mm)
    _total("Equity Value",               [txn["equity_value"]],   _fmt_mm)

    _sub("Transaction Assumptions")
    _row("LTM Exit Multiple",            [cfg.EXIT_MULTIPLE],        _fmt_mult)
    _row("Tax Rate",                     [cfg.TAX_RATE],             _fmt_pct)
    _row("Cash to B/S ($mm)",            [cfg.CASH_TO_BS],           _fmt_mm)
    _row("Transaction Expenses ($mm)",   [cfg.TRANSACTION_EXPENSES], _fmt_mm)
    _row("Rollover Equity (%)",          [cfg.ROLLOVER_EQUITY_PCT],  _fmt_pct)
    _row("Implied PE Ownership (%)",     [txn["pe_ownership"]],      _fmt_pct)

    _sub("Debt Sizing  (x Leverageable EBITDA = $300mm)")
    _row("Term Loan",  [txn["tl_leverage_x"],   txn["tl_principal"]],
         lambda v: (_fmt_mult(v) if v <= 10 else _fmt_mm(v)))
    _row("Mezzanine",  [txn["mezz_leverage_x"], txn["mezz_principal"]],
         lambda v: (_fmt_mult(v) if v <= 10 else _fmt_mm(v)))
    _total("Total Debt", [txn["total_leverage_x"],  txn["total_debt"]],
           lambda v: (_fmt_mult(v) if v <= 10 else _fmt_mm(v)))

    _sub("Sources & Uses")
    print(f"  {'USES':<{W_LABEL + W_COL - 8}}  {'SOURCES':>{W_LABEL // 2}}")
    pairs = [
        ("Purchase Equity Value",  txn["use_purchase_equity"],  "Term Loan",      txn["tl_principal"]),
        ("Refinancing Net Debt",   txn["use_refinancing_nd"],   "Mezzanine Debt", txn["mezz_principal"]),
        ("Cash to Balance Sheet",  txn["use_cash_to_bs"],       "Total Debt",     txn["total_debt"]),
        ("Underwriting Fees",      txn["use_fees"],             "Rollover Equity",txn["rollover_equity"]),
        ("Transaction Expenses",   txn["use_tx_expenses"],      "PE Investment",  txn["pe_equity_invest"]),
    ]
    for ul, uv, sl, sv in pairs:
        print(f"  {'  ' + ul:<{W_LABEL + W_COL - 4}}  {'  ' + sl:<{W_LABEL // 2}}  "
              f"{_fmt_mm(uv)}  {_fmt_mm(sv)}")
    print(SEP)
    print(f"  {'  Total Uses':<{W_LABEL + W_COL - 4}}  {'  Total Sources':<{W_LABEL // 2}}  "
          f"{_fmt_mm(txn['total_uses'])}  {_fmt_mm(txn['total_sources'])}")
    print(SEP)
    chk = txn["sources_uses_check"]
    flag = "✓ BALANCED" if abs(chk) < 0.01 else "✗ MISMATCH"
    print(f"  {flag}")


def print_forecast(results: dict):
    _hdr("2.  FINANCIAL FORECAST  ($ in millions)")
    N = results["N"]
    yr_labels = [f"Year {y}" for y in range(1, N + 1)]
    _yr_header(yr_labels, txn_label="Year 0")

    # Helper: values[0] = entry ref; values[1..N] = projections
    def yr(arr):
        return [arr[y] for y in range(N + 1)]

    def yr1N(arr):
        return [None] + [arr[y] for y in range(1, N + 1)]

    _row("Revenue",         yr(results["revenue"]),       _fmt_mm)
    growth = [None] + [results["revenue"][y] / results["revenue"][y-1] - 1
                       for y in range(1, N+1)]
    _row("% Growth",        growth,                       _fmt_pct, indent=4)
    _blank()

    _row("Gross Profit",    yr1N(results["gross_profit"]), _fmt_mm)
    gpm = [None] + [results["gross_profit"][y]/results["revenue"][y] for y in range(1, N+1)]
    _row("% Margin",        gpm,                           _fmt_pct, indent=4)
    _blank()

    _row("Less: SG&A",      [None] + [-results["sga"][y]   for y in range(1, N+1)], _fmt_mm)
    _row("Less: R&D",       [None] + [-results["rd"][y]    for y in range(1, N+1)], _fmt_mm)
    _total("EBITDA",        yr(results["ebitda"]),          _fmt_mm)
    _row("% Margin",        yr(results["ebitda_margin"]),   _fmt_pct, indent=4)
    _blank()

    _row("Less: D&A",       [None] + [-results["da"][y]    for y in range(1, N+1)], _fmt_mm)
    _total("EBIT",          yr1N(results["ebit"]),          _fmt_mm)
    _blank()

    _row("Less: Cash Interest",  [None]+[-results["total_cash_interest"][y] for y in range(1,N+1)],_fmt_mm)
    _row("Less: PIK Interest",   [None]+[-results["mezz_pik"][y]            for y in range(1,N+1)],_fmt_mm)
    _row("Less: Fee Amort",      [None]+[-results["fee_amort"][y]           for y in range(1,N+1)],_fmt_mm)
    _total("EBT",                yr1N(results["ebt"]),       _fmt_mm)
    _row("Less: Taxes",          [None]+[-results["taxes"][y]               for y in range(1,N+1)],_fmt_mm)
    _total("Net Income",         yr1N(results["net_income"]), _fmt_mm)
    _blank()

    _sub("Free Cash Flow Bridge")
    _yr_header(yr_labels, txn_label="Year 0")
    _row("Net Income",           yr1N(results["net_income"]),  _fmt_mm)
    _row("Plus: D&A",            yr1N(results["da"]),          _fmt_mm)
    _row("Plus: PIK Interest",   yr1N(results["mezz_pik"]),    _fmt_mm)
    _row("Plus: Fee Amort",      yr1N(results["fee_amort"]),   _fmt_mm)
    _row("Less: Capex",          [None]+[-results["capex"][y]     for y in range(1,N+1)],_fmt_mm)
    _row("Less: Increase in NWC",[None]+[-results["delta_nwc"][y] for y in range(1,N+1)],_fmt_mm)
    _total("Free Cash Flow",     yr1N(results["fcf"]),          _fmt_mm)
    _row("Less: Mandatory Amort",[None]+[-results["mandatory_amort"][y] for y in range(1,N+1)],_fmt_mm)
    _total("FCF Before Optional Paydown", yr1N(results["fcf_before_optional"]), _fmt_mm)


def print_debt_schedule(txn: dict, results: dict):
    _hdr("3.  DEBT SCHEDULE  ($ in millions)")
    N = results["N"]
    yr_labels = [f"Year {y}" for y in range(1, N + 1)]

    # Revolver
    _sub("Revolver")
    _yr_header(yr_labels, txn_label="Txn")
    rev_beg = [0.0] + [results["rev_bal"][y-1] for y in range(1, N+1)]
    _row("Beginning Balance",        rev_beg,                          _fmt_mm)
    _row("  Optional Draw",          [None]+[results["rev_draw"][y]  for y in range(1,N+1)],_fmt_mm)
    _row("  (Repayment)",            [None]+[-results["rev_repay"][y] for y in range(1,N+1)],_fmt_mm)
    _total("Ending Balance",         [0.0]+[results["rev_bal"][y]    for y in range(1,N+1)],_fmt_mm)
    _row("Commitment Fee (undrawn)", [None]+[results["rev_interest"][y] for y in range(1,N+1)],_fmt_mm)

    # Term Loan
    _sub("Term Loan")
    _yr_header(yr_labels, txn_label="Txn")
    tl_beg = [txn["tl_principal"]] + [results["tl_bal"][y-1] for y in range(1, N+1)]
    _row("Beginning Balance",        tl_beg,                                 _fmt_mm)
    _row("  (Mandatory Amort)",      [None]+[-results["mandatory_amort"][y]  for y in range(1,N+1)],_fmt_mm)
    _row("  (Discretionary Repmt.)", [None]+[-results["tl_optional"][y]      for y in range(1,N+1)],_fmt_mm)
    _total("Ending Balance",         [txn["tl_principal"]]+[results["tl_bal"][y] for y in range(1,N+1)],_fmt_mm)
    _row("All-in Rate",              [None]+[results["tl_rate"]]*N,          _fmt_pct)
    _row("Interest Expense",         [None]+[results["tl_interest"][y]       for y in range(1,N+1)],_fmt_mm)

    # Mezzanine
    _sub("Mezzanine (PIK) Debt")
    _yr_header(yr_labels, txn_label="Txn")
    mz_beg = [txn["mezz_principal"]] + [results["mezz_bal"][y-1] for y in range(1, N+1)]
    _row("Beginning Balance",        mz_beg,                                     _fmt_mm)
    _row("  Plus: PIK Accretion",    [None]+[results["mezz_pik"][y]              for y in range(1,N+1)],_fmt_mm)
    _row("  (Discretionary Repmt.)", [None]+[-results["mezz_optional"][y]         for y in range(1,N+1)],_fmt_mm)
    _total("Ending Balance",         [txn["mezz_principal"]]+[results["mezz_bal"][y] for y in range(1,N+1)],_fmt_mm)
    _row("Cash Rate",                [None]+[cfg.MEZZ_CASH_RATE]*N,             _fmt_pct)
    _row("PIK Rate",                 [None]+[cfg.MEZZ_PIK_RATE]*N,              _fmt_pct)
    _row("Cash Interest Expense",    [None]+[results["mezz_cash_interest"][y]   for y in range(1,N+1)],_fmt_mm)


def print_interest_summary(results: dict):
    _hdr("4.  INTEREST CALCULATION  ($ in millions)")
    N = results["N"]
    yr_labels = [f"Year {y}" for y in range(1, N + 1)]
    _yr_header(yr_labels, txn_label="Year 0")

    def y1N(arr):
        return [None] + [arr[y] for y in range(1, N+1)]

    _row("Revolver (incl. commitment fee)", y1N(results["rev_interest"]),       _fmt_mm)
    _row("Term Loan",                        y1N(results["tl_interest"]),        _fmt_mm)
    _row("Mezzanine — Cash Interest",        y1N(results["mezz_cash_interest"]), _fmt_mm)
    _row("Mezzanine — PIK Accretion",        y1N(results["mezz_pik"]),           _fmt_mm)
    _total("Total Cash Interest Expense",    y1N(results["total_cash_interest"]),_fmt_mm)
    _blank()
    _row("Amortization of Financing Fees",   y1N(results["fee_amort"]),          _fmt_mm)
    tot = [None] + [results["total_cash_interest"][y]+results["fee_amort"][y] for y in range(1,N+1)]
    _total("Total Int. + Fee Amort",         tot,                                _fmt_mm)


def print_balance_sheet(txn: dict, results: dict):
    _hdr("5.  BALANCE SHEET SUMMARY  ($ in millions)")
    N = results["N"]
    yr_labels = [f"Year {y}" for y in range(1, N + 1)]
    _yr_header(yr_labels, txn_label="Txn")

    cash_row    = [cfg.CASH_TO_BS] + [results["cash"][y] for y in range(1, N+1)]
    rev_row     = [0.0]             + [results["rev_bal"][y] for y in range(1, N+1)]
    mezz_row    = [txn["mezz_principal"]] + [results["mezz_bal"][y] for y in range(1, N+1)]
    tl_row      = [txn["tl_principal"]]   + [results["tl_bal"][y]   for y in range(1, N+1)]
    td_row      = [txn["total_debt"]] + [
        results["tl_bal"][y]+results["mezz_bal"][y]+results["rev_bal"][y]
        for y in range(1, N+1)]
    nd_row      = [td_row[i] - cash_row[i] for i in range(N + 1)]
    lev_row     = [td_row[i] / results["ebitda"][max(i,1)] if results["ebitda"][max(i,1)] else None
                   for i in range(N+1)]
    lev_row[0]  = td_row[0] / cfg.ENTRY_EBITDA
    nlev_row    = [nd_row[i] / results["ebitda"][max(i,1)] if results["ebitda"][max(i,1)] else None
                   for i in range(N+1)]
    nlev_row[0] = nd_row[0] / cfg.ENTRY_EBITDA

    _row("Cash",                 cash_row, _fmt_mm)
    _blank()
    _row("Revolver",             rev_row,  _fmt_mm)
    _row("Mezzanine (PIK)",      mezz_row, _fmt_mm)
    _row("Term Loan",            tl_row,   _fmt_mm)
    _total("Total Debt",         td_row,   _fmt_mm)
    pct_init = [None] + [td_row[y]/txn["total_debt"] for y in range(1, N+1)]
    _row("% of Initial Debt",    pct_init, _fmt_pct, indent=4)
    _blank()
    _total("Total Net Debt",     nd_row,   _fmt_mm)
    _blank()
    _row("Total Debt / EBITDA",  lev_row,  _fmt_mult)
    _row("Net Debt / EBITDA",    nlev_row, _fmt_mult)


def print_returns(txn: dict, returns_data: list):
    _hdr("6.  RETURNS CALCULATION  ($ in millions)")
    W_RC = W_COL
    yr_labels = [f"Exit Yr {r['exit_year']}" for r in returns_data]
    print(f"  {'':>{W_LABEL - 2}}  " +
          "  ".join(f"{y:>{W_RC}}" for y in yr_labels))
    print(SEP)

    def ret_row(label, vals, fmt_fn):
        cols = "  ".join(fmt_fn(v) for v in vals)
        print(f"  {label:<{W_LABEL - 2}}  {cols}")

    ret_row("LTM EBITDA",           [r["ltm_ebitda"]   for r in returns_data], _fmt_mm)
    ret_row("(x) Exit Multiple",    [cfg.EXIT_MULTIPLE for r in returns_data], _fmt_mult)
    ret_row("Total Enterprise Value",[r["exit_ev"]     for r in returns_data], _fmt_mm)
    ret_row("Less: Net Debt",       [-r["net_debt"]    for r in returns_data], _fmt_mm)
    print(SEP)
    ret_row("Equity Value",         [r["equity_value"] for r in returns_data], _fmt_mm)
    ret_row("(x) PE Ownership",     [txn["pe_ownership"] for _ in returns_data], _fmt_pct)
    print(SEP)
    ret_row("Value to PE Firm",     [r["value_to_pe"]  for r in returns_data], _fmt_mm)
    print()
    ret_row("PE Equity Investment", [txn["pe_equity_invest"]] * len(returns_data), _fmt_mm)
    print(SEP)
    ret_row("IRR",                  [r["irr"]  for r in returns_data], _fmt_pct)
    ret_row("MoM",                  [r["mom"]  for r in returns_data], _fmt_mult)
    print(SEP)


def print_sensitivity(sens: dict):
    _hdr("7.  SENSITIVITY TABLES  (Terminal Year Exit)")
    entry_mults = sens["entry_multiples"]
    exit_mults  = sens["exit_multiples"]

    def print_table(title, table, fmt_fn):
        print(f"\n  {title}")
        print(SEP)
        W_S = 10
        hdr = f"  {'Entry \\ Exit':<12}" + "".join(f"{em:.1f}x".rjust(W_S) for em in exit_mults)
        print(hdr)
        print(SEP)
        for i, em in enumerate(entry_mults):
            is_base_row = abs(float(em) - cfg.ENTRY_MULTIPLE) < 0.001
            row_str = f"  {em:.1f}x        " + "".join(
                fmt_fn(float(table[i, j])).rjust(W_S) for j in range(len(exit_mults))
            )
            if is_base_row:
                row_str = row_str.replace("  ", "→ ", 1)
            print(row_str)
        print(SEP)

    print_table("IRR Sensitivity", sens["irr_table"], _fmt_pct)
    print_table("MoM Sensitivity", sens["mom_table"], _fmt_mult)


# =============================================================================
# Entry point
# =============================================================================

def main():
    # ── 1. Collect inputs interactively ──────────────────────────────────────
    overrides = collect_inputs()
    _apply_overrides(overrides)

    # ── 2. Run model ──────────────────────────────────────────────────────────
    txn          = compute_transaction()
    results      = run_model(txn)
    returns_data = compute_returns(txn, results)
    sens         = compute_sensitivity(txn, results)

    # ── 3. Terminal output ────────────────────────────────────────────────────
    print_transaction(txn)
    print_forecast(results)
    print_debt_schedule(txn, results)
    print_interest_summary(results)
    print_balance_sheet(txn, results)
    print_returns(txn, returns_data)
    print_sensitivity(sens)

    # ── 4. Excel export (Inputs sheet + Model sheet with live formulas) ───────
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "..", "lbo_output.xlsx")
    write_excel(txn, results, returns_data, sens,
                overrides=overrides, output_path=output_path)
    print()


if __name__ == "__main__":
    main()
