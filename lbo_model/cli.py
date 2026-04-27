# =============================================================================
# cli.py — Interactive CLI prompts for all model inputs
# =============================================================================
#
# Walks the user through every assumption, grouped by section.
# Displays the current default in brackets — press Enter to accept.
# Returns a dict whose keys match the constant names in inputs.py.
#
# =============================================================================

import inputs as _defaults


# ── Prompt helpers ─────────────────────────────────────────────────────────────

def _ask(label: str, default, suffix: str = "", cast=float) -> object:
    """
    Prompt user with  '  {label} [{default}{suffix}]: '
    Strips $, commas, and trailing x/% before casting.
    If the user enters '%' we auto-divide by 100 (e.g. 38 → 0.38).
    """
    default_display = f"{default:,.4g}" if isinstance(default, float) else str(default)
    prompt_str = f"  {label} [{default_display}{suffix}]: "
    while True:
        raw = input(prompt_str).strip()
        if not raw:
            return default
        try:
            cleaned = raw.replace(",", "").replace("$", "").rstrip("x ")
            if cleaned.endswith("%"):
                cleaned = cleaned.rstrip("%")
                return float(cleaned) / 100.0
            return cast(cleaned)
        except ValueError:
            print(f"    ↳  Please enter a valid number (e.g. {default_display})")


def _ask_bool(label: str, default: bool) -> bool:
    default_str = "Y" if default else "N"
    raw = input(f"  {label} [{'Y' if default else 'N'}] (Y/N): ").strip().upper()
    if not raw:
        return default
    return raw.startswith("Y")


def _section(title: str):
    print()
    print(f"  {'─' * 50}")
    print(f"  {title}")
    print(f"  {'─' * 50}")


# ── Main collection function ───────────────────────────────────────────────────

def collect_inputs() -> dict:
    """
    Walk through every model assumption, grouped by section.
    Returns a dict of overrides whose keys match inputs.py constants.
    Only returns values that were explicitly changed OR all of them (safe to apply all).
    """
    d = _defaults

    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║   PEAK FRAMEWORKS — LBO MODEL INPUTS                ║")
    print("  ║   Press Enter to accept defaults shown in [ ]       ║")
    print("  ╚══════════════════════════════════════════════════════╝")

    cfg = {}

    # ── Entry Valuation ────────────────────────────────────────────────────────
    _section("1 / 5  ENTRY VALUATION")
    cfg["ENTRY_EBITDA"]    = _ask("Entry EBITDA ($mm)",        d.ENTRY_EBITDA)
    cfg["ENTRY_MULTIPLE"]  = _ask("Entry Multiple",            d.ENTRY_MULTIPLE, "x")
    cfg["EXISTING_DEBT"]   = _ask("Existing Debt ($mm)",       d.EXISTING_DEBT)
    cfg["EXISTING_CASH"]   = _ask("Existing Cash ($mm)",       d.EXISTING_CASH)

    # ── Transaction ────────────────────────────────────────────────────────────
    _section("2 / 5  TRANSACTION")
    cfg["EXIT_MULTIPLE"]         = _ask("Exit Multiple",              d.EXIT_MULTIPLE,        "x")
    cfg["TAX_RATE"]              = _ask("Tax Rate",                   d.TAX_RATE,             "%")
    cfg["CASH_TO_BS"]            = _ask("Cash to Balance Sheet ($mm)",d.CASH_TO_BS)
    cfg["TRANSACTION_EXPENSES"]  = _ask("Transaction Expenses ($mm)", d.TRANSACTION_EXPENSES)
    cfg["ROLLOVER_EQUITY_PCT"]   = _ask("Rollover Equity (%)",        d.ROLLOVER_EQUITY_PCT,  "%")

    # ── Debt ──────────────────────────────────────────────────────────────────
    _section("3 / 5  DEBT")
    cfg["BASE_RATE"]              = _ask("Flat Base Rate / SOFR",         d.BASE_RATE,              "%")
    cfg["LEVERAGEABLE_EBITDA"]    = _ask("Leverageable EBITDA ($mm)",     d.LEVERAGEABLE_EBITDA)
    print()
    print("  — Term Loan —")
    cfg["TERM_LOAN_LEVERAGE"]     = _ask("  Leverage (x EBITDA)",         d.TERM_LOAN_LEVERAGE,     "x")
    cfg["TERM_LOAN_SPREAD"]       = _ask("  Spread (L+)",                 d.TERM_LOAN_SPREAD,       "%")
    cfg["TERM_LOAN_FEE_PCT"]      = _ask("  Upfront Fee (%)",             d.TERM_LOAN_FEE_PCT,      "%")
    cfg["TERM_LOAN_AMORT_RATE"]   = _ask("  Mandatory Amort (% of orig)", d.TERM_LOAN_AMORT_RATE,   "%")
    print()
    print("  — Mezzanine / PIK Debt —")
    cfg["MEZZ_LEVERAGE"]          = _ask("  Leverage (x EBITDA)",         d.MEZZ_LEVERAGE,          "x")
    cfg["MEZZ_CASH_RATE"]         = _ask("  Cash Pay Rate",               d.MEZZ_CASH_RATE,         "%")
    cfg["MEZZ_PIK_RATE"]          = _ask("  PIK Rate",                    d.MEZZ_PIK_RATE,          "%")
    cfg["MEZZ_FEE_PCT"]           = _ask("  Upfront Fee (%)",             d.MEZZ_FEE_PCT,           "%")
    cfg["MEZZ_CASH_SWEEP"]        = _ask_bool("  Sweep Mezz after TL paid off?", d.MEZZ_CASH_SWEEP)
    print()
    print("  — Revolver —")
    cfg["REVOLVER_CAPACITY"]      = _ask("  Capacity ($mm)",              d.REVOLVER_CAPACITY)
    cfg["REVOLVER_COMMITMENT_FEE"]= _ask("  Commitment Fee (undrawn)",    d.REVOLVER_COMMITMENT_FEE,"%")
    cfg["REVOLVER_SPREAD"]        = _ask("  Spread (L+)",                 d.REVOLVER_SPREAD,        "%")
    print()
    cfg["MIN_CASH_BALANCE"]          = _ask("Min Cash Balance ($mm)",     d.MIN_CASH_BALANCE)
    cfg["FINANCING_FEE_AMORT_YEARS"] = _ask("Fee Amortization (years)",  d.FINANCING_FEE_AMORT_YEARS,
                                             cast=int)

    # ── Financial Forecast ────────────────────────────────────────────────────
    _section("4 / 5  FINANCIAL FORECAST (Simple Top-Line Model)")
    cfg["BASE_REVENUE"]         = _ask("Base Revenue — Year 0 ($mm)",  d.BASE_REVENUE)
    cfg["REVENUE_GROWTH_RATE"]  = _ask("Annual Revenue Growth Rate",   d.REVENUE_GROWTH_RATE,  "%")
    cfg["GROSS_MARGIN"]         = _ask("Gross Margin (% of Revenue)",  d.GROSS_MARGIN,         "%")
    cfg["SGA_PCT"]              = _ask("SG&A (% of Revenue)",          d.SGA_PCT,              "%")
    cfg["RD_PCT"]               = _ask("R&D  (% of Revenue)",          d.RD_PCT,               "%")
    cfg["DA_PCT"]               = _ask("D&A  (% of Revenue)",          d.DA_PCT,               "%")
    cfg["CAPEX_PCT"]            = _ask("Capex (% of Revenue)",         d.CAPEX_PCT,            "%")
    cfg["NWC_PCT"]              = _ask("Increase in NWC (% of Rev)",   d.NWC_PCT,              "%")

    # ── Hold Period & Sensitivity ──────────────────────────────────────────────
    _section("5 / 5  HOLD PERIOD & SENSITIVITY")
    cfg["HOLD_YEARS"]        = _ask("Hold Period (years)",     d.HOLD_YEARS,     cast=int)
    print()
    print("  Sensitivity table axes  (entry/exit multiple grid):")
    cfg["SENS_ENTRY_MIN"]    = _ask("  Entry axis min",        cfg["ENTRY_MULTIPLE"] - 1.0, "x")
    cfg["SENS_ENTRY_MAX"]    = _ask("  Entry axis max",        cfg["ENTRY_MULTIPLE"] + 1.0, "x")
    cfg["SENS_ENTRY_STEP"]   = _ask("  Entry axis step",       d.SENS_ENTRY_STEP,           "x")
    cfg["SENS_EXIT_MIN"]     = _ask("  Exit axis min",         cfg["EXIT_MULTIPLE"]  - 1.0, "x")
    cfg["SENS_EXIT_MAX"]     = _ask("  Exit axis max",         cfg["EXIT_MULTIPLE"]  + 1.0, "x")
    cfg["SENS_EXIT_STEP"]    = _ask("  Exit axis step",        d.SENS_EXIT_STEP,            "x")

    print()
    print("  ✓  All inputs collected — running model...")
    print()
    return cfg
