# =============================================================================
# inputs.py — All model assumptions
# Change values here; do not edit other modules for scenario analysis.
# =============================================================================

# ─── Entry Valuation ─────────────────────────────────────────────────────────
ENTRY_EBITDA      = 300      # $mm — used for deal sizing only
ENTRY_MULTIPLE    = 11.0     # x LTM EBITDA at entry
EXISTING_DEBT     = 500      # $mm — target company debt being refinanced
EXISTING_CASH     = 100      # $mm — target company cash netted in bridge

# ─── Transaction ─────────────────────────────────────────────────────────────
EXIT_MULTIPLE         = 11.0   # x LTM EBITDA at exit
TAX_RATE              = 0.38
CASH_TO_BS            = 50     # $mm — minimum cash left on balance sheet at close
TRANSACTION_EXPENSES  = 25     # $mm — legal, advisory, etc.
ROLLOVER_EQUITY_PCT   = 0.40   # fraction of total equity rolled by mgmt/sellers

# ─── Debt ────────────────────────────────────────────────────────────────────
LEVERAGEABLE_EBITDA   = 300    # $mm — EBITDA used to size debt tranches

# Floating base rate (flat; applied to all floating-rate tranches)
BASE_RATE             = 0.05   # e.g. SOFR; change to match market assumptions

# Term Loan
TERM_LOAN_LEVERAGE    = 2.0    # x LEVERAGEABLE_EBITDA → principal
TERM_LOAN_SPREAD      = 0.035  # L+350; all-in rate = BASE_RATE + SPREAD = 8.5%
TERM_LOAN_FEE_PCT     = 0.02   # OID / upfront fee as % of principal → fee $
TERM_LOAN_AMORT_RATE  = 0.00   # mandatory annual amort as % of original principal

# Mezzanine / PIK Debt
MEZZ_LEVERAGE         = 2.0    # x LEVERAGEABLE_EBITDA → principal
MEZZ_CASH_RATE        = 0.00   # cash-pay interest component
MEZZ_PIK_RATE         = 0.05   # PIK component — accretes to balance; flows through P&L
MEZZ_FEE_PCT          = 0.00   # upfront fee as % of principal
MEZZ_CASH_SWEEP       = False  # True → sweep mezz after TL fully repaid
                                # False → matches reference model (cash accumulates)

# Revolver
REVOLVER_CAPACITY        = 100   # $mm — total facility size
REVOLVER_COMMITMENT_FEE  = 0.01  # annual fee on UNDRAWN balance
REVOLVER_SPREAD          = 0.04  # L+400; all-in drawn rate = BASE_RATE + SPREAD
REVOLVER_FEE_PCT         = 0.00  # upfront facility fee as % of capacity

# Financing fee amortization (straight-line)
FINANCING_FEE_AMORT_YEARS = 7

# Minimum cash balance — revolver draws before cash falls below this
MIN_CASH_BALANCE = 50   # $mm

# ─── Financial Forecast ──────────────────────────────────────────────────────
BASE_REVENUE        = 2_000   # $mm — Year 0 (entry year) revenue
REVENUE_GROWTH_RATE = 0.11    # flat annual growth rate applied each year
GROSS_MARGIN        = 0.40    # Gross Profit / Revenue
SGA_PCT             = 0.10    # SG&A / Revenue
RD_PCT              = 0.10    # R&D / Revenue
DA_PCT              = 0.02    # D&A / Revenue
CAPEX_PCT           = 0.02    # Capex / Revenue
NWC_PCT             = 0.01    # Increase in NWC / Revenue

# ─── Hold Period ─────────────────────────────────────────────────────────────
HOLD_YEARS = 5     # number of projection years; returns shown for each exit year

# ─── Sensitivity Table Axes ──────────────────────────────────────────────────
# Both tables use terminal-year exit.  Edit min/max/step to resize the grid.
SENS_ENTRY_MIN   = ENTRY_MULTIPLE - 1.0   # 10.0x
SENS_ENTRY_MAX   = ENTRY_MULTIPLE + 1.0   # 12.0x
SENS_ENTRY_STEP  = 0.5

SENS_EXIT_MIN    = EXIT_MULTIPLE - 1.0    # 10.0x
SENS_EXIT_MAX    = EXIT_MULTIPLE + 1.0    # 12.0x
SENS_EXIT_STEP   = 0.5
