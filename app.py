# =============================================================================
# app.py — LBO Model Teacher | Streamlit Web App
# =============================================================================
#
#  Run:  streamlit run app.py
#
#  Layout: two-column — inputs left, live outputs right.
#  Every widget change triggers an automatic recompute (no submit button).
# =============================================================================

from __future__ import annotations

import contextlib
import io
import os
import sys
import tempfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import streamlit as st

# ── Model path ────────────────────────────────────────────────────────────────
ROOT = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(ROOT, "lbo_model"))

import inputs as cfg            # module-level defaults (we mutate via setattr)
import excel_export as xl_mod
import forecast as fc_mod
import returns as ret_mod
import sensitivity as sens_mod
import transaction as txn_mod

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LBO Model Teacher",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# DESIGN TOKENS
# ─────────────────────────────────────────────────────────────────────────────
NAVY   = "#1F4E79"
LBLUE  = "#D6E4F0"
YELLOW = "#FFF176"
WHITE  = "#FFFFFF"
GREY   = "#F7F9FC"

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  /* Page frame */
  .main .block-container {{
    padding-top: 1.1rem;
    padding-bottom: 2rem;
    max-width: 1700px;
  }}

  /* Section headers */
  .sh {{
    background: {NAVY};
    color: {WHITE};
    padding: 5px 12px;
    border-radius: 4px;
    font-size: .70rem;
    font-weight: 700;
    letter-spacing: .07em;
    text-transform: uppercase;
    margin: 12px 0 4px;
  }}

  /* Sub-headers (debt tranches) */
  .subh {{
    background: {LBLUE};
    color: {NAVY};
    padding: 3px 10px;
    border-radius: 3px;
    font-size: .68rem;
    font-weight: 700;
    margin: 10px 0 2px;
    letter-spacing: .04em;
    text-transform: uppercase;
  }}

  /* Tighter widget spacing in the input column */
  div[data-testid="stNumberInput"]   {{ margin-bottom: -4px; }}
  div[data-testid="stNumberInput"] label {{ font-size: .76rem; }}
  div[data-testid="stCheckbox"]  label {{ font-size: .76rem; }}
  div[data-testid="stColumns"]   {{ gap: 0.6rem; }}

  /* Metric cards */
  div[data-testid="metric-container"] {{
    background: {GREY};
    border: 1px solid {LBLUE};
    border-radius: 6px;
    padding: 10px 14px 8px;
  }}
  div[data-testid="metric-container"] [data-testid="stMetricValue"] {{
    font-size: 1.10rem;
    font-weight: 700;
    color: {NAVY};
  }}
  div[data-testid="metric-container"] [data-testid="stMetricLabel"] {{
    font-size: .72rem;
    color: #444;
  }}

  /* Download button */
  div[data-testid="stDownloadButton"] > button {{
    background: {NAVY};
    color: {WHITE};
    border: none;
    width: 100%;
    font-weight: 700;
    font-size: .85rem;
    padding: 10px 0;
    border-radius: 5px;
    margin-top: 6px;
  }}
  div[data-testid="stDownloadButton"] > button:hover {{
    background: #2e6fad;
    color: {WHITE};
  }}

  /* Scenario comparison table */
  .scen-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: .78rem;
    margin-top: 6px;
  }}
  .scen-table th {{
    background: {NAVY};
    color: {WHITE};
    padding: 6px 10px;
    text-align: center;
    font-weight: 700;
    letter-spacing: .04em;
  }}
  .scen-table th.row-hdr {{
    background: {LBLUE};
    color: {NAVY};
    text-align: left;
    width: 38%;
  }}
  .scen-table td {{
    padding: 5px 10px;
    text-align: center;
    border-bottom: 1px solid #e8edf2;
  }}
  .scen-table td.row-hdr {{
    text-align: left;
    color: #333;
    font-weight: 500;
  }}
  .scen-table tr.section-sep td,
  .scen-table tr.section-sep th {{
    background: {LBLUE};
    color: {NAVY};
    font-weight: 700;
    font-size: .70rem;
    letter-spacing: .05em;
    text-transform: uppercase;
    padding: 4px 10px;
  }}
  .scen-table tr:hover td {{ background: #f0f6fb; }}
  .scen-badge {{
    display: inline-block;
    background: {LBLUE};
    color: {NAVY};
    border-radius: 3px;
    padding: 1px 7px;
    font-size: .70rem;
    font-weight: 700;
    margin-left: 6px;
    vertical-align: middle;
  }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def sh(text: str) -> None:
    st.markdown(f'<div class="sh">{text}</div>', unsafe_allow_html=True)

def subh(text: str) -> None:
    st.markdown(f'<div class="subh">{text}</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# INPUT WIDGETS
# ─────────────────────────────────────────────────────────────────────────────

def mm_in(label: str, key: str, default: float, step: float = 10.0) -> float:
    """Dollar-million number input."""
    return st.number_input(label, value=float(default), step=step,
                           format="%.1f", key=key)

def x_in(label: str, key: str, default: float, step: float = 0.5,
         help: str | None = None) -> float:
    """Multiple (x) number input."""
    return st.number_input(label, value=float(default), step=step,
                           format="%.1f", key=key, min_value=0.0, help=help)

def pct_in(label: str, key: str, default_frac: float, step: float = 0.1,
           help: str | None = None) -> float:
    """
    Percentage input displayed as whole number (e.g. 38.00 for 38%).
    Returns a decimal fraction (0.38).
    """
    val = st.number_input(
        label, value=round(default_frac * 100, 4),
        step=step, format="%.2f", key=key, min_value=0.0, help=help,
    )
    return val / 100.0

def int_in(label: str, key: str, default: int,
           mn: int = 1, mx: int = 20) -> int:
    return int(st.number_input(label, value=int(default),
                               min_value=mn, max_value=mx, step=1, key=key))


# ─────────────────────────────────────────────────────────────────────────────
# PAGE HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    f'<h2 style="color:{NAVY};margin:0 0 0px;font-size:1.5rem">'
    "📊&nbsp; LBO Model Teacher</h2>",
    unsafe_allow_html=True,
)
st.caption(
    "Edit any input on the left — the model recalculates instantly.  "
    "Click **⬇ Download Excel** for a fully-linked audit trail with live formulas."
)

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE — scenario storage
# ─────────────────────────────────────────────────────────────────────────────
if "scenarios" not in st.session_state:
    st.session_state.scenarios: list[dict] = []   # max 3 items

MAX_SCENARIOS = 3

# ─────────────────────────────────────────────────────────────────────────────
# TWO-COLUMN LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
left_col, _, right_col = st.columns([1, 0.04, 1.45])

# ══════════════════════════════════════════════════════════════════════════════
# LEFT COLUMN — INPUTS
# ══════════════════════════════════════════════════════════════════════════════
with left_col:

    ov: dict = {}   # overrides dict — populated below, applied before model run

    # ── 1. Entry Valuation ────────────────────────────────────────────────────
    sh("1 · Entry Valuation")
    a, b = st.columns(2)
    with a:
        ov["ENTRY_EBITDA"]  = mm_in("Entry EBITDA ($mm)",   "e_ebitda",   cfg.ENTRY_EBITDA,  25.0)
        ov["EXISTING_DEBT"] = mm_in("Existing Debt ($mm)",  "e_ex_debt",  cfg.EXISTING_DEBT, 25.0)
    with b:
        ov["ENTRY_MULTIPLE"]= x_in( "Entry Multiple (x)",   "e_emult",    cfg.ENTRY_MULTIPLE,
            help=(
                "The price paid for the business expressed as a multiple of its LTM EBITDA — "
                "e.g. 11x means you paid 11× the company's annual EBITDA.\n\n"
                "📊 Typical range: 7x–14x depending on sector and market conditions; "
                "software deals often trade at 12x–20x, industrials at 6x–9x.\n\n"
                "📈 Effect on returns: a lower entry multiple means you pay less for the same "
                "business, directly increasing IRR and MoM — it is one of the single biggest drivers of PE returns."
            ))
        ov["EXISTING_CASH"] = mm_in("Existing Cash ($mm)",  "e_ex_cash",  cfg.EXISTING_CASH, 25.0)

    # ── 2. Transaction ────────────────────────────────────────────────────────
    sh("2 · Transaction")
    a, b = st.columns(2)
    with a:
        ov["EXIT_MULTIPLE"]        = x_in( "Exit Multiple (x)",          "t_xmult",   cfg.EXIT_MULTIPLE,
            help=(
                "The EBITDA multiple at which the PE firm sells the business at exit — "
                "determines the total enterprise value received.\n\n"
                "📊 Typical range: same as entry, 7x–14x; analysts often assume entry = exit "
                "(no multiple expansion) as a conservative base case.\n\n"
                "📈 Effect on returns: every 1x increase in exit multiple adds roughly "
                "EBITDA × (1 − net debt %) to equity value — highly sensitive at low leverage."
            ))
        ov["CASH_TO_BS"]           = mm_in("Cash to B/S ($mm)",          "t_cash_bs", cfg.CASH_TO_BS,           10.0)
        ov["ROLLOVER_EQUITY_PCT"]  = pct_in("Rollover Equity (%)",       "t_rollover",cfg.ROLLOVER_EQUITY_PCT,
            help=(
                "The percentage of total equity that existing management or sellers reinvest "
                "alongside the PE firm rather than cashing out — they 'roll' their equity into the new deal.\n\n"
                "📊 Typical range: 10%–40%; management rollover aligns incentives and "
                "reduces the cash the PE firm needs to write.\n\n"
                "📈 Effect on returns: higher rollover reduces the PE firm's equity check, "
                "which increases IRR if returns are positive — but also reduces the firm's absolute dollar gain."
            ))
    with b:
        ov["TAX_RATE"]             = pct_in("Tax Rate (%)",              "t_tax",     cfg.TAX_RATE)
        ov["TRANSACTION_EXPENSES"] = mm_in("Transaction Expenses ($mm)", "t_txexp",   cfg.TRANSACTION_EXPENSES,  5.0)

    # ── 3. Debt — General ─────────────────────────────────────────────────────
    sh("3 · Debt")
    a, b = st.columns(2)
    with a:
        ov["LEVERAGEABLE_EBITDA"]       = mm_in("Leverageable EBITDA ($mm)", "d_lev_ebitda", cfg.LEVERAGEABLE_EBITDA, 25.0)
        ov["MIN_CASH_BALANCE"]          = mm_in("Min Cash Balance ($mm)",    "d_min_cash",   cfg.MIN_CASH_BALANCE,    10.0)
    with b:
        ov["BASE_RATE"]                 = pct_in("Base Rate / SOFR (%)",     "d_base_rate",  cfg.BASE_RATE)
        ov["FINANCING_FEE_AMORT_YEARS"] = int_in("Fee Amort (years)",        "d_fee_amort",  cfg.FINANCING_FEE_AMORT_YEARS)

    subh("Term Loan")
    a, b = st.columns(2)
    with a:
        ov["TERM_LOAN_LEVERAGE"]   = x_in( "Leverage (x EBITDA)", "tl_lev",    cfg.TERM_LOAN_LEVERAGE,
            help=(
                "The size of the Term Loan expressed as a multiple of EBITDA — e.g. 3x means "
                "the TL principal equals 3× the company's annual EBITDA.\n\n"
                "📊 Typical range: 2x–5x for the Term Loan alone; total leverage (TL + Mezz) "
                "is typically 4x–6x in a healthy LBO market, constrained by lender appetite.\n\n"
                "📈 Effect on returns: more debt means a smaller equity check and higher IRR "
                "if the deal works — but also higher interest burden and greater downside risk."
            ))
        ov["TERM_LOAN_FEE_PCT"]    = pct_in("Upfront Fee (%)",    "tl_fee",    cfg.TERM_LOAN_FEE_PCT)
    with b:
        ov["TERM_LOAN_SPREAD"]     = pct_in("Spread L+ (%)",      "tl_spread", cfg.TERM_LOAN_SPREAD)
        ov["TERM_LOAN_AMORT_RATE"] = pct_in("Mandatory Amort (%)", "tl_amort", cfg.TERM_LOAN_AMORT_RATE)

    subh("Mezzanine")
    a, b = st.columns(2)
    with a:
        ov["MEZZ_LEVERAGE"]   = x_in( "Leverage (x EBITDA)", "mz_lev",  cfg.MEZZ_LEVERAGE)
        ov["MEZZ_PIK_RATE"]   = pct_in("PIK Rate (%)",        "mz_pik",  cfg.MEZZ_PIK_RATE,
            help=(
                "PIK stands for 'Payment in Kind' — instead of paying cash interest, "
                "the interest accrues and is added to the loan balance each year, compounding over time.\n\n"
                "📊 Typical range: 4%–8% for mezzanine PIK; total mezz return (cash + PIK) "
                "is usually 10%–15%, sitting between senior debt and equity in the capital structure.\n\n"
                "📈 Effect on returns: PIK reduces cash interest outflows (helping FCF and debt "
                "paydown) but grows the mezz balance at exit, increasing net debt and reducing equity value."
            ))
        ov["MEZZ_FEE_PCT"]    = pct_in("Upfront Fee (%)",     "mz_fee",  cfg.MEZZ_FEE_PCT)
    with b:
        ov["MEZZ_CASH_RATE"]  = pct_in("Cash Pay Rate (%)",   "mz_cash", cfg.MEZZ_CASH_RATE)
        ov["MEZZ_CASH_SWEEP"] = st.checkbox(
            "Sweep Mezz After TL Repaid?",
            value=bool(cfg.MEZZ_CASH_SWEEP), key="mz_sweep"
        )

    subh("Revolver")
    a, b = st.columns(2)
    with a:
        ov["REVOLVER_CAPACITY"]       = mm_in("Capacity ($mm)",      "rv_cap",     cfg.REVOLVER_CAPACITY,      25.0)
        ov["REVOLVER_SPREAD"]         = pct_in("Spread L+ (%)",      "rv_spread",  cfg.REVOLVER_SPREAD)
    with b:
        ov["REVOLVER_COMMITMENT_FEE"] = pct_in("Commitment Fee (%)","rv_commfee", cfg.REVOLVER_COMMITMENT_FEE)

    # ── 4. Financial Forecast ─────────────────────────────────────────────────
    sh("4 · Financial Forecast")
    a, b = st.columns(2)
    with a:
        ov["BASE_REVENUE"]        = mm_in("Base Revenue ($mm)",     "f_rev",    cfg.BASE_REVENUE,       100.0)
        ov["GROSS_MARGIN"]        = pct_in("Gross Margin (%)",      "f_gm",     cfg.GROSS_MARGIN)
        ov["RD_PCT"]              = pct_in("R&D (% Rev)",           "f_rd",     cfg.RD_PCT)
        ov["DA_PCT"]              = pct_in("D&A (% Rev)",           "f_da",     cfg.DA_PCT)
        ov["NWC_PCT"]             = pct_in("Incr NWC (% Rev)",      "f_nwc",    cfg.NWC_PCT)
    with b:
        ov["REVENUE_GROWTH_RATE"] = pct_in("Revenue Growth (%)",    "f_growth", cfg.REVENUE_GROWTH_RATE)
        ov["SGA_PCT"]             = pct_in("SG&A (% Rev)",          "f_sga",    cfg.SGA_PCT)
        ov["CAPEX_PCT"]           = pct_in("Capex (% Rev)",         "f_capex",  cfg.CAPEX_PCT)

    # ── 5. Hold Period ────────────────────────────────────────────────────────
    sh("5 · Hold Period")
    ov["HOLD_YEARS"] = int_in("Hold Period (years)", "h_years", cfg.HOLD_YEARS, mn=1, mx=10)

    # ── Scenario Manager ──────────────────────────────────────────────────────
    sh("6 · Scenario Manager")
    scen_name = st.text_input(
        "Scenario name",
        value="Base Case",
        placeholder="e.g. Base Case, Bull Case, Bear Case",
        key="scen_name_input",
    )
    sc1, sc2 = st.columns(2)
    save_clicked  = sc1.button("💾  Save Scenario",  use_container_width=True)
    clear_clicked = sc2.button("🗑  Clear All",       use_container_width=True)

    saved_count = len(st.session_state.scenarios)
    if saved_count:
        names = [s["name"] for s in st.session_state.scenarios]
        st.caption(f"Saved: **{' · '.join(names)}**")
    else:
        st.caption("No scenarios saved yet.")


# ─────────────────────────────────────────────────────────────────────────────
# APPLY OVERRIDES & RUN MODEL
# Every Streamlit rerun re-applies the current widget values.
# ─────────────────────────────────────────────────────────────────────────────
for k, v in ov.items():
    setattr(cfg, k, v)

# Auto-derive sensitivity axes (±1.0x around base, step 0.5x)
cfg.SENS_ENTRY_MIN  = cfg.ENTRY_MULTIPLE - 1.0
cfg.SENS_ENTRY_MAX  = cfg.ENTRY_MULTIPLE + 1.0
cfg.SENS_ENTRY_STEP = 0.5
cfg.SENS_EXIT_MIN   = cfg.EXIT_MULTIPLE - 1.0
cfg.SENS_EXIT_MAX   = cfg.EXIT_MULTIPLE + 1.0
cfg.SENS_EXIT_STEP  = 0.5

model_err: str | None = None
txn = results = rets = sens = None

try:
    txn     = txn_mod.compute_transaction()
    results = fc_mod.run_model(txn)
    rets    = ret_mod.compute_returns(txn, results)
    sens    = sens_mod.compute_sensitivity(txn, results)
except Exception as exc:
    model_err = str(exc)

# ── Handle scenario save / clear ─────────────────────────────────────────────
if clear_clicked:
    st.session_state.scenarios = []
    st.rerun()

if save_clicked and not model_err:
    name = (scen_name or "Scenario").strip()
    last_r = rets[-1]
    N_saved = ov["HOLD_YEARS"]
    record = {
        "name":           name,
        "ov":             dict(ov),          # snapshot of all inputs
        # key inputs (displayed in comparison)
        "entry_multiple": ov["ENTRY_MULTIPLE"],
        "exit_multiple":  ov["EXIT_MULTIPLE"],
        "revenue_growth": ov["REVENUE_GROWTH_RATE"],
        "tl_leverage":    ov["TERM_LOAN_LEVERAGE"],
        "mezz_leverage":  ov["MEZZ_LEVERAGE"],
        "hold_years":     N_saved,
        # key outputs (terminal year)
        "irr":            last_r["irr"],
        "mom":            last_r["mom"],
        "exit_ev":        last_r["exit_ev"],
        "equity_value":   last_r["equity_value"],
        "value_to_pe":    last_r["value_to_pe"],
        "pe_invest":      txn["pe_equity_invest"],
        "total_debt":     txn["total_debt"],
        "purchase_ev":    txn["purchase_ev"],
    }
    # Replace if same name already exists, else append (max 3)
    existing = [i for i, s in enumerate(st.session_state.scenarios) if s["name"] == name]
    if existing:
        st.session_state.scenarios[existing[0]] = record
    elif len(st.session_state.scenarios) < MAX_SCENARIOS:
        st.session_state.scenarios.append(record)
    else:
        # Drop oldest, add new
        st.session_state.scenarios.pop(0)
        st.session_state.scenarios.append(record)
    st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BYTES — cached so it only regenerates when inputs actually change
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Generating Excel…")
def _make_excel_bytes(_txn, _results, _rets, _sens, _ov_frozen) -> bytes:
    _ov = dict(_ov_frozen)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        tmp = f.name
    try:
        with contextlib.redirect_stdout(io.StringIO()):   # suppress print()
            xl_mod.write_excel(
                _txn, _results, _rets, _sens,
                overrides=_ov,
                output_path=tmp,
            )
        with open(tmp, "rb") as fh:
            return fh.read()
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


# ══════════════════════════════════════════════════════════════════════════════
# RIGHT COLUMN — OUTPUTS
# ══════════════════════════════════════════════════════════════════════════════
with right_col:

    # ── Error guard ───────────────────────────────────────────────────────────
    if model_err:
        st.error(f"⚠️ Model error — {model_err}")
        st.stop()

    N    = cfg.HOLD_YEARS
    last = rets[-1]   # terminal-year return dict

    # ── Input Validation ──────────────────────────────────────────────────────
    _warns   = []   # (level, message)   level: "red" | "yellow"

    # (1) Entry leverage > 6.0x
    entry_lev = txn["total_leverage_x"]
    if entry_lev > 6.0:
        _warns.append(("yellow",
            f"High leverage — total entry leverage is **{entry_lev:.1f}x EBITDA** "
            f"(threshold: 6.0x). Lenders may require tighter covenants or pricing."))

    # (2) Interest coverage < 1.5x in any projected year
    _icr_fails = []
    for _y in range(1, N + 1):
        _ebitda_y = results["ebitda"][_y]
        _cash_int  = (results["rev_interest"][_y]
                      + results["tl_interest"][_y]
                      + results["mezz_cash_interest"][_y])
        if _cash_int > 0:
            _icr = _ebitda_y / _cash_int
            if _icr < 1.5:
                _icr_fails.append((_y, _icr))
    if _icr_fails:
        _yr_strs = ", ".join(f"Year {y} ({icr:.2f}x)" for y, icr in _icr_fails)
        _warns.append(("yellow",
            f"Low interest coverage — ICR falls below 1.5x in: **{_yr_strs}**. "
            f"Risk of covenant breach or cash flow stress."))

    # (3) Negative IRR at terminal hold year
    if last["irr"] < 0:
        _warns.append(("red",
            f"Negative IRR — the model produces an IRR of **{last['irr']:.1%}** "
            f"at Year {N} exit. Check entry valuation, debt sizing, and exit assumptions."))

    # (4) Exit multiple > entry multiple + 3.0x
    _mult_delta = cfg.EXIT_MULTIPLE - cfg.ENTRY_MULTIPLE
    if _mult_delta > 3.0:
        _warns.append(("yellow",
            f"Aggressive exit assumption — exit multiple is **{_mult_delta:.1f}x higher** "
            f"than entry ({cfg.ENTRY_MULTIPLE:.1f}x → {cfg.EXIT_MULTIPLE:.1f}x). "
            f"Multiple expansion of this magnitude is rarely achievable."))

    if _warns:
        for _level, _msg in _warns:
            if _level == "red":
                st.markdown(
                    f'<div style="background:#FFEBEE;border-left:4px solid #C62828;'
                    f'padding:10px 14px;border-radius:4px;margin-bottom:6px;font-size:.82rem;">'
                    f'🔴&nbsp; {_msg}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div style="background:#FFFDE7;border-left:4px solid #F9A825;'
                    f'padding:10px 14px;border-radius:4px;margin-bottom:6px;font-size:.82rem;">'
                    f'🟡&nbsp; {_msg}</div>',
                    unsafe_allow_html=True,
                )
        st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

    # ── Hero metrics (terminal year) ──────────────────────────────────────────
    sh(f"Key Returns — Year {N} Exit")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"IRR (Year {N})",       f"{last['irr']:.1%}")
    c2.metric(f"MoM (Year {N})",       f"{last['mom']:.2f}x")
    c3.metric("Exit EV",               f"${last['exit_ev']:,.0f}mm")
    c4.metric("Value to PE",           f"${last['value_to_pe']:,.0f}mm")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Returns by exit year ──────────────────────────────────────────────────
    sh("Returns by Exit Year")

    year_labels = [f"Year {r['exit_year']}" for r in rets]
    ret_rows = {
        "IRR":               [f"{r['irr']:.1%}"         for r in rets],
        "MoM":               [f"{r['mom']:.2f}x"        for r in rets],
        "Exit EV ($mm)":     [f"${r['exit_ev']:,.0f}"   for r in rets],
        "Net Debt ($mm)":    [f"${r['net_debt']:,.0f}"  for r in rets],
        "Equity Value ($mm)":[f"${r['equity_value']:,.0f}" for r in rets],
        "Value to PE ($mm)": [f"${r['value_to_pe']:,.0f}"  for r in rets],
    }
    df_ret = pd.DataFrame(ret_rows, index=year_labels).T
    df_ret.index.name = ""
    st.dataframe(df_ret, use_container_width=True, height=246)

    # ── Value Creation Bridge ─────────────────────────────────────────────────
    sh("Value Creation Bridge")

    exit_yr_sel = st.selectbox(
        "Exit year",
        options=list(range(1, N + 1)),
        index=N - 1,
        format_func=lambda y: f"Year {y}",
        key="bridge_exit_yr",
    )

    # ── Bridge mechanics ──────────────────────────────────────────────────────
    # Entry equity value (post-close: EV minus net debt at close)
    entry_net_debt  = txn["total_debt"] - cfg.CASH_TO_BS
    entry_eq_val    = txn["purchase_ev"] - entry_net_debt

    # Exit equity value for selected year
    y = exit_yr_sel
    exit_ebitda = results["ebitda"][y]
    exit_total_debt = results["tl_bal"][y] + results["mezz_bal"][y] + results["rev_bal"][y]
    exit_net_debt   = exit_total_debt - results["cash"][y]
    exit_ev_sel     = exit_ebitda * cfg.EXIT_MULTIPLE
    exit_eq_val     = exit_ev_sel - exit_net_debt

    # Decomposition — the three components
    ebitda_growth   = (exit_ebitda - cfg.ENTRY_EBITDA) * cfg.ENTRY_MULTIPLE
    mult_expansion  = exit_ebitda * (cfg.EXIT_MULTIPLE - cfg.ENTRY_MULTIPLE)
    debt_paydown    = entry_net_debt - exit_net_debt

    total_created   = exit_eq_val - entry_eq_val    # sum of three components
    # Guard against division by zero
    pct_of = lambda v: (v / total_created * 100) if abs(total_created) > 1e-6 else 0.0

    # ── Waterfall chart ───────────────────────────────────────────────────────
    C_NAVY  = "#1F4E79"
    C_GREEN = "#2E7D32"
    C_RED   = "#C62828"
    C_TEAL  = "#00695C"
    C_AMBER = "#E65100"
    C_LGREY = "#ECEFF1"

    labels     = ["Entry\nEquity", "EBITDA\nGrowth", "Multiple\nExpansion", "Debt\nPaydown", "Exit\nEquity"]
    components = [ebitda_growth, mult_expansion, debt_paydown]
    values     = [entry_eq_val] + components + [exit_eq_val]

    # Running bottom for floating bars
    bottoms = [0.0, entry_eq_val, 0.0, 0.0, 0.0]
    bottoms[2] = entry_eq_val + ebitda_growth
    bottoms[3] = bottoms[2] + mult_expansion

    bar_colors = [
        C_NAVY,
        C_GREEN if ebitda_growth  >= 0 else C_RED,
        C_GREEN if mult_expansion >= 0 else C_RED,
        C_TEAL  if debt_paydown   >= 0 else C_AMBER,
        C_NAVY,
    ]

    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    bar_width = 0.52
    x = np.arange(len(labels))

    bars = ax.bar(x, values, bottom=bottoms, width=bar_width,
                  color=bar_colors, zorder=3, linewidth=0)

    # Connector lines between bars
    connector_pairs = [(0, 1), (1, 2), (2, 3), (3, 4)]
    tops = [b + v for b, v in zip(bottoms, values)]
    for left_i, right_i in connector_pairs:
        top_y = tops[left_i]
        ax.plot(
            [left_i + bar_width / 2, right_i - bar_width / 2],
            [top_y, top_y],
            color="#9E9E9E", lw=0.8, ls="--", zorder=4,
        )

    # Bar labels: $value + % of total (for middle three only)
    for i, (val, bot) in enumerate(zip(values, bottoms)):
        bar_top  = bot + val
        mid_y    = bot + val / 2

        if i == 0 or i == 4:
            # Solid bars — label at top, outside
            label_y  = bar_top + abs(exit_eq_val - entry_eq_val) * 0.03
            va       = "bottom"
            label_txt = f"${val:,.0f}mm"
            ax.text(x[i], label_y, label_txt, ha="center", va=va,
                    fontsize=8, fontweight="bold", color=C_NAVY, zorder=5)
        else:
            # Component bars — label inside if tall enough, else above
            comp     = components[i - 1]
            sign     = "+" if comp >= 0 else ""
            pct_txt  = f"{pct_of(comp):.0f}%"
            dollar   = f"{sign}${comp:,.0f}mm"
            full_lbl = f"{dollar}\n({pct_txt})"

            bar_height_pts = abs(val) / (ax.get_ylim()[1] - ax.get_ylim()[0] + 1e-9) * fig.get_size_inches()[1] * fig.dpi
            if abs(val) > abs(total_created) * 0.08:
                ax.text(x[i], mid_y, full_lbl, ha="center", va="center",
                        fontsize=7.5, fontweight="bold", color="white", zorder=5)
            else:
                offset  = abs(total_created) * 0.04
                label_y = bar_top + offset if comp >= 0 else bot - offset
                va2     = "bottom" if comp >= 0 else "top"
                ax.text(x[i], label_y, full_lbl, ha="center", va=va2,
                        fontsize=7.5, fontweight="bold",
                        color=bar_colors[i], zorder=5)

    # Axes formatting
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=8.5, color="#333333")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda v, _: f"${v:,.0f}mm"
    ))
    ax.tick_params(axis="y", labelsize=7.5, color="#BDBDBD")
    ax.tick_params(axis="x", length=0)
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#E0E0E0")
    ax.yaxis.grid(True, color="#F0F0F0", linewidth=0.7, zorder=0)
    ax.set_axisbelow(True)

    # Y-axis range with padding
    all_ys = [b + v for b, v in zip(bottoms, values)] + list(bottoms)
    y_min, y_max = min(all_ys), max(all_ys)
    pad = (y_max - y_min) * 0.18
    ax.set_ylim(min(0, y_min - pad), y_max + pad)

    # Zero line
    ax.axhline(0, color="#9E9E9E", linewidth=0.8, zorder=2)

    ax.set_title(
        f"Value Creation Bridge — Year {y} Exit  |  "
        f"IRR {rets[y-1]['irr']:.1%}  ·  MoM {rets[y-1]['mom']:.2f}x",
        fontsize=9, color=C_NAVY, fontweight="bold", pad=10,
    )

    fig.tight_layout(pad=1.0)
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    # ── Bridge summary row ────────────────────────────────────────────────────
    bc1, bc2, bc3, bc4 = st.columns(4)
    bc1.metric("EBITDA Growth",       f"${ebitda_growth:+,.0f}mm",  f"{pct_of(ebitda_growth):.0f}% of total")
    bc2.metric("Multiple Δ",          f"${mult_expansion:+,.0f}mm", f"{pct_of(mult_expansion):.0f}% of total")
    bc3.metric("Debt Paydown",        f"${debt_paydown:+,.0f}mm",   f"{pct_of(debt_paydown):.0f}% of total")
    bc4.metric("Total Value Created", f"${total_created:+,.0f}mm",  "100%")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Debt & Cash Balance Over Hold Period ─────────────────────────────────
    sh("Debt & Cash Balance Over Hold Period")

    years      = list(range(N + 1))
    rev_bals   = [results["rev_bal"][y]  for y in years]
    tl_bals    = [results["tl_bal"][y]   for y in years]
    mezz_bals  = [results["mezz_bal"][y] for y in years]
    cash_bals  = [results["cash"][y]     for y in years]
    x_pos      = np.arange(len(years))
    xlabels    = ["Close"] + [f"Year {y}" for y in range(1, N + 1)]

    C_TL   = "#1F4E79"   # navy  — Term Loan
    C_MEZZ = "#2E7D32"   # green — Mezzanine
    C_REV  = "#E65100"   # amber — Revolver
    C_CASH = "#FDD835"   # yellow — Cash line

    fig2, ax2 = plt.subplots(figsize=(7.2, 3.8))
    fig2.patch.set_facecolor("white")
    ax2.set_facecolor("white")

    bw = 0.55
    b_tl   = ax2.bar(x_pos, tl_bals,   width=bw, label="Term Loan",  color=C_TL,   zorder=3)
    b_mezz = ax2.bar(x_pos, mezz_bals,  width=bw, label="Mezzanine",  color=C_MEZZ, zorder=3,
                     bottom=tl_bals)
    b_rev  = ax2.bar(x_pos, rev_bals,   width=bw, label="Revolver",   color=C_REV,  zorder=3,
                     bottom=[t + m for t, m in zip(tl_bals, mezz_bals)])

    # Total debt labels at top of each stack
    for xi, (t, m, r) in enumerate(zip(tl_bals, mezz_bals, rev_bals)):
        total = t + m + r
        ax2.text(xi, total + max(cash_bals + [t + m + r for t, m, r in zip(tl_bals, mezz_bals, rev_bals)]) * 0.02,
                 f"${total:,.0f}", ha="center", va="bottom",
                 fontsize=7, color="#333333", fontweight="600", zorder=5)

    # Cash line on same axis
    ax2_r = ax2.twinx()
    ax2_r.plot(x_pos, cash_bals, color=C_CASH, linewidth=2.2,
               marker="o", markersize=5, markerfacecolor=C_CASH,
               markeredgecolor="white", markeredgewidth=1.2,
               label="Cash", zorder=6)
    for xi, c in enumerate(cash_bals):
        ax2_r.text(xi, c + max(cash_bals) * 0.07, f"${c:,.0f}",
                   ha="center", va="bottom", fontsize=7,
                   color="#7B6F00", fontweight="600", zorder=7)

    # Axes formatting — left (debt)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(xlabels, fontsize=8)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}mm"))
    ax2.tick_params(axis="y", labelsize=7.5)
    ax2.tick_params(axis="x", length=0)
    ax2.set_ylabel("Debt Balance ($mm)", fontsize=8, color="#333333")
    for spine in ["top", "right"]:
        ax2.spines[spine].set_visible(False)
    ax2.spines["bottom"].set_color("#E0E0E0")
    ax2.spines["left"].set_color("#E0E0E0")
    ax2.yaxis.grid(True, color="#F0F0F0", linewidth=0.7, zorder=0)
    ax2.set_axisbelow(True)

    # Right axis (cash)
    ax2_r.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:,.0f}mm"))
    ax2_r.tick_params(axis="y", labelsize=7.5, colors="#7B6F00")
    ax2_r.set_ylabel("Cash Balance ($mm)", fontsize=8, color="#7B6F00")
    ax2_r.spines["top"].set_visible(False)
    ax2_r.spines["left"].set_visible(False)
    ax2_r.spines["right"].set_color("#E0E0E0")

    # Combined legend
    handles_l, labels_l = ax2.get_legend_handles_labels()
    handles_r, labels_r = ax2_r.get_legend_handles_labels()
    ax2.legend(handles_l + handles_r, labels_l + labels_r,
               loc="upper right", fontsize=7.5, framealpha=0.9,
               edgecolor="#DDDDDD", ncol=2)

    ax2.set_title("Debt & Cash Balance Over Hold Period",
                  fontsize=9, fontweight="bold", color=C_NAVY, pad=10)
    fig2.tight_layout(pad=1.0)
    st.pyplot(fig2, use_container_width=True)
    plt.close(fig2)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Transaction snapshot ──────────────────────────────────────────────────
    sh("Transaction Snapshot")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Purchase EV",  f"${txn['purchase_ev']:,.0f}mm")
    c2.metric("Total Debt",   f"${txn['total_debt']:,.0f}mm")
    c3.metric("PE Equity",    f"${txn['pe_equity_invest']:,.0f}mm")
    c4.metric("Entry Leverage", f"{txn['total_leverage_x']:.1f}x EBITDA")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Sensitivity helpers ───────────────────────────────────────────────────
    entry_arr = sens["entry_multiples"]
    exit_arr  = sens["exit_multiples"]
    entry_labs = [f"{x:.1f}x" for x in entry_arr]
    exit_labs  = [f"{x:.1f}x" for x in exit_arr]

    # Find base-case cell indices for yellow highlight
    base_e = int(np.argmin(np.abs(entry_arr - cfg.ENTRY_MULTIPLE)))
    base_x = int(np.argmin(np.abs(exit_arr  - cfg.EXIT_MULTIPLE)))

    def _highlight_base(df: pd.DataFrame) -> pd.DataFrame:
        """Return same-shape DataFrame of CSS strings; yellow on base cell."""
        s = pd.DataFrame("", index=df.index, columns=df.columns)
        s.iloc[base_e, base_x] = (
            f"background-color:{YELLOW}; font-weight:bold; color:{NAVY}"
        )
        return s

    # ── IRR Sensitivity ───────────────────────────────────────────────────────
    sh("IRR Sensitivity — Entry Multiple (rows) vs Exit Multiple (cols)")

    irr_min = float(np.nanmin(sens["irr_table"]))
    irr_max = float(np.nanmax(sens["irr_table"]))

    df_irr = pd.DataFrame(
        sens["irr_table"], index=entry_labs, columns=exit_labs
    )
    df_irr.index.name = "Entry \\ Exit"

    st.dataframe(
        df_irr.style
              .background_gradient(cmap="RdYlGn", vmin=irr_min, vmax=irr_max)
              .format("{:.1%}")
              .apply(_highlight_base, axis=None),
        use_container_width=True,
        height=len(entry_labs) * 35 + 40,
    )

    # ── MoM Sensitivity ───────────────────────────────────────────────────────
    sh("MoM Sensitivity — Entry Multiple (rows) vs Exit Multiple (cols)")

    mom_min = float(np.nanmin(sens["mom_table"]))
    mom_max = float(np.nanmax(sens["mom_table"]))

    df_mom = pd.DataFrame(
        sens["mom_table"], index=entry_labs, columns=exit_labs
    )
    df_mom.index.name = "Entry \\ Exit"

    st.dataframe(
        df_mom.style
              .background_gradient(cmap="RdYlGn", vmin=mom_min, vmax=mom_max)
              .format("{:.2f}x")
              .apply(_highlight_base, axis=None),
        use_container_width=True,
        height=len(entry_labs) * 35 + 40,
    )

    # ── Download Excel ────────────────────────────────────────────────────────
    st.divider()

    try:
        excel_bytes = _make_excel_bytes(
            txn, results, rets, sens,
            tuple(sorted(ov.items())),   # hashable for cache key
        )
        st.download_button(
            label="⬇  Download Excel Model",
            data=excel_bytes,
            file_name="lbo_model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as _xl_err:
        st.error(f"Excel generation failed: {_xl_err}")

    st.caption(
        "The downloaded file has an **Inputs** sheet (blue editable cells) "
        "and a **Model** sheet with live Excel formulas — every cell recalculates "
        "when you change an input."
    )

    # ── Scenario Comparison ───────────────────────────────────────────────────
    scenarios = st.session_state.scenarios
    if scenarios:
        st.markdown("<br>", unsafe_allow_html=True)
        sh(f"Scenario Comparison  ({len(scenarios)}/{MAX_SCENARIOS} saved)")

        # Build column headers
        scen_cols = [s["name"] for s in scenarios]
        n_scen    = len(scen_cols)

        def _td(val: str, cls: str = "") -> str:
            return f'<td class="{cls}">{val}</td>'

        def _th(val: str, cls: str = "") -> str:
            return f'<th class="{cls}">{val}</th>'

        def _sep(label: str) -> str:
            span = n_scen + 1
            return (
                f'<tr class="section-sep">'
                f'<td colspan="{span}">{label}</td>'
                f'</tr>'
            )

        def _row(label: str, vals: list[str]) -> str:
            cells = _td(label, "row-hdr") + "".join(_td(v) for v in vals)
            return f"<tr>{cells}</tr>"

        # ── Header row
        header = (
            "<thead><tr>"
            + _th("", "row-hdr")
            + "".join(_th(n) for n in scen_cols)
            + "</tr></thead>"
        )

        # ── Body rows
        rows = []

        rows.append(_sep("INPUTS"))
        rows.append(_row("Entry Multiple",
                         [f"{s['entry_multiple']:.1f}x" for s in scenarios]))
        rows.append(_row("Exit Multiple",
                         [f"{s['exit_multiple']:.1f}x"  for s in scenarios]))
        rows.append(_row("Revenue Growth",
                         [f"{s['revenue_growth']:.1%}"  for s in scenarios]))
        rows.append(_row("TL Leverage",
                         [f"{s['tl_leverage']:.1f}x"    for s in scenarios]))
        rows.append(_row("Mezz Leverage",
                         [f"{s['mezz_leverage']:.1f}x"  for s in scenarios]))
        rows.append(_row("Hold Period",
                         [f"{s['hold_years']}yr"        for s in scenarios]))

        rows.append(_sep("TRANSACTION"))
        rows.append(_row("Purchase EV ($mm)",
                         [f"${s['purchase_ev']:,.0f}"   for s in scenarios]))
        rows.append(_row("Total Debt ($mm)",
                         [f"${s['total_debt']:,.0f}"    for s in scenarios]))
        rows.append(_row("PE Equity Invest ($mm)",
                         [f"${s['pe_invest']:,.0f}"     for s in scenarios]))

        rows.append(_sep("RETURNS  (terminal year exit)"))
        rows.append(_row("Exit EV ($mm)",
                         [f"${s['exit_ev']:,.0f}"       for s in scenarios]))
        rows.append(_row("Equity Value ($mm)",
                         [f"${s['equity_value']:,.0f}"  for s in scenarios]))
        rows.append(_row("Value to PE ($mm)",
                         [f"${s['value_to_pe']:,.0f}"   for s in scenarios]))
        rows.append(_row("IRR",
                         [f"{s['irr']:.1%}"             for s in scenarios]))
        rows.append(_row("MoM",
                         [f"{s['mom']:.2f}x"            for s in scenarios]))

        body = "<tbody>" + "".join(rows) + "</tbody>"
        table_html = (
            f'<table class="scen-table">{header}{body}</table>'
        )
        st.markdown(table_html, unsafe_allow_html=True)
        st.caption(
            "Save up to 3 named scenarios using the **Scenario Manager** "
            "in the inputs panel. Saving a name that already exists overwrites it."
        )

    # ── Reverse Solver ────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    sh("Reverse Solver — Back-Solve Exit Multiple for Target IRR")
    st.caption(
        "Given all current inputs, find the exit multiple required to hit a target IRR."
    )

    rs_c1, rs_c2 = st.columns(2)
    with rs_c1:
        rs_target_irr = st.number_input(
            "Target IRR (%)",
            value=20.0, min_value=0.1, max_value=100.0,
            step=0.5, format="%.1f", key="rs_target_irr",
        ) / 100.0
    with rs_c2:
        rs_exit_yr = st.selectbox(
            "Exit year",
            options=list(range(1, N + 1)),
            index=N - 1,
            format_func=lambda y: f"Year {y}",
            key="rs_exit_yr",
        )

    # ── Solver logic ──────────────────────────────────────────────────────────
    try:
        from scipy.optimize import brentq as _brentq

        _pe_invest = txn["pe_equity_invest"]
        _pe_own    = txn["pe_ownership"]
        _y         = rs_exit_yr

        # Terminal values fixed regardless of exit multiple
        _ebitda_y      = results["ebitda"][_y]
        _total_debt_y  = results["tl_bal"][_y] + results["mezz_bal"][_y] + results["rev_bal"][_y]
        _net_debt_y    = _total_debt_y - results["cash"][_y]

        def _irr_gap(xm: float) -> float:
            """IRR(xm) − target_irr.  Use CAGR formula — exact for single-entry/exit."""
            _exit_eq   = xm * _ebitda_y - _net_debt_y
            _v2pe      = _exit_eq * _pe_own
            if _pe_invest <= 0 or _v2pe <= 0:
                return -rs_target_irr   # force bracket to look higher
            _mom = _v2pe / _pe_invest
            return _mom ** (1.0 / _y) - 1.0 - rs_target_irr

        # Bracket search: try multiples from 0.1x to 50x
        _lo, _hi = 0.01, 50.0
        _f_lo = _irr_gap(_lo)
        _f_hi = _irr_gap(_hi)

        if _f_lo * _f_hi >= 0:
            # No sign change — target IRR is outside achievable range
            rs_solved_xm = None
            rs_err_msg   = (
                "Target IRR is not achievable with exit multiple in (0.01×, 50×). "
                "Try a lower target IRR or adjust other inputs."
            )
        else:
            rs_solved_xm = _brentq(_irr_gap, _lo, _hi, xtol=1e-6)
            rs_err_msg   = None

    except ImportError:
        rs_solved_xm = None
        rs_err_msg   = "scipy is not installed. Run: pip install scipy"
    except Exception as _exc:
        rs_solved_xm = None
        rs_err_msg   = str(_exc)

    # ── Display result ────────────────────────────────────────────────────────
    if rs_err_msg:
        st.warning(f"⚠️ {rs_err_msg}")
    else:
        _delta   = rs_solved_xm - cfg.EXIT_MULTIPLE
        _rel_lbl = (
            f"**{abs(_delta):.2f}x above** your base case exit multiple ({cfg.EXIT_MULTIPLE:.1f}x)"
            if _delta > 0 else
            f"**{abs(_delta):.2f}x below** your base case exit multiple ({cfg.EXIT_MULTIPLE:.1f}x)"
            if _delta < 0 else
            f"**equal to** your base case exit multiple ({cfg.EXIT_MULTIPLE:.1f}x)"
        )
        _color = "#E8F5E9" if _delta <= 0 else "#FFF8E1"
        _border = "#2E7D32" if _delta <= 0 else "#F9A825"
        _icon  = "✅" if _delta <= 0 else "⚠️"

        st.markdown(
            f'<div style="background:{_color};border-left:4px solid {_border};'
            f'padding:14px 18px;border-radius:5px;font-size:.88rem;margin-top:6px;">'
            f'{_icon}&nbsp; To achieve a <strong>{rs_target_irr:.0%} IRR</strong> '
            f'in <strong>Year {rs_exit_yr}</strong>, you need an exit multiple of '
            f'<strong style="font-size:1.1rem">{rs_solved_xm:.2f}x</strong> — {_rel_lbl}.'
            f'</div>',
            unsafe_allow_html=True,
        )

        # Metric breakdown
        _exit_ev_req  = rs_solved_xm * _ebitda_y
        _eq_val_req   = _exit_ev_req - _net_debt_y
        _v2pe_req     = _eq_val_req * _pe_own
        _mom_req      = _v2pe_req / _pe_invest if _pe_invest > 0 else float("nan")

        st.markdown("<br>", unsafe_allow_html=True)
        rm1, rm2, rm3, rm4 = st.columns(4)
        rm1.metric("Required Exit Multiple", f"{rs_solved_xm:.2f}x",
                   delta=f"{_delta:+.2f}x vs base case",
                   delta_color="inverse" if _delta > 0 else "normal")
        rm2.metric("Implied Exit EV",        f"${_exit_ev_req:,.0f}mm")
        rm3.metric("Implied Equity Value",   f"${_eq_val_req:,.0f}mm")
        rm4.metric("Implied MoM",            f"{_mom_req:.2f}x")
