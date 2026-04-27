# Peak Frameworks — LBO Model

An interactive leveraged buyout (LBO) model built with Python and Streamlit. The app presents a two-column layout: the left column exposes all deal assumptions (entry/exit valuation, debt sizing across Term Loan, Mezzanine, and Revolver tranches, financial forecast drivers, and hold period) as live number inputs pre-filled with sensible defaults; the right column instantly recalculates and displays key outputs including IRR and MoM for every exit year, a value creation bridge chart decomposing returns into EBITDA growth, multiple expansion, and debt paydown, a stacked debt & cash balance chart over the hold period, IRR and MoM sensitivity tables, input validation warnings (high leverage, thin interest coverage, aggressive exit assumptions), a reverse solver that back-solves the exit multiple required to hit a target IRR, and a scenario manager that saves up to three named scenarios side by side for comparison. A Download Excel button generates a fully-linked two-sheet workbook (blue editable Inputs sheet + formula-driven Model sheet) on demand.

## Running locally

```bash
# 1. Clone or download the repository
git clone <repo-url>
cd lbo-agent

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run app.py
```

The app will open at `http://localhost:8501`. No configuration is required — all defaults are loaded from `lbo_model/inputs.py`.

## Deploying to Streamlit Community Cloud

1. Push the repository to GitHub (ensure `app.py`, `requirements.txt`, and the `lbo_model/` directory are all committed).
2. Go to [share.streamlit.io](https://share.streamlit.io) and click **New app**.
3. Select your repository, set the branch, and set the main file path to `app.py`.
4. Click **Deploy** — Streamlit Cloud installs dependencies from `requirements.txt` automatically.

## Project structure

```
lbo-agent/
├── app.py                  # Streamlit web app
├── requirements.txt        # Python dependencies
├── README.md
└── lbo_model/
    ├── inputs.py           # All default assumptions (edit for scenario analysis)
    ├── transaction.py      # Entry valuation & Sources and Uses
    ├── forecast.py         # Year-by-year income statement, FCF, debt schedule
    ├── debt_schedule.py    # Debt sweep waterfall (Revolver → TL → Mezz)
    ├── returns.py          # IRR and MoM for each exit year
    ├── sensitivity.py      # 5x5 IRR and MoM sensitivity tables
    ├── excel_export.py     # Two-sheet Excel export with live formulas
    ├── cli.py              # Interactive CLI (alternative to the web app)
    └── main.py             # CLI entry point: python -X utf8 lbo_model/main.py
```
