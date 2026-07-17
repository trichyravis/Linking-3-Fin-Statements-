# Dynamic 3-Statement Financial Model

A classroom-ready Streamlit application for demonstrating how an income statement, balance sheet and cash-flow statement are interlinked over a three-year forecast period.

Designed for **The Mountain Path Academy** — [themountainpathacademy.com](https://themountainpathacademy.com/).

## Features

- Dynamic sidebar inputs for revenue, margins, tax, working capital, capex, debt and dividends
- Three-year income statement, balance sheet and cash-flow forecasts
- Fully linked working-capital, PPE, debt, retained-earnings and cash schedules
- Automatic balance-sheet integrity check
- Profitability and funding charts
- Formula-linked Excel model with editable assumptions and linked statements
- A dedicated teaching guide with suggested classroom scenarios
- Responsive academy-branded interface

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

On Windows, activate the environment using `.venv\\Scripts\\activate`.

## Deploy on Streamlit Community Cloud

1. Create a new GitHub repository.
2. Upload `app.py`, `requirements.txt`, and this `README.md` to the repository root.
3. Sign in at [share.streamlit.io](https://share.streamlit.io/).
4. Select **Create app**, choose your repository and branch, and set the main file path to `app.py`.
5. Click **Deploy**.

No API keys or secrets are required.

## Accounting relationships demonstrated

- Revenue growth drives sales and operating costs.
- DSO, DIO and DPO convert operating activity into balance-sheet working-capital balances.
- Capex increases PPE and investing cash outflow; depreciation reduces PPE and profit but is added back in CFO.
- Profit after tax increases equity; dividends reduce equity and financing cash flow.
- Borrowing affects debt, interest expense and financing cash flow.
- Closing cash from the cash-flow statement becomes cash on the balance sheet.

## Excel download conventions

- **Blue-font cells** on the Assumptions sheet are editable inputs.
- **Green-font cells** contain Excel formulas linked across the three statements.
- Changes made to assumptions in Excel recalculate the Income Statement, Balance Sheet and Cash Flow Statement.
- The Balance Check should remain zero in every forecast year.

## Educational disclaimer

This application is intended for classroom demonstration. It is not a substitute for professional accounting, investment, tax or audit advice.
