# Guided IFRS 9 ECL Computation App

This is a browser-based Expected Credit Loss app for loan schedules uploaded in Excel or CSV.

## How to run

Run the local server:

```powershell
C:\Users\ELITES\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe server.py
```

Then open:

```text
http://127.0.0.1:4174/
```

## What the user uploads

Use the **Download template** button for a starter workbook. The main sheet should contain a loan schedule with as many of these columns as available:

- Customer ID and Customer Name
- Disbursement Date
- Due Date or Maturity Date
- Loan Amount
- Outstanding Balance or Gross Carrying Amount
- Loan Classification
- Past Due Days / DPD
- EIR / Discount Rate
- Credit Limit, Undrawn Limit, or CCF for overdrafts and revolving facilities
- Collateral Type and Collateral Value
- Optional recovery cash flows in Month 1 to Month 12 columns
- Optional scorecard inputs such as Credit Search, Years In Business, Restructuring, collateral coverage, and cash flow coverage

The app can also accept an optional scorecard workbook with criteria, weights, and max scores. If no scorecard is uploaded, it uses the built-in scorecard adapted from the reference workbook.

## What the app computes

- IFRS 9 stage from DPD, default flag, restructuring, and loan classification.
- EAD/EOD from outstanding balance plus CCF-adjusted undrawn exposure.
- 12-month PD from loan classification, scorecard result, DPD, and macro stress.
- Lifetime PD from the 12-month PD over the remaining life.
- LGD from discounted expected recoveries, collateral haircut, or classification default.
- Scenario-weighted ECL using best/base/worse scenario weights.
- Portfolio totals, coverage ratio, and exportable result workbook.

## Macro data

The app includes a country selector and can fetch inflation and GDP growth from the World Bank World Development Indicators API:

- Inflation: `FP.CPI.TOTL.ZG`
- GDP growth: `NY.GDP.MKTP.KD.ZG`

If the online fetch is unavailable, the user can manually enter inflation and GDP growth in the app.

This tool is a computation aid. Final IFRS 9 model design, overrides, overlays, validation, and audit sign-off should follow the entity's approved accounting policy.
