# VAT Reconciliation

A Streamlit app that automates UK VAT working papers. Upload your Xero exports and get a fully formatted Excel working paper in seconds.

_[Screenshot placeholder — add an app screenshot here before publishing to GitHub]_

## What it does

- Parses Xero CSV/Excel exports (VAT Return, Account Transactions, Trial Balance, Balance Sheet, Aged Debtors/Creditors)
- Reconciles output VAT, input VAT, and net VAT payable against the VAT Return
- Generates a formatted Excel working paper with reconciliation schedules
- Supports standard, flat rate, accrual, and repayment VAT scenarios
- Annual summary: stack Q1–Q4 snapshots into a full-year reconciliation

## Install & run

```bash
cd ~/Claude/accounting-automations/vat-reconciliation
pip install -r requirements.txt
streamlit run app.py
```

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit front-end |
| `vat_engine.py` | Core parsing, reconciliation, and workbook-building logic |
| `demo_files/` | Sample Xero exports for Demo Company UK |
| `test_scenarios/` | 6 test scenarios covering accrual, flat rate, dormant, repayment, large numbers, annual summary |

## Test scenarios

```bash
cd test_scenarios
python run_scenarios.py
```

## Script hook

> "I produce a VAT working paper in 14 seconds. What used to take two hours."

## GitHub Topics

Suggested repository topics (set these on the GitHub repo once created):

`python` `finance` `accounting` `automation` `streamlit` `vat` `xero` `tax` `hmrc`
