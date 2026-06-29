# CLAUDE.md

## Overview
vat-reconciliation is a Streamlit app that automates UK VAT working papers.
Upload Xero exports (VAT Return, Account Transactions, Trial Balance, Balance
Sheet, Aged Debtors/Creditors) and it reconciles output/input/net VAT against the
VAT Return, then generates a fully formatted Excel working paper. Supports
standard, flat-rate, accrual, and repayment scenarios, plus an annual summary that
stacks Q1–Q4 snapshots into a full-year reconciliation.

## Quick Start
```bash
pip install -r requirements.txt
streamlit run app.py
```
Then upload the matching Xero exports in the app (sample set in `demo_files/`).

## Architecture
- `app.py` — Streamlit front-end only; imports everything from `vat_engine`.
- `vat_engine.py` — all logic, organised as:
  - **Parsers** (`VATReturnParser`, `TxnByBoxParser`, `TrialBalanceParser`,
    `BalanceSheetParser`, `AgedReportParser`, `AccountTxnParser`) coordinated by
    `ParseManager` — turn Xero exports into `ParsedData`.
  - `InputValidator` — validates the uploaded file set before parsing.
  - `ReconciliationEngine` — reconciles VAT boxes, produces `ReconciliationResults`.
  - `WorkbookBuilder` — renders the formatted Excel working paper.
  - `AnnualSummaryBuilder` + `QuarterSnapshot` — full-year stacking.
  - Config/result dataclasses: `JobConfig`, `JobResult`, `VATBoxes`, `ValidationResult`.

## Project Structure
- `app.py` — Streamlit UI / entrypoint
- `vat_engine.py` — parsing, reconciliation, workbook-building
- `demo_files/` — sample Xero exports for "Demo Company UK"
- `test_scenarios/` — 6 scenarios (accrual, flat rate, dormant, repayment, large
  numbers, annual summary) with `run_scenarios.py` driver

## Tests
```bash
cd test_scenarios
python run_scenarios.py
```

## Do Not
- Do not commit real client VAT data — `demo_files/` is synthetic ("Demo Company UK")
- Do not put parsing/reconciliation logic in `app.py` — it belongs in `vat_engine.py`
  so it stays testable without Streamlit
- Do not change the expected Xero export column/sheet layout the parsers depend on
  without updating both `vat_engine.py` and the `test_scenarios/` fixtures
