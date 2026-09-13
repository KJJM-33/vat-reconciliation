"""
test_workbook_builder.py
=========================
Direct assertions against WorkbookBuilder's Excel-formatting/formula
behaviour — cell colours, highlighted rows, flagged diffs. Built directly
against synthetic JobConfig/ParsedData/ReconciliationResults dataclasses
(no Xero fixtures needed), the same approach test_annual_summary_edge_cases.py
uses for AnnualSummaryBuilder. Fast, precise, and doesn't touch the
committed scenario fixtures.

Run:
    python test_scenarios/test_workbook_builder.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))

from vat_engine import (  # noqa: E402
    JobConfig,
    ValidationResult,
    Warning,
    ParsedData,
    ReconciliationResults,
    VATBoxes,
    WorkbookBuilder,
    _GREEN,
    _RED,
    _BLUE,
)

failures: list[str] = []


def check(label: str, actual, expected):
    if actual != expected:
        failures.append(f"{label}: expected {expected!r}, got {actual!r}")
    else:
        print(f"  ok  {label}")


def fg(cell) -> str:
    """Fill colour as a bare 6-hex-digit RGB string, or '' if unfilled.

    openpyxl reports an unfilled cell's fgColor.rgb as '00000000' (not
    None), so fill_type must be checked too, not just the rgb value.
    """
    if cell.fill.fill_type is None:
        return ""
    rgb = cell.fill.fgColor.rgb
    return rgb[-6:] if rgb else ""


def _cfg(**overrides) -> JobConfig:
    base = dict(client_name="Test Co", qe_end_date="28 Feb 2026", working_dir=".")
    base.update(overrides)
    return JobConfig(**base)


def _boxes(**overrides) -> VATBoxes:
    base = dict(period_start="01 Dec 2025", period_end="28 Feb 2026",
                vat_scheme="Standard", vat_number="GB123456789",
                period_str="For the period 01 Dec 2025 to 28 Feb 2026")
    base.update(overrides)
    return VATBoxes(**base)


# ─────────────────────────────────────────────────────────────────────────
# 1A. VAT Return sheet: Box 5 colouring (owe HMRC = red, repayment = green)
# ─────────────────────────────────────────────────────────────────────────
def test_box5_owed_is_red_repayment_is_green():
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes(box1=500.0, box3=500.0, box4=1500.0, box5=-1000.0))
    recs = ReconciliationResults()

    wb = Workbook()
    ws = wb.active
    builder._sheet_vat_return(ws, cfg, validation, data, recs)
    # Box 5 is the 5th row of the "VAT Calculations" block starting at row 18.
    box5_cell = ws.cell(row=22, column=3)
    check("VAT Return: Box5 negative (repayment) is green", fg(box5_cell), _GREEN)

    data2 = ParsedData(boxes=_boxes(box1=2000.0, box3=2000.0, box4=500.0, box5=1500.0))
    wb2 = Workbook()
    ws2 = wb2.active
    builder._sheet_vat_return(ws2, cfg, validation, data2, recs)
    box5_cell2 = ws2.cell(row=22, column=3)
    check("VAT Return: Box5 positive (owe HMRC) is red", fg(box5_cell2), _RED)


# ─────────────────────────────────────────────────────────────────────────
# 3C. Trial Balance sheet: VAT control nominal row is highlighted
# ─────────────────────────────────────────────────────────────────────────
def test_trial_balance_vat_control_row_highlighted():
    builder = WorkbookBuilder()
    cfg = _cfg(vat_control_nominal=820)
    validation = ValidationResult()
    tb = pd.DataFrame([
        {"Account Code": 200, "Account": "Sales", "Account Type": "Revenue",
         "Debit - Year to date": 0.0, "Credit - Year to date": 14986.0, "Prior Year": 0.0},
        {"Account Code": 820, "Account": "VAT Control", "Account Type": "Current Liability",
         "Debit - Year to date": 0.0, "Credit - Year to date": 1318.93, "Prior Year": 0.0},
    ])
    data = ParsedData(boxes=_boxes(), trial_balance=tb)
    recs = ReconciliationResults()

    wb = Workbook()
    ws = wb.active
    builder._sheet_trial_balance(ws, cfg, validation, data, recs)

    check("TB: nominal 820 (VAT control) row is highlighted blue",
          fg(ws.cell(row=12, column=1)), _BLUE)
    check("TB: non-control nominal 200 row is not highlighted",
          fg(ws.cell(row=11, column=1)), "")


# ─────────────────────────────────────────────────────────────────────────
# 3A/3B: Aged Payables/Receivables vs Balance Sheet agreement flag
# ─────────────────────────────────────────────────────────────────────────
def test_aged_bs_agreement_flag_green_when_reconciled_red_when_not():
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes())

    recs_ok = ReconciliationResults(ap_total=9362.69, bs_creditors=-9362.69, ap_bs_diff=0.0)
    wb = Workbook()
    ws = wb.active
    builder._sheet_aged_payables(ws, cfg, validation, data, recs_ok)
    check("Aged Payables: BS-agreement flag green when reconciled",
          fg(ws.cell(row=9, column=6)), _GREEN)

    recs_bad = ReconciliationResults(ap_total=9362.69, bs_creditors=-9000.00, ap_bs_diff=362.69)
    wb2 = Workbook()
    ws2 = wb2.active
    builder._sheet_aged_payables(ws2, cfg, validation, data, recs_bad)
    check("Aged Payables: BS-agreement flag red when out of tolerance",
          fg(ws2.cell(row=9, column=6)), _RED)

    recs_ar_ok = ReconciliationResults(ar_total=6049.90, bs_debtors=6049.90, ar_bs_diff=0.0)
    wb3 = Workbook()
    ws3 = wb3.active
    builder._sheet_aged_receivables(ws3, cfg, validation, data, recs_ar_ok)
    check("Aged Receivables: BS-agreement flag green when reconciled",
          fg(ws3.cell(row=9, column=6)), _GREEN)


# ─────────────────────────────────────────────────────────────────────────
# 0. File Register: Match? column reflects file_dates vs expected QE date
# ─────────────────────────────────────────────────────────────────────────
def test_file_register_match_column():
    builder = WorkbookBuilder()
    cfg = _cfg(qe_end_date="28 Feb 2026", vat_return_file="vr.xlsx",
               trial_balance_file="tb.xlsx", account_txn_file=None,
               aged_pay_file=None, aged_rec_file=None, balance_sheet_file="bs.xlsx")
    validation = ValidationResult(file_dates={
        "VAT Return": "28 Feb 2026",
        "Trial Balance": "01 Apr 2025",
    })
    data = ParsedData(boxes=_boxes())
    recs = ReconciliationResults()

    wb = Workbook()
    ws = wb.active
    builder._sheet_file_register(ws, cfg, validation, data, recs)

    # Row 8 = VAT Return (matches expected QE date) -> green tick.
    check("File Register: matching period is green", fg(ws.cell(row=8, column=4)), _GREEN)
    # Row 9 = Trial Balance (date mismatch) -> red cross.
    check("File Register: mismatched period is red", fg(ws.cell(row=9, column=4)), _RED)
    # Row 10 = Account Transactions (not configured, no file_dates entry) -> red cross too.
    check("File Register: unconfigured file also flagged red (no special-case)",
          fg(ws.cell(row=10, column=4)), _RED)
    check("File Register: unconfigured file shows 'Not configured'",
          ws.cell(row=10, column=2).value, "Not configured")


# ─────────────────────────────────────────────────────────────────────────
# VAT Checklist: all items render, count matches the source list
# ─────────────────────────────────────────────────────────────────────────
def test_checklist_all_items_rendered():
    builder = WorkbookBuilder()
    cfg = _cfg(preparer="J. Smith")
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes())
    recs = ReconciliationResults()

    wb = Workbook()
    ws = wb.active
    builder._sheet_checklist(ws, cfg, validation, data, recs)

    refs = [ws.cell(row=r, column=1).value for r in range(10, 30)
            if ws.cell(row=r, column=1).value not in (None, "")]
    check("Checklist: 20 items rendered", len(refs), 20)
    check("Checklist: first item ref is '1'", refs[0], "1")
    check("Checklist: last item ref is '14'", refs[-1], "14")
    check("Checklist: preparer name written to header",
          ws.cell(row=7, column=2).value, "J. Smith")


if __name__ == "__main__":
    test_box5_owed_is_red_repayment_is_green()
    test_trial_balance_vat_control_row_highlighted()
    test_aged_bs_agreement_flag_green_when_reconciled_red_when_not()
    test_file_register_match_column()
    test_checklist_all_items_rendered()

    print("\n" + "=" * 60)
    if failures:
        print(f"  FAIL — {len(failures)} assertion(s) failed:")
        for f in failures:
            print(f"    - {f}")
        sys.exit(1)
    else:
        print("  PASS — all workbook-builder checks passed")
