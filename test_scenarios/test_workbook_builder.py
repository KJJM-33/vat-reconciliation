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


# ─────────────────────────────────────────────────────────────────────────
# 1B. Txns by VAT Box: per-box totals + only non-empty sub-sections render
# ─────────────────────────────────────────────────────────────────────────
def test_txn_by_box_writes_totals_and_skips_empty_subsections():
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    b1_20 = pd.DataFrame([{"Date": "01/01/2026", "Account": "Sales", "Reference": "INV1",
                            "Details": "Sale", "VAT": 500.0, "Net": 2500.0}])
    b4_20 = pd.DataFrame([{"Date": "02/01/2026", "Account": "Expenses", "Reference": "BILL1",
                            "Details": "Purchase", "VAT": 100.0, "Net": 500.0}])
    data = ParsedData(
        boxes=_boxes(box1=2598.24, box4=1279.31, box6=14986.0, box7=0.0),
        txn_sections={
            "Box 1|20% (VAT on Income)": b1_20,
            "Box 4|20% (VAT on Expenses)": b4_20,
        },
    )
    recs = ReconciliationResults()

    wb = Workbook()
    ws = wb.active
    builder._sheet_txn_by_box(ws, cfg, validation, data, recs)

    check("Txn by box: Box 1 total written", ws.cell(row=10, column=6).value, 2598.24)
    check("Txn by box: Box 1 sub-header rendered", ws.cell(row=11, column=1).value,
          "20% (VAT on Income)")
    check("Txn by box: Box 1 transaction row written", ws.cell(row=13, column=3).value, "INV1")
    check("Txn by box: Box 4 total written", ws.cell(row=16, column=6).value, 1279.31)
    check("Txn by box: Box 4 transaction row written", ws.cell(row=19, column=3).value, "BILL1")
    # Box 6 has no configured txn_sections entries -> no sub-header/txn rows,
    # just the box-total banner row, before Box 7's banner follows immediately.
    check("Txn by box: Box 6 total written with no sub-sections", ws.cell(row=22, column=6).value,
          14986.0)
    check("Txn by box: Box 7 total written (zero, no sub-sections)",
          ws.cell(row=24, column=6).value, 0.0)


def test_txn_by_box_box8_box9_render_via_dynamic_subsections():
    """Box 8/9 (EU supplies/acquisitions, NI) previously had no entry at all in
    _sheet_txn_by_box's box list -- their transactions were parsed correctly by
    TxnByBoxParser but silently never rendered anywhere in the working paper.
    Unlike Box 1/4/6/7 (fixed, known sub-labels), Box 8/9's sub-label is
    whatever Xero grouped under that box for a given client, so this must be
    picked up dynamically from data.txn_sections rather than hard-coded."""
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    b8_ec = pd.DataFrame([{"Date": "23/02/2026", "Account": "Sales(200)", "Reference": "INV-0042",
                            "Details": "EC goods sale", "VAT": 0.0, "Net": 1995.0}])
    data = ParsedData(
        boxes=_boxes(box8=1995.0, box9=0.0),
        txn_sections={"Box 8|Zero Rated EC Goods Income": b8_ec},
    )
    recs = ReconciliationResults()

    wb = Workbook()
    ws = wb.active
    builder._sheet_txn_by_box(ws, cfg, validation, data, recs)

    check("Txn by box: Box 8 label written", ws.cell(row=18, column=1).value, "Box 8")
    check("Txn by box: Box 8 total written", ws.cell(row=18, column=6).value, 1995.0)
    check("Txn by box: Box 8 sub-header picked up from parsed data (not hard-coded)",
          ws.cell(row=19, column=1).value, "Zero Rated EC Goods Income")
    check("Txn by box: Box 8 transaction row written", ws.cell(row=21, column=3).value, "INV-0042")
    check("Txn by box: Box 9 label written", ws.cell(row=24, column=1).value, "Box 9")
    check("Txn by box: Box 9 total written (zero, no sub-sections)",
          ws.cell(row=24, column=6).value, 0.0)


# ─────────────────────────────────────────────────────────────────────────
# 2A. VAT Control: HMRC-payments table + closing reconciliation diff flag
# ─────────────────────────────────────────────────────────────────────────
def test_vat_control_hmrc_payments_table_and_diff_flag():
    builder = WorkbookBuilder()
    cfg = _cfg(opening_vat_balance=100.0, vat_control_nominal=820)
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes(box1=2598.24, box4=1279.31, box5=1318.93))

    hmrc_df = pd.DataFrame([{"Date": "15/02/2026", "Description": "HMRC VAT PAYMENT",
                              "Debit": 1200.0, "Credit": 0.0}])
    recs_ok = ReconciliationResults(hmrc_payments=hmrc_df, hmrc_total=1200.0,
                                     vat_control_tb=1318.93, vat_control_diff=0.5)
    wb = Workbook()
    ws = wb.active
    builder._sheet_vat_control(ws, cfg, validation, data, recs_ok)
    check("VAT Control: HMRC payment row written", ws.cell(row=24, column=1).value, "15/02/2026")
    check("VAT Control: HMRC payment description written",
          ws.cell(row=24, column=2).value, "HMRC VAT PAYMENT")
    check("VAT Control: HMRC payment amount written", ws.cell(row=24, column=3).value, 1200.0)
    check("VAT Control: total HMRC payments written", ws.cell(row=26, column=3).value, 1200.0)
    check("VAT Control: diff within tolerance is green", fg(ws.cell(row=31, column=3)), _GREEN)

    recs_bad = ReconciliationResults(hmrc_payments=pd.DataFrame(), hmrc_total=0.0,
                                      vat_control_tb=0.0, vat_control_diff=5000.0)
    wb2 = Workbook()
    ws2 = wb2.active
    builder._sheet_vat_control(ws2, cfg, validation, data, recs_bad)
    check("VAT Control: no HMRC payments shows italic placeholder message",
          ws2.cell(row=24, column=2).value,
          "No HMRC VAT payments identified in Account Transactions")
    check("VAT Control: nominal-not-in-TB note shown when vat_control_tb is falsy",
          ws2.cell(row=30, column=5).value,
          "◄ Nominal 820 not in TB — enter manually")
    check("VAT Control: diff outside tolerance is red", fg(ws2.cell(row=31, column=3)), _RED)


# ─────────────────────────────────────────────────────────────────────────
# 2B. Bank Rec: balances carried through, difference cell always shaded
# ─────────────────────────────────────────────────────────────────────────
def test_bank_rec_balances_and_difference_shading():
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes())
    recs = ReconciliationResults(bs_bank=5000.0)

    wb = Workbook()
    ws = wb.active
    builder._sheet_bank_rec(ws, cfg, validation, data, recs)

    check("Bank Rec: Balance per TB", ws.cell(row=12, column=3).value, 5000.0)
    check("Bank Rec: Balance per Statement", ws.cell(row=15, column=3).value, 5000.0)
    check("Bank Rec: Difference value", ws.cell(row=16, column=3).value, 0.0)
    check("Bank Rec: Difference row always shaded green", fg(ws.cell(row=16, column=3)), _GREEN)


# ─────────────────────────────────────────────────────────────────────────
# 2C. Box 6 Rec: proof-of-output-VAT block (Box 6 x 20% ~= Box 1)
# ─────────────────────────────────────────────────────────────────────────
def test_box6_rec_proof_of_output_vat_block():
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes(box1=2598.24, box6=14986.0))
    recs = ReconciliationResults(box6_tb_sales=10000.0, box6_tb_other=0.0, box6_tb_total=10000.0,
                                  box6_diff=4986.0, box6_zero_net=0.0, expected_output_vat=2997.2,
                                  vat_proof_diff=398.96)

    wb = Workbook()
    ws = wb.active
    builder._sheet_box6_rec(ws, cfg, validation, data, recs)

    check("Box6 Rec: VATable sales after removing zero-rated", ws.cell(row=28, column=3).value,
          14986.0)
    check("Box6 Rec: expected output VAT (x20%)", ws.cell(row=29, column=3).value, 2997.2)
    check("Box6 Rec: actual Box 1 shown for comparison", ws.cell(row=30, column=3).value, 2598.24)
    check("Box6 Rec: proof diff value", ws.cell(row=31, column=3).value, 398.96)
    check("Box6 Rec: proof diff outside default tolerance is red",
          fg(ws.cell(row=31, column=3)), _RED)
    check("Box6 Rec: QE-vs-YTD diff uses the wider tol_box6 tolerance (still red here)",
          fg(ws.cell(row=22, column=3)), _RED)


# ─────────────────────────────────────────────────────────────────────────
# 4A. Top 10 Box 4: rows + total. Note the `_vat` column contract below.
# ─────────────────────────────────────────────────────────────────────────
def test_top10_box4_rows_and_total():
    # _sheet_top10 reads recs.top10_box4["_vat"] directly (see vat_engine.py
    # _sheet_top10) -- that column only exists because
    # ReconciliationEngine._top10_box4 adds it before slicing nlargest(). A
    # ReconciliationResults built with a top10_box4 DataFrame that lacks
    # "_vat" (e.g. hand-built like everywhere else in this test file) raises
    # KeyError instead of silently doing something wrong -- confirmed while
    # writing this test. Not a production bug (the real pipeline always
    # populates both together) but the implicit contract is easy to trip
    # over when constructing ReconciliationResults directly, so it's
    # reproduced here deliberately rather than worked around.
    builder = WorkbookBuilder()
    cfg = _cfg()
    validation = ValidationResult()
    data = ParsedData(boxes=_boxes())

    top10 = pd.DataFrame([
        {"Date": "05/01/2026", "Account": "Equipment", "Reference": "INV100",
         "Details": "Laptop", "VAT": 400.0, "Net": 2000.0, "_vat": 400.0},
        {"Date": "06/01/2026", "Account": "Software", "Reference": "INV101",
         "Details": "Licence", "VAT": 100.0, "Net": 500.0, "_vat": 100.0},
    ])
    recs = ReconciliationResults(top10_box4=top10)
    wb = Workbook()
    ws = wb.active
    builder._sheet_top10(ws, cfg, validation, data, recs)

    check("Top10: first row reference", ws.cell(row=10, column=3).value, "INV100")
    check("Top10: second row VAT amount", ws.cell(row=11, column=5).value, 100.0)
    check("Top10: total sums the VAT column", ws.cell(row=13, column=5).value, 500.0)

    recs_empty = ReconciliationResults(top10_box4=pd.DataFrame())
    wb2 = Workbook()
    ws2 = wb2.active
    builder._sheet_top10(ws2, cfg, validation, data, recs_empty)
    check("Top10: empty top10_box4 renders zero total, no crash",
          ws2.cell(row=11, column=5).value, 0.0)


if __name__ == "__main__":
    test_box5_owed_is_red_repayment_is_green()
    test_trial_balance_vat_control_row_highlighted()
    test_aged_bs_agreement_flag_green_when_reconciled_red_when_not()
    test_file_register_match_column()
    test_checklist_all_items_rendered()
    test_txn_by_box_writes_totals_and_skips_empty_subsections()
    test_vat_control_hmrc_payments_table_and_diff_flag()
    test_bank_rec_balances_and_difference_shading()
    test_box6_rec_proof_of_output_vat_block()
    test_top10_box4_rows_and_total()

    print("\n" + "=" * 60)
    if failures:
        print(f"  FAIL — {len(failures)} assertion(s) failed:")
        for f in failures:
            print(f"    - {f}")
        sys.exit(1)
    else:
        print("  PASS — all workbook-builder checks passed")
