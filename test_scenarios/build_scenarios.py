"""
build_scenarios.py
===================
Generates synthetic test scenarios by copying & modifying the demo Xero
export files. Each scenario is a directory under test_scenarios/<name>/
containing the same 6 filenames as demo_files/, ready to be pointed at
by a JobConfig(working_dir=...).

Run:
    python test_scenarios/build_scenarios.py
"""

from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
DEMO = HERE.parent / "demo_files"
FILES = {
    "vat_return":    "Demo-Company-UK-VAT-Return.xlsx",
    "trial_balance": "Demo_Company__UK__-_Trial_Balance.xlsx",
    "balance_sheet": "Demo_Company__UK__-_Balance_Sheet.xlsx",
    "account_txns":  "Demo_Company__UK__-_Account_Transactions.xlsx",
    "aged_pay":      "Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx",
    "aged_rec":      "Demo_Company__UK__-_Aged_Receivables_Detail.xlsx",
}


def _copy_base(name: str) -> Path:
    out = HERE / name
    out.mkdir(parents=True, exist_ok=True)
    for fname in FILES.values():
        shutil.copy(DEMO / fname, out / fname)
    return out


def _load(out: Path, key: str):
    return openpyxl.load_workbook(out / FILES[key])


def _scale_numeric(ws, factor: float, skip_cols: set[int] = frozenset(), header_rows: int = 0):
    """Multiply every numeric cell by factor, skipping header rows and skip_cols (1-indexed)."""
    for row in ws.iter_rows(min_row=header_rows + 1):
        for cell in row:
            if cell.column in skip_cols:
                continue
            v = cell.value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.value = round(v * factor, 2)


# ─────────────────────────────────────────────────────────────────────────
# Scenario 1: Flat Rate VAT scheme
#   Box4 (input VAT) -> 0, Box1/Box5 recalculated at an illustrative 16.5%
#   flat rate on Box6. This intentionally breaks the "Box6 x 20% = Box1"
#   proof check -- the engine should flag it (red) without crashing.
# ─────────────────────────────────────────────────────────────────────────
def scenario_flat_rate():
    out = _copy_base("flat_rate")
    wb = _load(out, "vat_return")
    ws = wb["VAT Return"]
    ws["C8"] = "Flat Rate"          # VAT Scheme (pandas row 7, col 2)
    box6 = ws["C22"].value          # Box 6 net sales
    flat_rate_pct = 0.165
    new_box1 = round(box6 * (1 + 0.20) * flat_rate_pct, 2)  # FRS applies % to gross turnover
    ws["C15"] = new_box1            # Box 1
    ws["C17"] = new_box1            # Box 3 = Box1 + Box2(0)
    ws["C18"] = 0.0                 # Box 4 - no input VAT reclaim under FRS
    ws["C19"] = new_box1            # Box 5 = Box3 - Box4
    wb.save(out / FILES["vat_return"])


# ─────────────────────────────────────────────────────────────────────────
# Scenario 2: Standard accrual scheme (relabel only -- engine doesn't
#   branch on scheme name beyond display, but check it round-trips).
# ─────────────────────────────────────────────────────────────────────────
def scenario_accrual():
    out = _copy_base("accrual")
    wb = _load(out, "vat_return")
    ws = wb["VAT Return"]
    ws["C8"] = "Standard (Accrual)"
    wb.save(out / FILES["vat_return"])


# ─────────────────────────────────────────────────────────────────────────
# Scenario 3: Large numbers -- scale every monetary figure x50 to
#   simulate a much bigger client (~£700k turnover).
# ─────────────────────────────────────────────────────────────────────────
def scenario_large_numbers():
    out = _copy_base("large_numbers")
    FACTOR = 50

    wb = _load(out, "vat_return")
    ws = wb["VAT Return"]
    for r in range(15, 27):  # rows with box values, col C; skip col B (box numbers)
        c = ws.cell(row=r, column=3)
        if isinstance(c.value, (int, float)):
            c.value = round(c.value * FACTOR, 2)
    ws2 = wb["Transactions by VAT Box"]
    _scale_numeric(ws2, FACTOR, skip_cols=set(), header_rows=0)
    wb.save(out / FILES["vat_return"])

    wb = _load(out, "trial_balance")
    ws = wb["Trial Balance"]
    # Account Code (col A) must not be scaled; Debit/Credit/Prior Year (D,E,F) scaled
    for row in ws.iter_rows(min_row=6):
        for cell in row:
            if cell.column == 1:
                continue
            v = cell.value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.value = round(v * FACTOR, 2)
    wb.save(out / FILES["trial_balance"])

    wb = _load(out, "balance_sheet")
    ws = wb["Balance Sheet"]
    _scale_numeric(ws, FACTOR, skip_cols={1, 2}, header_rows=4)
    wb.save(out / FILES["balance_sheet"])

    wb = _load(out, "account_txns")
    ws = wb["Account Transactions"]
    _scale_numeric(ws, FACTOR, skip_cols={1, 2, 3, 4}, header_rows=4)
    wb.save(out / FILES["account_txns"])

    wb = _load(out, "aged_pay")
    ws = wb["Aged Payables Detail"]
    _scale_numeric(ws, FACTOR, skip_cols={1, 2, 3}, header_rows=5)
    wb.save(out / FILES["aged_pay"])

    wb = _load(out, "aged_rec")
    ws = wb["Aged Receivables Detail"]
    _scale_numeric(ws, FACTOR, skip_cols={1, 2, 3, 4}, header_rows=5)
    wb.save(out / FILES["aged_rec"])


# ─────────────────────────────────────────────────────────────────────────
# Scenario 4: VAT repayment (Box4 > Box1 -> negative Box5, HMRC owes client)
# ─────────────────────────────────────────────────────────────────────────
def scenario_vat_repayment():
    out = _copy_base("vat_repayment")
    wb = _load(out, "vat_return")
    ws = wb["VAT Return"]
    ws["C15"] = 500.00     # Box 1
    ws["C17"] = 500.00     # Box 3
    ws["C18"] = 1500.00    # Box 4 (input VAT > output VAT)
    ws["C19"] = -1000.00   # Box 5 (repayment due from HMRC)
    wb.save(out / FILES["vat_return"])

    # Reflect the repayment position in the VAT control nominal (820) on the TB
    wb = _load(out, "trial_balance")
    ws = wb["Trial Balance"]
    for row in ws.iter_rows(min_row=6):
        if row[0].value == 820:
            row[3].value = 0       # Debit
            row[4].value = 1000    # Credit -> Cr-Dr = +1000 (asset/repayment owed)
    wb.save(out / FILES["trial_balance"])


# ─────────────────────────────────────────────────────────────────────────
# Scenario 5: Dormant / zero-activity quarter -- every VAT box, TB nominal
#   and aged report figure is zero. Checks the engine handles all-zero
#   inputs without div/0 or KeyErrors.
# ─────────────────────────────────────────────────────────────────────────
def scenario_dormant():
    out = _copy_base("dormant")

    wb = _load(out, "vat_return")
    ws = wb["VAT Return"]
    for r in range(15, 27):
        c = ws.cell(row=r, column=3)
        if isinstance(c.value, (int, float)):
            c.value = 0
    wb.save(out / FILES["vat_return"])

    wb = _load(out, "trial_balance")
    ws = wb["Trial Balance"]
    for row in ws.iter_rows(min_row=6):
        for cell in row:
            if cell.column in (1,):
                continue
            v = cell.value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.value = 0
    wb.save(out / FILES["trial_balance"])

    wb = _load(out, "balance_sheet")
    ws = wb["Balance Sheet"]
    for row in ws.iter_rows(min_row=5):
        c = row[2]
        if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
            c.value = 0
    wb.save(out / FILES["balance_sheet"])

    # Aged reports / account txns: strip all data rows, keep headers only
    for key, sheet, hdr_row in [
        ("aged_pay", "Aged Payables Detail", 5),
        ("aged_rec", "Aged Receivables Detail", 5),
        ("account_txns", "Account Transactions", 4),
    ]:
        wb = _load(out, key)
        ws = wb[sheet]
        max_row = ws.max_row
        if max_row > hdr_row + 1:
            ws.delete_rows(hdr_row + 2, max_row - hdr_row - 1)
        wb.save(out / FILES[key])


# ─────────────────────────────────────────────────────────────────────────
# Scenario 6: Missing optional files (no account txns / aged reports)
# ─────────────────────────────────────────────────────────────────────────
def scenario_minimal_files():
    out = _copy_base("minimal_files")
    for key in ("account_txns", "aged_pay", "aged_rec"):
        (out / FILES[key]).unlink()


if __name__ == "__main__":
    scenarios = [
        scenario_flat_rate,
        scenario_accrual,
        scenario_large_numbers,
        scenario_vat_repayment,
        scenario_dormant,
        scenario_minimal_files,
    ]
    for fn in scenarios:
        fn()
        print(f"  ✓ built scenario: {fn.__name__}")
