"""
run_scenarios.py
=================
Runs VATWorkflowService against every scenario directory under
test_scenarios/ and reports pass/fail + key figures. Use to catch
crashes / bad assumptions before shipping.

Run:
    python test_scenarios/run_scenarios.py
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))

from vat_engine import JobConfig, VATWorkflowService, InputError, ParseError  # noqa: E402

SCENARIOS = [
    "flat_rate",
    "accrual",
    "large_numbers",
    "vat_repayment",
    "credit_note_heavy",
    "negative_box4",
    "missing_vat_return",
    "dormant",
    "minimal_files",
    "empty_trial_balance",
    "control_diff_at_tolerance",
]


def run(name: str) -> bool:
    wd = HERE / name
    cfg = JobConfig(
        client_name=f"Demo Company (UK) - {name}",
        qe_end_date="28 Feb 2026",
        working_dir=str(wd),
        vat_return_file="Demo-Company-UK-VAT-Return.xlsx",
        trial_balance_file="Demo_Company__UK__-_Trial_Balance.xlsx",
        account_txn_file="Demo_Company__UK__-_Account_Transactions.xlsx" if (wd / "Demo_Company__UK__-_Account_Transactions.xlsx").exists() else None,
        aged_pay_file="Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx" if (wd / "Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx").exists() else None,
        aged_rec_file="Demo_Company__UK__-_Aged_Receivables_Detail.xlsx" if (wd / "Demo_Company__UK__-_Aged_Receivables_Detail.xlsx").exists() else None,
        balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
    )
    try:
        result = VATWorkflowService().run_job(cfg)
        s = result.summary
        print(f"\n  >> {name}: OK")
        print(f"     box1={s['box1']} box4={s['box4']} box5={s['box5']} box6={s['box6']}")
        print(f"     vat_proof_diff={s['vat_proof_diff']} vat_control_diff={s['vat_control_diff']} box6_diff={s['box6_diff']}")
        print(f"     workbook={s['workbook']}")

        if name == "control_diff_at_tolerance":
            # Fixture is built so the reconstructed VAT control balance sits
            # exactly £1.00 (cfg.tol_general) away from the TB figure --
            # confirms the real pipeline's rounding doesn't nudge a
            # boundary case off of what _flag() considers "reconciled".
            diff = s["vat_control_diff"]
            assert abs(abs(diff) - 1.00) < 1e-9, (
                f"expected vat_control_diff to land exactly at £1.00 boundary, got {diff}"
            )

        return True
    except (InputError, ParseError) as e:
        print(f"\n  >> {name}: HANDLED ERROR — {e}")
        return True
    except Exception:
        print(f"\n  >> {name}: CRASH")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    results = {name: run(name) for name in SCENARIOS}
    print("\n" + "=" * 60)
    for name, ok in results.items():
        print(f"  {'PASS' if ok else 'FAIL'}  {name}")
    if not all(results.values()):
        sys.exit(1)
