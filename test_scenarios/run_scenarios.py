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
    "box2_ni_acquisitions",
    "box6_diff_at_tolerance",
    "box6_domestic_zero_rated",
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

        if name == "box2_ni_acquisitions":
            # Box2 (NI EU acquisitions VAT) is non-zero on the return, but
            # Box1/Box4/TB are unchanged from the base demo file -- documents
            # that _vat_control's reconstruction (opening + Box1 - Box4 -
            # HMRC payments) doesn't reference Box2 at all, so vat_control_diff
            # comes out identical to the unmodified base figures (1563.27,
            # same as e.g. the "accrual" scenario which shares that base).
            # See HOLIDAY_LOG.md open question -- not asserting this is
            # correct or incorrect, just pinning down current behaviour.
            diff = s["vat_control_diff"]
            assert abs(diff - 1563.27) < 1e-9, (
                f"expected vat_control_diff unaffected by Box2, got {diff}"
            )

        if name == "box6_diff_at_tolerance":
            # Fixture is built so box6_diff sits exactly at cfg.tol_box6
            # (£5.00) rather than the general £1.00 tolerance -- confirms
            # _sheet_box6_rec's wider tolerance for this specific diff holds
            # up through the real pipeline, not just the synthetic
            # WorkbookBuilder check.
            diff = s["box6_diff"]
            assert abs(abs(diff) - 5.00) < 1e-9, (
                f"expected box6_diff to land exactly at £5.00 boundary, got {diff}"
            )

        if name == "box6_domestic_zero_rated":
            # £1,000 of domestic zero-rated income was added to Box 6 (and to
            # a "Zero Rated Income" transaction sub-section). If _box6()
            # correctly excludes it from "vatable" (same as it already did
            # for the EC-flavoured zero-rated sub-section), vat_proof_diff
            # should be unchanged from the unmodified base scenario (-0.04,
            # same as e.g. box2_ni_acquisitions/box6_diff_at_tolerance above,
            # which share the same unmodified Box1/Box6 base figures). If the
            # £1,000 were still wrongly taxed at 20%, it would land at
            # -0.04 + 200.00 = 199.96 instead.
            diff = s["vat_proof_diff"]
            assert abs(diff - (-0.04)) < 1e-9, (
                f"expected vat_proof_diff unaffected by domestic zero-rated "
                f"income (same -0.04 as the unmodified base scenario), got {diff}"
            )

        if name == "dormant":
            # Every VAT box (1-9) is meant to be zeroed by build_scenarios.py's
            # scenario_dormant() -- assert directly against the parsed boxes
            # (not just box1/4/5/6, which summary exposes) so a future
            # off-by-one in that fixture generator's row range fails loudly
            # here rather than silently passing because a given box happened
            # to already be zero in the base demo file.
            b = result.parsed.boxes
            for n in range(1, 10):
                v = getattr(b, f"box{n}")
                assert v == 0.0, f"expected dormant scenario Box {n} == 0.0, got {v}"

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
