"""
test_annual_summary.py
=======================
Runs 4 quarters through VATWorkflowService (reusing the demo files, with
the opening VAT balance carried forward each quarter from the previous
quarter's reconstructed closing balance), then builds the FY Annual
Summary workbook from the resulting QuarterSnapshots.

Run:
    python test_scenarios/test_annual_summary.py
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))

from vat_engine import JobConfig, VATWorkflowService, QuarterSnapshot, AnnualSummaryBuilder  # noqa: E402

DEMO = HERE.parent / "demo_files"
OUT = HERE / "annual_summary_test"
OUT.mkdir(exist_ok=True)

QUARTER_ENDS = ["31 May 2025", "31 Aug 2025", "30 Nov 2025", "28 Feb 2026"]

if __name__ == "__main__":
    opening_balance = 0.0
    snapshots = []

    for i, qe in enumerate(QUARTER_ENDS, 1):
        qdir = OUT / f"Q{i}"
        qdir.mkdir(exist_ok=True)
        # reuse demo files for each quarter (mechanics test, not real accounting)
        for f in DEMO.iterdir():
            if f.suffix == ".xlsx":
                (qdir / f.name).write_bytes(f.read_bytes())

        cfg = JobConfig(
            client_name="Demo Company (UK)",
            qe_end_date="28 Feb 2026",  # demo files are fixed to this date
            working_dir=str(qdir),
            vat_return_file="Demo-Company-UK-VAT-Return.xlsx",
            trial_balance_file="Demo_Company__UK__-_Trial_Balance.xlsx",
            account_txn_file="Demo_Company__UK__-_Account_Transactions.xlsx",
            aged_pay_file="Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx",
            aged_rec_file="Demo_Company__UK__-_Aged_Receivables_Detail.xlsx",
            balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
            opening_vat_balance=opening_balance,
        )
        result = VATWorkflowService().run_job(cfg)
        snap = QuarterSnapshot(**result.summary["snapshot"])
        snap.period_end = f"Q{i} (label override: {qe})"
        snapshots.append(snap)
        opening_balance = result.summary["closing_vat_balance"]
        print(f"Q{i}: opening={snap.opening_vat_balance:.2f}  closing(carried fwd)={opening_balance:.2f}")

    out_path = OUT / "FY_VAT_Reconciliation_Demo_Company_UK.xlsx"
    AnnualSummaryBuilder().build("Demo Company (UK)", "FY 2025/26", snapshots, out_path)
    print(f"\nAnnual summary written to: {out_path}")

    # round-trip JSON serialization check
    for i, s in enumerate(snapshots, 1):
        j = s.to_json()
        s2 = QuarterSnapshot.from_json(j)
        assert s2 == s, f"Snapshot Q{i} round-trip mismatch"
    print("JSON round-trip OK for all quarters")
