"""
test_annual_summary_edge_cases.py
==================================
Direct checks on AnnualSummaryBuilder's quarter-continuity flag (the
"◄ Opening bal does not match Q{i-1} closing bal" note in the VAT
CONTROL ROLL-FORWARD table), built from hand-crafted QuarterSnapshot
objects rather than a full VATWorkflowService run -- fast, and pins
down the £1.00 tolerance boundary and out-of-order/missing-quarter
input precisely.

Run:
    python test_scenarios/test_annual_summary_edge_cases.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))

from vat_engine import QuarterSnapshot, AnnualSummaryBuilder  # noqa: E402

# Probe workbooks are throwaway (only their cell contents are asserted on),
# so they go to a temp dir rather than cluttering test_scenarios/ with
# committed output files that carry no review value of their own.
OUT = Path(tempfile.mkdtemp(prefix="vat_annual_summary_edge_"))

CHECKS = []


def check(name):
    def deco(fn):
        CHECKS.append((name, fn))
        return fn
    return deco


def _snap(period_end, opening, box1, box4, closing) -> QuarterSnapshot:
    return QuarterSnapshot(
        client_name="Demo Company (UK)",
        period_end=period_end,
        opening_vat_balance=opening,
        box1=box1, box4=box4,
        hmrc_total=0.0,
        vat_control_tb=closing,
        closing_vat_balance=closing,
        vat_control_diff=0.0,
    )


def _notes(quarters):
    """Build the FY workbook and return {quarter_label: note_text_or_None}
    for every row in the VAT CONTROL ROLL-FORWARD table (col 1 = label,
    col 10 = continuity note written by `_note`)."""
    from openpyxl import load_workbook
    out_path = OUT / "notes_probe.xlsx"
    AnnualSummaryBuilder().build("Demo Company (UK)", "FY 2025/26", quarters, out_path)
    wb = load_workbook(out_path)
    ws = wb["FY VAT Summary"]

    # Locate the roll-forward table by its header, then read exactly
    # len(quarters) data rows beneath it (col 1 = "Qi" label).
    hdr_row = next(
        r for r in range(1, ws.max_row + 1)
        if ws.cell(row=r, column=1).value == "Quarter"
        and ws.cell(row=r, column=7).value == "= Closing Bal"
    )
    notes = {}
    for offset in range(1, len(quarters) + 1):
        r = hdr_row + offset
        label = ws.cell(row=r, column=1).value
        note = ws.cell(row=r, column=10).value
        notes[f"{label}#{offset}"] = note
    return notes


@check("continuity: exactly £1.00 off does NOT flag (tolerance is inclusive)")
def _(  ):
    q1 = _snap("31 May 2025", 0.0, 1000.0, 200.0, 800.0)
    q2 = _snap("31 Aug 2025", 799.00, 500.0, 100.0, 1199.0)  # |800 - 799.00| = 1.00
    notes = _notes([q1, q2])
    assert notes["Q1#1"] is None, f"Q1 should never have a prior-quarter note, got {notes['Q1#1']!r}"
    assert notes["Q2#2"] is None, f"expected no continuity note at exactly £1.00, got {notes['Q2#2']!r}"


@check("continuity: £1.01 off DOES flag")
def _():
    q1 = _snap("31 May 2025", 0.0, 1000.0, 200.0, 800.0)
    q2 = _snap("31 Aug 2025", 798.99, 500.0, 100.0, 1198.99)  # |800 - 798.99| = 1.01
    notes = _notes([q1, q2])
    assert notes["Q2#2"] is not None and "does not match" in notes["Q2#2"], (
        f"expected a continuity note just past £1.00, got {notes['Q2#2']!r}"
    )


@check("continuity: negative-direction break is also flagged (abs, not signed)")
def _():
    q1 = _snap("31 May 2025", 0.0, 1000.0, 200.0, 800.0)
    q2 = _snap("31 Aug 2025", 850.0, 500.0, 100.0, 1250.0)  # opening > prior closing by 50
    notes = _notes([q1, q2])
    assert notes["Q2#2"] is not None and "does not match" in notes["Q2#2"]


@check("continuity: 4 consecutive quarters with real carry-forward -> no notes anywhere")
def _():
    q1 = _snap("31 May 2025", 0.0, 1000.0, 200.0, 800.0)
    q2 = _snap("31 Aug 2025", 800.0, 500.0, 100.0, 1200.0)
    q3 = _snap("30 Nov 2025", 1200.0, 300.0, 900.0, 600.0)
    q4 = _snap("28 Feb 2026", 600.0, 50.0, 50.0, 600.0)
    notes = _notes([q1, q2, q3, q4])
    assert all(v is None for v in notes.values()), f"expected a clean roll-forward, got {notes}"


@check("missing quarter in the middle: builder trusts list order/position, "
       "not calendar continuity -- documents current behaviour, not a fix")
def _():
    # Real Q1 and real Q3 (Q2 never uploaded). The builder has no concept
    # of calendar quarters -- it labels rows "Q1", "Q2", ... by list
    # position, so the real Q3 snapshot below is rendered as row "Q2".
    q1 = _snap("31 May 2025", 0.0, 1000.0, 200.0, 800.0)
    q3_real = _snap("30 Nov 2025", 1200.0, 300.0, 900.0, 600.0)  # ties to a real Q2 close, not Q1's
    notes = _notes([q1, q3_real])
    assert notes["Q1#1"] is None
    # Positional label is "Q2" even though this snapshot is calendar-Q3;
    # the continuity check still does its job numerically (opening 1200
    # vs prior closing 800 -> flagged), just under the wrong quarter number.
    assert notes["Q2#2"] is not None and "does not match" in notes["Q2#2"], (
        "expected the real gap between Q1 and Q3 to still be caught "
        "numerically even though the row is mislabelled Q2"
    )


@check("single-quarter FY (Q1 only) does not crash and has no continuity note")
def _():
    q1 = _snap("31 May 2025", 0.0, 1000.0, 200.0, 800.0)
    notes = _notes([q1])
    assert notes == {"Q1#1": None}


@check("empty quarters list does not crash")
def _():
    from vat_engine import AnnualSummaryBuilder
    out_path = OUT / "empty_probe.xlsx"
    AnnualSummaryBuilder().build("Demo Company (UK)", "FY 2025/26", [], out_path)
    assert out_path.exists()


if __name__ == "__main__":
    failures = 0
    for name, fn in CHECKS:
        try:
            fn()
            print(f"  ok  {name}")
        except AssertionError as e:
            failures += 1
            print(f"  FAIL  {name}\n        {e}")
        except Exception as e:
            failures += 1
            print(f"  ERROR  {name}\n        {type(e).__name__}: {e}")

    print("\n" + "=" * 60)
    if failures:
        print(f"  {failures} check(s) FAILED")
        sys.exit(1)
    print("  PASS — all annual-summary edge-case checks passed")
