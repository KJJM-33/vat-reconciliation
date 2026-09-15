"""
test_unit_helpers.py
=====================
Direct assertions against vat_engine's pure helper functions and
InputValidator — no Xero fixtures required, so these run fast and pin
down exact expected values (tolerance boundaries, parsing edge cases)
that the fixture-driven scenarios in run_scenarios.py can't easily check.

Run:
    python test_scenarios/test_unit_helpers.py
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))

from vat_engine import (  # noqa: E402
    JobConfig,
    InputValidator,
    VATWorkflowService,
    InputError,
    _flag,
    _safe_float,
    _norm_date,
)
from openpyxl import Workbook  # noqa: E402

DEMO = HERE.parent / "demo_files"

failures: list[str] = []


def check(label: str, actual, expected):
    if actual != expected:
        failures.append(f"{label}: expected {expected!r}, got {actual!r}")
    else:
        print(f"  ok  {label}")


# ─────────────────────────────────────────────────────────────────────────
# _flag: tolerance boundary is inclusive (abs(value) <= tol -> green)
# ─────────────────────────────────────────────────────────────────────────
def test_flag_tolerance_boundary():
    wb = Workbook()
    ws = wb.active

    from vat_engine import _GREEN, _RED

    def fg(cell) -> str:
        return "00" + cell.fill.fgColor.rgb[-6:] if cell.fill.fgColor.rgb else ""

    cell_at_tol = _flag(ws, 1, 1, 1.00, tol=1.00)
    check("_flag: value == tol is green", fg(cell_at_tol)[-6:], _GREEN)

    cell_over_tol = _flag(ws, 2, 1, 1.01, tol=1.00)
    check("_flag: value just over tol is red", fg(cell_over_tol)[-6:], _RED)

    cell_neg_at_tol = _flag(ws, 3, 1, -1.00, tol=1.00)
    check("_flag: negative value == tol is green", fg(cell_neg_at_tol)[-6:], _GREEN)

    cell_zero = _flag(ws, 4, 1, 0.0, tol=0.0)
    check("_flag: zero value, zero tol is green", fg(cell_zero)[-6:], _GREEN)


# ─────────────────────────────────────────────────────────────────────────
# _safe_float: currency formatting, blanks, dashes, NaN
# ─────────────────────────────────────────────────────────────────────────
def test_safe_float():
    check("_safe_float: plain number", _safe_float(123.45), 123.45)
    check("_safe_float: currency string", _safe_float("£1,234.56"), 1234.56)
    check("_safe_float: negative currency", _safe_float("-£1,234.56"), -1234.56)
    check("_safe_float: thousands separator only", _safe_float("1,000"), 1000.0)
    check("_safe_float: empty string", _safe_float(""), 0.0)
    check("_safe_float: dash placeholder", _safe_float("-"), 0.0)
    check("_safe_float: NaN", _safe_float(float("nan")), 0.0)
    check("_safe_float: garbage text", _safe_float("n/a"), 0.0)
    check("_safe_float: whitespace padded", _safe_float("  42.00  "), 42.0)


# ─────────────────────────────────────────────────────────────────────────
# _norm_date: the formats Xero exports actually use
# ─────────────────────────────────────────────────────────────────────────
def test_norm_date():
    check("_norm_date: 'DD Month YYYY'", _norm_date("28 February 2026"), "28 Feb 2026")
    check("_norm_date: 'DD Mon YYYY'", _norm_date("28 Feb 2026"), "28 Feb 2026")
    check("_norm_date: 'DD/MM/YYYY'", _norm_date("28/02/2026"), "28 Feb 2026")
    check("_norm_date: 'YYYY-MM-DD'", _norm_date("2026-02-28"), "28 Feb 2026")
    check("_norm_date: unparseable falls back to str", _norm_date("not a date"), "not a date")


# ─────────────────────────────────────────────────────────────────────────
# InputValidator: severity levels for missing/mismatched files
# ─────────────────────────────────────────────────────────────────────────
def test_validator_missing_optional_file_is_info_not_error():
    cfg = JobConfig(
        working_dir=str(DEMO),
        vat_return_file="Demo-Company-UK-VAT-Return.xlsx",
        trial_balance_file="Demo_Company__UK__-_Trial_Balance.xlsx",
        balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
        account_txn_file=None,
        aged_pay_file=None,
        aged_rec_file=None,
    )
    result = InputValidator().validate(cfg)
    check("validator: missing optional files still passes", result.passed, True)
    severities = {w.message.split(":")[0]: w.severity for w in result.warnings}
    check("validator: Account Transactions severity", severities.get("Account Transactions"), "info")


def test_validator_missing_required_file_is_error():
    cfg = JobConfig(
        working_dir=str(DEMO),
        vat_return_file="does-not-exist.xlsx",
        trial_balance_file="Demo_Company__UK__-_Trial_Balance.xlsx",
        balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
        account_txn_file=None,
        aged_pay_file=None,
        aged_rec_file=None,
    )
    result = InputValidator().validate(cfg)
    check("validator: missing required file fails validation", result.passed, False)
    vat_warnings = [w for w in result.warnings if w.section == "files" and "VAT Return" in w.message]
    check("validator: VAT Return warning severity is error", vat_warnings[0].severity if vat_warnings else None, "error")


def test_validator_date_mismatch_is_warning_not_error():
    cfg = JobConfig(
        working_dir=str(DEMO),
        qe_end_date="31 May 2099",  # deliberately wrong vs demo file contents
        vat_return_file="Demo-Company-UK-VAT-Return.xlsx",
        trial_balance_file="Demo_Company__UK__-_Trial_Balance.xlsx",
        balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
        account_txn_file=None,
        aged_pay_file=None,
        aged_rec_file=None,
    )
    result = InputValidator().validate(cfg)
    check("validator: date mismatch alone still passes", result.passed, True)
    date_warnings = [w for w in result.warnings if w.section == "dates"]
    check("validator: date mismatch produces a warning", len(date_warnings) > 0, True)
    if date_warnings:
        check("validator: date mismatch severity is warning", date_warnings[0].severity, "warning")


# ─────────────────────────────────────────────────────────────────────────
# InputValidator + ParseManager: a *required* file left unconfigured
# (attr is None, as opposed to configured-but-missing-on-disk above) is
# currently reported as "info" by InputValidator -- same severity as a
# genuinely optional file -- because _FILE_SPEC doesn't distinguish
# required vs. optional files by name, only by whether cfg.<attr> was set
# at all. validation.passed therefore stays True, and it's ParseManager
# that ultimately blocks the job with a clean InputError. This pins down
# that the two layers together still fail safely (no crash, no silent
# wrong output) even though the pre-flight validation message alone would
# undersell the problem. See HOLIDAY_LOG.md for the full note.
# ─────────────────────────────────────────────────────────────────────────
def test_validator_unconfigured_required_file_is_info_not_error():
    cfg = JobConfig(
        working_dir=str(DEMO),
        vat_return_file="Demo-Company-UK-VAT-Return.xlsx",
        trial_balance_file=None,
        balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
        account_txn_file=None,
        aged_pay_file=None,
        aged_rec_file=None,
    )
    result = InputValidator().validate(cfg)
    check("validator: unconfigured required file (None) still passes validation", result.passed, True)
    tb_warnings = [w for w in result.warnings if "Trial Balance" in w.message]
    check("validator: unconfigured required file severity is info", tb_warnings[0].severity if tb_warnings else None, "info")


def test_parse_manager_blocks_unconfigured_required_file():
    cfg = JobConfig(
        working_dir=str(DEMO),
        vat_return_file="Demo-Company-UK-VAT-Return.xlsx",
        trial_balance_file=None,
        balance_sheet_file="Demo_Company__UK__-_Balance_Sheet.xlsx",
        account_txn_file=None,
        aged_pay_file=None,
        aged_rec_file=None,
    )
    try:
        VATWorkflowService().run_job(cfg)
        check("workflow: unconfigured Trial Balance raises InputError", "no exception raised", "InputError")
    except InputError as e:
        check("workflow: unconfigured Trial Balance raises InputError", "InputError", "InputError")
        check("workflow: InputError message names the missing file", "Trial Balance" in str(e), True)
    except Exception as e:  # pragma: no cover
        check("workflow: unconfigured Trial Balance raises InputError", f"{type(e).__name__}: {e}", "InputError")


if __name__ == "__main__":
    test_flag_tolerance_boundary()
    test_safe_float()
    test_norm_date()
    test_validator_missing_optional_file_is_info_not_error()
    test_validator_missing_required_file_is_error()
    test_validator_date_mismatch_is_warning_not_error()
    test_validator_unconfigured_required_file_is_info_not_error()
    test_parse_manager_blocks_unconfigured_required_file()

    print("\n" + "=" * 60)
    if failures:
        print(f"  FAIL — {len(failures)} assertion(s) failed:")
        for f in failures:
            print(f"    - {f}")
        sys.exit(1)
    else:
        print("  PASS — all unit checks passed")
