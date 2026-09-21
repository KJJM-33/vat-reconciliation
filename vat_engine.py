"""
vat_engine.py — VAT Working Papers Engine v3
=============================================
Merges proven Xero parsing logic with clean architecture
designed for multi-client use and future app integration.

CURRENT USAGE (CLI):
    python vat_engine.py

FUTURE USAGE (app backend):
    from vat_engine import VATWorkflowService
    result = VATWorkflowService().run_job(job)

CONFIGURATION:
    Edit the JobConfig dataclass at the bottom (CLI mode)
    or pass a JobConfig to VATWorkflowService.run_job() (app mode).
"""

from __future__ import annotations

# ── stdlib ────────────────────────────────────────────────────────────────────
import json
import logging
import os
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ── third-party ───────────────────────────────────────────────────────────────
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
log = logging.getLogger("vat_engine")


# ══════════════════════════════════════════════════════════════════════════════
# EXCEPTIONS
# ══════════════════════════════════════════════════════════════════════════════

class VATEngineError(Exception):
    """Base exception."""

class InputError(VATEngineError):
    """Missing or inconsistent input."""

class ParseError(VATEngineError):
    """File cannot be parsed reliably."""


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION  ← only section that changes per client / per run
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class JobConfig:
    """
    One instance per job / client run.
    In app mode the frontend populates this from a form / file upload.
    In CLI mode edit the defaults at the bottom of this file.
    """
    # ── identity ──────────────────────────────────────────────────────────────
    client_name: str = "Demo Company (UK)"
    qe_end_date: str = "28 Feb 2026"       # Plain English, e.g. "28 Feb 2026"
    preparer: str = ""
    reviewer: str = ""

    # ── file paths ────────────────────────────────────────────────────────────
    # All paths relative to working_dir, or absolute.
    # Set to None to skip an optional file.
    working_dir: str = "."
    vat_return_file: str          = "Demo-Company-UK-VAT-Return.xlsx"
    trial_balance_file: str       = "Demo_Company__UK__-_Trial_Balance.xlsx"
    account_txn_file: str         = "Demo_Company__UK__-_Account_Transactions.xlsx"
    aged_pay_file: str            = "Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx"
    aged_rec_file: str            = "Demo_Company__UK__-_Aged_Receivables_Detail.xlsx"
    balance_sheet_file: str       = "Demo_Company__UK__-_Balance_Sheet.xlsx"

    # ── accounting config (customise per client if needed) ────────────────────
    sales_nominals: List[int]       = field(default_factory=lambda: [200])
    other_income_nominals: List[int]= field(default_factory=lambda: [270])
    vat_control_nominal: int        = 820
    opening_vat_balance: float      = 0.00   # Prior period closing VAT control (Cr)

    # ── tolerances ────────────────────────────────────────────────────────────
    tol_general: float = 1.00
    tol_box6:    float = 5.00


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class Warning:
    section: str
    severity: str   # "info" | "warning" | "error"
    message: str


@dataclass
class ValidationResult:
    passed: bool = True
    warnings: List[Warning] = field(default_factory=list)
    file_dates: Dict[str, str] = field(default_factory=dict)

    def warn(self, section: str, msg: str, severity: str = "warning") -> None:
        self.warnings.append(Warning(section, severity, msg))
        if severity == "error":
            self.passed = False


@dataclass
class VATBoxes:
    box1: float = 0.0
    box2: float = 0.0
    box3: float = 0.0
    box4: float = 0.0
    box5: float = 0.0
    box6: float = 0.0
    box7: float = 0.0
    box8: float = 0.0
    box9: float = 0.0
    period_start: str = ""
    period_end: str   = ""
    vat_scheme: str   = ""
    vat_number: str   = ""
    period_str: str   = ""   # full "For the period…" header string
    client_name: str  = ""


@dataclass
class ParsedData:
    boxes: VATBoxes                             = field(default_factory=VATBoxes)
    txn_sections: Dict[str, pd.DataFrame]       = field(default_factory=dict)
    trial_balance: pd.DataFrame                 = field(default_factory=pd.DataFrame)
    balance_sheet_raw: pd.DataFrame             = field(default_factory=pd.DataFrame)
    aged_pay_raw: pd.DataFrame                  = field(default_factory=pd.DataFrame)
    aged_rec_raw: pd.DataFrame                  = field(default_factory=pd.DataFrame)
    account_txns: pd.DataFrame                  = field(default_factory=pd.DataFrame)


@dataclass
class ReconciliationResults:
    # VAT control
    vat_control_diff: float                     = 0.0
    vat_control_tb: float                       = 0.0
    hmrc_payments: pd.DataFrame                 = field(default_factory=pd.DataFrame)
    hmrc_total: float                           = 0.0

    # Box 6
    box6_tb_sales: float                        = 0.0
    box6_tb_other: float                        = 0.0
    box6_tb_total: float                        = 0.0
    box6_diff: float                            = 0.0
    box6_zero_net: float                        = 0.0
    expected_output_vat: float                  = 0.0
    vat_proof_diff: float                       = 0.0

    # Aged balances
    ap_total: float                             = 0.0
    ap_overdue: float                           = 0.0
    ap_bs_diff: float                           = 0.0
    ar_total: float                             = 0.0
    ar_overdue: float                           = 0.0
    ar_bs_diff: float                           = 0.0

    # Balance sheet values
    bs_debtors: float                           = 0.0
    bs_creditors: float                         = 0.0
    bs_bank: float                              = 0.0

    # Top 10 box 4
    top10_box4: pd.DataFrame                    = field(default_factory=pd.DataFrame)


@dataclass
class JobResult:
    """Returned to the caller — app or CLI."""
    client_name: str
    period: str
    validation: ValidationResult
    parsed: ParsedData
    recs: ReconciliationResults
    workbook_path: str
    summary: Dict[str, Any]


@dataclass
class QuarterSnapshot:
    """
    Compact record of one quarter's VAT return + reconciliation, used to
    carry the VAT control balance forward into the next quarter and to
    build the annual (FY) VAT reconciliation.

    Saved/loaded as JSON sidecar files alongside the working papers.
    """
    client_name: str       = ""
    period_start: str       = ""
    period_end: str          = ""
    vat_scheme: str          = ""
    opening_vat_balance: float = 0.0
    box1: float = 0.0
    box2: float = 0.0
    box3: float = 0.0
    box4: float = 0.0
    box5: float = 0.0
    box6: float = 0.0
    box7: float = 0.0
    box8: float = 0.0
    box9: float = 0.0
    hmrc_total: float        = 0.0
    vat_control_tb: float    = 0.0     # VAT control nominal per TB at this QE (YTD)
    closing_vat_balance: float = 0.0   # reconstructed closing -> next quarter's opening
    vat_control_diff: float  = 0.0

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2)

    @staticmethod
    def from_json(text: str) -> "QuarterSnapshot":
        return QuarterSnapshot(**json.loads(text))

    @staticmethod
    def build(cfg: "JobConfig", boxes: "VATBoxes", recs: "ReconciliationResults") -> "QuarterSnapshot":
        """Build a snapshot from a completed job's config/parsed boxes/reconciliation."""
        closing = round(cfg.opening_vat_balance + boxes.box1 - boxes.box4 - recs.hmrc_total, 2)
        return QuarterSnapshot(
            client_name=cfg.client_name,
            period_start=boxes.period_start,
            period_end=boxes.period_end,
            vat_scheme=boxes.vat_scheme,
            opening_vat_balance=cfg.opening_vat_balance,
            box1=boxes.box1, box2=boxes.box2, box3=boxes.box3,
            box4=boxes.box4, box5=boxes.box5, box6=boxes.box6,
            box7=boxes.box7, box8=boxes.box8, box9=boxes.box9,
            hmrc_total=recs.hmrc_total,
            vat_control_tb=recs.vat_control_tb,
            closing_vat_balance=closing,
            vat_control_diff=recs.vat_control_diff,
        )


# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _norm_date(d: Any) -> str:
    for fmt in ["%d %B %Y", "%d %b %Y", "%d/%m/%Y", "%Y-%m-%d"]:
        try:
            return datetime.strptime(str(d).strip(), fmt).strftime("%d %b %Y")
        except ValueError:
            continue
    return str(d).strip()


def _safe_float(v: Any) -> float:
    if pd.isna(v):
        return 0.0
    txt = str(v).replace(",", "").replace("£", "").strip()
    if txt in {"", "-"}:
        return 0.0
    try:
        return float(txt)
    except ValueError:
        return 0.0


def _resolve(cfg: JobConfig, attr: str) -> Optional[Path]:
    val = getattr(cfg, attr, None)
    if val is None:
        return None
    p = Path(cfg.working_dir) / val
    return p if p.exists() else None


def _read_xl(path: Path, header=None, nrows=None) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=0, header=header, nrows=nrows)


def _extract_period(path: Path, row_i: int, col_i: int) -> str:
    """Pull a date from a known cell in the first sheet."""
    try:
        raw = _read_xl(path, nrows=15)
        cell = str(raw.iloc[row_i, col_i])
        m = re.search(r"\d{1,2}\s+\w+\s+\d{4}|\d{1,2}/\d{2}/\d{4}", cell)
        return _norm_date(m.group()) if m else cell
    except Exception:
        return "Unknown"


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 1 — VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

class InputValidator:
    """
    Checks all files exist and their period dates match the expected QE end date.
    Returns a ValidationResult with warnings — does NOT raise by default so the
    caller (app or CLI) can decide whether to proceed.
    """

    # (config_attr, row_in_file, col_in_file)
    _FILE_SPEC: List[Tuple[str, str, int, int]] = [
        ("vat_return_file",    "VAT Return",           11, 2),
        ("trial_balance_file", "Trial Balance",          2, 0),
        ("account_txn_file",   "Account Transactions",   2, 0),
        ("aged_pay_file",      "Aged Payables",           2, 0),
        ("aged_rec_file",      "Aged Receivables",        2, 0),
        ("balance_sheet_file", "Balance Sheet",           2, 0),
    ]

    def validate(self, cfg: JobConfig) -> ValidationResult:
        result = ValidationResult()
        qe = _norm_date(cfg.qe_end_date)

        for attr, label, ri, ci in self._FILE_SPEC:
            path = _resolve(cfg, attr)
            if path is None:
                fname = getattr(cfg, attr, None)
                if fname is None:
                    result.warn("files", f"{label}: not configured (skipped)", "info")
                else:
                    result.warn("files", f"{label}: file not found — {fname}", "error")
                continue

            period = _extract_period(path, ri, ci)
            result.file_dates[label] = period

            if period != qe:
                result.warn(
                    "dates",
                    f"{label}: period '{period}' does not match expected '{qe}' [{path.name}]",
                    "warning",
                )

        return result


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 2 — PARSING
# ══════════════════════════════════════════════════════════════════════════════

_VAT_SUB_LABELS = [
    "20% (VAT on Income)",
    "Zero Rated Income",
    "Zero Rated EC Goods Income",
    "20% (VAT on Expenses)",
    "20% (VAT on Expenses) - Adjusted",
    "5% (VAT on Expenses)",
    "Zero Rated Expenses",
    "Exempt Expenses",
    "Reverse Charge Expenses (20%)",
    "Reverse Charge Expenses (20%) Reclaimed VAT",
]


class VATReturnParser:
    """
    Parses the Xero VAT Return export.
    Uses known row positions as primary strategy (fast, reliable for Xero).
    Falls back to label-scanning if values are zero (handles layout drift).
    """

    # Xero row positions (0-indexed) in the 'VAT Return' sheet
    _ROW_MAP = {
        "client_name":  (1, 0),
        "period_str":   (2, 0),
        "vat_number":   (6, 2),
        "vat_scheme":   (7, 2),
        "period_start": (10, 2),
        "period_end":   (11, 2),
        "box1": (14, 2), "box2": (15, 2), "box3": (16, 2),
        "box4": (17, 2), "box5": (18, 2),
        "box6": (21, 2), "box7": (22, 2),
        "box8": (25, 2), "box9": (26, 2),
    }

    # Label fragments used as fallback (lower-case)
    _LABEL_MAP = {
        "box1": ["vat due in the period on sales"],
        "box2": ["acquisitions of goods made in northern ireland"],
        "box3": ["total vat due"],
        "box4": ["vat reclaimed in the period on purchases"],
        "box5": ["vat to pay hmrc", "net vat"],
        "box6": ["total value of sales and all other outputs"],
        "box7": ["total value of purchases and all other inputs"],
        "box8": ["total value of all supplies of goods"],
        "box9": ["total value of all acquisitions"],
    }

    def parse(self, path: Path) -> VATBoxes:
        raw = pd.read_excel(path, sheet_name="VAT Return", header=None)

        def gs(r, c):
            try:
                v = raw.iloc[r, c]
                return str(v) if pd.notna(v) else ""
            except IndexError:
                return ""

        def gf(r, c):
            try:
                v = raw.iloc[r, c]
                return _safe_float(v)
            except IndexError:
                return 0.0

        boxes = VATBoxes(
            client_name  = gs(*self._ROW_MAP["client_name"]),
            period_str   = gs(*self._ROW_MAP["period_str"]),
            vat_number   = gs(*self._ROW_MAP["vat_number"]),
            vat_scheme   = gs(*self._ROW_MAP["vat_scheme"]),
            period_start = gs(*self._ROW_MAP["period_start"]),
            period_end   = gs(*self._ROW_MAP["period_end"]),
            box1 = gf(*self._ROW_MAP["box1"]),
            box2 = gf(*self._ROW_MAP["box2"]),
            box3 = gf(*self._ROW_MAP["box3"]),
            box4 = gf(*self._ROW_MAP["box4"]),
            box5 = gf(*self._ROW_MAP["box5"]),
            box6 = gf(*self._ROW_MAP["box6"]),
            box7 = gf(*self._ROW_MAP["box7"]),
            box8 = gf(*self._ROW_MAP["box8"]),
            box9 = gf(*self._ROW_MAP["box9"]),
        )

        # Fallback: scan for labels if primary positions returned zero
        flat = raw.fillna("").astype(str)
        for box_key, labels in self._LABEL_MAP.items():
            if getattr(boxes, box_key) == 0.0:
                val = self._scan_for_value(flat, labels)
                if val != 0.0:
                    setattr(boxes, box_key, val)
                    log.debug("Fallback label scan used for %s: %.2f", box_key, val)

        return boxes

    @staticmethod
    def _scan_for_value(flat: pd.DataFrame, labels: List[str]) -> float:
        for r in range(len(flat)):
            for c in range(len(flat.columns)):
                cell = flat.iat[r, c].lower()
                if any(lbl in cell for lbl in labels):
                    # offset 1 is the box-number column (e.g. "1", "4", "6") in
                    # the standard Xero layout -- skip it so a genuinely zero
                    # box value isn't replaced by its own box number.
                    for offset in range(2, 5):
                        nc = c + offset
                        if nc < len(flat.columns):
                            v = _safe_float(flat.iat[r, nc])
                            if v != 0.0:
                                return v
        return 0.0


class TxnByBoxParser:
    """
    Parses the 'Transactions by VAT Box' sheet from the Xero VAT export.
    Returns a dict keyed "Box N|Sub-section label" → DataFrame.
    """

    def parse(self, path: Path) -> Dict[str, pd.DataFrame]:
        raw = pd.read_excel(path, sheet_name="Transactions by VAT Box", header=None)
        sections: Dict[str, pd.DataFrame] = {}
        cur_box: Optional[str] = None
        cur_sub: Optional[str] = None
        col_found = False
        rows: List[dict] = []

        def _save():
            if cur_box and cur_sub and rows:
                key = f"{cur_box}|{cur_sub}"
                sections[key] = pd.DataFrame(rows)

        for _, row in raw.iterrows():
            v0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""

            if re.match(r"^Box \d+$", v0):
                _save(); cur_box = v0; cur_sub = None; col_found = False; rows = []
                continue

            if v0 in _VAT_SUB_LABELS:
                _save(); cur_sub = v0; col_found = False; rows = []
                continue

            if v0 == "Date":
                col_found = True
                continue

            if col_found and cur_box and cur_sub and v0 not in ("", "nan"):
                try:
                    pd.to_datetime(v0, dayfirst=True)
                    rows.append({
                        "Date":      row.iloc[0],
                        "Account":   row.iloc[1],
                        "Reference": row.iloc[2],
                        "Details":   row.iloc[3],
                        "VAT":       row.iloc[4],
                        "Net":       row.iloc[5],
                    })
                except ValueError:
                    pass

        _save()
        return sections


class TrialBalanceParser:
    """
    Parses the Xero Trial Balance export.
    Detects the header row dynamically (looks for 'Account Code').
    Returns a clean DataFrame with standardised column names.
    """

    def parse(self, path: Path) -> pd.DataFrame:
        raw = _read_xl(path)
        # Find header row
        hdr_row = next(
            (i for i, r in raw.iterrows() if str(r.iloc[0]).strip() == "Account Code"),
            None,
        )
        if hdr_row is None:
            raise ParseError(f"Cannot find 'Account Code' header in {path.name}")

        df = pd.read_excel(path, sheet_name=0, header=hdr_row)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(subset=["Account Code"])
        df["Account Code"] = pd.to_numeric(df["Account Code"], errors="coerce")
        df = df.dropna(subset=["Account Code"])
        df["Account Code"] = df["Account Code"].astype(int)
        return df.reset_index(drop=True)


class BalanceSheetParser:
    def parse(self, path: Path) -> pd.DataFrame:
        return _read_xl(path)


class AgedReportParser:
    """Generic parser for both Aged Payables and Aged Receivables."""

    def parse(self, path: Path) -> pd.DataFrame:
        raw = _read_xl(path)
        hdr_row = next(
            (i for i, r in raw.iterrows() if str(r.iloc[0]).strip() == "Invoice Date"),
            None,
        )
        if hdr_row is None:
            return pd.DataFrame()
        df = pd.read_excel(path, sheet_name=0, header=hdr_row)
        df.columns = [str(c).strip() for c in df.columns]
        return df.reset_index(drop=True)


class AccountTxnParser:
    def parse(self, path: Path) -> pd.DataFrame:
        raw = _read_xl(path)
        hdr_row = next(
            (i for i, r in raw.iterrows() if str(r.iloc[0]).strip() == "Date"),
            None,
        )
        if hdr_row is None:
            return pd.DataFrame()
        df = pd.read_excel(path, sheet_name=0, header=hdr_row)
        df.columns = [str(c).strip() for c in df.columns]
        return df.reset_index(drop=True)


class ParseManager:
    """Orchestrates all parsers and returns a single ParsedData object."""

    def parse_all(self, cfg: JobConfig) -> ParsedData:
        log.info("Parsing source files")
        data = ParsedData()

        # VAT Return (required)
        vat_path = _resolve(cfg, "vat_return_file")
        if vat_path is None:
            raise InputError("VAT Return file not found — cannot proceed.")
        data.boxes = VATReturnParser().parse(vat_path)
        data.txn_sections = TxnByBoxParser().parse(vat_path)
        log.info("  VAT Return      ✓  Box1=%.2f  Box4=%.2f  Box5=%.2f  Box6=%.2f",
                 data.boxes.box1, data.boxes.box4, data.boxes.box5, data.boxes.box6)

        # Trial Balance (required)
        tb_path = _resolve(cfg, "trial_balance_file")
        if tb_path is None:
            raise InputError("Trial Balance file not found — cannot proceed.")
        data.trial_balance = TrialBalanceParser().parse(tb_path)
        log.info("  Trial Balance   ✓  %d nominals", len(data.trial_balance))

        # Balance Sheet (required)
        bs_path = _resolve(cfg, "balance_sheet_file")
        if bs_path is None:
            raise InputError("Balance Sheet file not found — cannot proceed.")
        data.balance_sheet_raw = BalanceSheetParser().parse(bs_path)
        log.info("  Balance Sheet   ✓")

        # Optional files
        for attr, label in [
            ("aged_pay_file",  "Aged Payables"),
            ("aged_rec_file",  "Aged Receivables"),
        ]:
            p = _resolve(cfg, attr)
            if p:
                df = AgedReportParser().parse(p)
                if attr == "aged_pay_file":
                    data.aged_pay_raw = df
                else:
                    data.aged_rec_raw = df
                log.info("  %-16s ✓", label)
            else:
                log.info("  %-16s —  skipped", label)

        acct_path = _resolve(cfg, "account_txn_file")
        if acct_path:
            data.account_txns = AccountTxnParser().parse(acct_path)
            log.info("  Account Txns    ✓  %d rows", len(data.account_txns))
        else:
            log.info("  Account Txns    —  skipped")

        return data


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 3 — RECONCILIATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════

class ReconciliationEngine:
    """
    All accounting checks live here.
    Pure functions of ParsedData + JobConfig → ReconciliationResults.
    No I/O, no formatting — easily unit-testable.
    """

    def run(self, cfg: JobConfig, data: ParsedData) -> ReconciliationResults:
        log.info("Running reconciliations")
        r = ReconciliationResults()

        r.bs_debtors   = self._bs_val(data, "Accounts Receivable")
        r.bs_creditors = self._bs_val(data, "Accounts Payable")
        r.bs_bank      = self._bs_val(data, "Business Bank Account")

        self._vat_control(cfg, data, r)
        self._box6(cfg, data, r)
        self._aged_payables(cfg, data, r)
        self._aged_receivables(cfg, data, r)
        self._top10_box4(data, r)

        return r

    # ── helpers ───────────────────────────────────────────────────────────────

    def _tb_balance(self, data: ParsedData, code: int) -> float:
        """Net balance for a nominal: Cr − Dr (positive = credit balance)."""
        tb = data.trial_balance
        row = tb[tb["Account Code"] == code]
        if row.empty:
            return 0.0
        dr = _safe_float(row["Debit - Year to date"].values[0])
        cr = _safe_float(row["Credit - Year to date"].values[0])
        return round(cr - dr, 2)

    def _bs_val(self, data: ParsedData, label: str) -> float:
        bs = data.balance_sheet_raw
        for _, row in bs.iterrows():
            if label.lower() in str(row.iloc[1]).lower():
                try:
                    return _safe_float(row.iloc[2])
                except Exception:
                    pass
        return 0.0

    def _aged_invoice_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        return df[
            df["Invoice Date"].astype(str).str.match(r"\d{4}-\d{2}-\d{2}", na=False)
        ].copy()

    def _aged_overdue(self, df: pd.DataFrame) -> float:
        cols = ["< 1 Month", "1 Month", "2 Months", "3 Months", "Older"]
        return float(sum(
            df[c].sum() for c in cols if c in df.columns
        ))

    # ── reconciliations ───────────────────────────────────────────────────────

    def _vat_control(self, cfg: JobConfig, data: ParsedData, r: ReconciliationResults):
        """
        Reconstructed VAT control account:
            Opening balance
          + Output VAT (Box 1)
          − Input VAT (Box 4)
          − HMRC payments
          = Closing balance  →  agree to TB nominal 820
        """
        r.vat_control_tb = self._tb_balance(data, cfg.vat_control_nominal)

        # Find HMRC VAT payments in account transactions
        if not data.account_txns.empty:
            desc_col = next(
                (c for c in data.account_txns.columns
                 if "desc" in c.lower() or "detail" in c.lower()),
                None,
            )
            if desc_col:
                mask = (
                    data.account_txns[desc_col].astype(str).str.contains("HMRC", case=False, na=False) &
                    data.account_txns[desc_col].astype(str).str.contains("VAT",  case=False, na=False)
                )
                r.hmrc_payments = data.account_txns[mask].copy()
                if "Debit" in r.hmrc_payments.columns:
                    r.hmrc_total = float(
                        pd.to_numeric(r.hmrc_payments["Debit"], errors="coerce").sum()
                    )

        # Reconstructed closing balance:
        #   Opening (Cr) + Box1 (Cr output) - Box4 (Dr input) - HMRC payments (Dr)
        reconstructed = (
            cfg.opening_vat_balance
            + data.boxes.box1
            - data.boxes.box4
            - r.hmrc_total
        )
        r.vat_control_diff = round(reconstructed - r.vat_control_tb, 2)

    def _box6(self, cfg: JobConfig, data: ParsedData, r: ReconciliationResults):
        """
        Box 6 reconciliation:
        1. Compare Box 6 to QE sales transactions (direct match)
        2. Compare Box 6 to YTD TB sales (expected to differ — document the gap)
        3. Proof: VATable sales × 20% should ≈ Box 1
        """
        # QE sales from transactions sheet
        sec_key = "Box 6|20% (VAT on Income)"
        sec_zero = "Box 6|Zero Rated EC Goods Income"
        box6_20  = data.txn_sections.get(sec_key, pd.DataFrame())
        box6_z   = data.txn_sections.get(sec_zero, pd.DataFrame())

        r.box6_zero_net = float(
            pd.to_numeric(box6_z["Net"], errors="coerce").sum()
        ) if not box6_z.empty else 0.0

        # TB totals (YTD)
        for code in cfg.sales_nominals:
            r.box6_tb_sales += abs(self._tb_balance(data, code))
        for code in cfg.other_income_nominals:
            r.box6_tb_other += abs(self._tb_balance(data, code))
        r.box6_tb_total = r.box6_tb_sales + r.box6_tb_other

        # QE vs YTD difference (document, don't fail)
        r.box6_diff = round(data.boxes.box6 - r.box6_tb_total, 2)

        # VAT proof
        vatable = data.boxes.box6 - r.box6_zero_net
        r.expected_output_vat = round(vatable * 0.20, 2)
        r.vat_proof_diff = round(r.expected_output_vat - data.boxes.box1, 2)

    def _aged_payables(self, cfg: JobConfig, data: ParsedData, r: ReconciliationResults):
        if data.aged_pay_raw.empty:
            return
        inv = self._aged_invoice_rows(data.aged_pay_raw)
        r.ap_total   = float(inv["Total"].sum()) if "Total" in inv.columns else 0.0
        r.ap_overdue = self._aged_overdue(inv)
        r.ap_bs_diff = round(r.ap_total - abs(r.bs_creditors), 2)

    def _aged_receivables(self, cfg: JobConfig, data: ParsedData, r: ReconciliationResults):
        if data.aged_rec_raw.empty:
            return
        inv = self._aged_invoice_rows(data.aged_rec_raw)
        r.ar_total   = float(inv["Total"].sum()) if "Total" in inv.columns else 0.0
        r.ar_overdue = self._aged_overdue(inv)
        r.ar_bs_diff = round(r.ar_total - abs(r.bs_debtors), 2)

    def _top10_box4(self, data: ParsedData, r: ReconciliationResults):
        frames = [
            data.txn_sections.get(k, pd.DataFrame())
            for k in data.txn_sections
            if k.startswith("Box 4")
        ]
        if not frames:
            return
        all_b4 = pd.concat(frames, ignore_index=True)
        all_b4["_vat"] = pd.to_numeric(all_b4["VAT"], errors="coerce").fillna(0)
        r.top10_box4 = all_b4.nlargest(10, "_vat").copy()


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 4 — WORKBOOK BUILDER
# ══════════════════════════════════════════════════════════════════════════════

# ── style constants ───────────────────────────────────────────────────────────
_GREY  = "D9D9D9"; _BLUE = "BDD7EE"; _GREEN = "E2EFDA"
_RED   = "FCE4D6"; _WHITE = "FFFFFF"; _BLACK = "000000"; _DKRED = "FF0000"
_NUM   = "#,##0.00"


def _hdr(ws, row, col, val, bold=True, bg=_GREY, size=10, align="left", span=1):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(bold=bold, size=size, name="Arial")
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    return c


def _cel(ws, row, col, v, bold=False, num_fmt=None, align="right",
         color=None, bg=None, italic=False):
    c = ws.cell(row=row, column=col, value=v)
    c.font = Font(bold=bold, size=10, name="Arial", color=color or _BLACK, italic=italic)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if num_fmt:
        c.number_format = num_fmt
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    return c


def _flag(ws, row, col, value, tol=1.00):
    bg = _GREEN if abs(value) <= tol else _RED
    return _cel(ws, row, col, value, bold=True, num_fmt=_NUM, bg=bg)


def _note(ws, row, col, text):
    _cel(ws, row, col, f"◄ {text}", align="left", color=_DKRED, italic=True)


def _txn_headers(ws, row):
    for i, h in enumerate(["Date", "Account", "Reference", "Details", "VAT (£)", "Net (£)"], 1):
        _hdr(ws, row, i, h, bg=_GREY, align="center" if i > 4 else "left")


def _write_txns(ws, start_row: int, df: pd.DataFrame) -> int:
    r = start_row
    for _, row in df.iterrows():
        d = row.get("Date", "")
        if pd.notna(d):
            try:
                d = pd.to_datetime(d, dayfirst=True).strftime("%d/%m/%Y")
            except Exception:
                d = str(d)
        _cel(ws, r, 1, d, align="left")
        _cel(ws, r, 2, str(row.get("Account", "") or ""), align="left")
        _cel(ws, r, 3, str(row.get("Reference", "") or ""), align="left")
        _cel(ws, r, 4, str(row.get("Details", "") or ""), align="left")
        for ci, key in [(5, "VAT"), (6, "Net")]:
            v = row.get(key)
            if pd.notna(v) and v != "":
                try:
                    _cel(ws, r, ci, float(v), num_fmt=_NUM)
                except (TypeError, ValueError):
                    pass
        r += 1
    return r


def _wp_header(ws, ref: str, detail: str, cfg: JobConfig, period: str):
    _cel(ws, 1, 1, f"Client: {cfg.client_name}", bold=True, align="left")
    _cel(ws, 2, 1, f"Period ended: {period}", bold=True, align="left")
    _cel(ws, 3, 1, f"Detail: {detail}", align="left")
    _cel(ws, 1, 6, f"ref: {ref}", align="right")
    _cel(ws, 2, 6, f"Prepared by: {cfg.preparer}", align="right")
    _cel(ws, 3, 6, f"Reviewed by: {cfg.reviewer}", align="right")


def _set_widths(ws, widths: List[int]):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


class WorkbookBuilder:
    """Builds the formatted Excel working papers file."""

    def _build_wb(
        self,
        cfg: JobConfig,
        validation: ValidationResult,
        data: ParsedData,
        recs: ReconciliationResults,
    ) -> Workbook:
        wb = Workbook()

        ws_list = [
            ("0. File Register",          self._sheet_file_register),
            ("1A. VAT Return",            self._sheet_vat_return),
            ("1B. Txns by VAT Box",       self._sheet_txn_by_box),
            ("2A. VAT Control",           self._sheet_vat_control),
            ("2B. Bank Rec",              self._sheet_bank_rec),
            ("2C. Box 6 Rec",             self._sheet_box6_rec),
            ("3A. Aged Payables",         self._sheet_aged_payables),
            ("3B. Aged Receivables",      self._sheet_aged_receivables),
            ("3C. Trial Balance",         self._sheet_trial_balance),
            ("4A. Top 10 Box 4",          self._sheet_top10),
            ("VAT Checklist",             self._sheet_checklist),
        ]

        first = True
        for title, builder_fn in ws_list:
            ws = wb.active if first else wb.create_sheet(title)
            if first:
                ws.title = title
                first = False
            builder_fn(ws, cfg, validation, data, recs)

        return wb

    def build(
        self,
        cfg: JobConfig,
        validation: ValidationResult,
        data: ParsedData,
        recs: ReconciliationResults,
        output_path: Path,
    ) -> Path:
        wb = self._build_wb(cfg, validation, data, recs)
        wb.save(output_path)
        log.info("Workbook saved: %s", output_path)
        return output_path

    def build_to_bytes(
        self,
        cfg: JobConfig,
        validation: ValidationResult,
        data: ParsedData,
        recs: ReconciliationResults,
    ) -> bytes:
        import io
        wb = self._build_wb(cfg, validation, data, recs)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()

    # ── Sheet 0: File Register ────────────────────────────────────────────────

    def _sheet_file_register(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [25, 55, 20, 10])
        _hdr(ws, 1, 1, "VAT WORKING PAPERS — FILE REGISTER", size=13, bg=_GREY, span=4, align="center")
        _cel(ws, 2, 1, f"Client:           {cfg.client_name}", bold=True, align="left")
        _cel(ws, 3, 1, f"Period:           {data.boxes.period_str}", bold=True, align="left")
        _cel(ws, 4, 1, f"Expected QE date: {cfg.qe_end_date}", bold=True, align="left")
        _cel(ws, 5, 1, f"Generated:        {datetime.now().strftime('%d %b %Y %H:%M')}", align="left")

        _hdr(ws, 7, 1, "Report");        _hdr(ws, 7, 2, "Filename")
        _hdr(ws, 7, 3, "Period in file", align="center")
        _hdr(ws, 7, 4, "Match?", align="center")

        qe = _norm_date(cfg.qe_end_date)
        file_labels = [
            ("vat_return_file",    "VAT Return"),
            ("trial_balance_file", "Trial Balance"),
            ("account_txn_file",   "Account Transactions"),
            ("aged_pay_file",      "Aged Payables"),
            ("aged_rec_file",      "Aged Receivables"),
            ("balance_sheet_file", "Balance Sheet"),
        ]
        for i, (attr, label) in enumerate(file_labels, 8):
            fname = getattr(cfg, attr, None)
            _cel(ws, i, 1, label, bold=True, align="left")
            _cel(ws, i, 2, fname or "Not configured", align="left")
            d = validation.file_dates.get(label, "—")
            _cel(ws, i, 3, d, align="center")
            ok = d == qe
            _cel(ws, i, 4, "✓" if ok else "✗", align="center", bold=True,
                 bg=_GREEN if ok else _RED)

        if validation.warnings:
            r = 16
            _hdr(ws, r, 1, "Warnings / Notices", bg=_GREY, span=4); r += 1
            for w in validation.warnings:
                bg = _RED if w.severity == "error" else _GREY if w.severity == "info" else "FCE4D6"
                _cel(ws, r, 1, w.severity.upper(), bold=True, align="center", bg=bg)
                _cel(ws, r, 2, w.section, align="left")
                _cel(ws, r, 3, w.message, align="left")
                ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
                r += 1

    # ── Sheet 1A: VAT Return ──────────────────────────────────────────────────

    def _sheet_vat_return(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [68, 8, 14, 0, 0, 28])
        b = data.boxes
        _wp_header(ws, "1A.", "1A. VAT Return", cfg, b.period_end)
        _hdr(ws, 7, 1, "VAT Return", size=13, bg=_WHITE, span=3)
        _cel(ws, 9, 1, b.period_str, bold=True, align="left")
        _hdr(ws, 11, 1, "VAT Return Details", bg=_GREY, span=3)
        for i, (k, v) in enumerate([("VAT Number", b.vat_number), ("VAT Scheme", b.vat_scheme),
                                     ("Period Start", b.period_start), ("Period End", b.period_end)], 12):
            _cel(ws, i, 1, k, align="left"); _cel(ws, i, 3, v, align="left")
        _hdr(ws, 17, 1, "VAT Calculations", bg=_GREY, span=3)
        for i, (n, desc, amt) in enumerate([
            (1, "VAT due on sales and other outputs", b.box1),
            (2, "VAT due on EU acquisitions (NI)", b.box2),
            (3, "Total VAT due (Box 1 + Box 2)", b.box3),
            (4, "VAT reclaimed on purchases and other inputs", b.box4),
            (5, "VAT to Pay HMRC", b.box5),
        ], 18):
            _cel(ws, i, 1, desc, align="left"); _cel(ws, i, 2, n, align="center")
            _cel(ws, i, 3, amt, bold=True, num_fmt=_NUM,
                 bg=_RED if n == 5 and amt > 0 else _GREEN if n == 5 else None)
        _hdr(ws, 24, 1, "Sales and Purchases Excluding VAT", bg=_GREY, span=3)
        for i, (n, desc, amt) in enumerate([
            (6, "Total net sales", b.box6), (7, "Total net purchases", b.box7),
            (8, "EU supplies (NI)", b.box8), (9, "EU acquisitions (NI)", b.box9),
        ], 25):
            _cel(ws, i, 1, desc, align="left"); _cel(ws, i, 2, n, align="center")
            _cel(ws, i, 3, amt, bold=True, num_fmt=_NUM)

    # ── Sheet 1B: Transactions by VAT Box ─────────────────────────────────────

    def _sheet_txn_by_box(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [14, 28, 20, 42, 12, 12])
        b = data.boxes
        _wp_header(ws, "1B.", "1B. Transactions by VAT Box", cfg, b.period_end)
        _hdr(ws, 7, 1, "Transactions by VAT Box", size=12, bg=_WHITE)
        _cel(ws, 8, 1, b.period_str, align="left")
        r = 10

        def _get(box, sub): return data.txn_sections.get(f"{box}|{sub}", pd.DataFrame())

        # Box 8/9 sub-labels vary by client (whatever Xero grouped under
        # that box), unlike 1/4/6/7 above where the sub-labels are known
        # and stable -- so pick up whatever TxnByBoxParser actually parsed
        # for that box rather than a hard-coded, possibly-wrong label.
        def _subs(box):
            prefix = f"{box}|"
            return [(k[len(prefix):], df) for k, df in data.txn_sections.items()
                    if k.startswith(prefix)]

        for box_lbl, desc, total, bg, subs in [
            ("Box 1", "VAT due on sales", b.box1, _BLUE,
             [("20% (VAT on Income)", _get("Box 1", "20% (VAT on Income)"))]),
            ("Box 4", "VAT reclaimed on purchases", b.box4, _GREEN,
             [("20% (VAT on Expenses)",          _get("Box 4", "20% (VAT on Expenses)")),
              ("20% (VAT on Expenses) - Adjusted", _get("Box 4", "20% (VAT on Expenses) - Adjusted")),
              ("5% (VAT on Expenses)",            _get("Box 4", "5% (VAT on Expenses)"))]),
            ("Box 6", "Net sales excluding VAT", b.box6, _BLUE,
             [("20% (VAT on Income)",       _get("Box 6", "20% (VAT on Income)")),
              ("Zero Rated EC Goods Income", _get("Box 6", "Zero Rated EC Goods Income"))]),
            ("Box 7", "Net purchases excluding VAT", b.box7, _GREEN,
             [("20% (VAT on Expenses)",          _get("Box 7", "20% (VAT on Expenses)")),
              ("20% (VAT on Expenses) - Adjusted", _get("Box 7", "20% (VAT on Expenses) - Adjusted")),
              ("5% (VAT on Expenses)",            _get("Box 7", "5% (VAT on Expenses)"))]),
            ("Box 8", "EU supplies (NI)", b.box8, _BLUE, _subs("Box 8")),
            ("Box 9", "EU acquisitions (NI)", b.box9, _GREEN, _subs("Box 9")),
        ]:
            _hdr(ws, r, 1, box_lbl, bg=bg, size=11, span=2)
            _cel(ws, r, 3, desc, bold=True, align="left")
            _cel(ws, r, 6, total, bold=True, num_fmt=_NUM); r += 1
            for sub_lbl, df in subs:
                if df.empty: continue
                _hdr(ws, r, 1, sub_lbl, bg=_GREY, span=6); r += 1
                _txn_headers(ws, r); r += 1
                r = _write_txns(ws, r, df); r += 1
            r += 1

    # ── Sheet 2A: VAT Control ─────────────────────────────────────────────────

    def _sheet_vat_control(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [18, 42, 14, 14, 32])
        b = data.boxes
        _wp_header(ws, "2A.", "2A. VAT Control", cfg, b.period_end)
        _hdr(ws, 7, 1, "VAT CONTROL", size=12, bg=_WHITE)
        _hdr(ws, 8, 1, f"VAT BASIS: {b.vat_scheme.upper()}", bg=_WHITE)

        _hdr(ws, 10, 1, "Opening Balance", bg=_GREY)
        _hdr(ws, 10, 3, "Dr", bg=_GREY, align="center")
        _hdr(ws, 10, 4, "Cr", bg=_GREY, align="center")
        _hdr(ws, 10, 5, "Comments", bg=_GREY)
        _cel(ws, 12, 2, "Bal b/fwd (prior period TB)", bold=True, align="left")
        _cel(ws, 12, 4, cfg.opening_vat_balance, num_fmt=_NUM)
        _note(ws, 12, 5, "Update with prior period closing VAT control balance")

        r = 14
        _hdr(ws, r, 1, "Output VAT (Box 1)", bg=_BLUE, span=2); r += 1
        _cel(ws, r, 4, "Cr £", bold=True, align="center"); r += 1
        _cel(ws, r, 2, f"QE {b.period_end} — per VAT return", align="left")
        _cel(ws, r, 4, b.box1, num_fmt=_NUM); r += 2

        _hdr(ws, r, 1, "Input VAT (Box 4)", bg=_GREEN, span=2); r += 1
        _cel(ws, r, 3, "Dr £", bold=True, align="center"); r += 1
        _cel(ws, r, 2, f"QE {b.period_end} — per VAT return", align="left")
        _cel(ws, r, 3, b.box4, num_fmt=_NUM); r += 2

        _hdr(ws, r, 1, "HMRC Payments (from Account Transactions)", bg=_GREY, span=4); r += 1
        _hdr(ws, r, 1, "Date", bg=_GREY, align="center")
        _hdr(ws, r, 2, "Description", bg=_GREY)
        _hdr(ws, r, 3, "Dr £", bg=_GREY, align="center"); r += 1
        if not recs.hmrc_payments.empty:
            desc_col = next((c for c in recs.hmrc_payments.columns
                             if "desc" in c.lower() or "detail" in c.lower()), None)
            for _, pr in recs.hmrc_payments.iterrows():
                try: d = pd.to_datetime(pr.get("Date", ""), dayfirst=True).strftime("%d/%m/%Y")
                except: d = str(pr.get("Date", ""))
                _cel(ws, r, 1, d, align="center")
                _cel(ws, r, 2, str(pr.get(desc_col, "")) if desc_col else "", align="left")
                try: _cel(ws, r, 3, float(pr.get("Debit", 0) or 0), num_fmt=_NUM)
                except: pass
                r += 1
        else:
            _cel(ws, r, 2, "No HMRC VAT payments identified in Account Transactions",
                 align="left", italic=True)
            _note(ws, r, 4, "Enter manually if description differs")
            r += 1

        r += 1
        _cel(ws, r, 2, "Total HMRC Payments", bold=True, align="left")
        _cel(ws, r, 3, recs.hmrc_total, bold=True, num_fmt=_NUM); r += 2

        _hdr(ws, r, 1, "Closing Balance & Reconciliation", bg=_GREY, span=5); r += 1
        _cel(ws, r, 2, "VAT to pay — Box 5", align="left"); _cel(ws, r, 3, b.box5, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, f"VAT Control per TB (nominal {cfg.vat_control_nominal})", align="left")
        _cel(ws, r, 3, recs.vat_control_tb, num_fmt=_NUM)
        if not recs.vat_control_tb:
            _note(ws, r, 5, f"Nominal {cfg.vat_control_nominal} not in TB — enter manually")
        r += 1
        _cel(ws, r, 2, "Difference", bold=True, align="left")
        _flag(ws, r, 3, recs.vat_control_diff, tol=cfg.tol_general)
        _note(ws, r, 5, "Green ≤ £1 | Red = investigate")

    # ── Sheet 2B: Bank Rec ────────────────────────────────────────────────────

    def _sheet_bank_rec(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [20, 38, 16, 32])
        _wp_header(ws, "2B.", "2B. Bank Rec", cfg, data.boxes.period_end)
        _hdr(ws, 7, 1, "BANK REC", size=12, bg=_WHITE)
        _hdr(ws, 11, 2, "Bank Account", bg=_GREY)
        _hdr(ws, 11, 3, "Business Bank Account", bg=_GREY, align="center")
        for i, (lbl, amt) in enumerate([
            ("Balance per TB", recs.bs_bank),
            ("Plus: unreconciled items", 0.0),
            ("Less: outstanding payments", 0.0),
            ("Balance per Statement", recs.bs_bank),
            ("Difference", 0.0),
        ], 12):
            bold = lbl in ("Balance per Statement", "Difference")
            _cel(ws, i, 2, lbl, bold=bold, align="left")
            bg = _GREEN if lbl == "Difference" else None
            _cel(ws, i, 3, amt, bold=bold, num_fmt=_NUM, bg=bg)
        _note(ws, 17, 4, "Attach bank statement / Xero bank rec screenshot")

    # ── Sheet 2C: Box 6 Rec ───────────────────────────────────────────────────

    def _sheet_box6_rec(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [25, 35, 14, 14, 35])
        b = data.boxes
        _wp_header(ws, "2C.", "2C. Box 6 Rec", cfg, b.period_end)
        _hdr(ws, 7, 1, "BOX 6 REC", size=12, bg=_WHITE)
        r = 9

        _hdr(ws, r, 1, "Turnover as per Box 6 (VAT return)", bg=_BLUE, span=4); r += 1
        _cel(ws, r, 1, f"QE {b.period_end}", align="left"); _cel(ws, r, 3, b.box6, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Add: Late claims", align="left"); _cel(ws, r, 3, 0.0, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Adjusted Box 6", bold=True, align="left")
        _cel(ws, r, 3, b.box6, bold=True, num_fmt=_NUM); r += 2

        _hdr(ws, r, 1, "Turnover as per TB (YTD)", bg=_GREEN, span=4); r += 1
        _cel(ws, r, 2, f"Sales ({', '.join(str(c) for c in cfg.sales_nominals)}) — per TB",
             align="left")
        _cel(ws, r, 3, recs.box6_tb_sales, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, f"Other Income ({', '.join(str(c) for c in cfg.other_income_nominals)}) — per TB",
             align="left")
        _cel(ws, r, 3, recs.box6_tb_other, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Total per TB", bold=True, align="left")
        _cel(ws, r, 3, recs.box6_tb_total, bold=True, num_fmt=_NUM); r += 2

        _hdr(ws, r, 1, "Reconciliation", bg=_GREY, span=4); r += 1
        _cel(ws, r, 2, "Box 6 (QE VAT return)", align="left"); _cel(ws, r, 3, b.box6, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Cumulative TB turnover (YTD)", align="left")
        _cel(ws, r, 3, recs.box6_tb_total, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Difference", bold=True, align="left")
        _flag(ws, r, 3, recs.box6_diff, tol=cfg.tol_box6)
        _note(ws, r, 5, "Difference expected if TB is YTD vs QE — review and explain"); r += 3

        _hdr(ws, r, 1, "Proof of Output VAT (Box 6 × 20% = Box 1)", bg=_BLUE, span=4); r += 1
        _cel(ws, r, 2, "Box 6 (net sales)", align="left"); _cel(ws, r, 3, b.box6, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Less: zero rated / EC sales", align="left")
        _cel(ws, r, 3, recs.box6_zero_net, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "VATable sales (20%)", bold=True, align="left")
        _cel(ws, r, 3, b.box6 - recs.box6_zero_net, bold=True, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Expected output VAT (×20%)", align="left")
        _cel(ws, r, 3, recs.expected_output_vat, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Actual Box 1", align="left"); _cel(ws, r, 3, b.box1, num_fmt=_NUM); r += 1
        _cel(ws, r, 2, "Difference", bold=True, align="left")
        _flag(ws, r, 3, recs.vat_proof_diff, tol=cfg.tol_general)

    # ── Sheet 3A: Aged Payables ───────────────────────────────────────────────

    def _sheet_aged_payables(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [16, 16, 22, 12, 12, 12, 12, 12, 12, 12])
        _wp_header(ws, "3A.", "3A. Aged Payables", cfg, data.boxes.period_end)
        _hdr(ws, 7, 1, "AGED PAYABLES", size=12, bg=_WHITE)
        _cel(ws, 8, 1,
             f"Total outstanding: £{recs.ap_total:,.2f}   Overdue: £{recs.ap_overdue:,.2f}",
             bold=True, align="left")
        _cel(ws, 9, 1, "Agree to Balance Sheet — Accounts Payable:", bold=True, align="left")
        _cel(ws, 9, 4, abs(recs.bs_creditors), num_fmt=_NUM)
        _cel(ws, 9, 5, recs.ap_total, num_fmt=_NUM)
        _flag(ws, 9, 6, recs.ap_bs_diff, tol=cfg.tol_general)
        _note(ws, 9, 7, "BS vs Aged Payables — should be nil")

        r = 11
        for i, h in enumerate(["Invoice Date", "Due Date", "Reference", "Current",
                                "< 1 Month", "1 Month", "2 Months", "3 Months", "Older", "Total"], 1):
            _hdr(ws, r, i, h, bg=_GREY, align="center" if i > 3 else "left")
        r += 1
        for _, row in data.aged_pay_raw.iterrows():
            v0 = str(row.iloc[0]).strip()
            if re.match(r"\d{4}-\d{2}-\d{2}", v0):
                try: inv_d = pd.to_datetime(v0).strftime("%d/%m/%Y")
                except: inv_d = v0
                try: due_d = pd.to_datetime(str(row.iloc[1])).strftime("%d/%m/%Y")
                except: due_d = str(row.iloc[1])
                _cel(ws, r, 1, inv_d, align="left"); _cel(ws, r, 2, due_d, align="left")
                _cel(ws, r, 3, str(row.iloc[2]), align="left")
                for ci, ci2 in enumerate(range(3, 10), 4):
                    try:
                        v = float(row.iloc[ci2])
                        if v != 0: _cel(ws, r, ci, v, num_fmt=_NUM)
                    except: pass
                r += 1
            elif (not v0.startswith("Total") and v0 not in ("", "nan", "Invoice Date") and
                  all(pd.isna(row.iloc[i]) or str(row.iloc[i]).strip() in ("", "nan")
                      for i in range(1, 5))):
                _hdr(ws, r, 1, v0, bg=_GREY, span=10); r += 1

    # ── Sheet 3B: Aged Receivables ────────────────────────────────────────────

    def _sheet_aged_receivables(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [16, 16, 16, 22, 12, 12, 12, 12, 12, 12, 12])
        _wp_header(ws, "3B.", "3B. Aged Receivables", cfg, data.boxes.period_end)
        _hdr(ws, 7, 1, "AGED RECEIVABLES", size=12, bg=_WHITE)
        _cel(ws, 8, 1,
             f"Total outstanding: £{recs.ar_total:,.2f}   Overdue: £{recs.ar_overdue:,.2f}",
             bold=True, align="left")
        _cel(ws, 9, 1, "Agree to Balance Sheet — Accounts Receivable:", bold=True, align="left")
        _cel(ws, 9, 4, recs.bs_debtors, num_fmt=_NUM)
        _cel(ws, 9, 5, recs.ar_total, num_fmt=_NUM)
        _flag(ws, 9, 6, recs.ar_bs_diff, tol=cfg.tol_general)
        _note(ws, 9, 7, "BS vs Aged Receivables — should be nil")

        r = 11
        for i, h in enumerate(["Invoice Date", "Due Date", "Invoice Number", "Reference",
                                "Current", "< 1 Month", "1 Month", "2 Months",
                                "3 Months", "Older", "Total"], 1):
            _hdr(ws, r, i, h, bg=_GREY, align="center" if i > 4 else "left")
        r += 1
        for _, row in data.aged_rec_raw.iterrows():
            v0 = str(row.iloc[0]).strip()
            if re.match(r"\d{4}-\d{2}-\d{2}", v0):
                try: inv_d = pd.to_datetime(v0).strftime("%d/%m/%Y")
                except: inv_d = v0
                try: due_d = pd.to_datetime(str(row.iloc[1])).strftime("%d/%m/%Y")
                except: due_d = str(row.iloc[1])
                _cel(ws, r, 1, inv_d, align="left"); _cel(ws, r, 2, due_d, align="left")
                _cel(ws, r, 3, str(row.iloc[2]), align="left")
                _cel(ws, r, 4, str(row.iloc[3]), align="left")
                for ci, ci2 in enumerate(range(4, 11), 5):
                    try:
                        v = float(row.iloc[ci2])
                        if v != 0: _cel(ws, r, ci, v, num_fmt=_NUM)
                    except: pass
                r += 1
            elif (not v0.startswith("Total") and v0 not in ("", "nan", "Invoice Date") and
                  all(pd.isna(row.iloc[i]) or str(row.iloc[i]).strip() in ("", "nan")
                      for i in range(1, 5))):
                _hdr(ws, r, 1, v0, bg=_GREY, span=11); r += 1

    # ── Sheet 3C: Trial Balance ───────────────────────────────────────────────

    def _sheet_trial_balance(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [14, 35, 18, 16, 16, 16])
        _wp_header(ws, "3C.", "3C. Trial Balance", cfg, data.boxes.period_end)
        _hdr(ws, 7, 1, "TRIAL BALANCE", size=12, bg=_WHITE)
        _cel(ws, 8, 1, f"As at {data.boxes.period_end}", bold=True, align="left")
        r = 10
        for i, h in enumerate(["Account Code", "Account", "Account Type",
                                "Debit YTD", "Credit YTD", "Prior Year"], 1):
            _hdr(ws, r, i, h, bg=_GREY, align="center" if i > 3 else "left")
        r += 1
        for _, row in data.trial_balance.iterrows():
            code = int(row["Account Code"])
            _cel(ws, r, 1, code, align="center")
            _cel(ws, r, 2, str(row["Account"]), align="left")
            _cel(ws, r, 3, str(row.get("Account Type", "")), align="left")
            dr = row.get("Debit - Year to date")
            cr = row.get("Credit - Year to date")
            py = row.iloc[5] if len(row) > 5 else None
            if pd.notna(dr) and dr: _cel(ws, r, 4, float(dr), num_fmt=_NUM)
            if pd.notna(cr) and cr: _cel(ws, r, 5, float(cr), num_fmt=_NUM)
            if pd.notna(py) and py:
                try: _cel(ws, r, 6, float(py), num_fmt=_NUM)
                except: pass
            if code == cfg.vat_control_nominal:
                for ci in range(1, 7):
                    ws.cell(row=r, column=ci).fill = PatternFill("solid", fgColor=_BLUE)
                    ws.cell(row=r, column=ci).font = Font(bold=True, name="Arial", size=10)
            r += 1

    # ── Sheet 4A: Top 10 Box 4 ────────────────────────────────────────────────

    def _sheet_top10(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [14, 28, 20, 42, 12, 12])
        _wp_header(ws, "4A.", "Top 10 Box 4 Transactions", cfg, data.boxes.period_end)
        _hdr(ws, 7, 1, "TOP 10 BOX 4 (INPUT VAT) — CHECK VAT INVOICES ON XERO",
             size=11, bg=_GREEN, span=6)
        r = 9
        _txn_headers(ws, r); r += 1
        r = _write_txns(ws, r, recs.top10_box4); r += 1
        _cel(ws, r, 4, "TOTAL", bold=True, align="right")
        _cel(ws, r, 5, float(recs.top10_box4["_vat"].sum()) if not recs.top10_box4.empty else 0.0,
             bold=True, num_fmt=_NUM, bg=_GREEN)

    # ── Sheet: VAT Checklist ──────────────────────────────────────────────────

    def _sheet_checklist(self, ws, cfg, validation, data, recs):
        _set_widths(ws, [10, 90, 8])
        b = data.boxes
        _hdr(ws, 1, 1, "VAT CHECKLIST", size=14, bg=_GREY, span=3, align="center")
        for i, (k, v) in enumerate([
            ("Client:", cfg.client_name), ("Period:", b.period_end),
            ("VAT Scheme:", b.vat_scheme), ("VAT no:", b.vat_number),
            ("Preparer:", cfg.preparer), ("Date:", ""),
        ], 3):
            _cel(ws, i, 1, k, bold=True, align="left"); _cel(ws, i, 2, v, align="left")

        items = [
            ("1",  "Prior period checklist and review points cleared"),
            ("2",  "Working papers completed"),
            ("3",  "Countries of zero rated sales confirmed"),
            ("4a", "VAT Return report on file"),
            ("4b", "VAT Detailed report (Transactions by VAT Box) on file"),
            ("4c", "VAT Control account reconciled — see 2A"),
            ("4d", "Aged Debtors reviewed, annotated, agreed to BS — see 3B"),
            ("4e", "Aged Creditors reviewed, annotated, agreed to BS — see 3A"),
            ("4f", "Bank reconciliations complete and statements on file — see 2B"),
            ("4g", "Trial balance referenced to workings — see 3C"),
            ("4h", "Nominal / General ledger reviewed for consistency"),
            ("6",  "Sales invoice numbers confirmed sequential"),
            ("7",  "Box 6 reconciled to TB cumulative sales — see 2C"),
            ("8",  "Proof of output VAT: Box 6 × 20% = Box 1 — see 2C"),
            ("9",  "Top 10 Box 4 transactions checked — VAT invoices on Xero — see 4A"),
            ("10", "Suspense account items scheduled"),
            ("11", "Items with No VAT analysed"),
            ("12", "Payroll journals posted"),
            ("13", "DLA breakdown sent to client if required"),
            ("14", "PAYE payments up to date"),
        ]
        for i, (ref, desc) in enumerate(items, 10):
            _cel(ws, i, 1, ref, align="center")
            _cel(ws, i, 2, desc, align="left")
            _cel(ws, i, 3, "", align="center")


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 4B — ANNUAL VAT RECONCILIATION
# ══════════════════════════════════════════════════════════════════════════════

class AnnualSummaryBuilder:
    """
    Builds a Financial Year VAT reconciliation workbook from a set of
    QuarterSnapshot records (one per quarter, produced by
    VATWorkflowService.run_job and saved as JSON sidecar files).

    Carries the VAT control balance forward quarter-to-quarter and rolls
    up Box 1-9 into FY totals -- i.e. the same VAT reconciliation done
    each quarter, accumulated for the whole financial year.
    """

    _BOX_COLS = [
        ("box1", "Box 1"), ("box2", "Box 2"), ("box3", "Box 3"),
        ("box4", "Box 4"), ("box5", "Box 5"), ("box6", "Box 6"),
        ("box7", "Box 7"), ("box8", "Box 8"), ("box9", "Box 9"),
    ]

    def build_to_bytes(self, client_name: str, fy_label: str,
                        quarters: List[QuarterSnapshot]) -> bytes:
        import io
        wb = self._build_wb(client_name, fy_label, quarters)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()

    def build(self, client_name: str, fy_label: str,
              quarters: List[QuarterSnapshot], output_path: Path) -> Path:
        wb = self._build_wb(client_name, fy_label, quarters)
        wb.save(output_path)
        log.info("Annual summary saved: %s", output_path)
        return output_path

    def _build_wb(self, client_name: str, fy_label: str,
                  quarters: List[QuarterSnapshot]) -> Workbook:
        wb = Workbook()
        ws = wb.active
        ws.title = "FY VAT Summary"
        _set_widths(ws, [16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16])

        _hdr(ws, 1, 1, "ANNUAL VAT RECONCILIATION", size=13, bg=_GREY, span=11, align="center")
        _cel(ws, 2, 1, f"Client: {client_name}", bold=True, align="left")
        _cel(ws, 3, 1, f"Financial Year: {fy_label}", bold=True, align="left")
        _cel(ws, 4, 1, f"Quarters included: {len(quarters)}", align="left")

        # ── Table 1: VAT Return boxes per quarter ───────────────────────────
        r = 6
        _hdr(ws, r, 1, "VAT RETURN BOXES BY QUARTER", bg=_BLUE, span=11); r += 1
        _hdr(ws, r, 1, "Quarter", align="left")
        _hdr(ws, r, 2, "Period End", align="left")
        for i, (_, label) in enumerate(self._BOX_COLS, 3):
            _hdr(ws, r, i, label, align="center")
        r += 1

        box_totals = {key: 0.0 for key, _ in self._BOX_COLS}
        for i, q in enumerate(quarters, 1):
            _cel(ws, r, 1, f"Q{i}", bold=True, align="left")
            _cel(ws, r, 2, q.period_end, align="left")
            for ci, (key, _) in enumerate(self._BOX_COLS, 3):
                v = getattr(q, key)
                box_totals[key] += v
                _cel(ws, r, ci, v, num_fmt=_NUM)
            r += 1

        _cel(ws, r, 1, "FY TOTAL", bold=True, align="left")
        _cel(ws, r, 2, fy_label, bold=True, align="left")
        for ci, (key, _) in enumerate(self._BOX_COLS, 3):
            _cel(ws, r, ci, round(box_totals[key], 2), bold=True, num_fmt=_NUM, bg=_GREEN)
        r += 3

        # ── Table 2: VAT control roll-forward ───────────────────────────────
        _hdr(ws, r, 1, "VAT CONTROL ROLL-FORWARD", bg=_BLUE, span=8); r += 1
        for ci, h in enumerate([
            "Quarter", "Period End", "Opening Bal", "+ Box 1 (output)",
            "- Box 4 (input)", "- HMRC paid", "= Closing Bal", "TB Bal (YTD)", "Diff",
        ], 1):
            _hdr(ws, r, ci, h, align="center" if ci > 2 else "left")
        r += 1

        prev_closing: Optional[float] = None
        for i, q in enumerate(quarters, 1):
            _cel(ws, r, 1, f"Q{i}", bold=True, align="left")
            _cel(ws, r, 2, q.period_end, align="left")
            _cel(ws, r, 3, q.opening_vat_balance, num_fmt=_NUM)
            _cel(ws, r, 4, q.box1, num_fmt=_NUM)
            _cel(ws, r, 5, q.box4, num_fmt=_NUM)
            _cel(ws, r, 6, q.hmrc_total, num_fmt=_NUM)
            _cel(ws, r, 7, q.closing_vat_balance, bold=True, num_fmt=_NUM)
            _cel(ws, r, 8, q.vat_control_tb, num_fmt=_NUM)
            _flag(ws, r, 9, q.vat_control_diff)

            if prev_closing is not None and abs(q.opening_vat_balance - prev_closing) > 1.00:
                _note(ws, r, 10,
                      f"Opening bal does not match Q{i-1} closing bal "
                      f"(£{prev_closing:,.2f}) — check carry-forward")
            prev_closing = q.closing_vat_balance
            r += 1

        r += 1
        if quarters:
            _cel(ws, r, 1, "FY closing VAT control balance", bold=True, align="left")
            _cel(ws, r, 7, quarters[-1].closing_vat_balance, bold=True, num_fmt=_NUM, bg=_GREEN)
            _cel(ws, r, 8, quarters[-1].vat_control_tb, bold=True, num_fmt=_NUM)
            _note(ws, r, 10, "FY closing balance should agree to year-end TB nominal (VAT control)")

        return wb


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 5 — SERVICE  (the layer a future app calls)
# ══════════════════════════════════════════════════════════════════════════════

class VATWorkflowService:
    """
    Single public entry point for the full pipeline.

    Future app integration:
        from vat_engine import VATWorkflowService, JobConfig
        cfg = JobConfig(
            client_name="Acme Ltd",
            qe_end_date="31 May 2026",
            vat_return_file=uploaded_paths["vat_return"],
            ...
        )
        result = VATWorkflowService().run_job(cfg)
        # result.summary  → for dashboard / API response
        # result.workbook_path  → for download link
        # result.recs  → for review cards / exception list in UI
    """

    def run_job(self, cfg: JobConfig) -> JobResult:
        log.info("=" * 60)
        log.info("  VAT ENGINE v3  |  %s  |  %s", cfg.client_name, cfg.qe_end_date)
        log.info("=" * 60)

        # 1. Validate
        validation = InputValidator().validate(cfg)
        self._print_validation(validation, cfg)
        if not validation.passed:
            raise InputError(
                "Validation failed — one or more required files are missing or "
                "have mismatched dates. See warnings above."
            )

        # 2. Parse
        data = ParseManager().parse_all(cfg)

        # 3. Reconcile
        recs = ReconciliationEngine().run(cfg, data)

        # 4. Build workbook
        safe_c = re.sub(r"[^\w\s-]", "", cfg.client_name).strip().replace(" ", "_")
        safe_p = cfg.qe_end_date.replace(" ", "_")
        out_path = Path(cfg.working_dir) / f"VAT_Working_Papers_{safe_c}_{safe_p}.xlsx"
        WorkbookBuilder().build(cfg, validation, data, recs, out_path)

        # 4b. Quarter snapshot — carry-forward into next quarter / annual summary
        snapshot = QuarterSnapshot.build(cfg, data.boxes, recs)
        closing_vat_balance = snapshot.closing_vat_balance
        snapshot_path = Path(cfg.working_dir) / f"VAT_Snapshot_{safe_c}_{safe_p}.json"
        snapshot_path.write_text(snapshot.to_json())

        # 5. Summary
        summary = {
            "client":           cfg.client_name,
            "period":           data.boxes.period_end,
            "vat_scheme":       data.boxes.vat_scheme,
            "box1":             data.boxes.box1,
            "box4":             data.boxes.box4,
            "box5":             data.boxes.box5,
            "box6":             data.boxes.box6,
            "vat_proof_diff":   recs.vat_proof_diff,
            "vat_control_diff": recs.vat_control_diff,
            "box6_diff":        recs.box6_diff,
            "ap_total":         recs.ap_total,
            "ap_bs_diff":       recs.ap_bs_diff,
            "ar_total":         recs.ar_total,
            "ar_bs_diff":       recs.ar_bs_diff,
            "validation_passed": validation.passed,
            "warnings":         len(validation.warnings),
            "workbook":         str(out_path),
            "snapshot":         asdict(snapshot),
            "snapshot_path":    str(snapshot_path),
            "closing_vat_balance": closing_vat_balance,
        }

        log.info("\n  ✓  Done — %s", out_path.name)
        self._print_summary(summary, cfg.tol_general)

        return JobResult(
            client_name=cfg.client_name,
            period=data.boxes.period_end,
            validation=validation,
            parsed=data,
            recs=recs,
            workbook_path=str(out_path),
            summary=summary,
        )

    @staticmethod
    def _print_validation(v: ValidationResult, cfg: JobConfig):
        qe = _norm_date(cfg.qe_end_date)
        print(f"\n  Expected QE date: {qe}\n")
        for label, date in v.file_dates.items():
            ok = date == qe
            print(f"  {'✓' if ok else '✗'} {label:<26}  {date}")
        if v.warnings:
            print()
            for w in v.warnings:
                print(f"  [{w.severity.upper()}] {w.message}")

    @staticmethod
    def _prompt_continue() -> bool:
        resp = input("\n  ⚠  Issues detected. Continue anyway? (y/n): ").strip().lower()
        return resp == "y"

    @staticmethod
    def _print_summary(s: dict, tol_general: float = 1.00):
        print("\n  Key figures:")
        print(f"    Box 1 (Output VAT):        £{s['box1']:>10,.2f}")
        print(f"    Box 4 (Input VAT):         £{s['box4']:>10,.2f}")
        print(f"    Box 5 (VAT to pay):        £{s['box5']:>10,.2f}")
        print(f"    Box 6 (Net sales):         £{s['box6']:>10,.2f}")
        print(f"    VAT proof diff (exp-act):  £{s['vat_proof_diff']:>10,.2f}  {'✓' if abs(s['vat_proof_diff'])<=tol_general else '✗'}")
        print(f"    VAT control diff:          £{s['vat_control_diff']:>10,.2f}  {'✓' if abs(s['vat_control_diff'])<=tol_general else '✗'}")
        print(f"    Box 6 vs TB diff:          £{s['box6_diff']:>10,.2f}  (YTD vs QE — expected)")
        print(f"    Aged Payables total:       £{s['ap_total']:>10,.2f}  BS diff: £{s['ap_bs_diff']:,.2f}  {'✓' if abs(s['ap_bs_diff'])<=tol_general else '✗'}")
        print(f"    Aged Receivables total:    £{s['ar_total']:>10,.2f}  BS diff: £{s['ar_bs_diff']:,.2f}  {'✓' if abs(s['ar_bs_diff'])<=tol_general else '✗'}")


# ══════════════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT  ← edit JobConfig here for CLI use
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    job = JobConfig(
        client_name         = "Demo Company (UK)",
        qe_end_date         = "28 Feb 2026",
        preparer            = "",
        reviewer            = "",
        working_dir         = ".",

        vat_return_file     = "Demo-Company-UK-VAT-Return.xlsx",
        trial_balance_file  = "Demo_Company__UK__-_Trial_Balance.xlsx",
        account_txn_file    = "Demo_Company__UK__-_Account_Transactions.xlsx",
        aged_pay_file       = "Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx",
        aged_rec_file       = "Demo_Company__UK__-_Aged_Receivables_Detail.xlsx",
        balance_sheet_file  = "Demo_Company__UK__-_Balance_Sheet.xlsx",

        # Accounting config — adjust per client
        sales_nominals          = [200],
        other_income_nominals   = [270],
        vat_control_nominal     = 820,
        opening_vat_balance     = 0.00,

        # Tolerances
        tol_general = 1.00,
        tol_box6    = 5.00,
    )

    try:
        VATWorkflowService().run_job(job)
    except InputError as e:
        print(f"\n  ERROR: {e}")
        sys.exit(1)
