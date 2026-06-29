import io
import json
import logging
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

logging.getLogger("vat_engine").setLevel(logging.WARNING)

from vat_engine import (
    AnnualSummaryBuilder,
    InputError,
    JobConfig,
    ParseError,
    ParseManager,
    QuarterSnapshot,
    ReconciliationEngine,
    ValidationResult,
    WorkbookBuilder,
    InputValidator,
)

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="VAT Working Papers",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DEMO_DIR = Path(__file__).parent / "demo_files"

DEMO_FILES = {
    "vat_return":   "Demo-Company-UK-VAT-Return.xlsx",
    "trial_bal":    "Demo_Company__UK__-_Trial_Balance.xlsx",
    "balance_sht":  "Demo_Company__UK__-_Balance_Sheet.xlsx",
    "account_txns": "Demo_Company__UK__-_Account_Transactions.xlsx",
    "aged_pay":     "Demo_Company__UK__-_Aged_Payables_Detail__2_.xlsx",
    "aged_rec":     "Demo_Company__UK__-_Aged_Receivables_Detail.xlsx",
}

# Session state defaults
for key in ("result", "workbook_bytes", "wb_filename", "demo_loaded",
            "snapshot_json", "snapshot_filename", "annual_bytes", "annual_filename"):
    if key not in st.session_state:
        st.session_state[key] = None
if "demo_loaded" not in st.session_state:
    st.session_state["demo_loaded"] = False
if "opening_vat_input" not in st.session_state:
    st.session_state["opening_vat_input"] = 0.00

# ── Header ────────────────────────────────────────────────────────────────────

st.title("📊 VAT Working Papers Generator")
st.caption(
    "Xero exports → formatted Excel working papers in one click. "
    "Data is processed locally — nothing leaves your machine."
)
st.caption(
    "📌 v1 supports **Xero** exports only. Support for other bookkeeping software "
    "(QuickBooks, Sage, FreeAgent) is on the roadmap."
)
st.divider()

tab_quarter, tab_annual = st.tabs(["📋 Quarterly Working Papers", "📅 Annual VAT Summary"])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — QUARTERLY WORKING PAPERS
# ════════════════════════════════════════════════════════════════════════════

with tab_quarter:

    # ── Carry forward from previous quarter ─────────────────────────────────
    with st.expander("🔁 Carry forward VAT control balance from previous quarter (optional)"):
        prior_snapshot_up = st.file_uploader(
            "Prior quarter VAT snapshot (JSON)",
            type=["json"], key="fu_prior_snapshot",
            help="Generated automatically when you ran last quarter's working papers — "
                 "sets this quarter's opening VAT balance to last quarter's closing balance.",
        )
        if prior_snapshot_up is not None:
            try:
                snap_data = json.loads(prior_snapshot_up.read().decode("utf-8"))
                carried = float(snap_data.get("closing_vat_balance", 0.0))
                st.success(
                    f"Loaded snapshot for **{snap_data.get('client_name', '')}** "
                    f"(period ending {snap_data.get('period_end', '')}). "
                    f"Closing VAT balance **£{carried:,.2f}** will be used as this "
                    f"quarter's opening balance below."
                )
                if st.session_state.get("_last_snapshot_name") != prior_snapshot_up.name:
                    st.session_state["opening_vat_input"] = carried
                    st.session_state["_last_snapshot_name"] = prior_snapshot_up.name
            except Exception as e:
                st.error(f"Could not read snapshot file: {e}")

    # ── Config form ───────────────────────────────────────────────────────────

    with st.form("config"):
        st.subheader("Engagement Details")
        col1, col2 = st.columns(2)
        with col1:
            client_name = st.text_input("Client name", value="Demo Company (UK)")
            preparer    = st.text_input("Preparer", value="")
        with col2:
            qe_end_date = st.text_input("Period end date", value="28 Feb 2026",
                                        help='Plain English, e.g. "28 Feb 2026"')
            reviewer    = st.text_input("Reviewer", value="")

        st.subheader("Accounting Config")
        col3, col4, col5, col6 = st.columns(4)
        with col3:
            opening_vat = st.number_input("Opening VAT balance (£)",
                                          format="%.2f", key="opening_vat_input",
                                          help="Prior period closing VAT control balance")
        with col4:
            vat_ctrl_nominal = st.number_input("VAT control nominal", value=820,
                                               step=1, format="%d")
        with col5:
            sales_nominals_raw = st.text_input("Sales nominals", value="200",
                                               help="Comma-separated GL codes")
        with col6:
            other_inc_raw = st.text_input("Other income nominals", value="270",
                                          help="Comma-separated GL codes")

        submitted = st.form_submit_button("Save config", type="secondary")

    # ── File uploaders ────────────────────────────────────────────────────────

    st.subheader("Upload Xero Exports")

    def _demo_bytes(filename: str) -> bytes:
        return (DEMO_DIR / filename).read_bytes()

    # Load demo files button
    if st.button("⚡ Load Demo Files", help="Pre-load all 6 demo files for a quick test run"):
        for k, fname in DEMO_FILES.items():
            st.session_state[f"upload_{k}"] = _demo_bytes(fname)
        st.session_state["demo_loaded"] = True
        st.rerun()

    col_req, col_opt = st.columns(2)

    with col_req:
        st.markdown("**Required files**")

        vat_return_up = st.file_uploader(
            "VAT Return (XLSX)",
            type=["xlsx"], key="fu_vat_return",
            help="Xero: Reports → VAT Return → Export",
        )
        trial_bal_up = st.file_uploader(
            "Trial Balance (XLSX)",
            type=["xlsx"], key="fu_trial_bal",
            help="Xero: Reports → Trial Balance → Export",
        )
        balance_sht_up = st.file_uploader(
            "Balance Sheet (XLSX)",
            type=["xlsx"], key="fu_balance_sht",
            help="Xero: Reports → Balance Sheet → Export",
        )

    with col_opt:
        st.markdown("**Optional files** *(enables additional reconciliations)*")

        account_txns_up = st.file_uploader(
            "Account Transactions (XLSX)",
            type=["xlsx"], key="fu_account_txns",
            help="Export from Xero for your Bank or VAT Control (820) account — "
                 "needed to extract HMRC VAT payment amounts for the VAT Control reconciliation",
        )
        aged_pay_up = st.file_uploader(
            "Aged Payables Detail (XLSX)",
            type=["xlsx"], key="fu_aged_pay",
            help="Needed for creditor reconciliation to Balance Sheet",
        )
        aged_rec_up = st.file_uploader(
            "Aged Receivables Detail (XLSX)",
            type=["xlsx"], key="fu_aged_rec",
            help="Needed for debtor reconciliation to Balance Sheet",
        )

    # Apply demo bytes to file uploaders via session state shim
    def _get_file(uploader_result, session_key: str):
        """Return UploadedFile or session-state bytes, whichever is available."""
        if uploader_result is not None:
            return uploader_result
        demo = st.session_state.get(f"upload_{session_key}")
        if demo:
            return io.BytesIO(demo)
        return None

    vat_return_file  = _get_file(vat_return_up,   "vat_return")
    trial_bal_file   = _get_file(trial_bal_up,    "trial_bal")
    balance_sht_file = _get_file(balance_sht_up,  "balance_sht")
    acct_txns_file   = _get_file(account_txns_up, "account_txns")
    aged_pay_file    = _get_file(aged_pay_up,      "aged_pay")
    aged_rec_file    = _get_file(aged_rec_up,      "aged_rec")

    if st.session_state.get("demo_loaded"):
        st.info("Demo files loaded. Click **Generate Working Papers** to run.")

    st.divider()

    # ── Generate ──────────────────────────────────────────────────────────────

    generate = st.button("Generate Working Papers", type="primary", use_container_width=True)

    if generate:
        # Client-side required file check
        missing = []
        if vat_return_file is None:  missing.append("VAT Return")
        if trial_bal_file is None:   missing.append("Trial Balance")
        if balance_sht_file is None: missing.append("Balance Sheet")

        if missing:
            st.error(f"Please upload the required file(s): {', '.join(missing)}")
        else:
            def _save(file_obj, name: str, tmpdir: str) -> str:
                path = Path(tmpdir) / name
                data = file_obj.read() if hasattr(file_obj, "read") else file_obj
                path.write_bytes(data)
                return str(path)

            try:
                with st.status("Generating working papers...", expanded=True) as status:
                    # Save uploads to temp dir
                    st.write("Saving uploaded files...")
                    tmpdir = tempfile.mkdtemp()
                    st.session_state["tmpdir"] = tmpdir

                    vat_path  = _save(vat_return_file,  "vat_return.xlsx",  tmpdir)
                    tb_path   = _save(trial_bal_file,   "trial_bal.xlsx",   tmpdir)
                    bs_path   = _save(balance_sht_file, "balance_sht.xlsx", tmpdir)
                    at_path   = _save(acct_txns_file,   "account_txns.xlsx", tmpdir) if acct_txns_file else None
                    ap_path   = _save(aged_pay_file,    "aged_pay.xlsx",    tmpdir) if aged_pay_file  else None
                    ar_path   = _save(aged_rec_file,    "aged_rec.xlsx",    tmpdir) if aged_rec_file  else None

                    # Parse nominals
                    def _parse_nominals(raw: str) -> list:
                        return [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]

                    cfg = JobConfig(
                        client_name            = client_name.strip(),
                        qe_end_date            = qe_end_date.strip(),
                        preparer               = preparer.strip(),
                        reviewer               = reviewer.strip(),
                        working_dir            = tmpdir,
                        vat_return_file        = "vat_return.xlsx",
                        trial_balance_file     = "trial_bal.xlsx",
                        account_txn_file       = "account_txns.xlsx" if at_path else None,
                        aged_pay_file          = "aged_pay.xlsx"     if ap_path else None,
                        aged_rec_file          = "aged_rec.xlsx"     if ar_path else None,
                        balance_sheet_file     = "balance_sht.xlsx",
                        sales_nominals         = _parse_nominals(sales_nominals_raw) or [200],
                        other_income_nominals  = _parse_nominals(other_inc_raw)      or [270],
                        vat_control_nominal    = int(vat_ctrl_nominal),
                        opening_vat_balance    = float(opening_vat),
                    )

                    # Validate
                    st.write("Validating files...")
                    validation = InputValidator().validate(cfg)

                    # Parse
                    st.write("Parsing VAT Return, Trial Balance, Balance Sheet...")
                    data = ParseManager().parse_all(cfg)

                    # Reconcile
                    st.write("Running reconciliations...")
                    recs = ReconciliationEngine().run(cfg, data)

                    # Build workbook bytes
                    st.write("Building workbook...")
                    wb_bytes = WorkbookBuilder().build_to_bytes(cfg, validation, data, recs)

                    # Quarter snapshot — carry-forward into next quarter / annual summary
                    snapshot = QuarterSnapshot.build(cfg, data.boxes, recs)

                    # Store in session state
                    safe_c = re.sub(r"[^\w\s-]", "", cfg.client_name).strip().replace(" ", "_")
                    safe_p = cfg.qe_end_date.replace(" ", "_")
                    st.session_state["workbook_bytes"] = wb_bytes
                    st.session_state["wb_filename"]    = f"VAT_Working_Papers_{safe_c}_{safe_p}.xlsx"
                    st.session_state["snapshot_json"]     = snapshot.to_json()
                    st.session_state["snapshot_filename"] = f"VAT_Snapshot_{safe_c}_{safe_p}.json"
                    st.session_state["result"]         = {
                        "boxes": data.boxes,
                        "recs":  recs,
                        "validation": validation,
                        "cfg": cfg,
                    }

                    status.update(label="Working papers generated successfully ✅", state="complete")

            except InputError as e:
                st.error(f"Input error: {e}")
                st.session_state["result"] = None
            except ParseError as e:
                st.error(f"Parse error: {e}")
                st.session_state["result"] = None
            except Exception as e:
                st.error(f"Unexpected error: {e}")
                st.session_state["result"] = None

    # ── Results ───────────────────────────────────────────────────────────────

    if st.session_state.get("result"):
        r   = st.session_state["result"]
        boxes = r["boxes"]
        recs  = r["recs"]
        val   = r["validation"]
        cfg   = r["cfg"]

        st.divider()

        # Validation notices
        if val.warnings:
            with st.expander(f"Validation notices ({len(val.warnings)})", expanded=not val.passed):
                for w in val.warnings:
                    icon = {"info": "ℹ️", "warning": "⚠️", "error": "❌"}.get(w.severity, "•")
                    st.markdown(f"{icon} **[{w.severity.upper()}]** {w.section}: {w.message}")

        # Summary metrics + reconciliation table
        col_l, col_r = st.columns(2)

        with col_l:
            st.subheader("VAT Return Summary")
            m1, m2 = st.columns(2)
            m1.metric("Box 1 — Output VAT",  f"£{boxes.box1:,.2f}")
            m2.metric("Box 4 — Input VAT",   f"£{boxes.box4:,.2f}")
            m3, m4 = st.columns(2)
            m3.metric("Box 5 — VAT to pay",  f"£{boxes.box5:,.2f}")
            m4.metric("Box 6 — Net sales",   f"£{boxes.box6:,.2f}")
            st.caption(f"Scheme: {boxes.vat_scheme} · Period: {boxes.period_end} · VAT No: {boxes.vat_number}")

        with col_r:
            st.subheader("Reconciliation Results")

            def _status(diff: float, tol: float, info_label: str = None) -> str:
                if info_label:
                    return f"ℹ️ {info_label}"
                return "✅ Reconciled" if abs(diff) <= tol else "❌ Difference"

            rows = [
                {
                    "Check":      "VAT Control",
                    "Difference": f"£{recs.vat_control_diff:,.2f}",
                    "Status":     _status(recs.vat_control_diff, cfg.tol_general),
                },
                {
                    "Check":      "VAT Proof (Box 6 × 20% vs Box 1)",
                    "Difference": f"£{recs.vat_proof_diff:,.2f}",
                    "Status":     _status(recs.vat_proof_diff, cfg.tol_general),
                },
                {
                    "Check":      "Box 6 vs TB YTD Sales",
                    "Difference": f"£{recs.box6_diff:,.2f}",
                    "Status":     _status(recs.box6_diff, cfg.tol_box6, "Expected (QE vs YTD)"),
                },
                {
                    "Check":      "Aged Payables vs Balance Sheet",
                    "Difference": f"£{recs.ap_bs_diff:,.2f}",
                    "Status":     _status(recs.ap_bs_diff, cfg.tol_general)
                                   if recs.ap_total != 0 else "ℹ️ File not provided",
                },
                {
                    "Check":      "Aged Receivables vs Balance Sheet",
                    "Difference": f"£{recs.ar_bs_diff:,.2f}",
                    "Status":     _status(recs.ar_bs_diff, cfg.tol_general)
                                   if recs.ar_total != 0 else "ℹ️ File not provided",
                },
            ]

            def _colour_status(val: str) -> str:
                if val.startswith("✅"):
                    return "color: #1a7a4a; font-weight: bold"
                if val.startswith("❌"):
                    return "color: #c0392b; font-weight: bold"
                return "color: #7d6608"

            df = pd.DataFrame(rows)
            styled = df.style.applymap(_colour_status, subset=["Status"])
            st.dataframe(styled, use_container_width=True, hide_index=True)

        st.divider()

        # Downloads
        st.download_button(
            label="📥 Download Working Papers (Excel)",
            data=st.session_state["workbook_bytes"],
            file_name=st.session_state["wb_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.download_button(
            label="📥 Download Quarter Snapshot (JSON) — for next quarter's carry-forward or Annual Summary",
            data=st.session_state["snapshot_json"],
            file_name=st.session_state["snapshot_filename"],
            mime="application/json",
            use_container_width=True,
        )

# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — ANNUAL VAT SUMMARY
# ════════════════════════════════════════════════════════════════════════════

with tab_annual:
    st.subheader("Annual VAT Reconciliation")
    st.caption(
        "Upload the Quarter Snapshot (JSON) files generated when you ran each quarter's "
        "working papers (Q1 → Q4) to build a full financial-year VAT reconciliation: "
        "Box 1-9 totals rolled up for the year, and the VAT control account carried "
        "forward quarter to quarter."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        annual_client = st.text_input("Client name", value="Demo Company (UK)", key="annual_client")
    with col_b:
        fy_label = st.text_input("Financial year", value="FY 2025/26", key="annual_fy_label")

    snapshot_ups = st.file_uploader(
        "Quarter snapshots (JSON) — select 1-4 files, in order Q1 → Q4",
        type=["json"], accept_multiple_files=True, key="fu_annual_snapshots",
    )

    if snapshot_ups:
        snapshots = []
        try:
            for f in snapshot_ups:
                snapshots.append(QuarterSnapshot(**json.loads(f.read().decode("utf-8"))))
        except Exception as e:
            st.error(f"Could not read snapshot file(s): {e}")
            snapshots = []

        if snapshots:
            st.write(f"Loaded {len(snapshots)} quarter(s):")
            preview_rows = [
                {
                    "Quarter":     f"Q{i}",
                    "Period End":  s.period_end,
                    "Opening Bal": f"£{s.opening_vat_balance:,.2f}",
                    "Box 1":       f"£{s.box1:,.2f}",
                    "Box 4":       f"£{s.box4:,.2f}",
                    "Box 5":       f"£{s.box5:,.2f}",
                    "Closing Bal": f"£{s.closing_vat_balance:,.2f}",
                }
                for i, s in enumerate(snapshots, 1)
            ]
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

            # Continuity check — does each quarter's opening balance match the
            # previous quarter's reconstructed closing balance?
            for i in range(1, len(snapshots)):
                prev_close = snapshots[i - 1].closing_vat_balance
                this_open  = snapshots[i].opening_vat_balance
                if abs(prev_close - this_open) > 1.00:
                    st.warning(
                        f"Q{i + 1} opening balance (£{this_open:,.2f}) does not match "
                        f"Q{i} closing balance (£{prev_close:,.2f}) — check carry-forward."
                    )

            if st.button("Generate FY VAT Reconciliation", type="primary", use_container_width=True):
                wb_bytes = AnnualSummaryBuilder().build_to_bytes(
                    annual_client.strip(), fy_label.strip(), snapshots
                )
                safe_c  = re.sub(r"[^\w\s-]", "", annual_client).strip().replace(" ", "_")
                safe_fy = re.sub(r"[^\w\s-]", "", fy_label).strip().replace(" ", "_")
                st.session_state["annual_bytes"]    = wb_bytes
                st.session_state["annual_filename"] = f"FY_VAT_Reconciliation_{safe_c}_{safe_fy}.xlsx"

    if st.session_state.get("annual_bytes"):
        st.download_button(
            label="📥 Download Annual VAT Reconciliation (Excel)",
            data=st.session_state["annual_bytes"],
            file_name=st.session_state["annual_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
