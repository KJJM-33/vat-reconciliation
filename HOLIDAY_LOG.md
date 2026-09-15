# Holiday Hardening Log

Unattended hardening work on `vat-reconciliation` while Keyaan is away.
Each session appends a dated entry. Read this file first, don't repeat
completed work. Branch: `claude/holiday-hardening`.

## Status summary (update this table each session)

| Item | Status |
|---|---|
| 1. Run existing suite, fix genuine test bugs | Done (2026-09-07, re-checked 2026-09-08, 2026-09-09, 2026-09-13, 2026-09-14, 2026-09-15) — no failures found |
| 2. Add edge-case coverage | In progress (2026-09-07, 2026-09-08, 2026-09-09, 2026-09-13, 2026-09-14, 2026-09-15) — see below for what's covered / still open |
| 3. Respect CLAUDE.md Do-Not list | Followed for all code changes. **Found (not fixed) a Do-Not violation in app.py — see open questions below.** |
| 4. Fix low-risk concrete bugs | 1 fixed (2026-09-08): stale account-code comparison in `build_scenarios.py`. No new bugs found with high confidence in Sessions 3-6 — see below |

## Open questions for Keyaan (do not act on these without his input)

- **app.py contains a few small pockets of logic that arguably belong in
  vat_engine.py, per the CLAUDE.md "Do Not" rule.** Found by audit on
  2026-09-08, not changed — these don't produce wrong output (the logic is
  correct, just duplicated/misplaced), and I didn't want to refactor a live
  Streamlit UI unattended without being able to visually verify it in a
  browser afterward. Specifics:
  1. `_parse_nominals()` (app.py ~line 244-245) — parses free-text GL
     nominal codes into ints before building `JobConfig`. Trivial to move
     to vat_engine.py as a module-level helper; zero behaviour change.
  2. `_status()` (app.py ~line 344-347) — reimplements the same
     `abs(diff) <= tol` reconciled/not-reconciled tolerance check that
     `_flag()` already does in vat_engine.py's `WorkbookBuilder` (used to
     colour the Excel cells). The Streamlit results table computes this
     independently instead of reading a precomputed status from
     `ReconciliationResults`.
  3. Quarter-continuity check (app.py ~line 458-467) — validates each
     quarter's opening VAT balance ties to the prior quarter's closing
     balance (hard-coded £1.00 tolerance), inline in the UI layer instead
     of in `AnnualSummaryBuilder`/`QuarterSnapshot`.
  All three are safe, mechanical "move this logic, don't change it" fixes
  in principle. Left for Keyaan's call (or a future session willing to
  spin up the Streamlit app and click through it after the change) rather
  than guessing this is fine to do unattended.
- **`InputValidator` doesn't flag an unconfigured *required* file
  (`trial_balance_file`/`balance_sheet_file`/`vat_return_file` left as
  `None`, as opposed to configured-but-missing-on-disk) as an error** —
  it's reported as `info`, same as a genuinely optional file, and
  `validation.passed` stays `True`. The job doesn't silently succeed or
  produce wrong output — `ParseManager` still raises a clean `InputError`
  a step later — but the pre-flight validation message undersells the
  problem. In practice this is unreachable from the real Streamlit app
  (app.py does its own client-side required-file check before ever
  building a `JobConfig`), so it only matters for direct `JobConfig`/API
  use. Pinned down with two new tests in `test_unit_helpers.py`
  (`test_validator_unconfigured_required_file_is_info_not_error`,
  `test_parse_manager_blocks_unconfigured_required_file`) rather than
  changed, since `InputValidator._FILE_SPEC`'s required/optional handling
  looked like it could be intentional (e.g. "not configured" might mean
  something different from "file missing" in some caller's mental model)
  and I wasn't confident enough to change validation-layer behaviour
  without checking with Keyaan first.
- **`WorkbookBuilder._sheet_file_register`'s "Match?" column flags an
  *unconfigured optional* file (e.g. no Account Transactions uploaded)
  with the same red ✗ as a file that's present but has the wrong period
  date on it.** Found and pinned down with a test
  (`test_workbook_builder.py`,
  `test_file_register_match_column`) on 2026-09-13. The code
  (`vat_engine.py`, `_sheet_file_register`) doesn't distinguish "not
  configured" from "configured but date mismatch" — both end up with
  `d = validation.file_dates.get(label, "—")` failing to equal the
  expected QE date, so both render red. This lives in `vat_engine.py`
  itself (not app.py), so it's in scope for this hardening work in
  principle, but changing it means deciding what an unconfigured-optional
  row *should* show instead (grey/neutral? blank? skip the column
  entirely?) — a genuine judgement call about what the working paper
  should communicate to a reviewer, not a mechanical fix. Left for
  Keyaan's call rather than guessing.
- **`ReconciliationEngine._vat_control` never references Box 2** (VAT due
  on acquisitions of goods made in Northern Ireland from EU member
  states). Found by audit on 2026-09-15. The reconstructed VAT-control
  closing balance is hard-coded as `opening + Box1 - Box4 - HMRC
  payments` (`vat_engine.py` `_vat_control`, ~line 727) -- Box 2 is
  parsed, displayed on the VAT Return sheet, and included in Box 3/Box 5
  (which *are* Box1+Box2 and Box3-Box4 respectively on the actual
  return), but the VAT control account reconciliation itself silently
  ignores it. For a GB-only client Box 2 is always 0 so this is
  invisible in practice; for an NI client with genuine EU goods
  acquisitions it would mean the VAT control tab's diff doesn't reflect
  a real Box 2 liability at all. Pinned down (not fixed) with a new
  fixture, `test_scenarios/box2_ni_acquisitions/` (Box2=200, Box3/Box5
  updated to stay internally consistent with the return, Box1/Box4/TB
  left untouched) -- confirms `vat_control_diff` comes out bit-for-bit
  identical to the same base figures with Box2=0 (1563.27), even though
  Box5 on the return changes from 1318.93 to 1518.93. **Not changed**:
  not confident what the "correct" reconstruction formula should be here
  -- whether NI postponed-accounting entries would (a) net to zero on
  the VAT control nominal because the offsetting Box4 reclaim is already
  posted through the same account, in which case today's omission is
  actually fine, or (b) need Box2 added to the Box1 term the way Box3
  already is on the return itself, in which case today's formula
  understates the reconstructed liability by exactly Box2 any time it's
  non-zero. This is a bookkeeping-treatment question, not a mechanical
  bug, and it's exactly the kind of live-client VAT-logic question this
  task says not to guess at. Flagging for Keyaan: if any current or
  future client operates under NI protocol postponed VAT accounting for
  EU goods movements, this is worth resolving before relying on the VAT
  Control tab for them.

---

## Session 1 — 2026-09-07 (Sunday)

**Starting state:** no prior branch or log existed — this is day 1.

**Ran:**
- `test_scenarios/run_scenarios.py` — all 6 original scenarios (flat_rate,
  accrual, large_numbers, vat_repayment, dormant, minimal_files) passed
  with no crashes.
- `test_scenarios/test_annual_summary.py` — Q1-Q4 stacking + FY summary +
  JSON round-trip all OK.
- No pytest-style tests existed (`test_annual_summary.py` is a standalone
  script, not pytest); `pytest test_scenarios/` collects 0 tests. Left this
  as-is rather than converting the harness, to avoid an unrequested
  refactor of the test style the repo already uses.

**No bugs found.** Read through `InputValidator`, `ReconciliationEngine`
(`_vat_control`, `_box6`, `_aged_payables`/`_aged_receivables`), `_flag`,
`_safe_float`, `_norm_date`. All rounding is `round(x, 2)` used only for
*display/diff* purposes (comparing reported Xero figures against
reconstructed ones) — the app never calculates statutory VAT itself, so
there's no rounding-direction bug to fix; Python's round-half-to-even on
a computed diff is not a compliance issue here. Nothing else looked
wrong with high confidence.

**Added test coverage** (all under `test_scenarios/`, following the
existing script-based convention, not pytest):

1. `test_unit_helpers.py` — new file. Direct assertions on pure helpers
   that the fixture-driven scenarios don't pin down precisely:
   - `_flag()` tolerance boundary is inclusive (`abs(diff) <= tol` →
     green; confirmed exactly-at-tolerance is green, one penny over is
     red, and it works for negative diffs too).
   - `_safe_float()` on currency strings, thousands separators, blanks,
     `"-"` placeholders, NaN, garbage text, negative currency.
   - `_norm_date()` across the date formats Xero exports actually use
     (`"28 February 2026"`, `"28 Feb 2026"`, `"28/02/2026"`,
     `"2026-02-28"`) plus an unparseable fallback.
   - `InputValidator.validate()`: missing *optional* file → `info`
     severity, validation still passes; missing *required* file (VAT
     Return) → `error` severity, validation fails; a period-date
     mismatch → `warning` severity only, validation still passes (i.e.
     confirms the documented "warn, don't block" behaviour actually
     holds).
   - Run: `python test_scenarios/test_unit_helpers.py`

2. Two new fixture scenarios in `build_scenarios.py` (+ wired into
   `run_scenarios.py`'s `SCENARIOS` list), both real UK VAT edge cases
   not covered by the existing 6:
   - `credit_note_heavy` — Box 1 (output VAT) goes **negative** because
     credit notes issued in the period exceed VAT on sales (a legitimate
     HMRC-permitted position, distinct from `vat_repayment` which only
     drives Box 5 negative via Box4 > Box1 with Box1 still positive).
     Confirms the engine handles a negative Box1/Box3 without crashing
     or mis-signing the VAT proof / control diffs.
   - `missing_vat_return` — the VAT Return file itself is absent (a
     *required* file, unlike the optional files `minimal_files` already
     covers). Confirms `InputValidator` flags it as an `error` and
     `VATWorkflowService.run_job()` raises a clean `InputError` rather
     than an unhandled file-not-found traceback.
   - Regenerate with `python test_scenarios/build_scenarios.py`, run with
     `python test_scenarios/run_scenarios.py`.

**Note on regenerated fixtures:** running `build_scenarios.py` /
`run_scenarios.py` re-saves *every* scenario's `.xlsx` files via
openpyxl, which rewrites zip-internal metadata (styles.xml, core.xml
timestamps, etc.) even when no cell value changes — this shows up as a
binary diff on files I didn't intend to touch. I reverted those
unrelated regenerated files before committing and kept only the two new
scenario directories, so this commit's diff is limited to what actually
changed. **Future sessions: after running the generator/runner scripts,
`git status` and `git checkout --` any `.xlsx` under `test_scenarios/`
you didn't deliberately add/change, before committing.**

**Still open for next session (item 2 continued):**
- Rounding *at VAT-box boundaries* specifically (e.g. a reconciliation
  diff landing exactly on a tolerance threshold via realistic Xero
  figures, not just the synthetic `_flag()` unit check already added).
- An "empty VAT period" scenario distinct from `dormant`: `dormant`
  zeroes out TB/VAT-return values but keeps the row structure; a
  genuinely empty period (e.g. a Trial Balance with zero data rows, not
  just zero-valued rows) hasn't been tried and might expose an
  assumption in `TrialBalanceParser`/`_tb_balance` about row presence.
- Partial file uploads through the full `VATWorkflowService` pipeline
  (only `InputValidator` was unit-tested directly above; haven't checked
  what happens if, say, only the VAT Return + Trial Balance are present
  but Balance Sheet is missing, end-to-end).
- Have not yet looked at `app.py` (Streamlit UI) at all this session —
  per CLAUDE.md it should contain no logic, worth a quick confirmation
  read next session rather than assuming.

**Commits this session:** see git log on `claude/holiday-hardening`.

---

## Session 2 — 2026-09-08 (Monday)

**Starting state:** picked up `claude/holiday-hardening` from Session 1
(commit `91fa957`). Re-ran everything first per item 1 before adding new
work.

**Ran (all green, no genuine failures):**
- `test_scenarios/run_scenarios.py` — all 8 existing scenarios still pass.
- `test_scenarios/test_unit_helpers.py` — all 24 existing checks pass.
- `test_scenarios/test_annual_summary.py` — still OK.

**Audited `app.py` for CLAUDE.md Do-Not compliance** (session 1's leftover
item — "haven't yet looked at app.py this session"). Found three small
pockets of logic that arguably belong in `vat_engine.py` (parsing GL
nominal codes, a duplicated tolerance/reconciled check, an inline
quarter-continuity balance check). None produce wrong output. **Did not
change app.py this session** — logged as an open item for Keyaan/a future
session with browser access to verify visually after the move, rather than
refactoring a live Streamlit UI unattended on judgement alone. Full detail
in the "Open questions" section above.

**Investigated the two remaining "still open" items from Session 1:**

1. *Empty Trial Balance (zero data rows, not just zero-valued rows).*
   Read through `TrialBalanceParser`, `_tb_balance`, `_box6`,
   `_vat_control`, `_aged_payables/receivables` — every TB nominal lookup
   already degrades to `0.0` on a missing row rather than raising. Added
   `test_scenarios/empty_trial_balance/` (new fixture: TB stripped to
   header + Total row only, real VAT Return figures kept) to lock this in
   as a regression test rather than just a code-read conclusion — confirms
   end-to-end (`VATWorkflowService.run_job`) it doesn't crash and degrades
   as expected (Box 6 vs TB diff = full Box 6 amount, VAT control diff =
   full reconstructed balance, aged reports unaffected).

2. *Partial file uploads through the full pipeline.* Already covered
   end-to-end by existing scenarios more than Session 1's note suggested:
   `missing_vat_return` (required file absent from disk) and
   `minimal_files` (optional files absent) both already run through
   `VATWorkflowService.run_job`, not just `InputValidator` directly. The
   one gap was a required file left *unconfigured* (`None`) rather than
   missing-on-disk — investigated, found the `InputValidator` behaviour
   noted above, and pinned it down with two new unit tests instead of
   building a full fixture scenario for it (no Xero file needed to
   reproduce, so a unit test is the right weight).

**Added, per item 2 (edge-case coverage):**
- `test_scenarios/empty_trial_balance/` — new fixture + scenario, wired
  into `run_scenarios.py`'s `SCENARIOS` list. See above.
- `test_scenarios/control_diff_at_tolerance/` — new fixture: VAT control
  diff engineered to land exactly on `cfg.tol_general` (£1.00) using real
  Xero-shaped figures (Box1=2598.24, Box4=1279.31, TB nominal 820 set so
  the reconstructed-vs-TB diff is exactly £1.00), not just the synthetic
  `_flag()` unit check Session 1 already added. `run_scenarios.py` asserts
  `vat_control_diff` lands at exactly the boundary so this doesn't
  silently stop testing what it's meant to test. This was the other
  concrete "still open" item from Session 1 (rounding at VAT-box
  boundaries with realistic figures).
- Two new checks in `test_unit_helpers.py` documenting the
  `InputValidator`/`ParseManager` behaviour for an unconfigured (not just
  missing-on-disk) required file — see open questions above.

**Fixed one concrete, low-risk bug (item 4) — in test fixture generation,
not `vat_engine.py`:** `test_scenarios/build_scenarios.py` compared the
Trial Balance's Account Code cell to the **int** `820`
(`row[0].value == 820`), but Xero exports store Account Code as **text**
(`'820'`), so the comparison was always `False` and silently no-opped.
This meant `scenario_vat_repayment()`'s intended TB nominal-820 override
(meant to reflect the repayment position on the VAT control account) never
actually applied — the fixture's Trial Balance silently kept its
unmodified base value the whole time, so that scenario's `vat_control_tb`/
`vat_control_diff` were wrong (`-244.34` / `-755.66` instead of the
intended `1000.0` / `-2000.0`). Not a `vat_engine.py` bug — the engine's
own TB parsing (`TrialBalanceParser`) already correctly coerces Account
Code to numeric via `pd.to_numeric` before comparing against
`cfg.vat_control_nominal` (an int), so real uploads were never affected.
Fixed by comparing `str(row[0].value) == "820"` instead (also applied to
the new `control_diff_at_tolerance` scenario, written before I'd found
this and initially hit the identical bug — good confirmation the fix
actually mattered: the assertion I added for that scenario failed until
this was fixed). Rebuilt all fixtures and reran the full suite — all still
pass, and `vat_repayment`'s figures now reflect what its own scenario
docstring says they should.

**Note on regenerated fixtures (same caveat as Session 1):** `git status`
after every `build_scenarios.py`/`run_scenarios.py` run and
`git checkout --` any `.xlsx` that only changed due to openpyxl metadata
churn, keeping the diff limited to deliberate changes. Did this
consistently this session — final diff before commit is exactly:
`build_scenarios.py`, `run_scenarios.py`, `test_unit_helpers.py`, the
`vat_repayment` fixture's Trial Balance + regenerated
snapshot/workbook (the bug-fix), and the two new scenario directories.

**Still open for next session:**
- The two "Open questions for Keyaan" above (app.py logic placement,
  InputValidator required-file severity) — need his input, don't act
  without it.
- Haven't yet tried a negative/credit *Box 4* scenario in isolation (only
  negative Box 1 via `credit_note_heavy` so far) — e.g. a large one-off
  capital purchase or VAT-inclusive adjustment driving input VAT reclaim
  unusually high relative to output VAT, distinct from `vat_repayment`
  (repayment scenario is Box4 > Box1, but Box4 itself is a plausible
  round number, not testing Box4 dominating for an unusual reason).
- Haven't looked closely at `AnnualSummaryBuilder`/quarter-stacking edge
  cases beyond the existing `test_annual_summary.py` happy path (e.g. a
  missing quarter in the middle of Q1-Q4, or the quarter-continuity
  £1.00-tolerance check itself under a boundary value) — ties into the
  app.py open question above since that specific check currently only
  exists in app.py, not vat_engine.py.
- No pytest-style tests still; left as-is per Session 1's reasoning
  (avoid an unrequested refactor of the existing test style).

**Commits this session:** see git log on `claude/holiday-hardening`.

---

## Session 3 — 2026-09-09 (Tuesday)

**Starting state:** picked up `claude/holiday-hardening` from Session 2
(commit `5a08b58`). Installed `requirements.txt` fresh (this container had
no packages installed) before running anything.

**Ran (all green, no genuine failures — item 1):**
- `test_scenarios/run_scenarios.py` — all 10 existing scenarios pass.
- `test_scenarios/test_unit_helpers.py` — all 24 existing checks pass.
- `test_scenarios/test_annual_summary.py` — still OK.
- Confirmed (again) no pytest-style tests exist anywhere in the repo
  (`pytest` isn't even installed) — same conclusion as Sessions 1-2,
  left as-is.

**Closed both of Session 2's "still open" items (item 2 — edge-case
coverage):**

1. **Negative Box 4 in isolation.** Added
   `test_scenarios/negative_box4/` (new fixture + scenario, wired into
   `run_scenarios.py`) — Box1=2598.24 unchanged, Box4=-150.00 (net input
   VAT negative, e.g. purchase credit notes/returns in the period
   exceeding VAT incurred on purchases — a legitimate UK VAT position),
   Box3=2598.24, Box5=2748.24 (=Box3-Box4). Distinct from
   `vat_repayment` (Box4 > Box1 but Box4 itself still positive) and
   `credit_note_heavy` (Box1 negative, not Box4). Confirms end-to-end
   the engine doesn't mis-sign `_vat_control`'s
   `opening + box1 - box4 - hmrc_total` arithmetic or crash when Box4
   itself is negative — ran clean (box4=-150.0, box5=2748.24 as
   expected, no crash).

2. **`AnnualSummaryBuilder` quarter-stacking edge cases.** Added
   `test_scenarios/test_annual_summary_edge_cases.py` — 7 checks built
   directly against hand-crafted `QuarterSnapshot` objects (fast, no
   need to reparse Xero files) plus the FY workbook's own
   `_note()`-written continuity-flag cells (VAT CONTROL ROLL-FORWARD
   table, col 10), not the app.py version:
   - Continuity tolerance boundary is inclusive: opening balance exactly
     £1.00 off the prior quarter's closing → **not** flagged; £1.01 off
     → flagged. Confirms `AnnualSummaryBuilder`'s own
     `abs(q.opening_vat_balance - prev_closing) > 1.00` check (this
     exists independently of the app.py duplicate logged below) matches
     `_flag()`'s `<=` convention used everywhere else in the workbook.
   - Break flagged in both directions (opening > prior close as well as
     opening < prior close) — confirms `abs()` is used, not a signed
     comparison.
   - 4 consecutive quarters with correct carry-forward → zero
     continuity notes anywhere (no false positives).
   - **Missing quarter in the middle** (e.g. only Q1 and real-Q3
     snapshots passed, Q2 never uploaded): `AnnualSummaryBuilder` has no
     concept of calendar quarters — it labels rows "Q1", "Q2", ... by
     **list position only**, so the real Q3 snapshot renders under the
     label "Q2". The continuity check still does its numeric job (the
     real gap between Q1's closing and Q3's opening balance is caught
     and flagged), just under a misleading row label. **Not changed** —
     documented and pinned down with a test rather than "fixed", because
     it's ambiguous whether calendar-quarter validation belongs in this
     builder at all (it takes whatever `QuarterSnapshot` list the caller
     hands it; the caller — app.py's Annual Summary tab — is presumably
     meant to be the one responsible for feeding it contiguous quarters
     in order) or whether an assumption like "index 0 is always Q1" is
     even meant to hold. Same reasoning as Session 1-2's other
     "pin down with a test, don't guess a fix" open items. Added to open
     questions below.
   - Single-quarter FY and empty quarters list both build without
     crashing.
   - Run: `python test_scenarios/test_annual_summary_edge_cases.py`.
     Probe workbooks are written to a `tempfile.mkdtemp()` dir, not
     under `test_scenarios/`, since they're pure assertion scaffolding
     with no review value of their own (unlike the scenario fixtures,
     which double as sample output for a human to eyeball) — avoids
     adding meaningless committed `.xlsx` files to the repo.

**No new low-risk bugs found this session (item 4).** Read through
`_vat_control`, `_top10_box4`, and `AnnualSummaryBuilder._build_wb` while
building the above two items; nothing else looked wrong with high
confidence.

**Note on regenerated fixtures (same caveat as Sessions 1-2, still
holds):** every `build_scenarios.py`/`run_scenarios.py` run rewrites
zip-internal metadata on every scenario's `.xlsx`, including ones this
session didn't touch. Did `git status` + `git checkout --` after every
run this session, same as before — final diff limited to
`build_scenarios.py`, `run_scenarios.py`, the new
`negative_box4/` fixture directory, and the new
`test_annual_summary_edge_cases.py` file.

**New item for "Open questions for Keyaan" (see section above — added
there, not a separate list):**
- `AnnualSummaryBuilder` labels VAT-control-roll-forward rows "Q1",
  "Q2", ... by list position, not by any calendar-quarter identifier on
  `QuarterSnapshot` itself (there isn't one — only `period_end`, a free
  string). If a quarter is skipped when uploading snapshots in the
  Streamlit Annual Summary tab, every subsequent row is mislabelled
  (off by one) even though the underlying £-value continuity check still
  fires correctly on the real gap. Locked in behaviour with a test
  (`test_annual_summary_edge_cases.py`, "missing quarter in the middle")
  rather than changing it — not clear whether it's in scope for
  `AnnualSummaryBuilder` to validate `period_end`s are consecutive
  calendar quarters (would need real date parsing of period_end, which
  is a free-text field, not guaranteed parseable) versus that being
  app.py's job as the upload UI, which is exactly the kind of "which
  layer should own this" question already flagged for Keyaan above.

**Still open for next session:**
- The three "Open questions for Keyaan" above — still need his input,
  don't act without it. Nothing this session changed their status.
- Session 2's other still-open items are now closed (see above); no new
  concrete gaps identified this session beyond what's logged as open
  questions. Worth a fresh look next session at `WorkbookBuilder`'s
  Excel-formatting/formula edge cases (untouched by any session so far —
  all prior + this session's work has been on parsing/reconciliation/
  annual-summary logic, not the workbook rendering layer itself) if
  nothing else stands out.
- No pytest-style tests still; left as-is per Session 1's reasoning.

**Commits this session:** see git log on `claude/holiday-hardening`.

---

## Session 4 — 2026-09-13 (Saturday)

*(Note from Session 5: this file's weekday labels don't always match the
actual calendar weekday for the stated date — e.g. 2026-09-13 is a Sunday,
not a Saturday. Not corrected retroactively since it doesn't affect the
content; just flagging in case a future session double-checks something
by weekday instead of date.)*

**Starting state:** picked up `claude/holiday-hardening` from Session 3
(commit `c13d6b3`). `pip install -r requirements.txt` fresh (this
container had no packages installed).

**Ran (all green, no genuine failures — item 1):**
- `test_scenarios/run_scenarios.py` — all 11 existing scenarios pass.
- `test_scenarios/test_unit_helpers.py` — all checks pass.
- `test_scenarios/test_annual_summary.py` — still OK.
- `test_scenarios/test_annual_summary_edge_cases.py` — all 7 checks pass.
- `pytest` still not installed / no pytest-style tests anywhere — same
  conclusion as every prior session, left as-is.

**Took up Session 3's suggested next focus — `WorkbookBuilder`'s
Excel-formatting/formula edge cases, untouched by any session so far
(item 2).** Read through the whole class (`vat_engine.py` ~880-1352):
sheet-by-sheet cell/column mapping (file register, VAT return, txn-by-box,
VAT control, bank rec, box 6 rec, aged payables/receivables, trial
balance, top 10 box 4, checklist) plus the shared helpers (`_hdr`, `_cel`,
`_flag`, `_note`, `_write_txns`). Checked the aged-payables/receivables
column-index arithmetic in particular (different header-column counts
between the two sheets) by hand against `data.aged_pay_raw`/
`aged_rec_raw`'s actual column layout — lines up correctly on both, no
off-by-one. Also re-confirmed `TrialBalanceParser` already guarantees
`Account Code` is a clean non-null `int` column before
`_sheet_trial_balance` does `int(row["Account Code"])` on it (relevant
after Session 2's fixture-generator bug in this exact area) — safe.

**Added `test_scenarios/test_workbook_builder.py`** — new file, same
"synthetic dataclasses, no Xero fixture needed" approach Session 3 used
for `test_annual_summary_edge_cases.py`, applied to `WorkbookBuilder` for
the first time. Calls the private `_sheet_*` methods directly against
hand-built `JobConfig`/`ParsedData`/`ReconciliationResults` objects and
asserts on the actual openpyxl cell fills/values written — this is
genuinely new coverage, not a restatement of what `run_scenarios.py`
already exercises end-to-end (that only checks the pipeline doesn't
crash and the key figures are right; it doesn't assert on cell colours or
which row gets highlighted). Checks added:
- VAT Return sheet: Box 5 negative (repayment) → green cell; Box 5
  positive (owe HMRC) → red cell.
- Trial Balance sheet: the row matching `cfg.vat_control_nominal` is
  highlighted blue across all 6 columns; a non-matching nominal isn't.
- Aged Payables/Receivables: the BS-agreement diff cell is green when
  within `_flag`'s tolerance and red when not (this needed
  `ap_bs_diff`/`ar_bs_diff` set explicitly on the synthetic
  `ReconciliationResults` — those fields aren't derived by the sheet
  method itself, only displayed, so a naive "set totals and expect the
  sheet to compute the diff" test would have silently tested nothing;
  caught this while writing the test, not a runtime bug).
- File Register: a file matching the expected QE date is green-ticked, a
  present-but-mismatched-date file is red-crossed, and an unconfigured
  optional file is *also* red-crossed — this last one is the open item
  logged above, pinned down as current behaviour rather than changed.
- VAT Checklist: all 20 items render with their reference codes, and the
  preparer name is written into the header block.
- Run: `python test_scenarios/test_workbook_builder.py`. Building the
  fill-colour assertion helper surfaced one thing worth noting for anyone
  extending this file: openpyxl reports an *unfilled* cell's
  `fill.fgColor.rgb` as the string `'00000000'`, not `None` — you have to
  check `fill.fill_type is None` first, or an "this cell should have no
  background" assertion will silently pass by accident (matched a literal
  `''` that dumping `.rgb` alone would never produce, so this would have
  failed loudly, not silently — still worth documenting for the next
  session touching this file).

**New item for "Open questions for Keyaan"** (see section above): the
File Register "Match?" red-flagging of unconfigured optional files,
found while writing the test above. Not changed — see reasoning there.

**No new low-risk bugs found this session (item 4)** beyond the File
Register item above, which is a UX/design judgement call, not a
mechanical bug (off-by-one, wrong rounding direction, etc.) — logged as
an open question per the same reasoning as prior sessions' similar
findings, not fixed unattended.

**Note on regenerated fixtures (same caveat as every prior session):**
running `run_scenarios.py`/`test_annual_summary.py` (and an ad-hoc probe
script used while investigating fill colours, pointed at the
`vat_repayment` scenario) rewrites zip-internal metadata on scenario
`.xlsx` files even when no cell value changes. `git status` +
`git checkout --` after every run this session — final diff is exactly
the new `test_scenarios/test_workbook_builder.py` file and this log.

**Still open for next session:**
- The four "Open questions for Keyaan" above (app.py logic placement,
  `InputValidator` required-file severity, `AnnualSummaryBuilder` Q-label
  positional numbering, File Register unconfigured-file red-flag) — still
  need Keyaan's input, don't act without it.
- `WorkbookBuilder` now has direct test coverage for the sheets most
  likely to hide a formatting bug (VAT Return colouring, TB highlight,
  aged BS-agreement flags, file register, checklist). Not yet covered:
  `_sheet_txn_by_box`, `_sheet_vat_control`'s HMRC-payments table
  rendering, `_sheet_bank_rec`, `_sheet_box6_rec`'s proof-of-output-VAT
  block, `_sheet_top10` — worth a look if nothing else stands out next
  session, same "synthetic dataclasses" approach as this session's new
  file.
- No pytest-style tests still; left as-is per Session 1's reasoning.

**Commits this session:** see git log on `claude/holiday-hardening`.

---

## Session 5 — 2026-09-14 (Monday)

**Starting state:** picked up `claude/holiday-hardening` from Session 4
(commit `8903537`). `pip install -r requirements.txt` fresh (this
container had no packages installed).

**Ran (all green, no genuine failures — item 1):**
- `test_scenarios/run_scenarios.py` — all 11 existing scenarios pass.
- `test_scenarios/test_unit_helpers.py` — all checks pass.
- `test_scenarios/test_annual_summary.py` — still OK.
- `test_scenarios/test_annual_summary_edge_cases.py` — all 7 checks pass.
- `test_scenarios/test_workbook_builder.py` — all 15 pre-existing checks
  pass (before this session's additions).
- `pytest` still not installed / no pytest-style tests anywhere — same
  conclusion as every prior session, left as-is.

**Took up Session 4's suggested next focus — the remaining
`WorkbookBuilder` sheets with no direct formatting/rendering coverage yet
(item 2):** `_sheet_txn_by_box`, `_sheet_vat_control`'s HMRC-payments
table, `_sheet_bank_rec`, `_sheet_box6_rec`'s proof-of-output-VAT block,
and `_sheet_top10`. Introspected each method's actual row layout with
synthetic `JobConfig`/`ParsedData`/`ReconciliationResults` objects (same
approach as Session 4) before writing assertions, rather than guessing
row numbers from reading the source.

**Added to `test_scenarios/test_workbook_builder.py`** (5 new test
functions, 24 new checks):
- `test_txn_by_box_writes_totals_and_skips_empty_subsections` — Box 1/4
  totals and transaction rows render correctly; confirmed Box 6/7 (no
  configured `txn_sections` entries in the synthetic data) still render
  their box-total banner row with no sub-header/transaction rows in
  between — the `if df.empty: continue` guard in `_sheet_txn_by_box`
  behaves as intended, not just "doesn't crash".
- `test_vat_control_hmrc_payments_table_and_diff_flag` — HMRC payment
  rows render (date/description/amount) when `recs.hmrc_payments` is
  populated; the italic placeholder message renders when it's empty; the
  "Nominal N not in TB — enter manually" note appears when
  `vat_control_tb` is falsy; the closing-difference cell flags green/red
  correctly either side of the £1 tolerance.
- `test_bank_rec_balances_and_difference_shading` — TB/statement balances
  both reflect `recs.bs_bank` (there's no independent statement figure
  fed in — see note below); the difference row is always green-shaded
  (it's hard-coded 0.0 with no computed reconciling items — see note
  below).
- `test_box6_rec_proof_of_output_vat_block` — the Box 6 × 20% ≈ Box 1
  proof block (VATable sales after removing zero-rated, expected output
  VAT, actual Box 1, and the flagged difference) renders and flags
  correctly; also confirmed the QE-vs-YTD difference above it uses the
  wider `cfg.tol_box6` (£5) tolerance rather than the general £1 one,
  matching `_sheet_box6_rec`'s explicit `tol=cfg.tol_box6` argument.
- `test_top10_box4_rows_and_total` — Top 10 Box 4 transaction rows and
  the VAT-column total render correctly; empty `top10_box4` renders a
  zero total without crashing.
- Run: `python test_scenarios/test_workbook_builder.py` (now 39 checks
  total, all passing).

**Worth flagging for whoever next touches `_sheet_top10` (documented in
the test itself, not changed — not a bug):** `_sheet_top10` reads
`recs.top10_box4["_vat"]` directly. That column only exists because
`ReconciliationEngine._top10_box4` adds it to the DataFrame before
`nlargest()` — it's not part of `top10_box4`'s "public" shape (Date/
Account/Reference/Details/VAT/Net, same as every other txn DataFrame in
this codebase). The real pipeline always produces both together, so this
isn't reachable as a bug in practice, but a `ReconciliationResults` built
directly with a `top10_box4` DataFrame in the "normal" shape (as every
other synthetic test in this file does for other fields) raises a bare
`KeyError: '_vat'` instead of a clearer error — hit this firsthand while
writing the test, before adding the `_vat` column deliberately to match
what the engine actually produces.

**No new "Open questions for Keyaan" this session** — nothing found while
building this coverage looked like a genuine judgement call the way
Sessions 2-4's items did; the `_vat`-column coupling above is an internal
implementation detail, not something with a UI-visible or judgement-call
dimension.

**No new low-risk bugs found this session (item 4).** Read
`InputValidator`, `VATReturnParser` (including its label-scan fallback),
`TxnByBoxParser`, `TrialBalanceParser`, `AgedReportParser`,
`AccountTxnParser`, and `ParseManager` end-to-end while looking for
something to hand off to the workbook-layer work above; all degrade
gracefully (missing header row → empty DataFrame, non-numeric Account
Code → dropped, primary VAT-return cell read → label-scan fallback with
an explicit skip-the-box-number-column comment already in place) and
nothing looked wrong with high confidence. `ReconciliationEngine._top10_box4`
picks `nlargest(10, "_vat")` on signed VAT values rather than by
magnitude — so a large *negative* Box 4 line (a purchase credit note)
would never surface in the "top 10 to check invoices for" list even
though it's exactly the kind of unusual line worth checking. Not changed:
plausible this is intentional (largest reclaims, not largest-magnitude
adjustments), and Session 3's `negative_box4` scenario shows a negative
Box 4 *total* is already a legitimate position the engine handles
correctly elsewhere — whether the Top 10 list should rank by magnitude
is a product judgement call, not a mechanical bug, so logging it here for
awareness rather than as a fifth "Open question" (no real client data has
surfaced this in practice, and it doesn't affect any figure that's relied
on for reconciliation, only which rows get manually spot-checked).

**Note on regenerated fixtures (same caveat as every prior session):**
`run_scenarios.py`, `test_annual_summary.py`, and this session's ad-hoc
row-layout introspection scripts (not committed — pure scratch, run
outside the repo's normal test files) rewrite zip-internal metadata on
scenario `.xlsx` files. `git status` + `git checkout --` on every touched
fixture directory before committing — final diff is exactly
`test_scenarios/test_workbook_builder.py` and this log.

**Still open for next session:**
- The four "Open questions for Keyaan" (app.py logic placement,
  `InputValidator` required-file severity, `AnnualSummaryBuilder` Q-label
  positional numbering, File Register unconfigured-file red-flag) — still
  need Keyaan's input, don't act without it.
- `WorkbookBuilder` now has direct coverage across every sheet method
  (`_sheet_file_register`, `_sheet_vat_return`, `_sheet_txn_by_box`,
  `_sheet_vat_control`, `_sheet_bank_rec`, `_sheet_box6_rec`,
  `_sheet_aged_payables`/`_receivables`, `_sheet_trial_balance`,
  `_sheet_top10`, `_sheet_checklist`) except the private helpers
  (`_hdr`/`_cel`/`_flag`/`_note`/`_write_txns`/`_wp_header`/`_set_widths`)
  which are exercised indirectly by every sheet test already and don't
  seem to need direct unit tests of their own.
- `_sheet_bank_rec` is worth a second look by a future session with more
  context on the real bank-rec workflow: `recs.bs_bank` (a single TB
  balance) is used for *both* "Balance per TB" and "Balance per
  Statement", and "unreconciled items"/"outstanding payments"/
  "Difference" are all hard-coded `0.0` — there's no actual bank
  statement figure or reconciling-items source fed into
  `ReconciliationResults` anywhere in the codebase. This might be
  entirely intentional (a placeholder scaffold for the preparer to fill
  in by hand in Excel, per the `_note` "Attach bank statement / Xero bank
  rec screenshot" already on that sheet) rather than a gap — flagging
  for awareness, not logging as an "open question" since nothing here
  looks like unintentional/wrong behaviour, just a sheet that's
  deliberately thinner than the others.
- No pytest-style tests still; left as-is per Session 1's reasoning.

**Commits this session:** see git log on `claude/holiday-hardening`.

---

## Session 6 — 2026-09-15 (Tuesday)

**Starting state:** picked up `claude/holiday-hardening` from Session 5
(commit `7b115ba`). `pip install -r requirements.txt` fresh (this
container had no packages installed).

**Ran (all green, no genuine failures — item 1):**
- `test_scenarios/run_scenarios.py` — all 11 existing scenarios pass.
- `test_scenarios/test_unit_helpers.py` — all checks pass.
- `test_scenarios/test_annual_summary.py` — still OK.
- `test_scenarios/test_annual_summary_edge_cases.py` — all 7 checks pass.
- `test_scenarios/test_workbook_builder.py` — all 39 checks pass.
- `pytest` still not installed / no pytest-style tests anywhere — same
  conclusion as every prior session, left as-is.

**New focus this session (item 2) — audited the parts of
`ReconciliationEngine` that read Box 2/Box 7/Box 8/Box 9 (the
EU-acquisitions/supplies boxes), since no prior session had looked at
these specifically (all prior scenario/bug work was on Box 1/4/5/6).**
Found that Box 2 (VAT due on NI-EU acquisitions) is parsed and displayed
correctly but never enters `_vat_control`'s reconstructed-balance
formula, even though it does enter Box 3/Box 5 on the return itself —
see new "Open questions for Keyaan" entry above. Confirmed Box 7/8/9 are
purely display fields (never fed into any reconciliation arithmetic), so
no equivalent gap there. Also re-confirmed (as `build_scenarios.py`
already documents for `scenario_accrual`/`scenario_flat_rate`) that the
engine never branches on `vat_scheme` beyond display — flat-rate and
accrual are relabels of the same reconciliation math, so there's no
separate "flat rate calculation" code path to audit for a scheme-specific
bug.

**Added two new fixture scenarios** (`test_scenarios/build_scenarios.py`
+ wired into `run_scenarios.py`'s `SCENARIOS` list):
- `box2_ni_acquisitions` — Box2=200 (NI-EU acquisition VAT), Box3/Box5
  updated to stay internally consistent with the return, Box1/Box4/TB
  left untouched. Pins down the Box 2 gap above as a regression-tested
  fact (`vat_control_diff` == the Box2=0 baseline, 1563.27) rather than
  just a one-off finding — if a future change makes `_vat_control` start
  referencing Box 2, this assertion will fail and need deliberate
  updating, not silently keep passing.
- `box6_diff_at_tolerance` — same idea as Session 2's
  `control_diff_at_tolerance` (which pins `vat_control_diff` to the
  general £1.00 tolerance boundary) but for `box6_diff` against its own
  wider £5.00 tolerance (`cfg.tol_box6`), using real Xero-shaped figures
  end-to-end rather than only the synthetic `WorkbookBuilder` check
  Session 5 already added. TB nominal 200 (Sales) set to a credit balance
  of 14,981.00 against the base demo's unmodified Box6=14,986.00, landing
  `box6_diff` at exactly 5.00. Confirms the real pipeline's rounding
  doesn't nudge this specific boundary off what `_flag(tol=cfg.tol_box6)`
  considers reconciled, closing the gap Session 3 left open ("rounding at
  VAT-box boundaries" was previously only pinned down for Box 1/4's
  general tolerance, not Box 6's wider one, through the full pipeline).

**No new low-risk bugs found this session (item 4)** — the Box 2 finding
above is a bookkeeping-treatment judgement call, not a mechanical bug
(no off-by-one, no wrong rounding direction, no contradiction of
documented logic — the code's own docstring for `_vat_control` states
exactly the formula it implements), so logged as an open question rather
than fixed, per this task's explicit instruction not to guess at VAT
rules.

**Note on regenerated fixtures (same caveat as every prior session):**
`run_scenarios.py` and the other test files rewrite zip-internal metadata
on every scenario `.xlsx` even when no cell value changes. `git status` +
`git checkout --` after every run this session — final diff is exactly
`test_scenarios/build_scenarios.py`, `test_scenarios/run_scenarios.py`,
the two new scenario directories, and this log.

**Still open for next session:**
- The five "Open questions for Keyaan" above (app.py logic placement,
  `InputValidator` required-file severity, `AnnualSummaryBuilder` Q-label
  positional numbering, File Register unconfigured-file red-flag, and
  this session's new Box 2/`_vat_control` gap) — still need Keyaan's
  input, don't act without them.
- Haven't looked closely yet at `TxnByBoxParser`'s handling of the EU
  supplies/acquisitions transaction sections (`Box 8`/`Box 9` sub-headers
  in the Transactions-by-VAT-Box sheet, if a real Xero export ever
  populates them) — everything checked this session was at the
  `ReconciliationEngine`/`VATBoxes` level, not the parser's row-grouping
  for those specific sections. Worth a look if nothing else stands out.
- No pytest-style tests still; left as-is per Session 1's reasoning.

**Commits this session:** see git log on `claude/holiday-hardening`.
