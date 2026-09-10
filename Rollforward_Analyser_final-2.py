"""
app.py — Workbook Roll-Forward Analyzer (Streamlit dashboard)

Generic spreadsheet roll-forward comparison tool. Point it at a previous
version and a current version of your workbooks (single files or entire
folders) and it produces one consolidated Excel report.

Run with:
    streamlit run app.py
"""

from __future__ import annotations

import io
import re
import time
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import (
    EffortAssumptions,
    build_workbook_report,
    load_workbooks,
    match_workbooks,
    rollforward_sheet_names,
    summarize_workbook,
    # --- ADDED: calibrate the manual-effort minutes-per-driver assumptions
    # against real reported hours instead of the fixed defaults ---
    effort_drivers,
    calibrate_effort_assumptions,
    # --- ADDED: cell-level roll-forward mapping for every roll-forward
    # sheet, not just Input/Process/Output (Req 4) ---
    extract_cell_level_rollforward,
    cell_rollforward_table,
    # --- FIX: this-year-only sheet overview & category roll-forward summary
    # (previously these totals wrongly combined previous + current year
    # sheet counts, and the Input/Process/Output breakdown wrongly used
    # cell-level diffs on one guessed sheet instead of counting sheets) ---
    current_year_overview,
    category_sheet_rollforward_summary,
    classify_current_sheets,
    # --- ADDED: full per-sheet Input/Process/Output detail, for the
    # separate complete Input/Process/Output workbook download ---
    ipo_sheet_detail_table,
    # --- ADDED: per-cell Input/Process/Output changes (incl. formula
    # changes), for the same complete workbook download ---
    ipo_cell_level_changes,
    # --- ADDED: formula-level roll-forward extraction (Previous Year cell
    # formula -> Current Year cell formula) + full raw workbook content
    # extract (every populated cell, every sheet, with sheet reference) ---
    extract_formula_level_rollforward,
    formula_rollforward_table,
    extract_full_workbook_content,
    # --- ADDED: hardened combined export — full content + Input/Process/
    # Output tagging, across every workbook, safe against the illegal-
    # character / duplicate-tab-name / row-limit failures that could make
    # the download button silently fail to appear ---
    extract_full_workbook_content_with_category,
    build_complete_workbook_bundle,
)

# --------------------------------------------------------------------------
# Page setup
# --------------------------------------------------------------------------

st.set_page_config(page_title="Roll-Forward Analyzer", layout="wide", page_icon="📊")

st.title("📊 Workbook Roll-Forward Analyzer")
st.caption("Compares a previous version and a current version of your workbooks, "
           "and reports what was rolled forward.")


def _clean_path_input(raw: str) -> str:
    """Strip stray quotes (common when using 'Copy as path' on Windows) and
    surrounding whitespace/newlines from a pasted path."""
    return raw.strip().strip('"').strip("'").strip()


def scan_folder_for_workbooks(folder: str):
    """Recursively find every .xlsx/.xlsm under `folder`.
    Returns (files_dict, message, level) where level is 'error'|'warning'|'ok'."""
    cleaned = _clean_path_input(folder)
    root = Path(cleaned).expanduser()

    if not root.exists():
        return {}, f"Path does not exist: `{cleaned}`", "error"
    if not root.is_dir():
        return {}, f"That path points to a file, not a folder: `{cleaned}`", "error"

    out = {}
    other_extensions = set()
    for path in sorted(root.rglob("*")):
        if path.is_dir():
            continue
        suffix = path.suffix.lower()
        if suffix in (".xlsx", ".xlsm") and not path.name.startswith("~$"):
            try:
                out[path.name] = path.read_bytes()
            except Exception as exc:  # noqa: BLE001
                st.warning(f"Could not read {path}: {exc}")
        elif suffix:
            other_extensions.add(suffix)

    if out:
        return out, f"Found **{len(out)}** workbook(s) in `{cleaned}`.", "ok"
    if ".xls" in other_extensions:
        return {}, (f"Found 0 `.xlsx`/`.xlsm` files in `{cleaned}`, but there are "
                     f"old-format `.xls` files there. Re-save those as `.xlsx` in "
                     f"Excel, or let me know and I can add `.xls` support."), "warning"
    if other_extensions:
        return {}, (f"Found 0 `.xlsx`/`.xlsm` files in `{cleaned}`. Other file types "
                     f"present: {', '.join(sorted(other_extensions))}."), "warning"
    return {}, f"Folder `{cleaned}` exists but appears to be empty.", "warning"


# --------------------------------------------------------------------------
# Sidebar — inputs
# --------------------------------------------------------------------------

with st.sidebar:
    st.header("1. Provide workbooks")
    input_mode = st.radio(
        "How do you want to supply the workbooks?",
        ["Enter folder paths (scans automatically)", "Upload files manually"],
        index=0,
    )

    prev_files, curr_files = None, None
    prev_bytes_folder, curr_bytes_folder = {}, {}

    if input_mode.startswith("Enter folder paths"):
        st.caption("Paste the full path to each folder. Every .xlsx/.xlsm inside "
                   "(including subfolders) is picked up automatically. You can "
                   "also point this at a single file's folder.")
        prev_folder = st.text_input("Previous version — folder path", key="prev_folder")
        curr_folder = st.text_input("Current version — folder path", key="curr_folder")
        if prev_folder:
            prev_bytes_folder, prev_msg, prev_level = scan_folder_for_workbooks(prev_folder)
            getattr(st, prev_level if prev_level != "ok" else "success")(prev_msg)
        if curr_folder:
            curr_bytes_folder, curr_msg, curr_level = scan_folder_for_workbooks(curr_folder)
            getattr(st, curr_level if curr_level != "ok" else "success")(curr_msg)
    else:
        prev_files = st.file_uploader(
            "Previous version — file(s)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="prev"
        )
        curr_files = st.file_uploader(
            "Current version — file(s)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="curr"
        )

    st.header("2. Manual-effort assumptions")
    st.caption("Replace with real timings from your process once available — "
               "these drive the hours/savings estimate.")
    min_cell = st.number_input("Minutes per changed cell", 0.1, 30.0, 1.5, 0.1)
    min_formula = st.number_input("Minutes per formula change", 0.1, 60.0, 3.0, 0.5)
    min_review = st.number_input("Minutes per sheet needing review", 1.0, 240.0, 20.0, 1.0)
    min_complex = st.number_input("Minutes per high-complexity sheet", 1.0, 240.0, 15.0, 1.0)

    st.header("3. Speed vs. thoroughness")
    speed_mode = st.select_slider(
        "How much of each sheet to scan",
        options=["Fast (500 rows × 60 cols)", "Balanced (2000 rows × 150 cols)", "Thorough (10000 rows × 300 cols)"],
        value="Balanced (2000 rows × 150 cols)",
    )
    st.caption("If your workbooks have data far beyond typical size, use Thorough — "
               "otherwise Fast/Balanced will be quicker and still cover almost all real files.")

    st.header("4. Run")
    run_clicked = st.button("▶ Run Analysis", type="primary", use_container_width=True)

_SPEED_CAPS = {
    "Fast (500 rows × 60 cols)": (500, 60),
    "Balanced (2000 rows × 150 cols)": (2000, 150),
    "Thorough (10000 rows × 300 cols)": (10000, 300),
}
scan_max_rows, scan_max_cols = _SPEED_CAPS[speed_mode]

assumptions = EffortAssumptions(
    minutes_per_changed_cell=min_cell,
    minutes_per_formula_change=min_formula,
    minutes_per_review_sheet=min_review,
    minutes_per_complex_formula_sheet=min_complex,
)

# --------------------------------------------------------------------------
# Run analysis
# --------------------------------------------------------------------------

if run_clicked:
    if input_mode.startswith("Enter folder paths"):
        prev_bytes = prev_bytes_folder
        curr_bytes = curr_bytes_folder
    else:
        prev_bytes = {f.name: f.read() for f in prev_files} if prev_files else {}
        curr_bytes = {f.name: f.read() for f in curr_files} if curr_files else {}

    if not prev_bytes or not curr_bytes:
        st.warning("Please provide at least one workbook for both previous and current "
                   "(check your folder paths, or upload files).")
    else:
        status_box = st.empty()
        progress_bar = st.progress(0)
        t_start = time.time()

        status_box.info("Opening workbooks...")
        prev_wbs = load_workbooks(prev_bytes)
        curr_wbs = load_workbooks(curr_bytes)

        wb_matches = match_workbooks(list(prev_bytes.keys()), list(curr_bytes.keys()))
        matched_only = [m for m in wb_matches if m.status == "Matched"]
        total = max(len(matched_only), 1)

        wb_sheet_map = {}  # pair_label -> list[SheetMatch]
        for i, m in enumerate(matched_only, start=1):
            elapsed = time.time() - t_start
            avg = elapsed / i if i > 1 else 0
            remaining = avg * (total - i + 1) if avg else None
            eta_txt = f" — est. {remaining:.0f}s remaining" if remaining else ""
            status_box.info(f"Comparing workbook {i}/{total}: "
                             f"**{m.prev_name} → {m.curr_name}**{eta_txt}")

            label = f"{m.prev_name} → {m.curr_name}"
            report = build_workbook_report(label, prev_wbs.get(m.prev_name),
                                            curr_wbs.get(m.curr_name),
                                            max_rows=scan_max_rows, max_cols=scan_max_cols)
            wb_sheet_map[label] = report
            progress_bar.progress(i / total)

        status_box.success(f"Done — analyzed {total} matched workbook(s) "
                            f"in {time.time() - t_start:.1f}s.")
        progress_bar.empty()

        st.session_state["wb_sheet_map"] = wb_sheet_map

        # --- ADDED: keep the loaded workbooks + match list around so the
        # Input/Process/Output cell-reference section below can reuse them
        # without re-parsing the files.
        st.session_state["ipo_prev_wbs"] = prev_wbs
        st.session_state["ipo_curr_wbs"] = curr_wbs
        st.session_state["ipo_wb_matches"] = {f"{m.prev_name} → {m.curr_name}": m for m in matched_only}

# --------------------------------------------------------------------------
# Recompute summaries live whenever the effort sliders move
# --------------------------------------------------------------------------

if "wb_sheet_map" not in st.session_state:
    st.info("Provide both versions' workbooks in the sidebar and click **Run Analysis**.")
    st.stop()

wb_sheet_map = st.session_state["wb_sheet_map"]

# --------------------------------------------------------------------------
# ADDED: calibrate "Estimated Manual Hours" against real reported hours
# (e.g. actual hours from whoever did the workbooks), instead of relying on
# the fixed sidebar-slider defaults. This fits the four minutes-per-driver
# assumptions (per changed cell / per formula change / per review sheet /
# per complex sheet) to best match the actual hours you supply, then that
# fitted set of assumptions can be applied to the whole report below.
# --------------------------------------------------------------------------

st.subheader("🎯 Calibrate \"Estimated Manual Hours\" Against Actual Reported Hours")
st.caption("Enter the real hours reported for workbooks you've already analyzed, and match each one "
           "to the corresponding analyzed workbook below. The tool then solves for the best-fit "
           "\"minutes per changed cell / formula change / review sheet / complex sheet\" that reproduce "
           "those actual hours — replacing guesswork with your own team's real timings.")

_default_calibration_rows = pd.DataFrame([
    {"Reported Workbook Name": "USN5_263a Inventory adj workpaper_2025", "Actual Hours": 16, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_2025_Tangibles", "Actual Hours": 80, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_C1_C6Depreciations_2025", "Actual Hours": 50, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_CapitalizedInterest_2025", "Actual Hours": 20, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_Consol IS, BS support_2025", "Actual Hours": 38, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_D&R-2025", "Actual Hours": 10, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_G and G_2025_Total SEC", "Actual Hours": 32, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_IDC_2025", "Actual Hours": 80, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_NPLH_2025", "Actual Hours": 40, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_PLH_2025", "Actual Hours": 80, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_Tangible J_Reclass analysis_2025", "Actual Hours": 10, "Matched To Analyzed Workbook": ""},
    {"Reported Workbook Name": "USN5_Wkpr_1120_2025", "Actual Hours": 304, "Matched To Analyzed Workbook": ""},
])

with st.expander("⚙️ Calibration data — edit rows, add your own, or clear the defaults", expanded=False):
    calibration_df = st.data_editor(
        _default_calibration_rows,
        num_rows="dynamic",
        use_container_width=True,
        key="calibration_editor",
        column_config={
            "Reported Workbook Name": st.column_config.TextColumn("Reported Workbook Name (from your team)"),
            "Actual Hours": st.column_config.NumberColumn("Actual Hours", min_value=0.0, step=0.5),
            "Matched To Analyzed Workbook": st.column_config.SelectboxColumn(
                "Matched To Analyzed Workbook",
                options=[""] + list(wb_sheet_map.keys()),
                help="Pick which analyzed workbook pair (below) this reported-hours row corresponds to.",
            ),
        },
    )

    calibrate_clicked = st.button("Calibrate Assumptions From These Actuals")

    if calibrate_clicked:
        matched_rows = calibration_df[
            (calibration_df["Matched To Analyzed Workbook"].fillna("") != "")
            & (calibration_df["Actual Hours"].fillna(0) > 0)
        ]
        if len(matched_rows) < 2:
            st.warning("Match at least 2 rows to an analyzed workbook (with hours > 0) to calibrate — "
                       "a single data point isn't enough to fit 4 assumptions reliably.")
        else:
            driver_rows, actual_hours_list, labels_used = [], [], []
            for _, row in matched_rows.iterrows():
                wb_label = row["Matched To Analyzed Workbook"]
                sheets = wb_sheet_map.get(wb_label, [])
                driver_rows.append(effort_drivers(sheets))
                actual_hours_list.append(float(row["Actual Hours"]))
                labels_used.append(wb_label)

            fitted, diagnostics = calibrate_effort_assumptions(driver_rows, actual_hours_list, base=assumptions)
            st.session_state["calibrated_assumptions"] = fitted
            st.session_state["calibration_diagnostics"] = diagnostics
            st.session_state["calibration_labels_used"] = labels_used

    if "calibrated_assumptions" in st.session_state:
        fitted = st.session_state["calibrated_assumptions"]
        diagnostics = st.session_state["calibration_diagnostics"]
        labels_used = st.session_state["calibration_labels_used"]

        st.markdown("**Fitted assumptions (minutes per unit):**")
        fit_c1, fit_c2, fit_c3, fit_c4 = st.columns(4)
        fit_c1.metric("Per changed cell", f'{fitted.minutes_per_changed_cell:.2f} min')
        fit_c2.metric("Per formula change", f'{fitted.minutes_per_formula_change:.2f} min')
        fit_c3.metric("Per review sheet", f'{fitted.minutes_per_review_sheet:.2f} min')
        fit_c4.metric("Per complex sheet", f'{fitted.minutes_per_complex_formula_sheet:.2f} min')

        if diagnostics.get("r_squared") is not None:
            st.caption(f"Fit quality (R²): {diagnostics['r_squared']}  "
                       f"(1.0 = perfect match to actuals; closer to 0 means the four drivers "
                       f"don't explain the actual hours well for this data).")

        diag_df = pd.DataFrame(diagnostics["per_workbook"])
        diag_df.insert(0, "Workbook", labels_used)
        st.dataframe(diag_df, use_container_width=True, hide_index=True)

        use_calibrated = st.checkbox(
            "Use these calibrated assumptions for the report below (instead of the sidebar sliders)",
            value=True, key="use_calibrated_assumptions",
        )
        if use_calibrated:
            assumptions = fitted
            st.info("Report below is using the CALIBRATED assumptions, not the sidebar sliders.")

summary_rows = [summarize_workbook(label, sheets, assumptions)
                for label, sheets in wb_sheet_map.items()]
df_summary = pd.DataFrame(summary_rows)

name_rows = []
for label, sheets in wb_sheet_map.items():
    name_rows.extend(rollforward_sheet_names(label, sheets))
df_names = pd.DataFrame(name_rows)

# --------------------------------------------------------------------------
# Display: workbook summary
# --------------------------------------------------------------------------

st.subheader("Workbook Summary")
st.dataframe(df_summary, use_container_width=True, hide_index=True)

# --------------------------------------------------------------------------
# ADDED: Sheet Overview — Hidden vs Non-Hidden, Total, Roll-Forward vs
# Non-Roll-Forward sheet counts per workbook pair (Requirements 1-3).
# --------------------------------------------------------------------------

st.subheader("🗂️ Sheet Overview — Hidden / Non-Hidden, Roll-Forward Counts (This Year Only)")
st.caption("Hidden sheets are shown ONLY as a plain count here and play no other part in any "
           "report below — they are excluded from every comparison, formula count, populated-cell "
           "count, and Input/Process/Output classification. Roll-forward counts and names below are "
           "scoped to NON-HIDDEN sheets only.")

_ipo_wb_matches_early = st.session_state.get("ipo_wb_matches", {})
_ipo_curr_wbs_early = st.session_state.get("ipo_curr_wbs", {})

overview_rows = []
for label, sheets in wb_sheet_map.items():
    m = _ipo_wb_matches_early.get(label)
    curr_name = m.curr_name if m else None
    wb_curr = _ipo_curr_wbs_early.get(curr_name) if curr_name else None
    overview_rows.append(current_year_overview(label, curr_name, wb_curr, sheets))
df_overview = pd.DataFrame(overview_rows)

if not df_overview.empty:
    total_hidden = int(df_overview["Hidden Sheets (This Year)"].sum())
    total_nonhidden = int(df_overview["Non-Hidden Sheets (This Year)"].sum())
    total_sheets_all = int(df_overview["Total Sheets (This Year, incl. Hidden)"].sum())
    total_rf_all = int(df_overview["Non-Hidden Sheets Rolled Forward"].sum())
    total_new_all = int(df_overview["Non-Hidden New This Year (Not Rolled Forward)"].sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Sheets (incl. Hidden)", total_sheets_all)
    c2.metric("Hidden Sheets", total_hidden)
    c3.metric("Non-Hidden Sheets", total_nonhidden)
    c4.metric("Non-Hidden Rolled Forward", total_rf_all)
    c5.metric("Non-Hidden New This Year", total_new_all)

# Display table without the raw sheet-name column (shown via expanders below instead)
_overview_display_cols = [c for c in df_overview.columns if c != "Non-Hidden Rolled-Forward Sheet Names"]
st.dataframe(df_overview[_overview_display_cols] if not df_overview.empty else df_overview,
             use_container_width=True, hide_index=True)

if not df_overview.empty:
    for _, row in df_overview.iterrows():
        names = row["Non-Hidden Rolled-Forward Sheet Names"]
        if names:
            with st.expander(f"📄 {row['Workbook']} — {row['Non-Hidden Sheets Rolled Forward']} "
                              f"non-hidden sheet(s) rolled forward"):
                for n in names.split(", "):
                    st.markdown(f"- {n}")

st.divider()

# --------------------------------------------------------------------------
# Display: roll-forward sheet names per workbook
# --------------------------------------------------------------------------

st.subheader("Roll-Forward Sheet Names")
if wb_sheet_map:
    chosen = st.selectbox("Choose a workbook", list(wb_sheet_map.keys()))
    rf_rows = rollforward_sheet_names(chosen, wb_sheet_map[chosen])

    left, right = st.columns([1, 2])
    with left:
        st.metric("Roll-Forward Sheets", len(rf_rows))
        if rf_rows:
            st.markdown("**Names:**")
            for i, row in enumerate(rf_rows, 1):
                tag = "  🔁 *(renamed)*" if row["Renamed?"] == "Yes" else ""
                st.markdown(f"{i}. {row['Current Sheet Name']}{tag}")
    with right:
        st.dataframe(pd.DataFrame(rf_rows), use_container_width=True, hide_index=True)

st.divider()
st.dataframe(df_names, use_container_width=True, hide_index=True)

# --------------------------------------------------------------------------
# ADDED: Cell-Level Roll-Forward Representation for EVERY Roll-Forward
# sheet (Requirement 4) — Previous Year cell -> Current Year cell mapping,
# e.g. "Sheet1!A1 -> Sheet1!A1", not limited to Input/Process/Output.
# --------------------------------------------------------------------------

st.subheader("🔁 Cell-Level Roll-Forward Mapping (All Roll-Forward Sheets)")

if wb_sheet_map:
    cell_rf_chosen = st.selectbox(
        "Choose a workbook for cell-level roll-forward detail",
        list(wb_sheet_map.keys()),
        key="cell_rf_chosen_workbook",
    )
    match_info_cell_rf = st.session_state.get("ipo_wb_matches", {}).get(cell_rf_chosen)
    if match_info_cell_rf is None:
        st.info("Re-run the analysis to populate cell-level roll-forward detail.")
        df_cell_rf = pd.DataFrame()
    else:
        cell_rf_mappings = extract_cell_level_rollforward(
            cell_rf_chosen,
            match_info_cell_rf.prev_name,
            match_info_cell_rf.curr_name,
            st.session_state.get("ipo_prev_wbs", {}).get(match_info_cell_rf.prev_name),
            st.session_state.get("ipo_curr_wbs", {}).get(match_info_cell_rf.curr_name),
            wb_sheet_map[cell_rf_chosen],
            max_rows=scan_max_rows,
            max_cols=scan_max_cols,
        )
        if not cell_rf_mappings:
            st.info("No sheets were classified as 'Roll-forward' for this workbook pair.")
            df_cell_rf = pd.DataFrame()
        else:
            for m in cell_rf_mappings:
                trunc_note = " (list truncated — showing a capped sample)" if m.truncated else ""
                with st.expander(f"{m.prev_sheet} → {m.curr_sheet}  "
                                  f"({m.retained_count} cell(s) rolled forward{trunc_note})"):
                    for prev_ref, curr_ref in m.mappings[:200]:
                        st.markdown(f"- Previous Year: `{prev_ref}` → Current Year: `{curr_ref}`")
                    if len(m.mappings) > 200:
                        st.caption(f"...and {len(m.mappings) - 200} more (full list in the download).")
            df_cell_rf = pd.DataFrame(cell_rollforward_table(cell_rf_mappings))
else:
    df_cell_rf = pd.DataFrame()

st.divider()

# --------------------------------------------------------------------------
# ADDED (per user request): Formula-Level Roll-Forward Extraction —
# for every roll-forward cell that holds a FORMULA, show the Previous Year
# cell reference and formula alongside the Current Year cell reference and
# formula, e.g. "Previous Year: Sheet1!B2 (=A2*1.1) -> Current Year:
# Sheet1!B2 (=A2*1.1)". Reuses the same workbook selected above for
# cell-level roll-forward detail.
# --------------------------------------------------------------------------

st.subheader("🧮 Formula Roll-Forward Mapping (Previous Year Formula → Current Year Formula)")
st.caption("For every Roll-forward sheet, every cell that holds a formula in either year, showing "
           "the exact formula text at the Previous Year cell and at the Current Year cell it rolled "
           "forward to — including where the formula changed between years.")

if wb_sheet_map:
    match_info_formula_rf = st.session_state.get("ipo_wb_matches", {}).get(cell_rf_chosen)
    if match_info_formula_rf is None:
        st.info("Re-run the analysis to populate formula-level roll-forward detail.")
        df_formula_rf = pd.DataFrame()
    else:
        formula_rf_mappings = extract_formula_level_rollforward(
            cell_rf_chosen,
            match_info_formula_rf.prev_name,
            match_info_formula_rf.curr_name,
            st.session_state.get("ipo_prev_wbs", {}).get(match_info_formula_rf.prev_name),
            st.session_state.get("ipo_curr_wbs", {}).get(match_info_formula_rf.curr_name),
            wb_sheet_map[cell_rf_chosen],
            max_rows=scan_max_rows,
            max_cols=scan_max_cols,
        )
        if not formula_rf_mappings:
            st.info("No formula cells found on the roll-forward sheets for this workbook pair.")
            df_formula_rf = pd.DataFrame()
        else:
            for m in formula_rf_mappings:
                trunc_note = " (list truncated — showing a capped sample)" if m.truncated else ""
                with st.expander(f"{m.prev_sheet} → {m.curr_sheet}  "
                                  f"({m.formula_count} formula cell(s){trunc_note})"):
                    for prev_ref, curr_ref, prev_formula, curr_formula, status in m.mappings[:200]:
                        tag = "🔁" if status == "Rolled Forward" else "✏️"
                        st.markdown(
                            f"- {tag} Previous Year: `{prev_ref}` = `{prev_formula}` "
                            f"→ Current Year: `{curr_ref}` = `{curr_formula}`  _( {status} )_"
                        )
                    if len(m.mappings) > 200:
                        st.caption(f"...and {len(m.mappings) - 200} more (full list in the download).")
            df_formula_rf = pd.DataFrame(formula_rollforward_table(formula_rf_mappings))
else:
    df_formula_rf = pd.DataFrame()

st.divider()

# --------------------------------------------------------------------------
# ADDED: Input / Process / Output roll-forward (cell-reference level)
# Extracts, for the workbook's own Input / Process / Output sheets (as laid
# out in the workbook map), what was Replaced vs Retained, split further
# into Formula changes vs General info changes, plus the actual A1-style
# cell ranges that were rolled forward — e.g. "A1 = ROLLFORWARD to J4 in
# the Report sheet".
# --------------------------------------------------------------------------

st.subheader("🔁 Input / Process / Output Roll-Forward (This Year's Sheets)")
st.caption("Every sheet in the CURRENT-year workbook is classified as Input, Process, "
           "or Output (using the sheet name, falling back to its content) and counted. "
           "'Rollforward' = that sheet also existed last year; 'not rollforwarded' = "
           "new this year. Sheet names can vary — classification isn't limited to exact "
           "matches from the reference list.")

ipo_curr_wbs = st.session_state.get("ipo_curr_wbs", {})
ipo_prev_wbs = st.session_state.get("ipo_prev_wbs", {})
ipo_wb_matches = st.session_state.get("ipo_wb_matches", {})

# --------------------------------------------------------------------------
# ADDED: category classification + full per-sheet detail, across EVERY
# workbook, computed once up front so it can feed (a) the aggregate
# Input/Process/Output totals + chart just below, (b) the per-workbook
# selector further down, and (c) the complete IPO workbook download.
# --------------------------------------------------------------------------

categories_by_workbook = {}
ipo_detail_rows_all = []
for label, sheets in wb_sheet_map.items():
    match_info = ipo_wb_matches.get(label)
    wb_curr_for_label = ipo_curr_wbs.get(match_info.curr_name) if match_info else None
    cats = classify_current_sheets(wb_curr_for_label) if wb_curr_for_label else {}
    categories_by_workbook[label] = cats
    ipo_detail_rows_all.extend(ipo_sheet_detail_table(label, sheets, wb_curr_for_label))
df_ipo_detail_all = pd.DataFrame(ipo_detail_rows_all)

st.markdown("**Portfolio totals — Input / Process / Output, across every analyzed workbook**")
if not df_ipo_detail_all.empty:
    totals_by_cat = (
        df_ipo_detail_all.groupby("Category")
        .agg(Total=("Sheet Name", "count"),
             Rollforward=("Rollforward Status", lambda s: (s == "Rollforward").sum()))
        .reindex(["Input", "Process", "Output", "Unclassified"])
        .fillna(0).astype(int).reset_index()
    )
    totals_by_cat["Not Rollforwarded (New This Year)"] = totals_by_cat["Total"] - totals_by_cat["Rollforward"]

    tot_c1, tot_c2, tot_c3 = st.columns(3)
    tot_c1.metric("Total Input Sheets", int(totals_by_cat.loc[totals_by_cat["Category"] == "Input", "Total"].sum()))
    tot_c2.metric("Total Process Sheets", int(totals_by_cat.loc[totals_by_cat["Category"] == "Process", "Total"].sum()))
    tot_c3.metric("Total Output Sheets", int(totals_by_cat.loc[totals_by_cat["Category"] == "Output", "Total"].sum()))

    st.bar_chart(totals_by_cat.set_index("Category")[["Total", "Rollforward"]])
    st.dataframe(totals_by_cat, use_container_width=True, hide_index=True)
else:
    st.info("Run the analysis to see portfolio-wide Input/Process/Output totals.")

st.divider()

if wb_sheet_map:
    ipo_chosen = st.selectbox(
        "Choose a workbook for Input/Process/Output roll-forward detail",
        list(wb_sheet_map.keys()),
        key="ipo_chosen_workbook",
    )
    match_info = ipo_wb_matches.get(ipo_chosen)
    if match_info is None:
        st.info("Re-run the analysis to populate Input/Process/Output detail.")
        df_ipo = pd.DataFrame()
    else:
        wb_curr_for_ipo = ipo_curr_wbs.get(match_info.curr_name)
        ipo_rows = category_sheet_rollforward_summary(
            ipo_chosen, wb_sheet_map[ipo_chosen], wb_curr_for_ipo,
        )
        if not ipo_rows:
            st.info("Could not read the current-year workbook for this pair.")
            df_ipo = pd.DataFrame()
        else:
            df_ipo = pd.DataFrame(ipo_rows)
            st.dataframe(df_ipo, use_container_width=True, hide_index=True)

            # Which sheets landed in each category, and which of those are new
            categories = categories_by_workbook.get(ipo_chosen, {})
            by_curr = {sm.curr_sheet: sm for sm in wb_sheet_map[ipo_chosen] if sm.curr_sheet}
            for category in ("Input", "Process", "Output", "Unclassified"):
                names = [n for n, c in categories.items()
                         if c == category or (category == "Unclassified" and c is None)]
                if not names:
                    continue
                with st.expander(f"{category} — {len(names)} sheet(s) this year"):
                    for n in sorted(names):
                        sm = by_curr.get(n)
                        tag = "🔁 Rollforward" if (sm and sm.prev_sheet) else "🆕 New this year"
                        st.markdown(f"- **{n}** — {tag}")
else:
    df_ipo = pd.DataFrame()

# --------------------------------------------------------------------------
# ADDED: per-cell Input/Process/Output changes, including formula changes,
# across EVERY workbook — the "what cells will be rolled forward, and what
# formula changes happened" detail. Feeds the complete IPO download below.
# --------------------------------------------------------------------------

ipo_cell_change_rows_all = []
for label, sheets in wb_sheet_map.items():
    match_info = ipo_wb_matches.get(label)
    wb_curr_for_label = ipo_curr_wbs.get(match_info.curr_name) if match_info else None
    wb_prev_for_label = ipo_prev_wbs.get(match_info.prev_name) if match_info else None
    ipo_cell_change_rows_all.extend(ipo_cell_level_changes(
        label, wb_prev_for_label, wb_curr_for_label, sheets, categories_by_workbook.get(label, {}),
    ))
df_ipo_cell_changes_all = pd.DataFrame(ipo_cell_change_rows_all)


# --------------------------------------------------------------------------

st.subheader("⬇️ Download Full Report")
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df_summary.to_excel(writer, sheet_name="Workbook_Summary", index=False)
    # --- ADDED: hidden/non-hidden + roll-forward vs non-roll-forward counts ---
    if 'df_overview' in dir() and isinstance(df_overview, pd.DataFrame) and not df_overview.empty:
        df_overview.to_excel(writer, sheet_name="Sheet_Overview", index=False)
    df_names.to_excel(writer, sheet_name="Rollforward_Sheet_Names", index=False)
    # --- ADDED: cell-level Previous Year -> Current Year mapping, all roll-forward sheets ---
    if 'df_cell_rf' in dir() and isinstance(df_cell_rf, pd.DataFrame) and not df_cell_rf.empty:
        df_cell_rf.to_excel(writer, sheet_name="Cell_Rollforward_Mapping", index=False)
    # --- FIX: Input/Process/Output roll-forward, this-year sheet counts only ---
    if 'df_ipo' in dir() and isinstance(df_ipo, pd.DataFrame) and not df_ipo.empty:
        df_ipo.to_excel(writer, sheet_name="IPO_Rollforward_ThisYear", index=False)
    # --- ADDED: formula-level Previous Year formula -> Current Year formula
    # mapping, for every roll-forward sheet's formula cells ---
    if 'df_formula_rf' in dir() and isinstance(df_formula_rf, pd.DataFrame) and not df_formula_rf.empty:
        df_formula_rf.to_excel(writer, sheet_name="Formula_Rollforward_Mapping", index=False)

st.download_button(
    "⬇️ Download Rollforward_Analysis.xlsx",
    data=buffer.getvalue(),
    file_name="Rollforward_Analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# --------------------------------------------------------------------------
# ADDED (per user request): full raw workbook content extract — every
# populated cell, from every sheet, in BOTH the Previous Year and Current
# Year workbook of every matched pair, each row carrying its own sheet
# reference (Workbook / Year / File / Sheet / Sheet Reference). This is a
# straight content dump, not a diff, so the entire workbook content can be
# pulled into one Excel file on demand.
# --------------------------------------------------------------------------

st.subheader("⬇️ Download Entire Workbook Content (All Sheets, With Sheet Reference)")
st.caption("Every populated cell from every sheet of every matched workbook pair — both the "
           "Previous Year and Current Year files — one row per cell, each tagged with its own "
           "Workbook / Year / File / Sheet / Sheet Reference (e.g. `Sheet1!B2`).")

full_content_rows_all = []
for label, sheets in wb_sheet_map.items():
    match_info_full = ipo_wb_matches.get(label)
    if match_info_full is None:
        continue
    wb_prev_for_full = ipo_prev_wbs.get(match_info_full.prev_name)
    wb_curr_for_full = ipo_curr_wbs.get(match_info_full.curr_name)
    full_content_rows_all.extend(extract_full_workbook_content(
        label,
        match_info_full.prev_name, match_info_full.curr_name,
        wb_prev_for_full, wb_curr_for_full,
        max_rows=scan_max_rows, max_cols=scan_max_cols,
    ))
df_full_content_all = pd.DataFrame(full_content_rows_all)

if not df_full_content_all.empty:
    full_content_buffer = io.BytesIO()
    with pd.ExcelWriter(full_content_buffer, engine="openpyxl") as full_content_writer:
        # One tab per workbook pair (Excel sheet-name length limit applies),
        # each tab holding every sheet's content for that pair, sheet-tagged.
        for label in df_full_content_all["Workbook"].unique():
            df_pair = df_full_content_all[df_full_content_all["Workbook"] == label]
            safe_name = re.sub(r'[\\/*?:\[\]]', "_", str(label))[:31]
            df_pair.to_excel(full_content_writer, sheet_name=safe_name or "Workbook", index=False)
        # Plus one combined "All_Content" tab with every workbook pair together.
        df_full_content_all.to_excel(full_content_writer, sheet_name="All_Content", index=False)

    st.download_button(
        "⬇️ Download Full_Workbook_Content.xlsx",
        data=full_content_buffer.getvalue(),
        file_name="Full_Workbook_Content.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Run the analysis to enable this download.")

# --------------------------------------------------------------------------
# ADDED: separate, complete Input/Process/Output workbook — every sheet
# from every workbook, split into its own tab by category, so each one can
# be drilled into individually. Kept as its own download (not merged into
# the report above) since it's a different kind of deliverable.
# --------------------------------------------------------------------------

st.subheader("⬇️ Download Complete Input / Process / Output Workbook")
st.caption("Every sheet from every workbook pair, split into its own tab by category "
           "(Input / Process / Output / Unclassified) — full detail, not just counts. Also "
           "includes a portfolio-wide totals tab and a cell-level tab showing exactly what will "
           "roll forward as-is vs what changed, including formula changes.")

if not df_ipo_detail_all.empty:
    ipo_buffer = io.BytesIO()
    with pd.ExcelWriter(ipo_buffer, engine="openpyxl") as ipo_writer:
        # Portfolio-wide Input/Process/Output totals (Requirement: totals + visual, in Excel too)
        totals_export = (
            df_ipo_detail_all.groupby("Category")
            .agg(Total=("Sheet Name", "count"),
                 Rollforward=("Rollforward Status", lambda s: (s == "Rollforward").sum()))
            .reindex(["Input", "Process", "Output", "Unclassified"])
            .fillna(0).astype(int).reset_index()
        )
        totals_export["Not Rollforwarded (New This Year)"] = totals_export["Total"] - totals_export["Rollforward"]
        totals_export.to_excel(ipo_writer, sheet_name="Totals_By_Category", index=False)

        df_ipo_detail_all.to_excel(ipo_writer, sheet_name="All_Sheets", index=False)
        for category in ("Input", "Process", "Output", "Unclassified"):
            df_cat = df_ipo_detail_all[df_ipo_detail_all["Category"] == category]
            if not df_cat.empty:
                df_cat.to_excel(ipo_writer, sheet_name=category, index=False)

        # --- ADDED: cell-level detail (retained cells + formula/value changes),
        # tagged by category -- "what cells will be rolled forward, incl. formula changes"
        if not df_ipo_cell_changes_all.empty:
            df_ipo_cell_changes_all.to_excel(ipo_writer, sheet_name="Cell_Level_Changes", index=False)
            for category in ("Input", "Process", "Output", "Unclassified"):
                df_cat_cells = df_ipo_cell_changes_all[df_ipo_cell_changes_all["Category"] == category]
                if not df_cat_cells.empty:
                    sheet_name = f"{category}_Cell_Changes"[:31]  # Excel sheet-name length limit
                    df_cat_cells.to_excel(ipo_writer, sheet_name=sheet_name, index=False)

    st.download_button(
        "⬇️ Download Input_Process_Output_Complete.xlsx",
        data=ipo_buffer.getvalue(),
        file_name="Input_Process_Output_Complete.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Run the analysis to enable this download.")

# ==========================================================================
# ADDED (per user request — nothing above this line was changed):
#
# Combined, hardened download: the entire raw workbook content (every
# populated cell, every sheet, both years, every matched workbook pair —
# same data as "Download Entire Workbook Content" above) PLUS an Input /
# Process / Output tab for every workbook, in one file.
#
# This is deliberately built with engine.build_complete_workbook_bundle()
# instead of a plain pd.ExcelWriter block, because the plain block above
# has no protection against three things that will silently kill the
# Streamlit script before st.download_button ever runs (which is exactly
# what "no download button appeared" looks like from the outside):
#   - a cell value containing a character Excel's file format forbids
#   - two workbook-pair names truncating to the same 31-character tab name
#   - a tab that ends up needing more than Excel's ~1.05M row limit
# build_complete_workbook_bundle() sanitizes, dedupes, and splits for all
# three, and returns a list of warnings instead of failing outright, which
# are surfaced below so it's clear if anything had to be adjusted.
# ==========================================================================

st.subheader("⬇️ Download Complete Workbook (All Content + Input/Process/Output, All Workbooks)")
st.caption("Every populated cell from every sheet of every matched workbook pair — both years — "
            "in one file: an All_Content tab, an Input / Process / Output / Unclassified tab, and "
            "one tab per workbook pair. Hardened against the failure modes that can silently stop "
            "the export from producing a download button (illegal characters, duplicate tab names, "
            "Excel's per-sheet row limit).")

complete_bundle_rows: list = []
for label, sheets in wb_sheet_map.items():
    match_info_bundle = ipo_wb_matches.get(label)
    if match_info_bundle is None:
        continue
    wb_prev_for_bundle = ipo_prev_wbs.get(match_info_bundle.prev_name)
    wb_curr_for_bundle = ipo_curr_wbs.get(match_info_bundle.curr_name)
    try:
        complete_bundle_rows.extend(extract_full_workbook_content_with_category(
            label,
            match_info_bundle.prev_name, match_info_bundle.curr_name,
            wb_prev_for_bundle, wb_curr_for_bundle,
            max_rows=scan_max_rows, max_cols=scan_max_cols,
        ))
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Could not extract content for **{label}**: {exc}")

if complete_bundle_rows:
    try:
        complete_bundle_bytes, complete_bundle_warnings = build_complete_workbook_bundle(complete_bundle_rows)
    except Exception as exc:  # noqa: BLE001
        complete_bundle_bytes, complete_bundle_warnings = b"", []
        st.error(f"Could not build the combined workbook: {exc}")
    else:
        for w in complete_bundle_warnings:
            st.caption(f"ℹ️ {w}")
        if complete_bundle_bytes:
            st.download_button(
                "⬇️ Download Complete_Workbook_Content.xlsx",
                data=complete_bundle_bytes,
                file_name="Complete_Workbook_Content.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Nothing to export yet.")
else:
    st.info("Run the analysis to enable this download.")
