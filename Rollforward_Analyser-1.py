"""
app.py — Workbook Roll-Forward Analyzer (Streamlit dashboard)

Generic spreadsheet roll-forward comparison tool. Point it at a previous
version and a current version of your workbooks (single files or entire
folders) and it produces one consolidated Excel report.

Run with:
    streamlit run app.py
"""

from __future__ import annotations

import hashlib
import io
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
    # --- ADDED: Input/Process/Output cell-reference roll-forward extraction ---
    extract_ipo_rollforward,
    ipo_rollforward_table,
    # --- ADDED: cell-level roll-forward mapping for every roll-forward
    # sheet, not just Input/Process/Output (Req 4) ---
    extract_cell_level_rollforward,
    cell_rollforward_table,
)

# --- ADDED: some engine versions (e.g. engine2244) do not ship workbook_sheet_overview();
# fall back to an equivalent built on the engine's own sheet_visibility() helper. ---
try:
    from engine import workbook_sheet_overview
except ImportError:
    from engine import sheet_visibility as _sheet_visibility

    def workbook_sheet_overview(label, prev_name, curr_name, wb_prev, wb_curr, sheets):
        pv = _sheet_visibility(wb_prev, prev_name, "Previous")
        cv = _sheet_visibility(wb_curr, curr_name, "Current")
        total = pv.hidden_count + pv.visible_count + cv.hidden_count + cv.visible_count
        n_rf = sum(1 for s_ in sheets if s_.status == "Roll-forward")
        return {
            "Workbook": label,
            "Hidden Sheets (Previous)": pv.hidden_count,
            "Non-Hidden Sheets (Previous)": pv.visible_count,
            "Hidden Sheets (Current)": cv.hidden_count,
            "Non-Hidden Sheets (Current)": cv.visible_count,
            "Total Sheets": total,
            "Total Roll-Forward Sheets": n_rf,
            "Total Non-Roll-Forward Sheets": total - n_rf,
        }

# --- ADDED: USN5 Input/Process(Calc)/Output classification + fast formula analysis ---
from formula_flow import scan_workbook, analyse_rollforward, match_retained, CATS
from usn5_classifier import classify_workbook, load_usn5_reference

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


@st.cache_resource(show_spinner=False, max_entries=8)
def _cached_scan_impl(_data: bytes, digest: str, max_rows: int, max_cols: int):
    return scan_workbook(_data, max_rows=max_rows, max_cols=max_cols)


def _cached_scan(data: bytes, max_rows: int, max_cols: int):
    """Streaming scan of every visible sheet, cached on file content (re-runs are instant)."""
    return _cached_scan_impl(data, hashlib.md5(data).hexdigest(), max_rows, max_cols)


def _engine_pairs(rf_rows, prev_scans, curr_scans):
    """(previous, current, how) tuples for the engine's roll-forward sheets, or None."""
    pairs = []
    for r in rf_rows or []:
        cur = r.get("Current Sheet Name")
        prv = r.get("Previous Sheet Name") or r.get("Prior Sheet Name") or cur
        if cur in curr_scans and prv in prev_scans:
            pairs.append((prv, cur, "same name" if prv == cur else "renamed"))
    return pairs or None


def _fx_frames(fx_results):
    """Flatten per-workbook results into export/display DataFrames (adds a Workbook column)."""
    frames = {k: [] for k in ("class", "main", "formulas", "incoming", "outgoing", "manual", "sheets", "edges")}
    for label, res in fx_results.items():
        if "report" not in res:
            continue
        r = res["report"]
        for key, rows in (("class", res["class_rows"]), ("main", r.summary_main),
                          ("formulas", r.summary_formulas), ("incoming", r.summary_incoming),
                          ("outgoing", r.summary_outgoing), ("manual", r.summary_manual),
                          ("sheets", r.per_sheet), ("edges", r.edges)):
            df = pd.DataFrame(rows)
            if not df.empty:
                df.insert(0, "Workbook", label)
                frames[key].append(df)
    return {k: (pd.concat(v, ignore_index=True) if v else pd.DataFrame()) for k, v in frames.items()}


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

    # --- ADDED: USN5 classification + formula analysis options ---
    st.header("3b. USN5 classification & formula analysis")
    usn5_ref_file = st.file_uploader(
        "USN5 Data Quantification file (optional)", type=["xlsx"], key="usn5_ref",
        help="If supplied, sheets already listed in it reuse its Classification + Reason "
             "(matched by workbook & sheet name, year-insensitive). Any other sheet is "
             "classified by the built-in USN5 rule engine.")
    fx_full_scan = st.checkbox(
        "Scan full sheets for formula analysis (recommended)", value=True,
        help="Formula scan streams the sheet XML (no openpyxl load) so full depth is still fast. "
             "Untick to reuse the row/column caps above.")
    fx_retained_mode = st.radio(
        "Retained sheets =",
        ["Roll-forward sheets from the analysis engine", "All sheets present in both versions"],
        index=0)

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

        # ------------------------------------------------------------------
        # --- ADDED: USN5 classification + roll-forward formula analysis.
        # One streaming XML pass per workbook (cached by file content), so this
        # adds only seconds even for workbooks with 400k+ cells.
        # ------------------------------------------------------------------
        fx_status = st.empty()
        fx_bar = st.progress(0)
        fx_t0 = time.time()
        usn5_ref = None
        if usn5_ref_file is not None:
            try:
                usn5_ref = load_usn5_reference(usn5_ref_file.getvalue())
            except Exception as exc:  # noqa: BLE001
                st.warning(f"Could not read the USN5 reference file ({exc}); using the rule engine only.")
        fx_rows, fx_caps = (1_048_576, 16_384) if fx_full_scan else (scan_max_rows, scan_max_cols)
        fx_results = {}
        for i, m in enumerate(matched_only, start=1):
            label = f"{m.prev_name} → {m.curr_name}"
            fx_status.info(f"Classifying & scanning formulas {i}/{len(matched_only)}: **{label}**")
            try:
                prev_sc = _cached_scan(prev_bytes[m.prev_name], fx_rows, fx_caps)
                curr_sc = _cached_scan(curr_bytes[m.curr_name], fx_rows, fx_caps)
                class_rows = classify_workbook(m.curr_name, curr_sc, usn5_ref)
                classes = {r["Sheet Name"]: r["Classification"] for r in class_rows}
                pairs = None
                if fx_retained_mode.startswith("Roll-forward"):
                    pairs = _engine_pairs(rollforward_sheet_names(label, wb_sheet_map[label]),
                                          prev_sc, curr_sc)
                    if pairs is None:  # engine gave nothing usable -> fall back to name matching
                        pairs = match_retained(list(prev_sc), list(curr_sc))
                rep_fx = analyse_rollforward(prev_sc, curr_sc, classes, pairs)
                truncated = sorted({n for n, s_ in list(prev_sc.items()) + list(curr_sc.items()) if s_.truncated})
                fx_results[label] = {"class_rows": class_rows, "report": rep_fx, "truncated": truncated,
                                     "n_retained": len(rep_fx.per_sheet)}
            except Exception as exc:  # noqa: BLE001
                fx_results[label] = {"error": f"{type(exc).__name__}: {exc}"}
            fx_bar.progress(i / max(len(matched_only), 1))
        st.session_state["fx_results"] = fx_results
        fx_bar.empty()
        fx_status.success(f"USN5 classification + formula analysis done in {time.time() - fx_t0:.1f}s.")

# --------------------------------------------------------------------------
# Recompute summaries live whenever the effort sliders move
# --------------------------------------------------------------------------

if "wb_sheet_map" not in st.session_state:
    st.info("Provide both versions' workbooks in the sidebar and click **Run Analysis**.")
    st.stop()

wb_sheet_map = st.session_state["wb_sheet_map"]

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

st.subheader("🗂️ Sheet Overview — Hidden / Roll-Forward Counts")

_ipo_wb_matches_early = st.session_state.get("ipo_wb_matches", {})
_ipo_prev_wbs_early = st.session_state.get("ipo_prev_wbs", {})
_ipo_curr_wbs_early = st.session_state.get("ipo_curr_wbs", {})

overview_rows = []
for label, sheets in wb_sheet_map.items():
    m = _ipo_wb_matches_early.get(label)
    prev_name = m.prev_name if m else None
    curr_name = m.curr_name if m else None
    wb_prev = _ipo_prev_wbs_early.get(prev_name) if prev_name else None
    wb_curr = _ipo_curr_wbs_early.get(curr_name) if curr_name else None
    overview_rows.append(workbook_sheet_overview(label, prev_name, curr_name, wb_prev, wb_curr, sheets))
df_overview = pd.DataFrame(overview_rows)

if not df_overview.empty:
    total_hidden = int(df_overview["Hidden Sheets (Previous)"].sum() + df_overview["Hidden Sheets (Current)"].sum())
    total_nonhidden = int(df_overview["Non-Hidden Sheets (Previous)"].sum() + df_overview["Non-Hidden Sheets (Current)"].sum())
    total_sheets_all = int(df_overview["Total Sheets"].sum())
    total_rf_all = int(df_overview["Total Roll-Forward Sheets"].sum())
    total_non_rf_all = int(df_overview["Total Non-Roll-Forward Sheets"].sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Hidden Sheets", total_hidden)
    c2.metric("Total Non-Hidden Sheets", total_nonhidden)
    c3.metric("Total Sheets (All Files)", total_sheets_all)
    c4.metric("Total Roll-Forward Sheets", total_rf_all)
    c5.metric("Total Non-Roll-Forward Sheets", total_non_rf_all)

st.dataframe(df_overview, use_container_width=True, hide_index=True)

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
# ADDED: Input / Process / Output roll-forward (cell-reference level)
# Extracts, for the workbook's own Input / Process / Output sheets (as laid
# out in the workbook map), what was Replaced vs Retained, split further
# into Formula changes vs General info changes, plus the actual A1-style
# cell ranges that were rolled forward — e.g. "A1 = ROLLFORWARD to J4 in
# the Report sheet".
# --------------------------------------------------------------------------

st.subheader("🔁 Input / Process / Output Roll-Forward (Cell References)")

ipo_prev_wbs = st.session_state.get("ipo_prev_wbs", {})
ipo_curr_wbs = st.session_state.get("ipo_curr_wbs", {})
ipo_wb_matches = st.session_state.get("ipo_wb_matches", {})

if wb_sheet_map:
    ipo_chosen = st.selectbox(
        "Choose a workbook for Input/Process/Output roll-forward detail",
        list(wb_sheet_map.keys()),
        key="ipo_chosen_workbook",
    )
    match_info = ipo_wb_matches.get(ipo_chosen)
    if match_info is None:
        st.info("Re-run the analysis to populate Input/Process/Output cell-reference detail.")
    else:
        ipo_details = extract_ipo_rollforward(
            ipo_chosen,
            ipo_prev_wbs.get(match_info.prev_name),
            ipo_curr_wbs.get(match_info.curr_name),
            wb_sheet_map[ipo_chosen],
            max_rows=scan_max_rows,
            max_cols=scan_max_cols,
            report_sheet_name="Report",
        )
        if not ipo_details:
            st.info("No sheets named/matching 'Input', 'Process' or 'Output' were found "
                     "in this workbook pair.")
        else:
            df_ipo = pd.DataFrame(ipo_rollforward_table(ipo_details))
            st.dataframe(df_ipo, use_container_width=True, hide_index=True)

            for d in ipo_details:
                with st.expander(f"{d.category} — cell reference detail "
                                  f"({d.prev_sheet or '—'} → {d.curr_sheet or '—'})"):
                    st.markdown(f"**Retained (rolled forward):** {d.retain_count} cell(s)")
                    if d.retain_ranges:
                        st.markdown(", ".join(d.retain_ranges))
                    st.markdown(f"**Replaced (changed):** {d.replace_count} cell(s)")
                    if d.replace_ranges:
                        st.markdown(", ".join(d.replace_ranges))
                    st.markdown(f"**Formula changes in Retain:** {d.formula_changes_in_retain}")
                    st.markdown(f"**General info changes in Retain:** {d.general_info_changes_in_retain}")
                    if d.report_sheet_note:
                        st.markdown("**Cell reference note:**")
                        st.markdown(d.report_sheet_note)
else:
    df_ipo = pd.DataFrame()

# --------------------------------------------------------------------------
# ADDED: USN5 classification (Input / Process-Calc / Output) with reasons
# --------------------------------------------------------------------------
fx_results = st.session_state.get("fx_results", {})
fx_df = _fx_frames(fx_results)

st.subheader("🏷️ USN5 Classification — Input / Process (Calc) / Output")
if not fx_results:
    st.info("Re-run the analysis to populate the USN5 classification.")
else:
    for lbl, res in fx_results.items():
        if "error" in res:
            st.error(f"{lbl}: {res['error']}")
    df_cls = fx_df["class"]
    if not df_cls.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Visible sheets classified", len(df_cls))
        c2.metric("Input", int((df_cls["Classification"] == "Input").sum()))
        c3.metric("Process (Calc)", int((df_cls["Classification"] == "Process").sum()))
        c4.metric("Output", int((df_cls["Classification"] == "Output").sum()))
        f1, f2 = st.columns([1, 2])
        pick_cls = f1.multiselect("Show classification", ["Input", "Process", "Output"],
                                  default=["Input", "Process", "Output"], key="fx_pick_cls")
        pick_txt = f2.text_input("Filter sheet name / reason contains", key="fx_pick_txt")
        view = df_cls[df_cls["Classification"].isin(pick_cls)]
        if pick_txt:
            mask = (view["Sheet Name"].astype(str).str.contains(pick_txt, case=False, regex=False)
                    | view["Classification Reason"].astype(str).str.contains(pick_txt, case=False, regex=False))
            view = view[mask]
        st.dataframe(view, use_container_width=True, hide_index=True,
                     column_config={"Classification Reason": st.column_config.TextColumn(width="large")})
        st.caption("Basis = 'USN5 reference file' when the sheet was found in the uploaded USN5 file, "
                   "otherwise 'Rule engine' (mandatory Process override → Output → defined Input "
                   "systems → generic workpaper rule). Only visible sheets are classified.")

# --------------------------------------------------------------------------
# ADDED: Roll-forward (retained) sheets — formulas, manual, incoming / outgoing
# --------------------------------------------------------------------------
st.subheader("🧮 Retained (Roll-Forward) Sheets — Formula, Manual, Incoming & Outgoing Analysis")
ok_labels = [k for k, v in fx_results.items() if "report" in v]
if not ok_labels:
    st.info("Run the analysis to populate the formula analysis.")
else:
    fx_pick = st.selectbox("Choose a workbook", ok_labels, key="fx_pick_wb")
    res = fx_results[fx_pick]
    rep_fx = res["report"]
    if res["truncated"]:
        st.warning("These sheets were larger than the scan limit, so counts are partial: "
                   + ", ".join(res["truncated"][:15]) + (" …" if len(res["truncated"]) > 15 else "")
                   + ". Tick 'Scan full sheets' in the sidebar and re-run.")
    st.caption(f"{res['n_retained']} retained sheet(s) analysed · "
               f"{len(rep_fx.not_retained_prev)} previous-only · {len(rep_fx.not_retained_curr)} current-only sheet(s).")

    def _show(title, rows):
        st.markdown(f"**{title}**")
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    _show("Retained sheets analysis — Summary", rep_fx.summary_main)
    _show("Retained sheets analysis — Formulas", rep_fx.summary_formulas)
    _show("Retained sheets analysis — Formulas — Incoming analysis", rep_fx.summary_incoming)
    _show("Retained sheets analysis — Formulas — Outgoing analysis", rep_fx.summary_outgoing)
    _show("Retained sheets analysis — Formulas — Manual", rep_fx.summary_manual)

    with st.expander("Sheet-by-sheet detail (incoming / outgoing sheet names & types)", expanded=False):
        cat_pick = st.multiselect("Category", CATS, default=CATS, key="fx_cat_pick")
        df_sh = pd.DataFrame(rep_fx.per_sheet)
        if not df_sh.empty:
            st.dataframe(df_sh[df_sh["Category"].isin(cat_pick)], use_container_width=True, hide_index=True)
    with st.expander("Formula flow — provider sheet → consumer sheet", expanded=False):
        if rep_fx.edges:
            st.dataframe(pd.DataFrame(rep_fx.edges), use_container_width=True, hide_index=True)
        else:
            st.info("No cross-sheet formulas touching retained sheets were found.")
    with st.expander("Definitions used"):
        st.markdown(
            "- **Incoming formula** — formula inside sheet *S* that reads another sheet (or an external workbook).\n"
            "- **Outgoing formula** — formula in another sheet that reads sheet *S*.\n"
            "- **Manual** — hard-typed numeric constant (text labels are not counted).\n"
            "- **Changes** — compared cell-by-cell at the same address, previous vs current: modified + added + removed. "
            "References to a renamed sheet (e.g. `'Calc 2024'!A1` → `'Calc 2025'!A1`) are **not** counted as changes.\n"
            "- Calc = sheets classified *Process* by the USN5 logic.\n"
            "- Not resolved: `INDIRECT()` text references and defined names that point to other sheets.")

# --------------------------------------------------------------------------
# Excel export — exactly two sheets, everything combined
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
    # --- ADDED: Input/Process/Output roll-forward cell-reference sheet ---
    if 'df_ipo' in dir() and isinstance(df_ipo, pd.DataFrame) and not df_ipo.empty:
        df_ipo.to_excel(writer, sheet_name="IPO_Rollforward_CellRefs", index=False)
    # --- ADDED: USN5 classification + formula / incoming / outgoing analysis ---
    for _key, _sheet in (("class", "USN5_Classification"), ("main", "RF_Summary"),
                         ("formulas", "RF_Formulas"), ("incoming", "RF_Incoming"),
                         ("outgoing", "RF_Outgoing"), ("manual", "RF_Manual"),
                         ("sheets", "RF_Sheet_Detail"), ("edges", "RF_Formula_Flow")):
        if not fx_df[_key].empty:
            fx_df[_key].to_excel(writer, sheet_name=_sheet, index=False)

st.download_button(
    "⬇️ Download Rollforward_Analysis.xlsx",
    data=buffer.getvalue(),
    file_name="Rollforward_Analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
