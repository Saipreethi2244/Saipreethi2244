"""
engine.py
=========
Core analysis engine for a generic workbook Roll-Forward Analyzer.

This module contains NO UI code. It is pure logic so it can be unit-tested
and reused outside of Streamlit if needed. Nothing here is specific to any
industry, company, or file-naming convention — it works on any pair of
"previous version" / "current version" spreadsheet folders.

Pipeline
--------
1. load_workbooks()          -> read raw bytes into openpyxl Workbook objects
2. match_workbooks()         -> pair up previous-version / current-version files
3. match_sheets()            -> pair up sheets within a matched workbook
                                 (handles renames, e.g. "Summary Calculation"
                                 -> "Summary Calc")
4. compare_sheet_cells()     -> cell-level diff (unchanged/changed/new/cleared)
                                 and formula-level diff, incl. a dedicated
                                 "year/period-reference-only" change category
                                 (e.g. '2024 Data'!F20 -> '2025 Data'!F20)
5. sheet_complexity()        -> formula counts, cross-sheet/external refs,
                                 expressed as a 0-100 complexity percentage
6. build_workbook_report()   -> ties everything together per workbook pair
7. summarize_workbook()      -> one consolidated row per workbook: sheet
                                 counts, roll-forward count, formula counts,
                                 populated cells, complexity %, automation
                                 opportunity, and manual-effort hours — all
                                 in a single place (nothing reported separately)
8. estimate_manual_effort()  -> converts diff counts into hours, using
                                 user-adjustable minute assumptions
9. automation_opportunity()  -> High / Medium / Low label per sheet
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

YEAR_TOKEN = re.compile(r"(19|20)\d{2}")


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

def _strip_year(name: str) -> str:
    """Replace 4-digit years with a placeholder so '...2024...' and
    '...2025...' compare as similar names/formulas."""
    return YEAR_TOKEN.sub("YYYY", name)


def _similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    return SequenceMatcher(None, _strip_year(a).lower(), _strip_year(b).lower()).ratio()


def is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _cell_is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


# --------------------------------------------------------------------------
# Data classes
# --------------------------------------------------------------------------

@dataclass
class WorkbookMatch:
    prev_name: Optional[str]
    curr_name: Optional[str]
    match_pct: float
    status: str  # "Matched" | "Previous Only" | "Current Only"


@dataclass
class CellDiff:
    unchanged: int = 0
    changed: int = 0
    new: int = 0
    cleared: int = 0
    value_changed: int = 0  # subset of `changed` where NEITHER side is a formula
    # (i.e. a plain "general info" value change, not a formula change)
    changed_samples: List[Tuple[str, object, object]] = field(default_factory=list)

    @property
    def total_populated(self) -> int:
        return self.unchanged + self.changed + self.new + self.cleared

    @property
    def unchanged_ratio(self) -> float:
        denom = self.unchanged + self.changed
        return self.unchanged / denom if denom else 1.0


@dataclass
class FormulaDiff:
    unchanged: int = 0
    changed: int = 0
    year_ref_change: int = 0
    new: int = 0
    removed: int = 0
    changed_samples: List[Tuple[str, str, str]] = field(default_factory=list)

    @property
    def total(self) -> int:
        return self.unchanged + self.changed + self.year_ref_change + self.new + self.removed

    @property
    def unchanged_ratio(self) -> float:
        denom = self.unchanged + self.year_ref_change + self.changed
        return (self.unchanged + self.year_ref_change) / denom if denom else 1.0


@dataclass
class ComplexityInfo:
    formula_count: int = 0
    cross_sheet_refs: int = 0
    external_refs: int = 0
    nested_formulas: int = 0
    max_formula_length: int = 0
    merged_cells: int = 0
    hidden: bool = False
    percentage: float = 0.0  # 0-100 complexity score, normalized


@dataclass
class SheetMatch:
    workbook_pair: str
    prev_sheet: Optional[str]
    curr_sheet: Optional[str]
    name_similarity: float
    structure_similarity: float
    cell_diff: CellDiff
    formula_diff: FormulaDiff
    complexity: ComplexityInfo
    roll_forward_score: float
    status: str  # "Roll-forward" | "Review" | "New" | "Removed"
    automation: str  # "High" | "Medium" | "Low"
    renamed: bool = False


# --------------------------------------------------------------------------
# Step 1: loading
# --------------------------------------------------------------------------

def load_workbooks(files: Dict[str, bytes], read_only: bool = True) -> Dict[str, openpyxl.Workbook]:
    """files: {filename: raw_bytes}. Returns {filename: Workbook}.

    read_only=True (default) is noticeably faster and lighter on memory for
    large files, since openpyxl doesn't materialize every cell as an
    editable object. The only cost is that ws.merged_cells isn't available
    in read-only mode; sheet_complexity() degrades gracefully (reports 0)
    when that happens rather than failing."""
    out = {}
    for name, raw in files.items():
        try:
            out[name] = openpyxl.load_workbook(BytesIO(raw), data_only=False, read_only=read_only)
        except Exception as exc:  # noqa: BLE001
            out[name] = None
            print(f"Could not open {name}: {exc}")
    return out


# --------------------------------------------------------------------------
# Step 2: workbook matching
# --------------------------------------------------------------------------

def match_workbooks(prev_names: List[str], curr_names: List[str],
                     threshold: float = 0.5) -> List[WorkbookMatch]:
    pairs = []
    for p in prev_names:
        for c in curr_names:
            pairs.append((p, c, _similarity(p, c)))
    pairs.sort(key=lambda x: x[2], reverse=True)

    used_prev, used_curr = set(), set()
    matches: List[WorkbookMatch] = []
    for p, c, score in pairs:
        if p in used_prev or c in used_curr:
            continue
        if score >= threshold:
            used_prev.add(p)
            used_curr.add(c)
            matches.append(WorkbookMatch(p, c, round(score * 100, 1), "Matched"))

    for p in prev_names:
        if p not in used_prev:
            matches.append(WorkbookMatch(p, None, 0.0, "Previous Only"))
    for c in curr_names:
        if c not in used_curr:
            matches.append(WorkbookMatch(None, c, 0.0, "Current Only"))

    matches.sort(key=lambda m: (m.status != "Matched", -(m.match_pct)))
    return matches


# --------------------------------------------------------------------------
# Step 3: sheet matching (handles renames)
# --------------------------------------------------------------------------

def _sheet_signature(ws: Worksheet, max_rows=200, max_cols=60):
    """Cheap structural fingerprint: dims, populated-cell count, header row."""
    max_r = min(ws.max_row or 1, max_rows)
    max_c = min(ws.max_column or 1, max_cols)
    populated = 0
    header_vals = []
    for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True):
        for v in row:
            if not _cell_is_blank(v):
                populated += 1
    if ws.max_row:
        header_vals = [str(v) for v in next(
            ws.iter_rows(min_row=1, max_row=1, max_col=max_c, values_only=True), []
        ) if not _cell_is_blank(v)]
    return (ws.max_row or 0, ws.max_column or 0, populated, header_vals)


def _structure_similarity(sig1, sig2) -> float:
    r1, c1, pop1, hdr1 = sig1
    r2, c2, pop2, hdr2 = sig2
    dim_sim = 1 - (abs(r1 - r2) / max(r1, r2, 1) + abs(c1 - c2) / max(c1, c2, 1)) / 2
    pop_sim = 1 - abs(pop1 - pop2) / max(pop1, pop2, 1)
    hdr_sim = _similarity(" ".join(hdr1), " ".join(hdr2))
    return max(0.0, 0.3 * dim_sim + 0.3 * pop_sim + 0.4 * hdr_sim)


def _visible_sheetnames(wb: Optional[openpyxl.Workbook]) -> List[str]:
    """Sheet names with openpyxl sheet_state == 'visible' only -- excludes
    'hidden' and 'veryHidden' tabs. Used everywhere sheets are matched/
    analyzed so hidden sheets never enter any comparison, count, or export
    other than the explicit Hidden vs Non-Hidden totals in
    current_year_overview()."""
    if wb is None:
        return []
    return [s for s in wb.sheetnames if getattr(wb[s], "sheet_state", "visible") == "visible"]


def match_sheets(wb_prev: openpyxl.Workbook, wb_curr: openpyxl.Workbook,
                  name_threshold: float = 0.35) -> List[dict]:
    """Returns list of dicts: {prev, curr, name_sim, struct_sim, renamed}.
    Hidden sheets (on either side) are excluded entirely -- they are never
    matched, compared, or reported on anywhere except the Hidden/Non-Hidden
    counts in current_year_overview()."""
    prev_sheets = _visible_sheetnames(wb_prev)
    curr_sheets = _visible_sheetnames(wb_curr)

    sigs_prev = {s: _sheet_signature(wb_prev[s]) for s in prev_sheets}
    sigs_curr = {s: _sheet_signature(wb_curr[s]) for s in curr_sheets}

    candidates = []
    for p in prev_sheets:
        for c in curr_sheets:
            n_sim = _similarity(p, c)
            s_sim = _structure_similarity(sigs_prev[p], sigs_curr[c])
            combined = 0.5 * n_sim + 0.5 * s_sim
            candidates.append((p, c, n_sim, s_sim, combined))
    candidates.sort(key=lambda x: x[4], reverse=True)

    used_p, used_c = set(), set()
    results = []
    for p, c, n_sim, s_sim, combined in candidates:
        if p in used_p or c in used_c:
            continue
        # exact name match always wins even if structure drifted a lot
        if p == c or combined >= name_threshold:
            used_p.add(p)
            used_c.add(c)
            results.append({
                "prev": p, "curr": c, "name_sim": n_sim, "struct_sim": s_sim,
                "renamed": (p != c and n_sim < 0.98),
            })

    for p in prev_sheets:
        if p not in used_p:
            results.append({"prev": p, "curr": None, "name_sim": 0.0,
                             "struct_sim": 0.0, "renamed": False})
    for c in curr_sheets:
        if c not in used_c:
            results.append({"prev": None, "curr": c, "name_sim": 0.0,
                             "struct_sim": 0.0, "renamed": False})
    return results


# --------------------------------------------------------------------------
# Step 4 & 5: cell + formula comparison
# --------------------------------------------------------------------------

def classify_formula_pair(f1: str, f2: str) -> str:
    if f1 == f2:
        return "unchanged"
    if _strip_year(f1) == _strip_year(f2):
        return "year_ref_change"
    return "changed"


def compare_sheet_cells(ws_prev: Optional[Worksheet], ws_curr: Optional[Worksheet],
                         max_rows=2000, max_cols=150,
                         sample_cap=300) -> Tuple[CellDiff, FormulaDiff]:
    """Row-at-a-time comparison using iter_rows(values_only=True), which is
    far faster than looking up ws.cell(row, col) one cell at a time (that
    approach re-materializes a Cell object per call and is a common
    performance trap for large workbooks)."""
    cell_diff = CellDiff()
    formula_diff = FormulaDiff()

    if ws_prev is None and ws_curr is None:
        return cell_diff, formula_diff

    max_r = max(ws_prev.max_row if ws_prev else 0, ws_curr.max_row if ws_curr else 0)
    max_c = max(ws_prev.max_column if ws_prev else 0, ws_curr.max_column if ws_curr else 0)
    max_r = min(max_r, max_rows)
    max_c = min(max_c, max_cols)
    if max_r == 0 or max_c == 0:
        return cell_diff, formula_diff

    from itertools import zip_longest
    from openpyxl.utils import get_column_letter

    empty_row = (None,) * max_c
    rows_prev = ws_prev.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True) \
        if ws_prev else iter(())
    rows_curr = ws_curr.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True) \
        if ws_curr else iter(())
    col_letters = [get_column_letter(c) for c in range(1, max_c + 1)]

    for r, (row_prev, row_curr) in enumerate(
            zip_longest(rows_prev, rows_curr, fillvalue=empty_row), start=1):
        row_prev = row_prev or empty_row
        row_curr = row_curr or empty_row
        for c in range(max_c):
            v_prev = row_prev[c] if c < len(row_prev) else None
            v_curr = row_curr[c] if c < len(row_curr) else None
            blank_prev, blank_curr = _cell_is_blank(v_prev), _cell_is_blank(v_curr)

            if blank_prev and blank_curr:
                continue

            addr = f"{col_letters[c]}{r}"

            # --- formula-specific tracking ---
            f_prev = v_prev if is_formula(v_prev) else None
            f_curr = v_curr if is_formula(v_curr) else None
            if f_prev or f_curr:
                if f_prev and f_curr:
                    cls = classify_formula_pair(f_prev, f_curr)
                    if cls == "unchanged":
                        formula_diff.unchanged += 1
                    elif cls == "year_ref_change":
                        formula_diff.year_ref_change += 1
                        if len(formula_diff.changed_samples) < sample_cap:
                            formula_diff.changed_samples.append((addr, f_prev, f_curr))
                    else:
                        formula_diff.changed += 1
                        if len(formula_diff.changed_samples) < sample_cap:
                            formula_diff.changed_samples.append((addr, f_prev, f_curr))
                elif f_curr and not f_prev:
                    formula_diff.new += 1
                elif f_prev and not f_curr:
                    formula_diff.removed += 1

            # --- generic cell tracking (values + formulas together) ---
            if blank_prev and not blank_curr:
                cell_diff.new += 1
            elif blank_curr and not blank_prev:
                cell_diff.cleared += 1
            elif v_prev == v_curr:
                cell_diff.unchanged += 1
            else:
                cell_diff.changed += 1
                if not (f_prev or f_curr):
                    cell_diff.value_changed += 1
                if len(cell_diff.changed_samples) < sample_cap:
                    cell_diff.changed_samples.append((addr, v_prev, v_curr))

    return cell_diff, formula_diff


# --------------------------------------------------------------------------
# Step 6: complexity
# --------------------------------------------------------------------------

def sheet_complexity(ws: Optional[Worksheet], max_rows: int = 2000, max_cols: int = 150) -> ComplexityInfo:
    info = ComplexityInfo()
    if ws is None:
        return info
    info.hidden = ws.sheet_state != "visible"
    try:
        info.merged_cells = len(ws.merged_cells.ranges)
    except AttributeError:
        info.merged_cells = 0  # not available on read-only worksheets

    max_r = min(ws.max_row or 1, max_rows)
    max_c = min(ws.max_column or 1, max_cols)
    for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True):
        for v in row:
            if is_formula(v):
                info.formula_count += 1
                info.max_formula_length = max(info.max_formula_length, len(v))
                if "!" in v:
                    info.cross_sheet_refs += 1
                if "[" in v:
                    info.external_refs += 1
                if v.count("(") >= 2:
                    info.nested_formulas += 1

    raw_score = info.formula_count + info.cross_sheet_refs * 2 + info.external_refs * 3
    # Normalize to a 0-100 percentage; a raw score of 300+ is treated as maximum
    # complexity (100%). Tune CAP below if your workbooks run much larger/smaller.
    CAP = 300
    info.percentage = round(min(raw_score / CAP, 1.0) * 100, 1)
    return info


# --------------------------------------------------------------------------
# Step 7-9: roll-forward score, automation opportunity, effort
# --------------------------------------------------------------------------

def roll_forward_score(struct_sim: float, cell_diff: CellDiff, formula_diff: FormulaDiff) -> float:
    score = (0.3 * struct_sim
             + 0.4 * cell_diff.unchanged_ratio
             + 0.3 * formula_diff.unchanged_ratio)
    return round(score * 100, 1)


def classify_status(score: float, has_both_sides: bool) -> str:
    if not has_both_sides:
        return "New"
    if score >= 80:
        return "Roll-forward"
    if score >= 50:
        return "Review"
    return "Review"  # low similarity but still matched -> needs human review


def automation_opportunity(score: float, formula_diff: FormulaDiff, complexity: ComplexityInfo) -> str:
    """High: strong roll-forward pattern & regular formulas.
       Medium: some repeated pattern but real changes exist.
       Low: irregular / heavily judgment-based."""
    is_high_complexity = complexity.percentage >= 70
    if score >= 85 and not is_high_complexity:
        return "High"
    if score >= 85 and is_high_complexity and formula_diff.unchanged_ratio >= 0.7:
        return "High"
    if score >= 55:
        return "Medium"
    return "Low"


@dataclass
class EffortAssumptions:
    minutes_per_changed_cell: float = 1.5
    minutes_per_formula_change: float = 3.0
    minutes_per_review_sheet: float = 20.0
    minutes_per_complex_formula_sheet: float = 15.0
    automation_pct_high: float = 0.85
    automation_pct_medium: float = 0.45
    automation_pct_low: float = 0.10


# ADDED: pulled out of estimate_manual_effort so the same four "effort
# driver" counts can be reused by the calibration routine below without
# duplicating the counting logic.
def effort_drivers(sheet_matches: List[SheetMatch]) -> dict:
    return {
        "changed_cells": sum(sm.cell_diff.changed + sm.cell_diff.new + sm.cell_diff.cleared
                              for sm in sheet_matches),
        "formula_changes": sum(sm.formula_diff.changed + sm.formula_diff.year_ref_change
                                + sm.formula_diff.new + sm.formula_diff.removed
                                for sm in sheet_matches),
        "review_sheets": sum(1 for sm in sheet_matches if sm.status == "Review"),
        "complex_sheets": sum(1 for sm in sheet_matches if sm.complexity.percentage >= 70),
    }


def estimate_manual_effort(sheet_matches: List[SheetMatch], assumptions: EffortAssumptions) -> dict:
    drivers = effort_drivers(sheet_matches)
    changed_cells = drivers["changed_cells"]
    formula_changes = drivers["formula_changes"]
    review_sheets = drivers["review_sheets"]
    complex_sheets = drivers["complex_sheets"]

    minutes = (changed_cells * assumptions.minutes_per_changed_cell
               + formula_changes * assumptions.minutes_per_formula_change
               + review_sheets * assumptions.minutes_per_review_sheet
               + complex_sheets * assumptions.minutes_per_complex_formula_sheet)
    manual_hours = minutes / 60

    high = sum(1 for sm in sheet_matches if sm.automation == "High")
    medium = sum(1 for sm in sheet_matches if sm.automation == "Medium")
    low = sum(1 for sm in sheet_matches if sm.automation == "Low")
    total_auto_sheets = max(high + medium + low, 1)
    blended_automation_pct = (
        high * assumptions.automation_pct_high
        + medium * assumptions.automation_pct_medium
        + low * assumptions.automation_pct_low
    ) / total_auto_sheets

    potential_hours_saved = manual_hours * blended_automation_pct

    return {
        "changed_cells": changed_cells,
        "formula_changes": formula_changes,
        "review_sheets": review_sheets,
        "complex_sheets": complex_sheets,
        "estimated_manual_hours": round(manual_hours, 1),
        "potential_automation_pct": round(blended_automation_pct * 100, 1),
        "potential_hours_saved": round(potential_hours_saved, 1),
        "high_automation_sheets": high,
        "medium_automation_sheets": medium,
        "low_automation_sheets": low,
    }


# ADDED: calibrate the four minutes-per-driver assumptions above against
# real reported hours (e.g. actuals from the people who did the work),
# instead of relying on the fixed defaults (1.5 / 3.0 / 20.0 / 15.0).
# Solves:  minutes_per_changed_cell * changed_cells
#        + minutes_per_formula_change * formula_changes
#        + minutes_per_review_sheet * review_sheets
#        + minutes_per_complex_formula_sheet * complex_sheets  ~=  actual_minutes
# as a non-negative least squares fit across every workbook supplied, so
# the fitted assumptions are the best-fit "typical minutes per unit of
# work" implied by the actuals, and can then be reused to recompute
# Estimated Manual Hours for every workbook going forward.
def calibrate_effort_assumptions(driver_rows: List[dict], actual_hours: List[float],
                                  base: Optional[EffortAssumptions] = None) -> Tuple[EffortAssumptions, dict]:
    """driver_rows: one dict per workbook, each with the four effort_drivers()
    keys (changed_cells, formula_changes, review_sheets, complex_sheets).
    actual_hours: the reported actual hours for that same workbook, same order.
    Returns (fitted EffortAssumptions, diagnostics dict) -- diagnostics
    includes per-workbook predicted-vs-actual hours and an R^2 fit quality
    so the fit can be sanity-checked before being applied."""
    import numpy as np

    keys = ["changed_cells", "formula_changes", "review_sheets", "complex_sheets"]
    X = np.array([[row.get(k, 0) for k in keys] for row in driver_rows], dtype=float)
    y = np.array(actual_hours, dtype=float) * 60.0  # actual hours -> actual minutes

    if X.shape[0] < 1 or X.shape[1] == 0:
        return (base or EffortAssumptions()), {"error": "No data to calibrate against."}

    try:
        from scipy.optimize import nnls
        coeffs, _residual = nnls(X, y)
    except Exception:
        # Fallback if scipy isn't available: ordinary least squares, then
        # clip to non-negative (a driver can't have *negative* minutes).
        coeffs, *_ = np.linalg.lstsq(X, y, rcond=None)
        coeffs = np.clip(coeffs, 0, None)

    fitted = EffortAssumptions(
        minutes_per_changed_cell=round(float(coeffs[0]), 3),
        minutes_per_formula_change=round(float(coeffs[1]), 3),
        minutes_per_review_sheet=round(float(coeffs[2]), 3),
        minutes_per_complex_formula_sheet=round(float(coeffs[3]), 3),
        automation_pct_high=(base or EffortAssumptions()).automation_pct_high,
        automation_pct_medium=(base or EffortAssumptions()).automation_pct_medium,
        automation_pct_low=(base or EffortAssumptions()).automation_pct_low,
    )

    predicted_minutes = X @ coeffs
    predicted_hours = predicted_minutes / 60.0
    actual_hours_arr = np.array(actual_hours, dtype=float)
    ss_res = float(np.sum((actual_hours_arr - predicted_hours) ** 2))
    ss_tot = float(np.sum((actual_hours_arr - actual_hours_arr.mean()) ** 2)) if len(actual_hours_arr) > 1 else 0.0
    r_squared = round(1 - ss_res / ss_tot, 3) if ss_tot > 0 else None

    diagnostics = {
        "r_squared": r_squared,
        "per_workbook": [
            {
                "actual_hours": round(float(a), 1),
                "predicted_hours": round(float(p), 1),
                "difference": round(float(p) - float(a), 1),
            }
            for a, p in zip(actual_hours, predicted_hours)
        ],
    }
    return fitted, diagnostics


# --------------------------------------------------------------------------
# Orchestration for one matched workbook pair
# --------------------------------------------------------------------------

def build_workbook_report(pair_label: str,
                           wb_prev: Optional[openpyxl.Workbook],
                           wb_curr: Optional[openpyxl.Workbook],
                           max_rows: int = 2000, max_cols: int = 150) -> List[SheetMatch]:
    sheet_pairs = match_sheets(wb_prev, wb_curr) if (wb_prev or wb_curr) else []
    results: List[SheetMatch] = []

    for sp in sheet_pairs:
        ws_prev = wb_prev[sp["prev"]] if (wb_prev and sp["prev"]) else None
        ws_curr = wb_curr[sp["curr"]] if (wb_curr and sp["curr"]) else None

        cell_diff, formula_diff = compare_sheet_cells(ws_prev, ws_curr,
                                                        max_rows=max_rows, max_cols=max_cols)
        complexity = sheet_complexity(ws_curr or ws_prev, max_rows=max_rows, max_cols=max_cols)
        has_both = ws_prev is not None and ws_curr is not None

        if not has_both and ws_prev is not None:
            status = "Removed"
            score = 0.0
        elif not has_both and ws_curr is not None:
            status = "New"
            score = 0.0
        else:
            score = roll_forward_score(sp["struct_sim"], cell_diff, formula_diff)
            status = classify_status(score, has_both)

        auto = automation_opportunity(score, formula_diff, complexity) if has_both else "Low"

        results.append(SheetMatch(
            workbook_pair=pair_label,
            prev_sheet=sp["prev"],
            curr_sheet=sp["curr"],
            name_similarity=round(sp["name_sim"] * 100, 1),
            structure_similarity=round(sp["struct_sim"] * 100, 1),
            cell_diff=cell_diff,
            formula_diff=formula_diff,
            complexity=complexity,
            roll_forward_score=score,
            status=status,
            automation=auto,
            renamed=sp["renamed"],
        ))
    return results


# --------------------------------------------------------------------------
# Consolidated per-workbook summary (single row, everything combined)
# --------------------------------------------------------------------------

def summarize_workbook(pair_label: str, sheet_matches: List[SheetMatch],
                        assumptions: "EffortAssumptions") -> dict:
    """One consolidated dict per workbook: sheet counts, roll-forward count,
    formula counts, populated cells, complexity %, automation opportunity,
    and manual-effort hours — all together, nothing split out separately."""
    total_prev_sheets = sum(1 for s in sheet_matches if s.prev_sheet)
    total_curr_sheets = sum(1 for s in sheet_matches if s.curr_sheet)
    rolled_forward = [s for s in sheet_matches if s.status == "Roll-forward"]

    total_formulas = sum(s.complexity.formula_count for s in sheet_matches)
    formulas_in_rollforward = sum(s.complexity.formula_count for s in rolled_forward)
    populated_cells = sum(s.cell_diff.total_populated for s in sheet_matches)

    avg_complexity_pct = (
        round(sum(s.complexity.percentage for s in sheet_matches) / len(sheet_matches), 1)
        if sheet_matches else 0.0
    )

    effort = estimate_manual_effort(sheet_matches, assumptions)

    return {
        "Workbook": pair_label,
        "Total Sheets (Previous)": total_prev_sheets,
        "Total Sheets (Current)": total_curr_sheets,
        "Sheets Rolled Forward": len(rolled_forward),
        "Total Formulas": total_formulas,
        "Formulas in Rolled-Forward Sheets": formulas_in_rollforward,
        "Populated Cells": populated_cells,
        "Complexity %": avg_complexity_pct,
        "Automation Opportunity %": effort["potential_automation_pct"],
        "Estimated Manual Hours": effort["estimated_manual_hours"],
        "Potential Hours Saved": effort["potential_hours_saved"],
    }


def rollforward_sheet_names(pair_label: str, sheet_matches: List[SheetMatch]) -> List[dict]:
    """One row per rolled-forward sheet: workbook + previous/current sheet name."""
    return [{
        "Workbook": pair_label,
        "Previous Sheet Name": s.prev_sheet,
        "Current Sheet Name": s.curr_sheet,
        "Renamed?": "Yes" if s.renamed else "No",
    } for s in sheet_matches if s.status == "Roll-forward"]


# --------------------------------------------------------------------------
# ADDED: Input / Process / Output roll-forward extraction (cell-reference
# level). This section is purely additive -- nothing above this point was
# changed. It targets the workbook's own "Input", "Process" and "Output"
# sheets (as laid out in the workbook map) and reports, per sheet:
#   - how many cells were REPLACED vs RETAINED (rolled forward)
#   - the compact A1-style cell range(s) that were retained/replaced
#   - of the retained cells, how many carried a formula change vs a plain
#     general-info (value) change
#   - a plain-English cell-reference note in the requested format, e.g.
#     "A1 = ROLLFORWARD to J4 in the Report sheet"
# --------------------------------------------------------------------------

IPO_LABELS = ("Input", "Process", "Output")


def _is_ipo_sheet(name: Optional[str], label: str) -> bool:
    """Loose matcher: sheet literally named 'Input'/'Process'/'Output', or
    containing that word (case-insensitive) -- e.g. 'Input Data', 'Process
    (Tax Tools)', 'Output - Filings'."""
    if not name:
        return False
    return label.lower() in name.lower()


def _compress_to_ranges(cells: List[Tuple[int, int]]) -> List[str]:
    """Turn a list of (row, col) cell coordinates into a compact list of
    A1-style ranges. Cells are grouped into contiguous same-column-span
    row-runs, then stacked vertically where the column span repeats on
    consecutive rows. Good enough for reporting purposes (not a general
    rectangle-packing solver)."""
    from openpyxl.utils import get_column_letter
    if not cells:
        return []

    by_row: Dict[int, List[int]] = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    for r in by_row:
        by_row[r].sort()

    row_runs: List[Tuple[int, int, int]] = []
    for r, cols in by_row.items():
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
                continue
            row_runs.append((r, start, prev))
            start = prev = c
        row_runs.append((r, start, prev))

    by_colspan: Dict[Tuple[int, int], List[int]] = {}
    for (r, cs, ce) in row_runs:
        by_colspan.setdefault((cs, ce), []).append(r)

    ranges: List[Tuple[int, int, int, int]] = []
    for (cs, ce), rows in by_colspan.items():
        rows = sorted(rows)
        start_r = prev_r = rows[0]
        for rr in rows[1:]:
            if rr == prev_r + 1:
                prev_r = rr
                continue
            ranges.append((start_r, cs, prev_r, ce))
            start_r = prev_r = rr
        ranges.append((start_r, cs, prev_r, ce))

    out = []
    for (r1, c1, r2, c2) in ranges:
        a1 = f"{get_column_letter(c1)}{r1}"
        a2 = f"{get_column_letter(c2)}{r2}"
        out.append(a1 if a1 == a2 else f"{a1}:{a2}")
    return sorted(out)


@dataclass
class IPORollforwardDetail:
    workbook_pair: str
    category: str  # "Input" | "Process" | "Output"
    prev_sheet: Optional[str]
    curr_sheet: Optional[str]
    replace_count: int
    retain_count: int
    retain_ranges: List[str] = field(default_factory=list)
    replace_ranges: List[str] = field(default_factory=list)
    formula_changes_in_retain: int = 0
    formula_change_refs: List[str] = field(default_factory=list)
    general_info_changes_in_retain: int = 0
    general_info_change_refs: List[str] = field(default_factory=list)
    report_sheet_note: str = ""


def extract_ipo_rollforward(pair_label: str,
                             wb_prev: Optional[openpyxl.Workbook],
                             wb_curr: Optional[openpyxl.Workbook],
                             sheet_matches: List[SheetMatch],
                             max_rows: int = 2000, max_cols: int = 150,
                             report_sheet_name: str = "Report") -> List[IPORollforwardDetail]:
    """For the workbook's Input / Process / Output sheets, walk cell-by-cell
    and report which cells were RETAINED (rolled forward, unchanged) vs
    REPLACED (changed), as compact A1-style cell ranges, plus a split of the
    replaced cells into 'formula changes' vs 'general info changes' --
    mirroring the Rollforward summary table (Replace / Retain / Formula
    changes in Retain / General info changes in Retain)."""
    from openpyxl.utils import get_column_letter
    from itertools import zip_longest

    details: List[IPORollforwardDetail] = []

    for category in IPO_LABELS:
        # --- CHANGED: uses the keyword-aware matcher (Requirement 5) so
        # abbreviations, full names, and content-only references are still
        # recognized, not just an exact/substring sheet-name match.
        sm = _find_ipo_sheet_match(category, sheet_matches, wb_prev, wb_curr)
        if sm is None:
            continue

        ws_prev = wb_prev[sm.prev_sheet] if (wb_prev and sm.prev_sheet) else None
        ws_curr = wb_curr[sm.curr_sheet] if (wb_curr and sm.curr_sheet) else None
        if ws_prev is None and ws_curr is None:
            continue

        max_r = min(max(ws_prev.max_row if ws_prev else 0, ws_curr.max_row if ws_curr else 0), max_rows)
        max_c = min(max(ws_prev.max_column if ws_prev else 0, ws_curr.max_column if ws_curr else 0), max_cols)
        if max_r == 0 or max_c == 0:
            continue

        empty_row = (None,) * max_c
        rows_prev = ws_prev.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True) if ws_prev else iter(())
        rows_curr = ws_curr.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True) if ws_curr else iter(())

        retain_cells: List[Tuple[int, int]] = []
        replace_cells: List[Tuple[int, int]] = []
        formula_change_cells: List[Tuple[int, int]] = []
        general_change_cells: List[Tuple[int, int]] = []

        for r, (row_prev, row_curr) in enumerate(
                zip_longest(rows_prev, rows_curr, fillvalue=empty_row), start=1):
            row_prev = row_prev or empty_row
            row_curr = row_curr or empty_row
            for c in range(max_c):
                v_prev = row_prev[c] if c < len(row_prev) else None
                v_curr = row_curr[c] if c < len(row_curr) else None
                blank_prev, blank_curr = _cell_is_blank(v_prev), _cell_is_blank(v_curr)
                if blank_prev and blank_curr:
                    continue

                col = c + 1
                if v_prev == v_curr:
                    retain_cells.append((r, col))
                else:
                    replace_cells.append((r, col))
                    if is_formula(v_prev) or is_formula(v_curr):
                        formula_change_cells.append((r, col))
                    else:
                        general_change_cells.append((r, col))

        retain_ranges = _compress_to_ranges(retain_cells)
        replace_ranges = _compress_to_ranges(replace_cells)
        formula_refs = [f"{get_column_letter(c)}{r}" for (r, c) in formula_change_cells]
        general_refs = [f"{get_column_letter(c)}{r}" for (r, c) in general_change_cells]

        note_parts = []
        for rng in retain_ranges:
            if ":" in rng:
                start_addr, end_addr = rng.split(":")
                note_parts.append(f"{start_addr} = ROLLFORWARD to {end_addr} in the {report_sheet_name} sheet")
            else:
                note_parts.append(f"{rng} = ROLLFORWARD in the {report_sheet_name} sheet")
        note = "; ".join(note_parts)

        details.append(IPORollforwardDetail(
            workbook_pair=pair_label,
            category=category,
            prev_sheet=sm.prev_sheet,
            curr_sheet=sm.curr_sheet,
            replace_count=len(replace_cells),
            retain_count=len(retain_cells),
            retain_ranges=retain_ranges,
            replace_ranges=replace_ranges,
            formula_changes_in_retain=len(formula_change_cells),
            formula_change_refs=formula_refs,
            general_info_changes_in_retain=len(general_change_cells),
            general_info_change_refs=general_refs,
            report_sheet_note=note,
        ))

    return details


# --------------------------------------------------------------------------
# ADDED (Requirement 1 & 2): Hidden vs Non-Hidden sheet counts, and overall
# sheet counts (total / roll-forward / non-roll-forward) per workbook pair.
# --------------------------------------------------------------------------

@dataclass
class SheetVisibilitySummary:
    file_label: str  # "Previous" | "Current"
    file_name: Optional[str]
    hidden_sheets: List[str] = field(default_factory=list)
    visible_sheets: List[str] = field(default_factory=list)

    @property
    def hidden_count(self) -> int:
        return len(self.hidden_sheets)

    @property
    def visible_count(self) -> int:
        return len(self.visible_sheets)


def sheet_visibility(wb: Optional[openpyxl.Workbook], file_name: Optional[str],
                      file_label: str) -> SheetVisibilitySummary:
    """Splits a workbook's sheets into Hidden vs Non-Hidden (visible),
    using openpyxl's sheet_state ('visible' / 'hidden' / 'veryHidden')."""
    hidden, visible = [], []
    if wb is not None:
        for name in wb.sheetnames:
            state = getattr(wb[name], "sheet_state", "visible")
            (hidden if state != "visible" else visible).append(name)
    return SheetVisibilitySummary(file_label, file_name, hidden, visible)


# --------------------------------------------------------------------------
# ADDED (Requirement 4): Cell-level roll-forward mapping for EVERY
# roll-forward sheet (not just Input/Process/Output) — records the exact
# Previous-Year cell -> Current-Year cell for every retained/unchanged
# cell, in the requested "Sheet1!A1 -> Sheet1!A1" form.
# --------------------------------------------------------------------------

@dataclass
class CellRollforwardMapping:
    workbook_pair: str
    prev_file: Optional[str]
    curr_file: Optional[str]
    prev_sheet: str
    curr_sheet: str
    mappings: List[Tuple[str, str]] = field(default_factory=list)  # (prev_ref, curr_ref)
    retained_count: int = 0
    truncated: bool = False


def extract_cell_level_rollforward(pair_label: str,
                                    prev_file: Optional[str], curr_file: Optional[str],
                                    wb_prev: Optional[openpyxl.Workbook],
                                    wb_curr: Optional[openpyxl.Workbook],
                                    sheet_matches: List[SheetMatch],
                                    max_rows: int = 2000, max_cols: int = 150,
                                    max_mappings_per_sheet: int = 1000) -> List[CellRollforwardMapping]:
    """For every sheet classified as 'Roll-forward', walks cell-by-cell and
    records the Previous-Year cell -> Current-Year cell mapping for every
    retained (unchanged) cell -- e.g. 'Sheet1!A1 -> Sheet1!A1'. Capped per
    sheet via max_mappings_per_sheet so a huge sheet doesn't produce an
    unusable wall of references (retained_count still reflects the true
    total; `truncated` flags when the cap was hit)."""
    from itertools import zip_longest
    from openpyxl.utils import get_column_letter

    out: List[CellRollforwardMapping] = []
    for sm in sheet_matches:
        if sm.status != "Roll-forward" or not (sm.prev_sheet and sm.curr_sheet):
            continue
        ws_prev = wb_prev[sm.prev_sheet] if wb_prev else None
        ws_curr = wb_curr[sm.curr_sheet] if wb_curr else None
        if ws_prev is None or ws_curr is None:
            continue

        max_r = min(max(ws_prev.max_row or 0, ws_curr.max_row or 0), max_rows)
        max_c = min(max(ws_prev.max_column or 0, ws_curr.max_column or 0), max_cols)
        if max_r == 0 or max_c == 0:
            continue

        empty_row = (None,) * max_c
        rows_prev = ws_prev.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)
        rows_curr = ws_curr.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)

        mapping = CellRollforwardMapping(pair_label, prev_file, curr_file, sm.prev_sheet, sm.curr_sheet)
        for r, (row_prev, row_curr) in enumerate(
                zip_longest(rows_prev, rows_curr, fillvalue=empty_row), start=1):
            row_prev = row_prev or empty_row
            row_curr = row_curr or empty_row
            for c in range(max_c):
                v_prev = row_prev[c] if c < len(row_prev) else None
                v_curr = row_curr[c] if c < len(row_curr) else None
                if _cell_is_blank(v_prev) and _cell_is_blank(v_curr):
                    continue
                if v_prev == v_curr:
                    mapping.retained_count += 1
                    if len(mapping.mappings) < max_mappings_per_sheet:
                        addr = f"{get_column_letter(c + 1)}{r}"
                        mapping.mappings.append((f"{sm.prev_sheet}!{addr}", f"{sm.curr_sheet}!{addr}"))
                    else:
                        mapping.truncated = True
        out.append(mapping)
    return out


def cell_rollforward_table(mappings: List[CellRollforwardMapping]) -> List[dict]:
    """Flattens CellRollforwardMapping objects into the requested
    'Previous Year: Sheet1!A1 -> Current Year: Sheet1!A1' representation,
    one row per rolled-forward cell."""
    rows = []
    for m in mappings:
        for prev_ref, curr_ref in m.mappings:
            rows.append({
                "Workbook": m.workbook_pair,
                "Previous File": m.prev_file,
                "Current File": m.curr_file,
                "Previous Sheet": m.prev_sheet,
                "Current Sheet": m.curr_sheet,
                "Previous Year Cell": prev_ref,
                "Current Year Cell": curr_ref,
                "Roll-Forward Mapping": f"Previous Year: {prev_ref} \u2192 Current Year: {curr_ref}",
            })
    return rows


# ADDED: per-cell Input/Process/Output roll-forward detail, including
# formula changes -- not just the retained/unchanged cells above. For
# every ROLL-FORWARD sheet (any category), walks cell-by-cell and records
# what will roll forward as-is vs what changed (a formula edit, a
# year-reference-only edit, or a plain value edit), tagged with that
# sheet's Input/Process/Output category. This is the row-level answer to
# "what cells will be rolled forward, including formula changes" -- meant
# for the separate complete Input/Process/Output workbook download.
def ipo_cell_level_changes(pair_label: str,
                            wb_prev: Optional[openpyxl.Workbook],
                            wb_curr: Optional[openpyxl.Workbook],
                            sheet_matches: List[SheetMatch],
                            categories: Dict[str, Optional[str]],
                            max_rows: int = 2000, max_cols: int = 150,
                            max_rows_per_sheet: int = 2000) -> List[dict]:
    from itertools import zip_longest
    from openpyxl.utils import get_column_letter

    rows: List[dict] = []
    for sm in sheet_matches:
        if sm.status != "Roll-forward" or not (sm.prev_sheet and sm.curr_sheet):
            continue
        ws_prev = wb_prev[sm.prev_sheet] if wb_prev else None
        ws_curr = wb_curr[sm.curr_sheet] if wb_curr else None
        if ws_prev is None or ws_curr is None:
            continue

        category = categories.get(sm.curr_sheet) or "Unclassified"

        max_r = min(max(ws_prev.max_row or 0, ws_curr.max_row or 0), max_rows)
        max_c = min(max(ws_prev.max_column or 0, ws_curr.max_column or 0), max_cols)
        if max_r == 0 or max_c == 0:
            continue

        empty_row = (None,) * max_c
        rows_prev = ws_prev.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)
        rows_curr = ws_curr.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)

        sheet_row_count = 0
        for r, (row_prev, row_curr) in enumerate(
                zip_longest(rows_prev, rows_curr, fillvalue=empty_row), start=1):
            if sheet_row_count >= max_rows_per_sheet:
                break
            row_prev = row_prev or empty_row
            row_curr = row_curr or empty_row
            for c in range(max_c):
                v_prev = row_prev[c] if c < len(row_prev) else None
                v_curr = row_curr[c] if c < len(row_curr) else None
                if _cell_is_blank(v_prev) and _cell_is_blank(v_curr):
                    continue

                addr = f"{get_column_letter(c + 1)}{r}"
                prev_ref = f"{sm.prev_sheet}!{addr}"
                curr_ref = f"{sm.curr_sheet}!{addr}"

                if v_prev == v_curr:
                    change_type = "Retained (Rolled Forward)"
                elif is_formula(v_prev) or is_formula(v_curr):
                    ftype = classify_formula_pair(str(v_prev or ""), str(v_curr or ""))
                    change_type = {"changed": "Formula Changed",
                                   "year_ref_change": "Formula Changed (Year Reference Only)",
                                   "unchanged": "Retained (Rolled Forward)"}.get(ftype, "Formula Changed")
                else:
                    change_type = "Value Changed"

                rows.append({
                    "Workbook": pair_label,
                    "Category": category,
                    "Previous Sheet": sm.prev_sheet,
                    "Current Sheet": sm.curr_sheet,
                    "Previous Year Cell": prev_ref,
                    "Current Year Cell": curr_ref,
                    "Change Type": change_type,
                    "Previous Value": v_prev,
                    "Current Value": v_curr,
                })
                sheet_row_count += 1
    return rows


# --------------------------------------------------------------------------
# ADDED (Requirement 5): keyword-driven Input/Process/Output classifier.
# Goes beyond an exact/substring sheet-name match: it also recognizes
# abbreviations, full names, and references found *inside* the sheet
# content, using the vocabulary from the reference Input/Process/Output
# picture. Sheet-name matching is tried first (cheap, precise); scanning
# cell content is only used as a fallback when no sheet name matches.
# --------------------------------------------------------------------------

IPO_KEYWORDS: Dict[str, List[str]] = {
    "Input": [
        "input", "book data", "ifrs", "gaap", "sap", "sap bw", "sap/bw", "bw report",
        "s4", "s4 gr", "s/4", "ers", "sac", "sac report", "tdrt", "kas",
        "tmn02", "tmn 02", "non tmn02", "non-tmn02", "list of election", "election",
        "legal entity", "tax law", "tax regulation",
    ],
    "Process": [
        "process", "tax tool", "excel workpaper", "corptax", "longview", "uta",
        "dttu", "soda", "tars", "bna tangible", "bna workpaper", "bna", "ptms",
    ],
    "Output": [
        "output", "annual federal tax return", "federal tax return", "multi-state tax return",
        "multi state tax return", "state tax return", "property tax return",
        "property tax payment", "federal audit", "state audit", "audit requirement",
        "quarterly estimated filing", "estimated filing", "extension filing",
        "tax payment", "tax provision", "group reporting", "dep notice",
        "tax advisory", "tax filing",
    ],
}


def _normalize_token(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


# Tax-return "line" worksheets -- e.g. 'L16', 'L23.3', 'L16 Rent',
# 'L23.3_Depreciation', 'L 16 - Rent'. These never contain the word
# "process" or any of the IPO_KEYWORDS vocabulary, but they are the
# per-line calculation sheets that feed the return -> Process.
_LINE_ITEM_RE = re.compile(r"^l\s*[-_]?\s*\d+(\.\d+)*(\b|[_\s-])", re.IGNORECASE)


def _looks_like_line_item_sheet(name: Optional[str]) -> bool:
    if not name:
        return False
    return bool(_LINE_ITEM_RE.match(name.strip()))


def classify_ipo_by_keywords(name: Optional[str] = None, sample_text: str = "") -> Optional[str]:
    """Returns 'Input' | 'Process' | 'Output' | None. Checks the sheet name
    against the plain label and the keyword vocabulary first (handles short
    names/abbreviations and full names), then a naming-pattern check (e.g.
    'L16', 'L23.3' line-item sheets -> Process), then falls back to scanning
    the supplied cell-content sample for keyword hits, picking whichever
    category scores the most hits."""
    if name:
        for label in IPO_LABELS:
            if label.lower() in name.lower():
                return label
        if _looks_like_line_item_sheet(name):
            return "Process"

    combined = " ".join(_normalize_token(x) for x in (name or "", sample_text) if x)
    if not combined.strip():
        return None

    scores = {label: sum(1 for kw in kws if _normalize_token(kw) in combined)
              for label, kws in IPO_KEYWORDS.items()}
    scores = {k: v for k, v in scores.items() if v > 0}
    if not scores:
        return None
    return max(scores, key=scores.get)


def _sheet_sample_text(ws: Optional[Worksheet], max_rows: int = 100, max_cols: int = 40) -> str:
    """Bounded sample of a sheet's text (cell values + formula text), used
    for keyword scanning. Cheap enough to run per-sheet as a fallback when
    the sheet name alone doesn't identify it as Input/Process/Output."""
    if ws is None:
        return ""
    max_r = min(ws.max_row or 1, max_rows)
    max_c = min(ws.max_column or 1, max_cols)
    parts = []
    for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True):
        for v in row:
            if isinstance(v, str) and v.strip():
                parts.append(v)
    return " ".join(parts)


def _sheet_formula_count(ws: Optional[Worksheet], max_rows: int = 100, max_cols: int = 40) -> int:
    """Bounded count of formula cells in a sheet (used as a 'this tab does
    work' signal for generic-workpaper classification)."""
    if ws is None:
        return 0
    max_r = min(ws.max_row or 1, max_rows)
    max_c = min(ws.max_column or 1, max_cols)
    count = 0
    for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True):
        for v in row:
            if is_formula(v):
                count += 1
    return count


# --------------------------------------------------------------------------
# ADDED: reference-grounded Input/Process/Output classifier. This logic is
# derived directly from a real tax data-quantification workbook's "Applied
# Logic" sheet (source-system priority rules the client already uses to
# classify tabs), rather than an ad-hoc keyword list. It focuses primarily
# on the SHEET NAME (per the reference: "sheet-name priority first"), and
# falls back to scanning visible cell content only when the name alone
# doesn't identify a source system.
#
# Priority, matching the reference logic:
#   1. Identify a "source system" for the sheet -- by name first, then by
#      content indicators (header/label text found on the sheet).
#   2. If that source system is one of the four *mandatory Process*
#      systems (UTA, DTTU, Longview, CorpTax) -> Process, regardless of
#      anything else on the sheet.
#   3. If that source system is one of the *defined Input* systems (SAP,
#      SAP Tax Module, ERS, ERS - APLI, TMN02, KAS, Election Statement,
#      GEMS Extract, Data Request from R&A, Tax Law / Regulations)
#      -> Input.
#   4. No source system identified (a generic workpaper tab): only a
#      named final-deliverable tab (Taxable Income Summary, TI Summary,
#      Filing Package, Audit Package) counts as Output. Otherwise, a tab
#      that calculates, maps, reconciles, allocates, adjusts, validates,
#      combines, or supports schedules (formula cells present, or those
#      action words appear on the sheet) is Process; a tab that looks like
#      plain source data with no such indicators is Input.
# --------------------------------------------------------------------------

# Mandatory-Process source systems: identified by sheet name first. Per
# the reference data, the actual name-priority trigger for CorpTax tabs is
# "TBBS" (a compound name like "Corptax load_2025" turned out to be an
# ERS-sourced Input tab, not a CorpTax working tab) -- so "corptax" itself
# is deliberately not a name trigger here; a literal "CorpTax" name still
# matches via the plain-label check in classify_ipo_by_keywords.
_PROCESS_MANDATORY_NAME_TOKENS: Dict[str, List[str]] = {
    "CorpTax": ["tbbs"],
    "Longview": ["longview", "lv", "a9", "5100"],
    "UTA": ["uta"],
    "DTTU": ["dttu"],
}
# Content indicators (visible cell text) for the same mandatory systems,
# used only when the name doesn't already identify one.
_PROCESS_MANDATORY_CONTENT_INDICATORS: Dict[str, List[str]] = {
    "CorpTax": ["tbbs", "combined balance sheet"],
    "Longview": ["lvl4", "taxcolumns", "timeper", "5100"],
}

# Defined-Input source systems: identified by sheet name first. Checked in
# this order so a more specific token (e.g. "apli") wins over a more
# generic one that could also be present (e.g. "ers").
_INPUT_NAME_TOKENS: Dict[str, List[str]] = {
    "ERS - APLI": ["apli"],
    "ERS": ["ers"],
    "SAP": ["sap"],
    "TMN02": ["tmn02", "tmn 02"],
    "KAS": ["kas"],
    "Election Statement": ["election"],
    "GEMS Extract": ["gems"],
    "Data Request from R&A": ["a202"],
    "Tax Law / Regulations": ["tax law", "tax regulation"],
}
# Content indicators for defined-Input systems, used only when the name
# doesn't already identify one.
_INPUT_CONTENT_INDICATORS: Dict[str, List[str]] = {
    "SAP Tax Module": ["depreciation key", "sub-number", "sub number",
                        "asset description", "asset class"],
    "SAP": ["company code", "account number description", "account number",
            "posting legal entity", "local jv code"],
    "ERS - APLI": ["doc # fi", "doc type fi", "le:ou legal entity",
                   "line item text orig", "posting period"],
    "ERS": ["g/l account key", "itd balance amount"],
    "TMN02": ["tmn02", "tmn 02"],
}

# Named final-deliverable tabs -> Output (only these count, per the
# reference logic -- everything else that isn't a recognized source
# system falls to the generic Process/Input rule below).
_OUTPUT_NAME_TOKENS = [
    "taxable income summary", "ti summary", "filing package", "audit package",
]

# ADDED: IRS-style tax-form-number tabs -- e.g. "F4797", "1120",
# "Form 1120", "WP 1120", "Workpaper 1120" -- are the filed/finished form
# itself, so they are always Output, regardless of formulas or other
# content on the tab. Recognizes:
#   - a single letter glued directly to 3-4 digits ("F4797", "F1120")
#   - a bare 3-4 digit number as the whole sheet name ("1120", "4797")
#   - a 3-4 digit number after a "Form"/"WP"/"Workpaper"/"F" prefix word
#     ("Form 1120", "WP 1120", "Workpaper 1120", "F 4797")
# 4-digit tokens starting with 19/20 are excluded so real years (2024,
# 2025, ...) are never mistaken for a form number.
_FORM_NUMBER_TOKEN_RE = re.compile(r"^\d{3,4}[a-z]?$")
_FORM_PREFIX_WORDS = ("form", "wp", "workpaper", "f")


def _is_form_number_token(tok: str) -> bool:
    tok = tok.strip().lower()
    if not _FORM_NUMBER_TOKEN_RE.match(tok):
        return False
    digits = re.match(r"^\d+", tok).group()
    if len(digits) == 4 and digits[:2] in ("19", "20"):
        return False  # looks like a year, not a form number
    return True


def _looks_like_tax_form_sheet(name: Optional[str]) -> bool:
    if not name:
        return False
    norm = name.strip().lower()

    # Glued single letter + digits: "F4797", "1120A". Skip a leading "L"
    # here -- "L####" is the separate line-item-tab convention (handled by
    # _looks_like_line_item_sheet, classified as Process), not a form number.
    m = re.match(r"^([a-z])(\d{3,4}[a-z]?)$", norm)
    if m and m.group(1) != "l" and _is_form_number_token(m.group(2)):
        return True

    tokens = [t for t in re.split(r"[\s_\-]+", norm) if t]
    if not tokens:
        return False

    # Bare form number as the entire sheet name: "1120", "4797"
    if len(tokens) == 1 and _is_form_number_token(tokens[0]):
        return True

    # Prefixed form number: "Form 1120", "WP 1120", "Workpaper 1120", "F 4797"
    if len(tokens) >= 2 and tokens[0] in _FORM_PREFIX_WORDS \
       and _is_form_number_token(tokens[1]):
        return True

    return False

# Action words that indicate a generic workpaper tab does calculation
# work (-> Process) rather than just holding source data (-> Input).
_PROCESS_ACTION_WORDS = [
    "adjustment", "schedule", "analysis", "summary", "calculation", "calculate",
    "rollforward", "roll-forward", "reconciliation", "reconcile", "mapping",
    "map", "workpaper", "validate", "validation", "combine", "combined",
    "allocate", "allocation", "support",
]


def _name_matches_any(name: str, tokens: List[str]) -> bool:
    norm = _normalize_token(name)
    words = set(norm.split())
    for t in tokens:
        nt = _normalize_token(t)
        if " " in nt:
            # multi-word phrase: substring match on the normalized name
            if nt and nt in norm:
                return True
        else:
            # single token: whole-word match only, so short tokens like
            # "lv" or "a9" don't false-hit inside unrelated words
            # (e.g. "Salvage" contains "lv" as a substring but not as a word)
            if nt in words:
                return True
    return False


def _content_hits(text: str, indicators: List[str]) -> int:
    norm = _normalize_token(text)
    return sum(1 for kw in indicators if _normalize_token(kw) in norm)


def classify_ipo_reference_logic(name: Optional[str] = None,
                                  sample_text: str = "",
                                  formula_count: int = 0) -> Optional[str]:
    """Input/Process/Output classifier following a real client's applied
    classification logic (sheet-name priority, then content indicators,
    then generic calculate/map/reconcile/adjust/validate/combine/support
    fallback). See the block comment above for the full priority order."""
    name = name or ""

    # Step 1a: mandatory-Process source system, by name.
    for system, tokens in _PROCESS_MANDATORY_NAME_TOKENS.items():
        if _name_matches_any(name, tokens):
            return "Process"

    # Step 1b: defined-Input source system, by name.
    for system, tokens in _INPUT_NAME_TOKENS.items():
        if _name_matches_any(name, tokens):
            return "Input"

    # Step 1c (ADDED): IRS-style tax-form-number tabs -- e.g. "F4797",
    # "1120", "Workpaper 1120" -- are always Output. Checked before the
    # content-based and generic-fallback steps below so a form tab full of
    # formulas doesn't get misclassified as Process.
    if _looks_like_tax_form_sheet(name):
        return "Output"

    if sample_text:
        # Step 2a: mandatory-Process source system, by content indicators.
        best_system, best_hits = None, 0
        for system, indicators in _PROCESS_MANDATORY_CONTENT_INDICATORS.items():
            hits = _content_hits(sample_text, indicators)
            if hits > best_hits:
                best_system, best_hits = system, hits
        if best_system and best_hits >= 1:
            return "Process"

        # Step 2b: defined-Input source system, by content indicators.
        best_system, best_hits = None, 0
        for system, indicators in _INPUT_CONTENT_INDICATORS.items():
            hits = _content_hits(sample_text, indicators)
            if hits > best_hits:
                best_system, best_hits = system, hits
        if best_system and best_hits >= 1:
            return "Input"

    # Step 3: tax-return "line item" naming pattern -- checked only after
    # name- and content-based source-system evidence, since many numbered
    # "L##" tabs are actually source-system Input tabs (e.g. an SAP or
    # ERS - APLI extract that happens to be numbered by return line).
    if _looks_like_line_item_sheet(name):
        return "Process"

    if not sample_text:
        return None

    # Step 4: no source system identified -- generic workpaper tab.
    # Only a named final-deliverable tab is Output.
    if _name_matches_any(name, _OUTPUT_NAME_TOKENS) or \
       _content_hits(sample_text, _OUTPUT_NAME_TOKENS) >= 1:
        return "Output"

    # Otherwise: does work (formulas / action words) -> Process;
    # looks like plain source data -> Input.
    if formula_count > 0 or _content_hits(sample_text, _PROCESS_ACTION_WORDS) >= 1:
        return "Process"

    return "Input"


# --------------------------------------------------------------------------
# ADDED (fix): sheet-COUNT based Input/Process/Output summary, scoped to
# the CURRENT YEAR workbook only (not a combined previous+current total,
# and not a single-sheet cell-level guess). Every sheet that exists in the
# current-year file is classified into Input/Process/Output/Unclassified;
# a sheet counts as "Rollforward" if it also existed in the previous-year
# file (i.e. match_sheets() paired it with a previous-year sheet),
# otherwise it counts as "not rollforwarded" (new this year). This matches
# the "category / Total / not rollforwarded / Rollforward / Formula
# changes in Rollforwarding / General info changes in Rollforwarded" report
# layout.
# --------------------------------------------------------------------------

IPO_LABELS_WITH_UNCLASSIFIED = IPO_LABELS + ("Unclassified",)


def classify_current_sheets(wb_curr: Optional[openpyxl.Workbook]) -> Dict[str, Optional[str]]:
    """Classifies every sheet in the CURRENT-year workbook into
    'Input' / 'Process' / 'Output', using the reference-grounded
    classify_ipo_reference_logic (sheet-name source-system priority, then
    content indicators, then generic calculate/adjust/reconcile/etc.
    fallback -- see that function's docstring for the full rule order)."""
    result: Dict[str, Optional[str]] = {}
    if wb_curr is None:
        return result
    for name in _visible_sheetnames(wb_curr):
        ws = wb_curr[name]
        cat = classify_ipo_reference_logic(name)
        if cat is None:
            cat = classify_ipo_reference_logic(
                name, _sheet_sample_text(ws), _sheet_formula_count(ws))
        result[name] = cat
    return result


def category_sheet_rollforward_summary(pair_label: str,
                                        sheet_matches: List[SheetMatch],
                                        wb_curr: Optional[openpyxl.Workbook]) -> List[dict]:
    """THIS-YEAR-ONLY sheet-count summary per Input/Process/Output category:
      - Total: how many current-year sheets fall in this category
      - not rollforwarded: of those, how many are new this year (no
        matching previous-year sheet)
      - Rollforward: of those, how many existed last year too
      - Formula changes in Rollforwarding: total formula changes across
        that category's rolled-forward sheets
      - General info changes in Rollforwarded: total non-formula (value)
        changes across that category's rolled-forward sheets
    Ends with a Total row. Nothing here sums previous-year sheet counts
    into the totals -- Total/not-rollforwarded/Rollforward all count
    CURRENT-year sheets only."""
    if wb_curr is None:
        return []

    categories = classify_current_sheets(wb_curr)
    by_curr = {sm.curr_sheet: sm for sm in sheet_matches if sm.curr_sheet}

    rows = []
    grand = {"Total": 0, "not rollforwarded": 0, "Rollforward": 0,
             "Formula changes in Rollforwarding": 0,
             "General info changes in Rollforwarded": 0}

    for category in IPO_LABELS_WITH_UNCLASSIFIED:
        sheet_names = [n for n, c in categories.items()
                       if c == category or (category == "Unclassified" and c is None)]
        rf_sheets, not_rf = [], 0
        for n in sheet_names:
            sm = by_curr.get(n)
            if sm is not None and sm.prev_sheet is not None:
                rf_sheets.append(sm)
            else:
                not_rf += 1

        formula_changes = sum(sm.formula_diff.changed + sm.formula_diff.year_ref_change
                               for sm in rf_sheets)
        general_changes = sum(sm.cell_diff.value_changed for sm in rf_sheets)

        row = {
            "Workbook": pair_label,
            "category": category,
            "Total": len(sheet_names),
            "not rollforwarded": not_rf,
            "Rollforward": len(rf_sheets),
            "Formula changes in Rollforwarding": formula_changes,
            "General info changes in Rollforwarded": general_changes,
        }
        rows.append(row)
        for k in grand:
            grand[k] += row[k]

    if rows:
        rows.append({"Workbook": pair_label, "category": "Total", **grand})
    return rows


# ADDED: full per-sheet Input/Process/Output detail (one row per sheet,
# not just category counts) -- so every sheet, in every workbook, can be
# drilled into individually and exported as a complete Input/Process/
# Output breakdown, in addition to the count-only summary above.

def ipo_sheet_detail_table(pair_label: str, sheet_matches: List[SheetMatch],
                            wb_curr: Optional[openpyxl.Workbook]) -> List[dict]:
    """One row per CURRENT-year sheet: its Input/Process/Output category,
    whether it was rolled forward from last year, and (for rolled-forward
    sheets) its formula/general-info change counts. This is the row-level
    companion to category_sheet_rollforward_summary()'s count-only rows."""
    if wb_curr is None:
        return []

    categories = classify_current_sheets(wb_curr)
    by_curr = {sm.curr_sheet: sm for sm in sheet_matches if sm.curr_sheet}

    rows = []
    for name in _visible_sheetnames(wb_curr):
        category = categories.get(name) or "Unclassified"
        sm = by_curr.get(name)
        rolled_forward = bool(sm and sm.prev_sheet is not None)
        rows.append({
            "Workbook": pair_label,
            "Sheet Name": name,
            "Category": category,
            "Rollforward Status": "Rollforward" if rolled_forward else "New This Year",
            "Previous Sheet Name": sm.prev_sheet if (sm and rolled_forward) else "",
            "Formula changes": (sm.formula_diff.changed + sm.formula_diff.year_ref_change)
                                if (sm and rolled_forward) else "",
            "General info changes": sm.cell_diff.value_changed if (sm and rolled_forward) else "",
        })
    return rows


def current_year_overview(pair_label: str, curr_name: Optional[str],
                           wb_curr: Optional[openpyxl.Workbook],
                           sheet_matches: List[SheetMatch]) -> dict:
    """THIS-YEAR-ONLY sheet overview.
    Hidden sheets are surfaced ONLY as a plain count here -- their names are
    not listed and they play no other part in this report (they were
    already excluded upstream, in match_sheets(), from every comparison,
    formula count, populated-cell count, IPO classification, etc.).
    Everything else on this row -- rolled-forward count, its sheet names,
    and the new-this-year count -- is scoped to NON-HIDDEN sheets only,
    since sheet_matches never contains a hidden sheet in the first place."""
    vis = sheet_visibility(wb_curr, curr_name, "Current")
    total_sheets_all = vis.hidden_count + vis.visible_count  # true total, incl. hidden

    # sheet_matches only ever contains non-hidden current-year sheets (hidden
    # ones are filtered out in match_sheets()), so this is already scoped to
    # "among non-hidden sheets" without any further filtering needed here.
    curr_only = [s for s in sheet_matches if s.curr_sheet]
    rolled_forward = [s for s in curr_only if s.prev_sheet is not None]
    new_this_year = [s for s in curr_only if s.prev_sheet is None]

    return {
        "Workbook": pair_label,
        "Current File": curr_name,
        "Total Sheets (This Year, incl. Hidden)": total_sheets_all,
        "Hidden Sheets (This Year)": vis.hidden_count,
        "Non-Hidden Sheets (This Year)": vis.visible_count,
        "Non-Hidden Sheets Rolled Forward": len(rolled_forward),
        "Non-Hidden Rolled-Forward Sheet Names": ", ".join(s.curr_sheet for s in rolled_forward),
        "Non-Hidden New This Year (Not Rolled Forward)": len(new_this_year),
    }


def _find_ipo_sheet_match(category: str, sheet_matches: List[SheetMatch],
                           wb_prev: Optional[openpyxl.Workbook],
                           wb_curr: Optional[openpyxl.Workbook]) -> Optional[SheetMatch]:
    """Locates the SheetMatch that represents `category` ('Input' /
    'Process' / 'Output'), trying progressively looser checks:
      1) exact/substring sheet-name match (original behavior)
      2) keyword-based classification of the sheet name (abbreviations,
         full names, alternate naming conventions)
      3) keyword-based classification of the sheet's cell content, for
         sheets whose name gives no hint at all
    """
    for s in sheet_matches:
        if _is_ipo_sheet(s.curr_sheet, category) or _is_ipo_sheet(s.prev_sheet, category):
            return s

    for s in sheet_matches:
        for name in (s.curr_sheet, s.prev_sheet):
            if name and classify_ipo_by_keywords(name) == category:
                return s

    for s in sheet_matches:
        ws_curr = wb_curr[s.curr_sheet] if (wb_curr and s.curr_sheet) else None
        ws_prev = wb_prev[s.prev_sheet] if (wb_prev and s.prev_sheet) else None
        text = _sheet_sample_text(ws_curr) or _sheet_sample_text(ws_prev)
        if text and classify_ipo_by_keywords(None, text) == category:
            return s
    return None


def ipo_rollforward_table(details: List[IPORollforwardDetail]) -> List[dict]:
    """Flatten IPORollforwardDetail objects into the exact row/column shape
    shown in the Rollforward example screenshot: Input/Process/Output rows
    with Replace, Retain, Formula changes in Retain, General info changes
    in Retain -- plus the cell-reference ranges and note."""
    rows = []
    for d in details:
        rows.append({
            "Workbook": d.workbook_pair,
            "": d.category,
            "Replace": d.replace_count,
            "Retain": d.retain_count,
            "Formula changes in Retain": d.formula_changes_in_retain,
            "General info changes in Retain": d.general_info_changes_in_retain,
            "Retain Cell Ranges": ", ".join(d.retain_ranges) if d.retain_ranges else "",
            "Replace Cell Ranges": ", ".join(d.replace_ranges) if d.replace_ranges else "",
            "Cell Reference Note": d.report_sheet_note,
        })
    if rows:
        rows.append({
            "Workbook": "",
            "": "Total",
            "Replace": sum(r["Replace"] for r in rows),
            "Retain": sum(r["Retain"] for r in rows),
            "Formula changes in Retain": sum(r["Formula changes in Retain"] for r in rows),
            "General info changes in Retain": sum(r["General info changes in Retain"] for r in rows),
            "Retain Cell Ranges": "",
            "Replace Cell Ranges": "",
            "Cell Reference Note": "",
        })
    return rows


# ==========================================================================
# ADDED (per user request — do not alter anything above this line):
#
#   1. Formula-level Roll-Forward extraction: for every cell that rolls
#      forward from the Previous Year sheet to the Current Year sheet and
#      holds a FORMULA, capture the formula text itself alongside the
#      "Sheet!Cell -> Sheet!Cell" reference — i.e. the actual formula used
#      in roll-forwarding, shown at that Previous Year cell and that
#      Current Year cell.
#
#   2. Full raw workbook content extract: every populated cell, from every
#      sheet, in both the Previous Year and Current Year workbook, each row
#      tagged with its sheet reference — so the entire workbook content can
#      be pulled into one Excel report, not just the diffs.
#
# Both are purely additive: new dataclasses/functions only, nothing above
# is modified or called differently.
# ==========================================================================

@dataclass
class FormulaRollforwardMapping:
    workbook_pair: str
    prev_file: Optional[str]
    curr_file: Optional[str]
    prev_sheet: str
    curr_sheet: str
    # each tuple: (prev_ref, curr_ref, prev_formula, curr_formula, status)
    mappings: List[Tuple[str, str, str, str, str]] = field(default_factory=list)
    formula_count: int = 0
    truncated: bool = False


def extract_formula_level_rollforward(pair_label: str,
                                       prev_file: Optional[str], curr_file: Optional[str],
                                       wb_prev: Optional[openpyxl.Workbook],
                                       wb_curr: Optional[openpyxl.Workbook],
                                       sheet_matches: List[SheetMatch],
                                       max_rows: int = 2000, max_cols: int = 150,
                                       max_mappings_per_sheet: int = 1000) -> List[FormulaRollforwardMapping]:
    """For every sheet classified as 'Roll-forward', walks cell-by-cell and,
    for every cell where EITHER the Previous Year or Current Year value is a
    formula, records the Previous Year cell -> Current Year cell reference
    together with both formula strings -- e.g. Previous Year 'Sheet1!B2'
    (=A2*1.1) -> Current Year 'Sheet1!B2' (=A2*1.1) when the formula rolled
    forward unchanged, or with the differing Current Year formula when it
    was updated. Status is 'Rolled Forward' when the formula text is
    identical between years, 'Formula Changed' otherwise. Capped per sheet
    via max_mappings_per_sheet (formula_count still reflects the true
    total; `truncated` flags when the cap was hit)."""
    from itertools import zip_longest
    from openpyxl.utils import get_column_letter

    out: List[FormulaRollforwardMapping] = []
    for sm in sheet_matches:
        if sm.status != "Roll-forward" or not (sm.prev_sheet and sm.curr_sheet):
            continue
        ws_prev = wb_prev[sm.prev_sheet] if wb_prev else None
        ws_curr = wb_curr[sm.curr_sheet] if wb_curr else None
        if ws_prev is None or ws_curr is None:
            continue

        max_r = min(max(ws_prev.max_row or 0, ws_curr.max_row or 0), max_rows)
        max_c = min(max(ws_prev.max_column or 0, ws_curr.max_column or 0), max_cols)
        if max_r == 0 or max_c == 0:
            continue

        empty_row = (None,) * max_c
        rows_prev = ws_prev.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)
        rows_curr = ws_curr.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)

        mapping = FormulaRollforwardMapping(pair_label, prev_file, curr_file, sm.prev_sheet, sm.curr_sheet)
        for r, (row_prev, row_curr) in enumerate(
                zip_longest(rows_prev, rows_curr, fillvalue=empty_row), start=1):
            row_prev = row_prev or empty_row
            row_curr = row_curr or empty_row
            for c in range(max_c):
                v_prev = row_prev[c] if c < len(row_prev) else None
                v_curr = row_curr[c] if c < len(row_curr) else None
                if not (is_formula(v_prev) or is_formula(v_curr)):
                    continue

                addr = f"{get_column_letter(c + 1)}{r}"
                prev_ref = f"{sm.prev_sheet}!{addr}"
                curr_ref = f"{sm.curr_sheet}!{addr}"
                status = "Rolled Forward" if v_prev == v_curr else "Formula Changed"

                mapping.formula_count += 1
                if len(mapping.mappings) < max_mappings_per_sheet:
                    mapping.mappings.append(
                        (prev_ref, curr_ref, str(v_prev or ""), str(v_curr or ""), status)
                    )
                else:
                    mapping.truncated = True
        if mapping.formula_count:
            out.append(mapping)
    return out


def formula_rollforward_table(mappings: List[FormulaRollforwardMapping]) -> List[dict]:
    """Flattens FormulaRollforwardMapping objects into one row per formula
    cell, in the requested 'Previous Year: Sheet1!B2 -> Current Year:
    Sheet1!B2' representation, with both formula strings shown alongside
    each cell reference."""
    rows = []
    for m in mappings:
        for prev_ref, curr_ref, prev_formula, curr_formula, status in m.mappings:
            rows.append({
                "Workbook": m.workbook_pair,
                "Previous File": m.prev_file,
                "Current File": m.curr_file,
                "Previous Sheet": m.prev_sheet,
                "Current Sheet": m.curr_sheet,
                "Previous Year Cell": prev_ref,
                "Current Year Cell": curr_ref,
                "Previous Year Formula": prev_formula,
                "Current Year Formula": curr_formula,
                "Status": status,
                "Roll-Forward Formula Mapping": (
                    f"Previous Year: {prev_ref} ({prev_formula}) \u2192 "
                    f"Current Year: {curr_ref} ({curr_formula})"
                ),
            })
    return rows


def extract_full_workbook_content(pair_label: str,
                                   prev_file: Optional[str], curr_file: Optional[str],
                                   wb_prev: Optional[openpyxl.Workbook],
                                   wb_curr: Optional[openpyxl.Workbook],
                                   max_rows: int = 2000, max_cols: int = 150) -> List[dict]:
    """Walks every visible sheet of both the Previous Year and Current Year
    workbook and records every populated cell, tagged with its Year, File
    name, and Sheet name (a full 'sheet reference' on every row), plus
    whether that cell holds a formula. This is a straight content dump (not
    a diff) so the entire workbook content can be extracted into Excel,
    not just what changed."""
    from openpyxl.utils import get_column_letter

    rows: List[dict] = []

    def _walk(year_label, file_name, wb):
        if wb is None:
            return
        for sheet_name in _visible_sheetnames(wb):
            ws = wb[sheet_name]
            max_r = min(ws.max_row or 0, max_rows)
            max_c = min(ws.max_column or 0, max_cols)
            if max_r == 0 or max_c == 0:
                continue
            for r_idx, row in enumerate(
                    ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True), start=1):
                for c_idx, value in enumerate(row, start=1):
                    if _cell_is_blank(value):
                        continue
                    addr = f"{get_column_letter(c_idx)}{r_idx}"
                    rows.append({
                        "Workbook": pair_label,
                        "Year": year_label,
                        "File": file_name,
                        "Sheet": sheet_name,
                        "Sheet Reference": f"{sheet_name}!{addr}",
                        "Cell": addr,
                        "Is Formula": bool(is_formula(value)),
                        "Value": value,
                    })

    _walk("Previous Year", prev_file, wb_prev)
    _walk("Current Year", curr_file, wb_curr)
    return rows


# ==========================================================================
# ADDED (per user request — do not alter anything above this line):
#
# Fixes the "Download Entire Workbook Content" button sometimes not
# appearing at all, and adds the requested Input/Process/Output tagging on
# top of the full raw content dump.
#
# Root causes addressed (any one of these, on its own, silently kills the
# whole Streamlit script before it reaches st.download_button — Streamlit
# shows a traceback, but on a long page it's easy to miss and it just looks
# like "no button appeared"):
#
#   1. openpyxl refuses to write a string containing certain control
#      characters (e.g. stray \x0b / \x1c bytes that sneak into cells when
#      data was pasted in from another system) — raises IllegalCharacterError.
#   2. Two different workbook-pair labels can collapse to the same 31-char
#      Excel tab name once sanitized/truncated (Excel's own limit) —
#      openpyxl then raises "Sheet ... already exists".
#   3. A single tab holding literally every cell from every sheet of every
#      workbook, times two years, can exceed Excel's hard 1,048,576-row
#      limit per sheet once you have more than a handful of real-sized
#      workbooks — pandas/openpyxl then raises on write.
#
# The functions below make the export defensive against all three, and tag
# every row with its Input / Process / Output / Unclassified category so
# the same download also answers "give me every Input/Process/Output sheet,
# for every workbook, in one file."
# ==========================================================================

# openpyxl's own definition of characters that are illegal in an .xlsx cell
# (control characters outside the small set XML allows). Re-declared here
# rather than imported so this still works on older openpyxl versions that
# don't expose it at the same import path.
_ILLEGAL_XLSX_CHARS_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]"
)
_EXCEL_CELL_CHAR_LIMIT = 32000        # Excel's actual limit is 32,767; leave headroom
_EXCEL_MAX_ROWS_PER_SHEET = 1_000_000  # Excel's actual limit is 1,048,576; leave headroom
_EXCEL_MAX_TAB_LEN = 31                # Excel worksheet-name character limit


def safe_excel_value(value):
    """Makes a single cell value safe to write to .xlsx: strips characters
    openpyxl/Excel reject outright (would otherwise raise
    IllegalCharacterError and silently abort the whole export), and
    truncates strings that are too long for a single Excel cell. Leaves
    numbers, dates, booleans, and None untouched."""
    if isinstance(value, str):
        cleaned = _ILLEGAL_XLSX_CHARS_RE.sub("", value)
        if len(cleaned) > _EXCEL_CELL_CHAR_LIMIT:
            cleaned = cleaned[:_EXCEL_CELL_CHAR_LIMIT] + " …(truncated)"
        return cleaned
    return value


def sanitize_rows_for_excel(rows: List[dict]) -> List[dict]:
    """Applies safe_excel_value() to every value in every row dict. Returns
    new dicts; the input list/dicts are left untouched."""
    return [{k: safe_excel_value(v) for k, v in row.items()} for row in rows]


def unique_excel_sheet_name(base: str, used_names: "set[str]") -> str:
    """Turns `base` into a valid, UNIQUE Excel tab name: strips characters
    Excel forbids in tab names, truncates to 31 chars, and — if that
    collides with a name already used in this workbook — appends a short
    numeric suffix (trimming further as needed so the result still fits in
    31 chars). Mutates `used_names` to include the returned name."""
    cleaned = re.sub(r'[\\/*?:\[\]]', "_", str(base)).strip() or "Sheet"
    candidate = cleaned[:_EXCEL_MAX_TAB_LEN]
    n = 2
    while candidate in used_names:
        suffix = f"_{n}"
        candidate = cleaned[:_EXCEL_MAX_TAB_LEN - len(suffix)] + suffix
        n += 1
    used_names.add(candidate)
    return candidate


def chunk_rows_for_excel_sheet(rows: List[dict], max_rows: int = _EXCEL_MAX_ROWS_PER_SHEET) -> List[List[dict]]:
    """Splits `rows` into chunks no larger than Excel's per-sheet row limit,
    so a category or 'all content' tab that would otherwise overflow a
    single sheet is written as Part 1 / Part 2 / ... instead of crashing
    the export."""
    if not rows:
        return []
    return [rows[i:i + max_rows] for i in range(0, len(rows), max_rows)]


def extract_full_workbook_content_with_category(pair_label: str,
                                                  prev_file: Optional[str], curr_file: Optional[str],
                                                  wb_prev: Optional[openpyxl.Workbook],
                                                  wb_curr: Optional[openpyxl.Workbook],
                                                  max_rows: int = 2000, max_cols: int = 150) -> List[dict]:
    """Same full raw content dump as extract_full_workbook_content() (every
    populated cell, every visible sheet, both years, sheet-referenced) but
    with an added 'Category' column classifying each row's sheet as Input /
    Process / Output / Unclassified — using the same classify_current_sheets()
    logic already used for the current-year Input/Process/Output views
    (applied to whichever workbook, previous or current, that row came
    from). Built on top of extract_full_workbook_content() rather than
    duplicating its cell-walking logic."""
    rows = extract_full_workbook_content(
        pair_label, prev_file, curr_file, wb_prev, wb_curr,
        max_rows=max_rows, max_cols=max_cols,
    )
    cats_prev = classify_current_sheets(wb_prev)
    cats_curr = classify_current_sheets(wb_curr)
    for row in rows:
        cats = cats_curr if row["Year"] == "Current Year" else cats_prev
        row["Category"] = cats.get(row["Sheet"]) or "Unclassified"
    return rows


def build_complete_workbook_bundle(all_rows: List[dict]) -> Tuple[bytes, List[str]]:
    """Builds ONE .xlsx file (as bytes) containing:
      - 'All_Content'          — every populated cell, every sheet, every
                                  workbook pair, both years (split into
                                  All_Content_1/_2/... if it would exceed
                                  Excel's per-sheet row limit)
      - 'Input' / 'Process' / 'Output' / 'Unclassified'
                                  — the same rows, filtered by Category, so
                                  every Input/Process/Output sheet across
                                  every workbook is in its own tab
                                  (also split into _1/_2/... if needed)
      - one tab per workbook pair — that pair's content only, so a single
        workbook can still be inspected in isolation (name deduped/
        truncated to fit Excel's 31-character tab-name limit; split into
        _p1/_p2/... if a single pair alone would overflow a sheet)

    `all_rows` is the concatenation of extract_full_workbook_content_with_
    category() across every matched workbook pair.

    Every value is passed through sanitize_rows_for_excel() first so
    stray control characters from source data can never raise
    IllegalCharacterError and silently abort the whole export, and every
    tab name is deduped via unique_excel_sheet_name() so two workbook pairs
    that truncate to the same 31-character tab name can never collide.

    Returns (xlsx_bytes, warnings) — `warnings` lists anything that was
    split, truncated, or renamed, so the caller can surface that instead of
    just handing back a silent, possibly-incomplete file. Raises no
    exceptions for the conditions described above; only re-raises if the
    underlying write genuinely fails for an unrelated reason."""
    import pandas as pd

    warnings: List[str] = []
    if not all_rows:
        return b"", ["No workbook content was extracted — nothing to export."]

    rows = sanitize_rows_for_excel(all_rows)
    dropped_illegal = sum(
        1 for orig, clean in zip(all_rows, rows)
        if isinstance(orig.get("Value"), str) and orig.get("Value") != clean.get("Value")
    )
    if dropped_illegal:
        warnings.append(
            f"{dropped_illegal} cell value(s) contained characters Excel can't store "
            f"and had those characters stripped (the rest of the value was kept)."
        )

    buffer = BytesIO()
    used_names: set = set()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        def _write_split(rows_for_tab: List[dict], base_name: str):
            chunks = chunk_rows_for_excel_sheet(rows_for_tab)
            if not chunks:
                return
            if len(chunks) > 1:
                warnings.append(
                    f"'{base_name}' had {len(rows_for_tab):,} rows — split into "
                    f"{len(chunks)} tabs to stay under Excel's per-sheet row limit."
                )
            for i, chunk in enumerate(chunks, start=1):
                label = base_name if len(chunks) == 1 else f"{base_name}_{i}"
                sheet_name = unique_excel_sheet_name(label, used_names)
                if sheet_name != label:
                    warnings.append(f"Tab '{label}' renamed to '{sheet_name}' to avoid a duplicate/invalid tab name.")
                pd.DataFrame(chunk).to_excel(writer, sheet_name=sheet_name, index=False)

        _write_split(rows, "All_Content")

        for category in ("Input", "Process", "Output", "Unclassified"):
            cat_rows = [r for r in rows if r.get("Category") == category]
            if cat_rows:
                _write_split(cat_rows, category)

        pairs = list(dict.fromkeys(r["Workbook"] for r in rows))
        for pair_label in pairs:
            pair_rows = [r for r in rows if r["Workbook"] == pair_label]
            _write_split(pair_rows, pair_label)

    return buffer.getvalue(), warnings
