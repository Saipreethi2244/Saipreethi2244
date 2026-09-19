"""
formula_flow.py — fast formula / manual-value / incoming-outgoing analysis
for roll-forward (retained) sheets.

WHY THIS IS FAST
    Workbooks are NOT loaded through openpyxl.  Each visible worksheet's XML is
    streamed once (lxml iterparse if available, stdlib otherwise) and only the
    facts we need are kept:  formula text, numeric constants, header text and
    cross-sheet references.  Shared strings are resolved lazily (only the few
    header cells), so a 400k-cell sheet scans in about a second.

DEFINITIONS (used consistently everywhere)
    Formula cell     a cell containing a formula (normal, shared, array).
    Manual cell      a cell holding a hard-typed NUMERIC constant (not a formula,
                     not text).  Text labels are deliberately not counted.
    Incoming formula a formula located IN sheet S that reads from another sheet
                     T (T != S) or from an external workbook.
                     -> "sheet S receives data from T".
    Outgoing formula a formula located in another sheet X that reads from sheet S.
                     -> "sheet S sends data to X".
    Retained sheet   a sheet present in both the previous and the current
                     workbook (matched by name, then by name-with-year-masked).
    Change           compared cell-by-cell at the same address, previous vs current:
                     modified + added + removed.
"""
from __future__ import annotations

import io
import posixpath
import re
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from functools import lru_cache
from typing import Callable, Dict, Iterable, List, Optional, Tuple

try:  # lxml is ~2x faster, but optional
    from lxml import etree as ET  # type: ignore
    _LXML = True
except Exception:  # pragma: no cover
    import xml.etree.ElementTree as ET  # type: ignore
    _LXML = False

EXTERNAL = "[External workbook]"
_REF = re.compile(r"([A-Z]+)(\d+)$")
_NS_RE = re.compile(rb'xmlns="([^"]+)"')


# ----------------------------------------------------------------------------
# small helpers
# ----------------------------------------------------------------------------
@lru_cache(maxsize=4096)
def _col_to_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _split_ref(ref: str) -> Tuple[int, int]:
    m = _REF.match(ref)
    return int(m.group(2)), _col_to_num(m.group(1))


# Sheet references inside a formula: 'My Sheet'!A1   Sheet1!A1   [1]Sheet!A1   'C:\x\[f.xlsx]S'!A1
_QUOTED = re.compile(r"'((?:[^']|'')+)'!")
_PLAIN = re.compile(r"(?<![\w.'\]])([A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*)!")
_STRLIT = re.compile(r'"(?:[^"]|"")*"')


@lru_cache(maxsize=200_000)
def _sheet_refs(formula: str) -> Tuple[str, ...]:
    """Distinct sheet names referenced by a formula text (string literals ignored)."""
    if "!" not in formula:
        return ()
    f = _STRLIT.sub('""', formula)
    out = []
    for m in _QUOTED.finditer(f):
        out.append(m.group(1).replace("''", "'"))
    f2 = _QUOTED.sub("", f)
    for m in _PLAIN.finditer(f2):
        out.append(m.group(1))
    return tuple(dict.fromkeys(out))


def _norm_formula(text: str) -> str:
    return text.replace(" ", "").replace("\n", "").replace("'", "").upper()


def _apply_rename(norm_text: str, rename: Dict[str, str]) -> str:
    """Rewrite references to renamed sheets (previous name -> current name) in a normalised formula."""
    for old, new in rename.items():
        if old in norm_text:
            norm_text = re.sub(r"(?<![A-Z0-9_.])" + re.escape(old) + "!", new + "!", norm_text)
    return norm_text


# ----------------------------------------------------------------------------
# data classes
# ----------------------------------------------------------------------------
@dataclass
class SheetScan:
    name: str
    state: str = "visible"
    used_rows: int = 0
    used_cols: int = 0
    nonblank: int = 0
    formulas: Dict[str, tuple] = field(default_factory=dict)   # addr -> key
    manuals: Dict[str, float] = field(default_factory=dict)    # addr -> number
    xref: Dict[str, Tuple[str, ...]] = field(default_factory=dict)  # addr -> target sheets (cross-sheet formulas only)
    header_idx: set = field(default_factory=set)               # shared-string idx of first rows
    header_text: List[str] = field(default_factory=list)       # upper-case header strings
    truncated: bool = False

    @property
    def n_formula(self) -> int:
        return len(self.formulas)

    @property
    def n_manual(self) -> int:
        return len(self.manuals)


# ----------------------------------------------------------------------------
# workbook scanner
# ----------------------------------------------------------------------------
def _read_workbook_parts(z: zipfile.ZipFile):
    """Return list of (sheet_name, state, xml_path) for worksheets, in workbook order."""
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_target = {}
    for rel in rels_xml:
        typ = rel.get("Type", "")
        if typ.endswith("/worksheet"):
            tgt = rel.get("Target", "")
            path = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join("xl", tgt))
            rid_target[rel.get("Id")] = path
    out = []
    for el in wb_xml.iter():
        if el.tag.endswith("}sheet") or el.tag == "sheet":
            rid = None
            for k, v in el.attrib.items():
                if k.endswith("}id") or k == "id":
                    rid = v
            if rid in rid_target:
                out.append((el.get("name"), el.get("state", "visible"), rid_target[rid]))
    return out


def _load_shared_strings(z: zipfile.ZipFile, wanted: set) -> Dict[int, str]:
    """Resolve only the shared-string indexes we need; stop reading as soon as we have them."""
    if not wanted or "xl/sharedStrings.xml" not in z.namelist():
        return {}
    last = max(wanted)
    out: Dict[int, str] = {}
    with z.open("xl/sharedStrings.xml") as fh:
        i = -1
        for _, el in ET.iterparse(fh, events=("end",)):
            if el.tag.endswith("}si"):
                i += 1
                if i in wanted:
                    out[i] = "".join(t.text or "" for t in el.iter() if t.tag.endswith("}t"))
                el.clear()
                if i >= last:
                    break
    return out


def _scan_sheet(z: zipfile.ZipFile, path: str, scan: SheetScan, known_lower: Dict[str, str],
                max_rows: int, max_cols: int, header_rows: int = 15, header_cols: int = 60,
                keep_values: bool = True) -> None:
    with z.open(path) as fh:
        head = fh.read(4096)
    m = _NS_RE.search(head)
    ns = m.group(1).decode() if m else "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ROW, C, F, V, IS, T = (f"{{{ns}}}{x}" for x in ("row", "c", "f", "v", "is", "t"))
    self_lower = scan.name.lower()

    formulas, manuals, xref = scan.formulas, scan.manuals, scan.xref
    shared: Dict[str, Tuple[str, int, int]] = {}
    used_r = used_c = nonblank = 0
    header_idx, header_text = scan.header_idx, scan.header_text

    with z.open(path) as fh:
        for _, row in ET.iterparse(fh, events=("end",)):
            if row.tag != ROW:
                continue
            rnum = int(row.get("r") or 0)
            if rnum > max_rows:
                scan.truncated = True
                break
            ccount = 0
            for c in row:
                ccount += 1
                ref = c.get("r")
                if ref:
                    mm = _REF.match(ref)
                    if not mm:
                        continue
                    col = _col_to_num(mm.group(1))
                    r = rnum or int(mm.group(2))
                else:  # very rare: no address attribute
                    col, r = ccount, rnum
                    ref = f"{'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[col-1] if col<=26 else 'A'}{r}"
                if col > max_cols:
                    scan.truncated = True
                    continue
                f = c.find(F)
                v = c.find(V)
                t = c.get("t")
                if f is not None:
                    nonblank += 1
                    ftxt = f.text
                    ftype = f.get("t")
                    if ftype == "shared":
                        si = f.get("si")
                        if ftxt:
                            shared[si] = (ftxt, r, col)
                            key = ftxt
                            base = ftxt
                        else:
                            mst = shared.get(si)
                            if mst is None:
                                key, base = ("?", si), ""
                            else:
                                key, base = ("S", mst[0], r - mst[1], col - mst[2]), mst[0]
                    else:
                        base = ftxt or ""
                        key = base
                    formulas[ref] = key
                    if base and "!" in base or (base and "[" in base):
                        tg = []
                        for nm in _sheet_refs(base):
                            real = known_lower.get(nm.lower())
                            if real is not None and real.lower() != self_lower:
                                tg.append(real)
                            elif real is None and re.match(r"\[\d+\]", nm):
                                tg.append(EXTERNAL)
                        if "[" in base and re.search(r"'?\[\d+\][^!]*!", base):
                            if EXTERNAL not in tg:
                                tg.append(EXTERNAL)
                        if tg:
                            xref[ref] = tuple(dict.fromkeys(tg))
                elif v is not None and v.text not in (None, ""):
                    nonblank += 1
                    if t is None or t == "n":
                        try:
                            val = float(v.text)
                            if keep_values:
                                manuals[ref] = val
                            else:
                                manuals[ref] = 0.0
                            if r <= header_rows and col <= header_cols:
                                header_text.append(str(int(val)) if val == int(val) else str(val))
                        except ValueError:
                            pass
                    elif t == "s":
                        if r <= header_rows and col <= header_cols:
                            try:
                                header_idx.add(int(v.text))
                            except ValueError:
                                pass
                    elif t in ("str",) and r <= header_rows and col <= header_cols:
                        header_text.append(v.text.upper())
                elif t == "inlineStr":
                    is_ = c.find(IS)
                    if is_ is not None:
                        txt = "".join(x.text or "" for x in is_.iter(T))
                        if txt:
                            nonblank += 1
                            if r <= header_rows and col <= header_cols:
                                header_text.append(txt.upper())
                        else:
                            continue
                    else:
                        continue
                else:
                    continue
                if r > used_r:
                    used_r = r
                if col > used_c:
                    used_c = col
            row.clear()
    scan.used_rows, scan.used_cols, scan.nonblank = used_r, used_c, nonblank


def scan_workbook(data: bytes, max_rows: int = 2000, max_cols: int = 150,
                  visible_only: bool = True, only_sheets: Optional[Iterable[str]] = None,
                  keep_values: bool = True,
                  progress: Optional[Callable[[str], None]] = None) -> Dict[str, SheetScan]:
    """Scan a workbook (bytes) -> {sheet_name: SheetScan}.  Hidden/veryHidden sheets are skipped
    by default (USN5 scope = visible sheets only)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    parts = _read_workbook_parts(z)
    if visible_only:
        parts = [p for p in parts if p[1] == "visible"]
    known_lower = {n.lower(): n for n, _, _ in parts}
    only = {s.lower() for s in only_sheets} if only_sheets else None
    scans: Dict[str, SheetScan] = {}
    for name, state, path in parts:
        sc = SheetScan(name=name, state=state)
        if progress:
            progress(name)
        _scan_sheet(z, path, sc, known_lower, max_rows, max_cols,
                    keep_values=(only is None or name.lower() in only) and keep_values)
        scans[name] = sc
    wanted = set().union(*(s.header_idx for s in scans.values())) if scans else set()
    sst = _load_shared_strings(z, wanted)
    for s in scans.values():
        s.header_text.extend(sst[i].upper() for i in s.header_idx if i in sst)
        s.header_idx = set()
    return scans


# ----------------------------------------------------------------------------
# effective formula text (only needed when two "keys" differ)
# ----------------------------------------------------------------------------
def _effective(key, addr: str) -> str:
    if isinstance(key, tuple):
        if key[0] == "S":
            _, master, dr, dc = key
            r, c = _split_ref(addr)
            origin_c = c - dc
            origin_r = r - dr
            try:
                from openpyxl.formula.translate import Translator
                from openpyxl.utils import get_column_letter
                origin = f"{get_column_letter(origin_c)}{origin_r}"
                return _norm_formula(Translator("=" + master, origin=origin).translate_formula(addr)[1:])
            except Exception:
                return _norm_formula(master)
        return "?" + str(key)
    return _norm_formula(key)


def _formula_equal(k1, a1: str, k2, a2: str, rename: Optional[Dict[str, str]] = None) -> bool:
    """k1/a1 = previous formula, k2/a2 = current formula.  `rename` maps normalised previous
    sheet names to their current names so a renamed sheet is not reported as a formula change."""
    if k1 == k2:
        return True
    e1 = _effective(k1, a1)
    if rename:
        e1 = _apply_rename(e1, rename)
    return e1 == _effective(k2, a2)


# ----------------------------------------------------------------------------
# sheet matching (retained sheets)
# ----------------------------------------------------------------------------
_YEAR = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)|(?<=[_\-\s])\d{2}(?![\d])")


def _mask(name: str) -> str:
    return re.sub(r"\s+", " ", _YEAR.sub("#", name.strip().lower()))


def match_retained(prev_names: List[str], curr_names: List[str]) -> List[Tuple[str, str, str]]:
    """Return [(prev_name, curr_name, how)].  how = 'same name' | 'renamed (year/case)'."""
    out, used_prev = [], set()
    pl = {n.strip().lower(): n for n in prev_names}
    for cn in curr_names:
        k = cn.strip().lower()
        if k in pl:
            out.append((pl[k], cn, "same name"))
            used_prev.add(pl[k])
    matched_curr = {c for _, c, _ in out}
    pm = defaultdict(list)
    for pn in prev_names:
        if pn not in used_prev:
            pm[_mask(pn)].append(pn)
    for cn in curr_names:
        if cn in matched_curr:
            continue
        cands = pm.get(_mask(cn), [])
        if len(cands) == 1:
            out.append((cands[0], cn, "renamed (year/case)"))
            used_prev.add(cands[0])
    order = {n: i for i, n in enumerate(curr_names)}
    out.sort(key=lambda x: order[x[1]])
    return out


# ----------------------------------------------------------------------------
# comparison of a retained sheet pair
# ----------------------------------------------------------------------------
@dataclass
class PairDiff:
    f_same: int = 0
    f_modified: int = 0
    f_added: int = 0
    f_removed: int = 0
    m_same: int = 0
    m_modified: int = 0
    m_added: int = 0
    m_removed: int = 0
    in_changed: int = 0   # incoming (cross-sheet) formulas modified/added/removed
    changed_addrs: set = field(default_factory=set)  # current-sheet addresses of modified/added formulas

    @property
    def f_changes(self):
        return self.f_modified + self.f_added + self.f_removed

    @property
    def m_changes(self):
        return self.m_modified + self.m_added + self.m_removed


def compare_sheet(prev: SheetScan, curr: SheetScan, rename: Optional[Dict[str, str]] = None) -> PairDiff:
    d = PairDiff()
    pf, cf = prev.formulas, curr.formulas
    for addr, key in cf.items():
        pk = pf.get(addr)
        if pk is None:
            d.f_added += 1
            d.changed_addrs.add(addr)
        elif _formula_equal(pk, addr, key, addr, rename):
            d.f_same += 1
        else:
            d.f_modified += 1
            d.changed_addrs.add(addr)
    d.f_removed = sum(1 for a in pf if a not in cf)
    pm, cm = prev.manuals, curr.manuals
    for addr, val in cm.items():
        pv = pm.get(addr)
        if pv is None:
            d.m_added += 1
        elif pv == val or abs(pv - val) <= 1e-9 * max(1.0, abs(pv)):
            d.m_same += 1
        else:
            d.m_modified += 1
    d.m_removed = sum(1 for a in pm if a not in cm)
    # incoming (cross-sheet) formulas that changed
    inc_changed = sum(1 for a in curr.xref if a in d.changed_addrs)
    inc_removed = sum(1 for a in prev.xref if a not in cf)
    d.in_changed = inc_changed + inc_removed
    return d


# ----------------------------------------------------------------------------
# full report
# ----------------------------------------------------------------------------
CATS = ["Input", "Calc", "Output"]


def to_cat(classification: str) -> str:
    return "Calc" if classification.lower().startswith("process") else classification.capitalize()


@dataclass
class FormulaReport:
    per_sheet: List[dict]          # one row per retained sheet
    edges: List[dict]              # provider -> consumer formula flows
    summary_main: List[dict]       # Table 1
    summary_formulas: List[dict]   # Table 2
    summary_incoming: List[dict]   # Table 3
    summary_outgoing: List[dict]   # Table 4
    summary_manual: List[dict]     # Table 5
    not_retained_prev: List[str]
    not_retained_curr: List[str]


def analyse_rollforward(prev_scans: Dict[str, SheetScan], curr_scans: Dict[str, SheetScan],
                        classes: Dict[str, str],
                        pairs: Optional[List[Tuple[str, str, str]]] = None) -> FormulaReport:
    """classes: {current sheet name -> 'Input'|'Process'|'Output'} (all visible current sheets).
    pairs: optional explicit [(prev, curr, how)] retained list (e.g. from the main engine)."""
    if pairs is None:
        pairs = match_retained(list(prev_scans), list(curr_scans))
    retained_curr = {c for _, c, _ in pairs}

    # ---- provider -> consumer flows over ALL visible current sheets ----
    flow: Dict[Tuple[str, str], int] = Counter()
    for cname, sc in curr_scans.items():
        for addr, targets in sc.xref.items():
            for t in targets:
                flow[(t, cname)] += 1

    pair_by_curr = {c: (p, how) for p, c, how in pairs}
    # previous-name -> current-name for renamed sheets (so 'Calc 2024'!A1 == 'Calc 2025'!A1)
    rename = {_norm_formula(p): _norm_formula(c) for p, c, _ in pairs if _norm_formula(p) != _norm_formula(c)}
    diffs: Dict[str, PairDiff] = {}
    for p, c, _ in pairs:
        diffs[c] = compare_sheet(prev_scans[p], curr_scans[c], rename)

    # changed formulas per (provider, consumer) edge (only where consumer is retained)
    edge_changed: Dict[Tuple[str, str], int] = Counter()
    for cname in retained_curr:
        sc, d = curr_scans[cname], diffs[cname]
        for addr in d.changed_addrs:
            for t in sc.xref.get(addr, ()):
                edge_changed[(t, cname)] += 1

    def cls_of(name: str) -> str:
        return to_cat(classes.get(name, "Process")) if name != EXTERNAL else "External"

    per_sheet = []
    for p, c, how in pairs:
        sc, d = curr_scans[c], diffs[c]
        incoming = Counter()
        for targets in sc.xref.values():
            for t in targets:
                incoming[t] += 1
        outgoing = {cons: n for (prov, cons), n in flow.items() if prov == c}
        n_in = len(sc.xref)
        n_out = sum(outgoing.values())
        per_sheet.append({
            "Sheet (current)": c, "Sheet (previous)": p, "Matched by": how,
            "Category": to_cat(classes.get(c, "Process")),
            "Number of Formulas": sc.n_formula,
            "Number of Manual": sc.n_manual,
            "Formula changes": d.f_changes,
            "Formula modified": d.f_modified, "Formula added": d.f_added, "Formula removed": d.f_removed,
            "Manual changes": d.m_changes,
            "Manual modified": d.m_modified, "Manual added": d.m_added, "Manual removed": d.m_removed,
            "Manual not changing": d.m_same,
            "Number of Income formulas": n_in,
            "Income formula changes": d.in_changed,
            "Number of Outgoing formulas": n_out,
            "Incoming sheet name": ", ".join(sorted(incoming, key=lambda x: (-incoming[x], x))),
            "Incoming sheet type": ", ".join(sorted({cls_of(t) for t in incoming})),
            "Outgoing sheet name": ", ".join(sorted(outgoing, key=lambda x: (-outgoing[x], x))),
            "Outgoing sheet type": ", ".join(sorted({cls_of(t) for t in outgoing})),
        })

    edges = []
    for (prov, cons), n in sorted(flow.items(), key=lambda kv: (-kv[1], kv[0])):
        if prov in retained_curr or cons in retained_curr:
            edges.append({
                "Provider sheet (outgoing from)": prov, "Provider type": cls_of(prov),
                "Consumer sheet (incoming to)": cons, "Consumer type": cls_of(cons),
                "Formulas": n,
                "Formulas changed vs previous": edge_changed.get((prov, cons), 0) if cons in retained_curr else "n/a (sheet not retained)",
                "Provider retained?": "Yes" if prov in retained_curr else "No",
                "Consumer retained?": "Yes" if cons in retained_curr else "No",
            })

    # ------------------------- category summaries -------------------------
    def by_cat(cat):
        return [r for r in per_sheet if r["Category"] == cat]

    def add_total(rows, key_sum):
        tot = {"Category": "Total"}
        for k in key_sum:
            tot[k] = sum(r[k] for r in rows if r["Category"] in CATS)
        rows.append(tot)
        return rows

    t1, t2, t3, t4, t5 = [], [], [], [], []
    for cat in CATS:
        rs = by_cat(cat)
        t1.append({"Category": cat, "Number of Sheets": len(rs),
                   "Number of Formulas": sum(r["Number of Formulas"] for r in rs),
                   "Number of Manual": sum(r["Number of Manual"] for r in rs),
                   "Number of formula changes": sum(r["Formula changes"] for r in rs),
                   "Number of Manual changes": sum(r["Manual changes"] for r in rs)})
        inc_names = sorted({n for r in rs for n in r["Incoming sheet name"].split(", ") if n})
        out_names = sorted({n for r in rs for n in r["Outgoing sheet name"].split(", ") if n})
        out_types = Counter()
        for r in rs:
            for n in (x for x in r["Outgoing sheet name"].split(", ") if x):
                out_types[cls_of(n)] += 1
        t2.append({"Category": cat, "Number of Sheets": len(rs),
                   "Number of Formulas": sum(r["Number of Formulas"] for r in rs),
                   "Number of Income formulas": sum(r["Number of Income formulas"] for r in rs),
                   "Number of Outgoing formulas": sum(r["Number of Outgoing formulas"] for r in rs),
                   "Number of income formulas changes": sum(r["Income formula changes"] for r in rs),
                   "Incoming sheet name": ", ".join(inc_names),
                   "Outgoing sheet Name": ", ".join(out_names),
                   "Outgoing sheet type": ", ".join(f"{k}: {v}" for k, v in sorted(out_types.items()))})
        inc = [r for r in rs if r["Number of Income formulas"] > 0]
        t3.append({"Category": cat, "Number of Sheets": len(inc),
                   "Number of Formulas": sum(r["Number of Formulas"] for r in inc),
                   "Number of Income formulas": sum(r["Number of Income formulas"] for r in inc),
                   "Changes": sum(r["Income formula changes"] for r in inc)})
        out = [r for r in rs if r["Number of Outgoing formulas"] > 0]
        t4.append({"Category": cat, "Number of Sheets": len(out),
                   "Number of Formulas": sum(r["Number of Formulas"] for r in out),
                   "Number of Outgoing formulas": sum(r["Number of Outgoing formulas"] for r in out)})
        man = [r for r in rs if r["Number of Manual"] > 0 or r["Manual removed"] > 0]
        t5.append({"Category": cat, "Number of Sheets": len(man),
                   "Number of Manual": sum(r["Number of Manual"] for r in man),
                   "Changing number": sum(r["Manual changes"] for r in man),
                   "Not Changing number": sum(r["Manual not changing"] for r in man),
                   "of which modified": sum(r["Manual modified"] for r in man),
                   "of which added": sum(r["Manual added"] for r in man),
                   "of which removed": sum(r["Manual removed"] for r in man)})
    for tbl, cols in ((t1, ["Number of Sheets", "Number of Formulas", "Number of Manual",
                            "Number of formula changes", "Number of Manual changes"]),
                      (t3, ["Number of Sheets", "Number of Formulas", "Number of Income formulas", "Changes"]),
                      (t4, ["Number of Sheets", "Number of Formulas", "Number of Outgoing formulas"]),
                      (t5, ["Number of Sheets", "Number of Manual", "Changing number", "Not Changing number",
                            "of which modified", "of which added", "of which removed"])):
        tot = {"Category": "Total"}
        for k in cols:
            tot[k] = sum(r[k] for r in tbl)
        tbl.append(tot)
    tot2 = {"Category": "Total"}
    for k in ["Number of Sheets", "Number of Formulas", "Number of Income formulas",
              "Number of Outgoing formulas", "Number of income formulas changes"]:
        tot2[k] = sum(r[k] for r in t2)
    t2.append(tot2)

    return FormulaReport(
        per_sheet=per_sheet, edges=edges,
        summary_main=t1, summary_formulas=t2, summary_incoming=t3,
        summary_outgoing=t4, summary_manual=t5,
        not_retained_prev=[n for n in prev_scans if n not in {p for p, _, _ in pairs}],
        not_retained_curr=[n for n in curr_scans if n not in retained_curr],
    )
