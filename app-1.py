"""
NIBT / FS Items consolidator  --  run with:  streamlit run app.py

pip install streamlit pandas openpyxl pillow pytesseract
(image mode also needs the Tesseract OCR program installed on the PC)
"""
import io
import json
import re

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Workbook Consolidator", layout="wide")

SEG_RE = re.compile(r"^(S\d+|#)$", re.I)          # S1, S3, S8, #
FS_RE = re.compile(r"^\d+[A-Z]+$")                # 849TT
NUM_RE = re.compile(r"^\d{4,}$")                  # 6002000
ID_FIND = re.compile(r"\b(?:\d+[A-Z]{1,4}|\d{5,})\b")
BAD = {"", "#N/A", "NAN", "NONE", "NAT", "0"}


# ---------------------------------------------------------------- helpers
def norm(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    return "" if s in BAD else s


def to_num(v) -> float:
    """'-1,15,70,34,325.31 USD' -> -115703432531.31 ; handles (123) negatives."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[A-Za-z\s,]", "", str(v))
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        x = float(s)
    except ValueError:
        return 0.0
    return -x if neg else x


def score(series: pd.Series, rx) -> float:
    vals = [norm(v) for v in series]
    vals = [v for v in vals if v]
    return sum(bool(rx.match(v)) for v in vals) / len(vals) if vals else 0.0


@st.cache_data(show_spinner=False)
def sheet_names(data: bytes):
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner="Reading main workbook...")
def load_main(data: bytes, sheet: str):
    raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, dtype=object)

    # header row = the row that mentions "Corporate Nibit/Nibt"
    hdr = None
    for i in range(min(25, len(raw))):
        if any(re.search(r"corporate\s*nib", str(c), re.I) for c in raw.iloc[i]):
            hdr = i
            break
    if hdr is None:
        return None

    # segment row = row (at/above header+1) holding the most S1/S3/.../# labels
    best, seg_row = 0, None
    for i in range(0, min(hdr + 3, len(raw))):
        n = sum(bool(SEG_RE.match(str(c).strip())) for c in raw.iloc[i])
        if n > best:
            best, seg_row = n, i
    seg_map = {}
    if seg_row is not None:
        for j, c in enumerate(raw.iloc[seg_row]):
            if SEG_RE.match(str(c).strip()):
                seg_map[j] = str(c).strip().upper()

    start = max(hdr, seg_row if seg_row is not None else hdr) + 1
    body = raw.iloc[start:].reset_index(drop=True)
    heads = [str(c).strip() for c in raw.iloc[hdr]]
    return body, heads, seg_map


def find_cols(body, heads):
    def pick(pattern, rx=None):
        cands = [j for j, h in enumerate(heads) if re.search(pattern, h, re.I)]
        if not cands:
            return None
        if rx is not None:
            return max(cands, key=lambda j: score(body[j], rx))
        return cands[0]

    return {
        "FS items": pick(r"fs\s*item", FS_RE),     # the one with 849TT-style values
        "Corporate NIBT": pick(r"corporate\s*nib"),
        "Associate NIBT": pick(r"associate\s*nib"),
    }


def ids_from_workbook(data: bytes) -> list[str]:
    found = []
    for df in pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object).values():
        for v in df.to_numpy().ravel():
            n = norm(v)
            if n and (FS_RE.match(n) or NUM_RE.match(n)):
                found.append(n)
    return list(dict.fromkeys(found))


def ids_from_image(data: bytes) -> list[str]:
    from PIL import Image
    import pytesseract

    text = pytesseract.image_to_string(Image.open(io.BytesIO(data)))
    return list(dict.fromkeys(ID_FIND.findall(text.upper())))


def consolidate(body, heads, seg_map, key_cols, ids):
    labels = list(dict.fromkeys(seg_map.values()))
    amounts = pd.DataFrame({lab: 0.0 for lab in labels}, index=body.index)
    for j, lab in seg_map.items():
        amounts[lab] += body[j].map(to_num)

    out, missing = {}, {}
    idset = set(ids)
    for name, j in key_cols.items():
        if j is None:
            continue
        keys = body[j].map(norm)
        mask = keys.isin(idset)
        res = amounts[mask].groupby(keys[mask]).sum()
        res.index.name = name
        out[name] = res.reset_index()
        missing[name] = sorted(idset - set(res.index))
    return out, missing


# ---------------------------------------------------------------- Step 3 helpers
ACCT_RE = re.compile(r"^\d{7}$")                  # accounts are exactly 7 digits (no #, no blanks)
KEYS = ["FS items", "Corporate NIBT", "Associate NIBT"]


def seg_order(labels):
    """S1..S8 in numeric order, '#' last."""
    return sorted(labels, key=lambda l: (1, 0) if l == "#" else (0, int(l[1:])))


def build_rows(body, seg_map, key_cols):
    """One row per source row: FS item, Corporate acct, Associate acct + amount per segment."""
    labels = seg_order(set(seg_map.values()))
    df = pd.DataFrame(0.0, index=body.index, columns=labels)
    for j, lab in seg_map.items():
        df[lab] += body[j].map(to_num)

    def clean(name, rx):
        j = key_cols.get(name)
        if j is None:
            return pd.Series("", index=body.index)
        s = body[j].map(norm)
        return s.where(s.map(lambda v: bool(rx.match(v))), "")

    df.insert(0, "Associate NIBT", clean("Associate NIBT", ACCT_RE))
    df.insert(0, "Corporate NIBT", clean("Corporate NIBT", ACCT_RE))
    df.insert(0, "FS items", clean("FS items", FS_RE))
    df = df[(df[KEYS] != "").any(axis=1)]
    return df.reset_index(drop=True), labels


def summarize(df, key, labels):
    """Consolidate by key across S1-S8 and #, add Total column and a TOTAL row."""
    d = df[df[key] != ""]
    g = d.groupby(key)[labels].sum()
    g["Total"] = g.sum(axis=1)
    g = g.sort_index()
    g.loc["TOTAL"] = g.sum()
    g.index.name = key
    return g.reset_index()


DASH_HTML = r"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>NIBT Consolidation Dashboard</title>
<style>
:root{--bg:#f4f6fb;--card:#fff;--ink:#1c2340;--mut:#6b7494;--line:#e3e7f2;--a:#5b5bf0;--b:#14b8a6}
:root[data-theme=dark]{--bg:#0f1426;--card:#171e36;--ink:#e8ecff;--mut:#8f9ac0;--line:#27304f}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 system-ui,"Segoe UI",Arial,sans-serif;background:var(--bg);color:var(--ink)}
header{padding:22px 28px;background:linear-gradient(120deg,#5b5bf0,#14b8a6);color:#fff;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}
header h1{margin:0;font-size:22px}header p{margin:2px 0 0;opacity:.85;font-size:13px}
button,select,input{font:inherit;color:inherit}
.btn{border:1px solid var(--line);background:var(--card);padding:7px 14px;border-radius:999px;cursor:pointer}
.btn:hover{border-color:var(--a)}header .btn{background:rgba(255,255,255,.18);border-color:rgba(255,255,255,.4);color:#fff}
main{padding:20px 28px;max-width:1400px;margin:auto}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;box-shadow:0 2px 10px rgba(30,40,90,.05)}
.controls{display:flex;gap:14px;flex-wrap:wrap;align-items:center}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:999px;overflow:hidden}
.seg button{border:0;background:transparent;padding:7px 16px;cursor:pointer}
.seg button.on{background:var(--a);color:#fff}
.picker{flex:1;min-width:240px;display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.picker input{flex:1;min-width:180px;padding:8px 12px;border:1px solid var(--line);border-radius:10px;background:var(--bg)}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}
.chip{background:var(--a);color:#fff;border-radius:999px;padding:3px 6px 3px 12px;display:inline-flex;gap:6px;align-items:center}
.chip b{cursor:pointer;background:rgba(255,255,255,.25);border-radius:50%;width:18px;height:18px;text-align:center;line-height:18px;font-size:12px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin:16px 0}
.kpi{border-left:5px solid var(--a)}.kpi:nth-child(2){border-color:var(--b)}.kpi:nth-child(3){border-color:#f59e0b}.kpi:nth-child(4){border-color:#ec4899}
.kpi small{color:var(--mut);display:block}.kpi div{font-size:24px;font-weight:700;margin-top:2px;word-break:break-all}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px}
@media(max-width:900px){.grid{grid-template-columns:1fr}}
h3{margin:0 0 8px;font-size:15px}svg{width:100%;height:auto;display:block}svg text{fill:var(--mut);font-size:11px}
.bar{cursor:pointer;transition:opacity .15s}.bar:hover{opacity:.7}
.tabs{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
.tabs .btn.on{background:var(--a);color:#fff;border-color:var(--a)}.tabs .sp{flex:1}
.wrap{overflow:auto;max-height:520px;border:1px solid var(--line);border-radius:10px}
table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
th,td{padding:7px 10px;text-align:right;white-space:nowrap;border-bottom:1px solid var(--line)}
th{position:sticky;top:0;background:var(--card);cursor:pointer;user-select:none;z-index:1}
th:first-child,td:first-child{text-align:left;position:sticky;left:0;background:var(--card);font-weight:600}
tbody tr:hover td{outline:1px solid var(--a)}tbody tr{cursor:pointer}
tr.tot td{font-weight:700;background:var(--bg);position:sticky;bottom:0}
.empty{padding:30px;text-align:center;color:var(--mut)}
</style></head><body>
<header><div><h1>NIBT Consolidation Dashboard</h1><p>Segments S1-S8 and # consolidated per FS item, Corporate NIBT and Associate NIBT account</p></div>
<button class="btn" id="theme">Dark / Light</button></header>
<main>
<div class="card"><div class="controls">
<div class="seg"><button id="mFs">FS item(s)</button><button id="mAc">Account</button></div>
<div class="picker"><input id="inp" list="dl" placeholder="Type or pick to add..." autocomplete="off"><datalist id="dl"></datalist>
<button class="btn" id="all">Select all</button><button class="btn" id="clr">Clear (show all)</button></div></div>
<div class="chips" id="chips"></div></div>
<div class="kpis" id="kpis"></div>
<div class="grid"><div class="card"><h3>Total by segment</h3><div id="c1"></div></div>
<div class="card"><h3 id="c2t">Top 10</h3><div id="c2"></div></div></div>
<div class="card"><div class="tabs" id="tabs"></div><div class="wrap" id="tbl"></div>
<p style="color:var(--mut);margin:8px 0 0;font-size:12px">Click a column header to sort. Click a row or bar to drill into it. Accounts are 7-digit codes only.</p></div>
</main>
<script>
const D=__DATA__, L=D.labels, KEYS=["FS items","Corporate NIBT","Associate NIBT"];
let mode=D.pre.mode, sel=new Set(D.pre.sel), tab=0, sk=null, sd=-1;
const $=id=>document.getElementById(id);
const fmt=v=>v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
const abbr=v=>{const a=Math.abs(v),s=v<0?"-":"";return a>=1e9?s+(a/1e9).toFixed(1)+"B":a>=1e6?s+(a/1e6).toFixed(1)+"M":a>=1e3?s+(a/1e3).toFixed(1)+"K":s+a.toFixed(0)};
const sum=a=>a.reduce((x,y)=>x+y,0);
function options(){const s=new Set();D.rows.forEach(r=>{if(mode=="fs"){if(r[0])s.add(r[0])}else{if(r[1])s.add(r[1]);if(r[2])s.add(r[2])}});return [...s].sort()}
function filtered(){if(!sel.size)return D.rows;return D.rows.filter(r=>mode=="fs"?sel.has(r[0]):(sel.has(r[1])||sel.has(r[2])))}
function agg(rows,i){const m=new Map();rows.forEach(r=>{const k=r[i];if(!k)return;let a=m.get(k);if(!a){a=new Array(L.length).fill(0);m.set(k,a)}r[3].forEach((v,j)=>a[j]+=v)});
 return [...m].map(([k,a])=>({k,a,t:sum(a)}))}
function setMode(m){if(m!=mode){mode=m;sel=new Set()}render()}
function drill(k){mode=tab==0?"fs":"ac";sel=new Set([k]);render()}
function chart1(vals){const W=640,H=250,p=30,mx=Math.max(0,...vals),mn=Math.min(0,...vals),R=(mx-mn)||1;
 const y=v=>p+(mx-v)/R*(H-2*p-16),y0=y(0),bw=W/vals.length;let s=`<svg viewBox="0 0 ${W} ${H}"><line x1="0" x2="${W}" y1="${y0}" y2="${y0}" stroke="var(--line)"/>`;
 vals.forEach((v,i)=>{const x=i*bw+bw*.15,h=Math.abs(y(v)-y0),t=v>=0?y(v):y0,c=v>=0?"var(--a)":"#e04848";
 s+=`<rect class="bar" x="${x}" y="${t}" width="${bw*.7}" height="${Math.max(h,1)}" rx="4" fill="${c}"><title>${L[i]}: ${fmt(v)}</title></rect>`;
 s+=`<text x="${x+bw*.35}" y="${v>=0?t-5:t+h+12}" text-anchor="middle">${abbr(v)}</text><text x="${x+bw*.35}" y="${H-4}" text-anchor="middle" style="font-weight:700">${L[i]}</text>`});
 return s+"</svg>"}
function chart2(rows){const top=[...rows].sort((a,b)=>Math.abs(b.t)-Math.abs(a.t)).slice(0,10);if(!top.length)return'<div class="empty">No data</div>';
 const W=640,rh=24,H=top.length*rh+10,lx=84,mx=Math.max(0,...top.map(r=>r.t)),mn=Math.min(0,...top.map(r=>r.t)),R=(mx-mn)||1,x=v=>lx+(v-mn)/R*(W-lx-70),x0=x(0);
 let s=`<svg viewBox="0 0 ${W} ${H}">`;
 top.forEach((r,i)=>{const yy=i*rh+4,xv=x(r.t);s+=`<text x="${lx-8}" y="${yy+15}" text-anchor="end">${r.k}</text><rect class="bar" data-k="${r.k}" x="${Math.min(x0,xv)}" y="${yy}" width="${Math.max(Math.abs(xv-x0),1)}" height="${rh-6}" rx="4" fill="${r.t>=0?"var(--b)":"#e04848"}"><title>${r.k}: ${fmt(r.t)}</title></rect><text x="${r.t>=0?xv+5:xv-5}" y="${yy+15}" text-anchor="${r.t>=0?"start":"end"}">${abbr(r.t)}</text>`});
 return s+"</svg>"}
function table(rows){if(!rows.length)return'<div class="empty">No matching rows</div>';
 const cols=[KEYS[tab],...L,"Total"],mxc=cols.slice(1).map((_,i)=>Math.max(1,...rows.map(r=>Math.abs(i<L.length?r.a[i]:r.t))));
 let data=[...rows];if(sk!=null)data.sort((a,b)=>sk==0?a.k.localeCompare(b.k)*-sd:((sk-1<L.length?a.a[sk-1]:a.t)-(sk-1<L.length?b.a[sk-1]:b.t))*sd);
 const tot=new Array(L.length).fill(0);rows.forEach(r=>r.a.forEach((v,i)=>tot[i]+=v));
 const cell=(v,i)=>{const al=Math.min(.4,Math.abs(v)/mxc[i]*.4);return`<td style="background:${v<0?"rgba(224,72,72,":"rgba(20,184,166,"}${al})">${fmt(v)}</td>`};
 let h="<table><thead><tr>"+cols.map((c,i)=>`<th data-i="${i}">${c}${sk==i?(sd<0?" ▼":" ▲"):""}</th>`).join("")+"</tr></thead><tbody>";
 data.forEach(r=>{h+=`<tr data-k="${r.k}"><td>${r.k}</td>`+r.a.map((v,i)=>cell(v,i)).join("")+cell(r.t,L.length)+"</tr>"});
 h+=`<tr class="tot"><td>TOTAL</td>`+tot.map(v=>`<td>${fmt(v)}</td>`).join("")+`<td>${fmt(sum(tot))}</td></tr></tbody></table>`;return h}
function render(){
 $("mFs").className=mode=="fs"?"on":"";$("mAc").className=mode=="ac"?"on":"";
 $("dl").innerHTML=options().map(o=>`<option value="${o}">`).join("");
 $("chips").innerHTML=sel.size?[...sel].map(k=>`<span class="chip">${k}<b data-x="${k}">×</b></span>`).join(""):'<span style="color:var(--mut)">No filter - showing everything</span>';
 const rows=filtered();
 const A=[agg(rows,0),agg(rows,1),agg(rows,2)],segs=L.map((_,i)=>sum(rows.map(r=>r[3][i])));
 $("kpis").innerHTML=[["Grand total",fmt(sum(segs))],["FS items",A[0].length],["Corporate NIBT accounts",A[1].length],["Associate NIBT accounts",A[2].length]].map(k=>`<div class="card kpi"><small>${k[0]}</small><div>${k[1]}</div></div>`).join("");
 $("c1").innerHTML=chart1(segs);$("c2t").textContent="Top 10 - "+KEYS[tab];$("c2").innerHTML=chart2(A[tab]);
 $("tabs").innerHTML=KEYS.map((k,i)=>`<button class="btn ${i==tab?"on":""}" data-t="${i}">${k} (${A[i].length})</button>`).join("")+'<span class="sp"></span><button class="btn" id="csv">Download CSV</button>';
 $("tbl").innerHTML=table(A[tab]);
 document.querySelectorAll("#tbl th").forEach(th=>th.onclick=()=>{const i=+th.dataset.i;sd=sk==i?-sd:-1;sk=i;render()});
 document.querySelectorAll("#tbl tbody tr:not(.tot)").forEach(tr=>tr.onclick=()=>drill(tr.dataset.k));
 document.querySelectorAll("#c2 .bar").forEach(b=>b.onclick=()=>drill(b.dataset.k));
 document.querySelectorAll("[data-t]").forEach(b=>b.onclick=()=>{tab=+b.dataset.t;sk=null;render()});
 document.querySelectorAll("[data-x]").forEach(b=>b.onclick=()=>{sel.delete(b.dataset.x);render()});
 $("csv").onclick=()=>{const cols=[KEYS[tab],...L,"Total"];const out=[cols.join(",")];A[tab].forEach(r=>out.push([r.k,...r.a.map(v=>v.toFixed(2)),r.t.toFixed(2)].join(",")));
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([out.join("\n")],{type:"text/csv"}));a.download=KEYS[tab].replace(" ","_")+".csv";a.click()}}
$("mFs").onclick=()=>setMode("fs");$("mAc").onclick=()=>setMode("ac");
$("inp").onchange=e=>{const v=e.target.value.trim().toUpperCase();if(options().includes(v)){sel.add(v);e.target.value="";render()}};
$("all").onclick=()=>{sel=new Set(options());render()};$("clr").onclick=()=>{sel=new Set();render()};
$("theme").onclick=()=>{const r=document.documentElement;r.dataset.theme=r.dataset.theme=="dark"?"light":"dark"};
render();
</script></body></html>"""


def build_dashboard(rows, labels, mode="fs", sel=()):
    g = rows.groupby(KEYS, sort=False)[labels].sum().reset_index()
    payload = {
        "labels": labels,
        "rows": [[a, b, c, [round(float(x), 2) for x in vals]]
                 for a, b, c, *vals in g[KEYS + labels].itertuples(index=False, name=None)],
        "pre": {"mode": mode, "sel": list(sel)},
    }
    return DASH_HTML.replace("__DATA__", json.dumps(payload).replace("</", "<\\/"))


# ---------------------------------------------------------------- UI
st.title("Workbook Consolidator")

# ---- Step 1
st.header("Step 1 - Main workbook")
main_file = st.file_uploader("Upload the main Excel workbook", type=["xlsx", "xlsm", "xls"], key="main")

if main_file:
    data = main_file.getvalue()
    sheet = st.selectbox("Sheet", sheet_names(data))
    loaded = load_main(data, sheet)
    if loaded is None:
        st.error("Couldn't find a 'Corporate Nibit' header on this sheet. Try another sheet.")
        st.stop()
    body, heads, seg_map = loaded
    key_cols = find_cols(body, heads)

    with st.expander("Detected columns (change if wrong)", expanded=False):
        opts = [None] + list(range(len(heads)))
        fmt = lambda j: "-- none --" if j is None else f"Col {j + 1}: {heads[j] or '(blank)'}"
        for name in list(key_cols):
            key_cols[name] = st.selectbox(
                name, opts, index=opts.index(key_cols[name]), format_func=fmt, key=f"c_{name}"
            )
        st.write("Segment columns:", {f"Col {j + 1}": lab for j, lab in seg_map.items()})

    st.success(
        f"Loaded {len(body):,} rows | segments: {', '.join(dict.fromkeys(seg_map.values())) or 'none found'}"
    )

    # ---- Step 2
    st.header("Step 2 - What to search")
    mode = st.selectbox("Choose input type", ["Excel workbook", "Image"])
    if mode == "Excel workbook":
        up = st.file_uploader("Upload search workbook", type=["xlsx", "xlsm", "xls"], key="s_x")
    else:
        up = st.file_uploader("Upload search image", type=["png", "jpg", "jpeg"], key="s_i")
        if up:
            st.image(up, width=400)

    if up:
        try:
            found = ids_from_workbook(up.getvalue()) if mode == "Excel workbook" else ids_from_image(up.getvalue())
        except Exception as e:  # OCR missing, bad file, ...
            st.error(f"Could not read the file: {e}")
            found = []
        txt = st.text_area(
            "IDs to search (one per line - edit/fix OCR mistakes here)",
            "\n".join(found),
            height=150,
        )
        ids = [norm(x) for x in txt.splitlines() if norm(x)]

        if st.button("Consolidate", type="primary") and ids:
            result, missing = consolidate(body, heads, seg_map, key_cols, ids)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as xw:
                for name, df in result.items():
                    st.subheader(f"{name}  ({len(df)} matched)")
                    st.dataframe(df.style.format({c: "{:,.2f}" for c in df.columns[1:]}), use_container_width=True)
                    if missing[name]:
                        st.caption("Not found: " + ", ".join(missing[name]))
                    df.to_excel(xw, sheet_name=name.replace(" ", "_"), index=False)
            if result:
                st.download_button(
                    "Download consolidated Excel",
                    buf.getvalue(),
                    "consolidated.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    # ---- Step 3
    st.header("Step 3 - FS item / Account lookup")
    rows, labels = build_rows(body, seg_map, key_cols)

    by = st.radio("Look up by", ["FS item(s)", "Account"], horizontal=True)
    if by == "FS item(s)":
        opts_ = sorted(x for x in rows["FS items"].unique() if x)
        pick = st.multiselect("FS item(s)", opts_, placeholder="Choose one or more FS items")
        mask = rows["FS items"].isin(pick)
        dmode = "fs"
    else:
        opts_ = sorted((set(rows["Corporate NIBT"]) | set(rows["Associate NIBT"])) - {""})
        pick = st.multiselect("Account (7-digit)", opts_, placeholder="Choose one or more accounts")
        mask = rows["Corporate NIBT"].isin(pick) | rows["Associate NIBT"].isin(pick)
        dmode = "ac"

    if pick:
        sub = rows[mask]
        buf3 = io.BytesIO()
        with pd.ExcelWriter(buf3, engine="openpyxl") as xw:
            for key in KEYS:
                tbl = summarize(sub, key, labels)
                st.subheader(f"{key}  ({len(tbl) - 1} items)")
                st.dataframe(
                    tbl.style.format({c: "{:,.2f}" for c in tbl.columns[1:]}),
                    use_container_width=True,
                    hide_index=True,
                )
                tbl.to_excel(xw, sheet_name=key.replace(" ", "_"), index=False)
        st.download_button(
            "Download these tables (Excel)", buf3.getvalue(), "lookup_tables.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_lookup",
        )
    else:
        st.info("Pick at least one FS item or account to see the tables.")

    st.subheader("Interactive dashboard")
    html = build_dashboard(rows, labels, dmode, pick)
    st.download_button(
        "Download interactive dashboard (HTML)", html.encode("utf-8"), "NIBT_dashboard.html",
        "text/html", key="dl_dash",
    )
    with st.expander("Preview dashboard here", expanded=False):
        components.html(html, height=1000, scrolling=True)
else:
    st.info("Upload the main workbook to begin.")
