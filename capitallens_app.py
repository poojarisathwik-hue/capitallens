"""
CapitalLens — Financial Model Analyzer
A Streamlit-based investment research dashboard.

Install dependencies:
    pip install streamlit pandas openpyxl plotly xlrd

Run:
    streamlit run capitallens_app.py
"""

import re
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CapitalLens — Financial Model Analyzer",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# THEME / CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500&family=Syne:wght@400;500;600;700&display=swap');

:root {
  --ink: #0A0B0D;
  --ink2: #12151A;
  --ink3: #1C2130;
  --gold: #C9A84C;
  --emerald: #2DBD8A;
  --crimson: #E05252;
  --sapphire: #5B8FE8;
  --violet: #9B7EE8;
  --amber: #E0943A;
  --text: #EDF0F7;
  --text2: #8A95AA;
}

html, body, [class*="css"] {
  font-family: 'Syne', sans-serif !important;
  background-color: #0A0B0D !important;
  color: #EDF0F7 !important;
}

/* Hide Streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }

/* App container */
.block-container {
  padding: 1.5rem 2rem 4rem !important;
  max-width: 1280px !important;
}

/* Brand header */
.brand-header {
  display: flex;
  align-items: center;
  gap: 14px;
  padding: 18px 0 22px;
  border-bottom: 1px solid rgba(255,255,255,0.07);
  margin-bottom: 28px;
}
.brand-mark {
  width: 38px; height: 38px;
  background: rgba(201,168,76,0.12);
  border: 1px solid rgba(201,168,76,0.4);
  border-radius: 9px;
  display: flex; align-items: center; justify-content: center;
  font-size: 18px;
}
.brand-name {
  font-family: 'Playfair Display', Georgia, serif !important;
  font-size: 22px;
  color: #EDF0F7;
}
.brand-tag {
  font-size: 9px;
  letter-spacing: 2.5px;
  text-transform: uppercase;
  color: #C9A84C;
}

/* KPI Cards */
.kpi-card {
  background: #12151A;
  border: 1px solid rgba(255,255,255,0.06);
  border-radius: 10px;
  padding: 17px 18px;
  position: relative;
  overflow: hidden;
  margin-bottom: 2px;
}
.kpi-accent-gold::before    { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:#C9A84C; }
.kpi-accent-emerald::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:#2DBD8A; }
.kpi-accent-crimson::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:#E05252; }
.kpi-accent-sapphire::before{ content:''; position:absolute; top:0; left:0; right:0; height:2px; background:#5B8FE8; }
.kpi-accent-violet::before  { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:#9B7EE8; }
.kpi-accent-amber::before   { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:#E0943A; }
.kpi-label { font-size: 9px; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; color: #8A95AA; margin-bottom: 9px; }
.kpi-val   { font-family: 'IBM Plex Mono', monospace !important; font-size: 22px; font-weight: 500; line-height: 1; margin-bottom: 5px; }
.kpi-foot  { font-size: 10px; color: #475060; }
.val-pos   { color: #2DBD8A !important; }
.val-neg   { color: #E05252 !important; }
.val-gold  { color: #C9A84C !important; }
.val-sapp  { color: #5B8FE8 !important; }
.val-vio   { color: #9B7EE8 !important; }
.foot-pos  { color: #2DBD8A !important; }
.foot-neg  { color: #E05252 !important; }

/* Verdict banner */
.verdict {
  border-radius: 10px;
  padding: 18px 24px;
  margin: 20px 0 28px;
  display: flex;
  align-items: flex-start;
  gap: 18px;
  border: 1px solid;
}
.verdict-undervalued { background: rgba(45,189,138,0.12); border-color: rgba(45,189,138,0.25); }
.verdict-overvalued  { background: rgba(224,82,82,0.12);  border-color: rgba(224,82,82,0.25); }
.verdict-fair        { background: rgba(201,168,76,0.12); border-color: rgba(201,168,76,0.25); }
.verdict-nodcf       { background: #12151A; border-color: rgba(255,255,255,0.06); }
.verdict-icon { font-size: 30px; flex-shrink: 0; margin-top: 2px; }
.verdict-title { font-family: 'Playfair Display', Georgia, serif !important; font-size: 18px; margin-bottom: 3px; color: #EDF0F7; }
.verdict-sub { font-size: 12.5px; color: #8A95AA; line-height: 1.6; }

/* Section titles */
.sec-title {
  font-size: 9.5px;
  font-weight: 700;
  letter-spacing: 1.8px;
  text-transform: uppercase;
  color: #8A95AA;
  margin-bottom: 16px;
  padding-bottom: 10px;
  border-bottom: 1px solid rgba(255,255,255,0.06);
}

/* Insight rows */
.insight-row {
  display: flex;
  align-items: flex-start;
  gap: 14px;
  padding: 15px 0;
  border-bottom: 1px solid rgba(255,255,255,0.06);
}
.insight-row:last-child { border-bottom: none; }
.insight-icon {
  width: 36px; height: 36px;
  border-radius: 8px;
  display: flex; align-items: center; justify-content: center;
  font-size: 16px; flex-shrink: 0;
}
.insight-body { flex: 1; line-height: 1.7; font-size: 13.5px; color: #EDF0F7; }
.insight-body strong { font-weight: 700; }

/* Badge */
.badge {
  display: inline-flex; align-items: center;
  padding: 2px 8px; border-radius: 4px;
  font-size: 10px; font-weight: 700; letter-spacing: 0.5px;
  text-transform: uppercase; margin-left: 6px;
}
.badge-green  { background: rgba(45,189,138,0.12); color: #2DBD8A; border: 1px solid rgba(45,189,138,0.2); }
.badge-red    { background: rgba(224,82,82,0.12);  color: #E05252; border: 1px solid rgba(224,82,82,0.2); }
.badge-gold   { background: rgba(201,168,76,0.12); color: #C9A84C; border: 1px solid rgba(201,168,76,0.25); }
.badge-blue   { background: rgba(91,143,232,0.12); color: #5B8FE8; border: 1px solid rgba(91,143,232,0.2); }

/* Styled table */
.styled-table {
  width: 100%; border-collapse: collapse;
  font-size: 12.5px; font-family: 'Syne', sans-serif;
}
.styled-table th {
  text-align: left; padding: 9px 12px;
  font-size: 9px; font-weight: 700; letter-spacing: 1.2px;
  text-transform: uppercase; color: #8A95AA;
  border-bottom: 1px solid rgba(255,255,255,0.07);
  background: #12151A;
}
.styled-table td {
  padding: 10px 12px;
  border-bottom: 1px solid rgba(255,255,255,0.04);
  color: #EDF0F7;
}
.styled-table td.num { text-align: right; font-family: 'IBM Plex Mono', monospace; font-size: 12px; }
.styled-table tr.hl td { background: rgba(201,168,76,0.06); font-weight: 600; }
.styled-table tr:hover td { background: rgba(255,255,255,0.015); }
.td-pos { color: #2DBD8A !important; }
.td-neg { color: #E05252 !important; }
.td-gold{ color: #C9A84C !important; }

/* sensitivity table */
.sens-table { width: 100%; border-collapse: collapse; font-size: 11.5px; }
.sens-table th { padding: 8px 10px; text-align: center; font-size: 9px; letter-spacing: 0.8px; text-transform: uppercase; color: #8A95AA; border: 1px solid rgba(255,255,255,0.06); background: #1C2130; font-weight: 700; }
.sens-table td { padding: 8px 10px; text-align: center; border: 1px solid rgba(255,255,255,0.06); font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #EDF0F7; }
.sens-good { background: rgba(45,189,138,0.15) !important; color: #2DBD8A !important; }
.sens-bad  { background: rgba(224,82,82,0.15) !important;  color: #E05252 !important; }
.sens-mid  { background: rgba(201,168,76,0.1) !important;  color: #C9A84C !important; }
.sens-current { border: 2px solid #C9A84C !important; color: #C9A84C !important; }

/* Assumption cards */
.assumption-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 12px; margin-bottom: 20px; }
.assumption-card { background: #1C2130; border: 1px solid rgba(255,255,255,0.1); border-radius: 6px; padding: 14px 16px; }
.assumption-label { font-size: 9px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: #8A95AA; margin-bottom: 6px; }
.assumption-val { font-family: 'IBM Plex Mono', monospace !important; font-size: 19px; font-weight: 500; color: #C9A84C; }
.assumption-note { font-size: 10.5px; color: #475060; margin-top: 3px; }

/* margin progress bars */
.margin-bar-wrap { margin-bottom: 12px; }
.margin-bar-label { display: flex; justify-content: space-between; font-size: 11.5px; color: #8A95AA; margin-bottom: 4px; }
.margin-bar-track { height: 6px; background: #1C2130; border-radius: 3px; overflow: hidden; }
.margin-bar-fill { height: 100%; border-radius: 3px; }

/* File uploader tweak */
[data-testid="stFileUploader"] {
  background: #12151A;
  border: 1.5px dashed rgba(255,255,255,0.11);
  border-radius: 14px;
  padding: 20px;
}
[data-testid="stFileUploader"]:hover {
  border-color: #C9A84C;
}

/* Tabs */
[data-testid="stTabs"] button {
  font-family: 'Syne', sans-serif !important;
  font-size: 12px !important;
  font-weight: 600 !important;
  color: #8A95AA !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
  color: #EDF0F7 !important;
  border-bottom-color: #C9A84C !important;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# COLOUR PALETTE (for Plotly)
# ─────────────────────────────────────────────
C = {
    "sapphire": "#5B8FE8",
    "emerald":  "#2DBD8A",
    "crimson":  "#E05252",
    "gold":     "#C9A84C",
    "violet":   "#9B7EE8",
    "amber":    "#E0943A",
    "text2":    "#8A95AA",
}
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="IBM Plex Mono, monospace", color="#8A95AA", size=11),
    xaxis=dict(gridcolor="rgba(255,255,255,0.05)", tickfont=dict(size=11)),
    yaxis=dict(gridcolor="rgba(255,255,255,0.05)", tickfont=dict(size=11)),
    margin=dict(l=10, r=10, t=30, b=10),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt(n, dec=0):
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "—"
    try:
        return f"{n:,.{dec}f}"
    except Exception:
        return "—"

def pct(n):
    if n is None or pd.isna(n):
        return "—"
    sign = "+" if n >= 0 else ""
    return f"{sign}{n:.1f}%"

def fmtpct(n):
    if n is None or pd.isna(n):
        return "—"
    return f"{n:.1f}%"

def fmtx(n):
    if n is None or pd.isna(n):
        return "—"
    return f"{n:.2f}x"

def safe(n, default=0):
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return default
    return n

def clamp(v, mn, mx):
    return max(mn, min(mx, v))


# ─────────────────────────────────────────────
# EXCEL PARSING
# ─────────────────────────────────────────────
def load_sheet_raw(wb, sheet_name):
    """Load a worksheet as a list-of-lists (raw)."""
    try:
        df = wb.parse(sheet_name, header=None, dtype=str)
        return df.fillna("").values.tolist()
    except Exception:
        return []

def find_row(raw, *terms):
    """Find first row whose first cell contains any of the terms (case-insensitive)."""
    terms_lower = [t.lower() for t in terms]
    for row in raw:
        c = str(row[0]).lower().strip() if row else ""
        if any(t in c for t in terms_lower):
            return row
    return None

def find_row_exact(raw, term):
    for row in raw:
        c = str(row[0]).lower().strip()
        if c == term.lower():
            return row
    return None

def find_pl_date_row(raw):
    """Find the report-date row under the P&L section."""
    # Look for PROFIT & LOSS header first, then find Report Date beneath it
    in_pl = False
    for row in raw:
        c = str(row[0]).lower().strip()
        if "profit" in c and ("loss" in c or "&" in c or "l" in c):
            in_pl = True
            continue
        if in_pl and "report date" in c:
            return row
        # Stop if we hit another major section
        if in_pl and c and any(s in c for s in ["balance sheet", "cash flow", "quarter", "derived", "price"]):
            break

    # fallback: find first 'Report Date' row that has annual dates (not quarterly)
    for row in raw:
        c = str(row[0]).lower().strip()
        if "report date" in c:
            years = get_years(row)
            # Annual data: years should be distinct and spaced ~1yr apart
            if len(years) >= 3:
                diffs = [years[i+1] - years[i] for i in range(len(years)-1)]
                if all(d == 1 for d in diffs):
                    return row

    # last fallback: first row with 4+ years in columns
    for row in raw:
        cells = [str(x) for x in row[1:]]
        if sum(1 for x in cells if re.search(r"\b20\d{2}\b", x)) >= 4:
            return row
    return None

def get_years(date_row):
    if not date_row:
        return []
    years = []
    for d in date_row[1:]:
        d = str(d).strip()
        if not d:
            continue
        m = re.search(r"(20\d{2}|19\d{2})", d)
        if m:
            years.append(int(m.group(1)))
    return years

def get_vals(row, count):
    if row is None:
        return [None] * count
    vals = []
    for v in row[1:count+1]:
        v = str(v).replace(",", "").strip()
        try:
            vals.append(float(v))
        except (ValueError, TypeError):
            vals.append(None)
    while len(vals) < count:
        vals.append(None)
    return vals

def find_dcf_row(dcf_raw, *terms):
    terms_lower = [t.lower() for t in terms]
    for row in dcf_raw:
        c0 = str(row[0] if len(row) > 0 else "").lower()
        c1 = str(row[1] if len(row) > 1 else "").lower()
        if any(t in c0 or t in c1 for t in terms_lower):
            return row
    return None

def get_dcf_val(row):
    if not row:
        return None
    label_in_b = str(row[0]).strip() == ""
    raw_val = row[2] if (label_in_b and len(row) > 2) else (row[1] if len(row) > 1 else None)
    try:
        return float(str(raw_val).replace(",", "").replace("%", "").strip())
    except (ValueError, TypeError):
        return None

def find_first_large_val(row):
    if not row:
        return None
    for cell in row[1:]:
        try:
            v = float(str(cell).replace(",", "").strip())
            if v > 100:
                return v
        except (ValueError, TypeError):
            continue
    return None


# ─────────────────────────────────────────────
# PLOTLY CHART HELPERS
# ─────────────────────────────────────────────
def bar_chart(labels, datasets, is_pct=False):
    fig = go.Figure()
    for ds in datasets:
        fig.add_trace(go.Bar(
            name=ds["label"], x=labels, y=ds["data"],
            marker_color=ds["color"],
            marker=dict(line=dict(width=0)),
        ))
    layout = {**PLOTLY_LAYOUT, "barmode": "group", "height": 280}
    layout["yaxis"] = dict(ticksuffix="%" if is_pct else "", gridcolor="rgba(255,255,255,0.05)")
    fig.update_layout(**layout)
    return fig

def line_chart(labels, datasets, is_pct=False):
    fig = go.Figure()
    for ds in datasets:
        fig.add_trace(go.Scatter(
            name=ds["label"], x=labels, y=ds["data"],
            line=dict(color=ds["color"], width=2.2),
            mode="lines+markers",
            marker=dict(size=5, color=ds["color"]),
            fill=ds.get("fill", None),
            fillcolor=ds.get("fillcolor", "rgba(0,0,0,0)"),
        ))
    layout = {**PLOTLY_LAYOUT, "height": 280}
    layout["yaxis"] = dict(ticksuffix="%" if is_pct else "", gridcolor="rgba(255,255,255,0.05)")
    fig.update_layout(**layout)
    return fig

def waterfall_chart(labels, values, colors=None):
    if colors is None:
        colors = [C["emerald"] if v >= 0 else C["crimson"] for v in values]
    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=colors,
        marker=dict(line=dict(width=0)),
    ))
    fig.update_layout(**PLOTLY_LAYOUT, height=260)
    return fig


# ─────────────────────────────────────────────
# KPI CARD HTML
# ─────────────────────────────────────────────
def kpi_card(label, val, footer="", val_class="", foot_class="", accent="gold"):
    return f"""
<div class="kpi-card kpi-accent-{accent}">
  <div class="kpi-label">{label}</div>
  <div class="kpi-val {val_class}">{val}</div>
  <div class="kpi-foot {foot_class}">{footer}</div>
</div>"""


# ─────────────────────────────────────────────
# STYLED TABLE HTML
# ─────────────────────────────────────────────
def styled_table(headers, rows, highlight_last=True):
    th = "".join(f'<th class="{"num" if i>0 else ""}">{h}</th>' for i, h in enumerate(headers))
    body = ""
    for ri, row in enumerate(rows):
        row_class = "hl" if (highlight_last and ri == len(rows)-1) else ""
        cells = "".join(
            f'<td class="{" ".join(filter(None, ["num" if ci>0 else "", row.get("cls_"+str(ci), "")])) }">'
            f'{row["vals"][ci]}</td>'
            for ci in range(len(headers))
        )
        body += f'<tr class="{row_class}">{cells}</tr>'
    return f'<table class="styled-table"><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>'


# ─────────────────────────────────────────────
# PARSE MAIN DATA SHEET
# ─────────────────────────────────────────────
def parse_data(raw):
    pl_date_row = find_pl_date_row(raw)
    if not pl_date_row:
        return None, "Could not find P&L Report Date row."
    years = get_years(pl_date_row)
    n = len(years)
    if not n:
        return None, "No year columns found."

    # Company info
    comp_row = next((r for r in raw if "COMPANY NAME" in str(r[0]).upper()), None)
    comp_name = str(comp_row[1]).strip() if comp_row else ""

    mcap_row  = find_row(raw, "market capitalization", "market cap", "mkt cap")
    price_row = find_row(raw, "current price", "cmp", "ltp")
    try: mcap  = float(str(mcap_row[1]).replace(",","")) if mcap_row else None
    except: mcap = None
    try: price = float(str(price_row[1]).replace(",","")) if price_row else None
    except: price = None

    # P&L
    sales_row    = find_row(raw, "sales", "revenue", "total revenue", "net sales")
    if not sales_row:
        return None, "Revenue/Sales row not found."
    rm_row       = find_row(raw, "raw material cost", "raw material", "cogs", "cost of goods")
    chginv_row   = find_row(raw, "change in inventory", "chg in inv")
    power_row    = find_row(raw, "power and fuel", "power & fuel", "energy cost")
    othermfr_row = find_row(raw, "other mfr", "other manufacturing", "mfr expenses")
    emp_row      = find_row(raw, "employee cost", "staff cost", "employee benefit", "salary")
    selling_row  = find_row(raw, "selling and admin", "selling & admin", "sg&a", "selling general")
    otherexp_row = find_row(raw, "other expenses", "other expenditure")
    otherinc_row = find_row(raw, "other income", "non-operating income")
    depr_row     = find_row(raw, "depreciation", "amortization", "d&a")
    int_row      = find_row(raw, "interest", "finance cost", "finance charges")
    pbt_row      = find_row(raw, "profit before tax", "pbt", "ebt")
    tax_row      = find_row_exact(raw, "tax") or find_row(raw, "income tax", "tax expense")
    np_row       = find_row(raw, "net profit", "profit after tax", "pat", "net income")
    if not np_row:
        return None, "Net Profit row not found."

    sales    = get_vals(sales_row, n)
    rm       = get_vals(rm_row, n)
    chginv   = get_vals(chginv_row, n)
    power    = get_vals(power_row, n)
    othermfr = get_vals(othermfr_row, n)
    emp      = get_vals(emp_row, n)
    selling  = get_vals(selling_row, n)
    otherexp = get_vals(otherexp_row, n)
    otherinc = get_vals(otherinc_row, n) if otherinc_row else [0]*n
    depr     = get_vals(depr_row, n)
    interest = get_vals(int_row, n)
    pbt      = get_vals(pbt_row, n)
    tax      = get_vals(tax_row, n) if tax_row else [None]*n
    np_      = get_vals(np_row, n)

    ebitda = []
    for i in range(n):
        if sales[i] is None:
            ebitda.append(None)
        else:
            opex = (safe(rm[i]) + safe(chginv[i]) + safe(power[i]) +
                    safe(othermfr[i]) + safe(emp[i]) + safe(selling[i]) + safe(otherexp[i]))
            ebitda.append(sales[i] - opex + safe(otherinc[i]))

    gross_profit   = [safe(sales[i]) - safe(rm[i]) - safe(chginv[i]) if sales[i] is not None else None for i in range(n)]
    gp_margin      = [gp/sales[i]*100 if gp is not None and sales[i] else None for i, gp in enumerate(gross_profit)]
    ebitda_margin  = [e/sales[i]*100 if e is not None and sales[i] else None for i, e in enumerate(ebitda)]
    np_margin      = [np_[i]/sales[i]*100 if np_[i] is not None and sales[i] else None for i in range(n)]

    # Balance Sheet
    equity      = get_vals(find_row(raw, "equity share capital", "share capital"), n)
    reserves    = get_vals(find_row(raw, "reserves", "retained earnings"), n)
    borrowings  = get_vals(find_row(raw, "borrowings", "total debt", "long term debt"), n)
    cash        = get_vals(find_row(raw, "cash & bank", "cash and bank", "cash equivalents"), n)
    inventory   = get_vals(find_row(raw, "inventory", "inventories"), n)
    receivables = get_vals(find_row(raw, "receivables", "trade receivables", "debtors"), n)
    net_block   = get_vals(find_row(raw, "net block", "net fixed assets", "ppe"), n)

    # Cash Flow
    cf_op  = get_vals(find_row(raw, "cash from operating", "operating activities", "cfo"), n)
    cf_inv = get_vals(find_row(raw, "cash from investing", "investing activities", "cfi"), n)
    cf_fin = get_vals(find_row(raw, "cash from financing", "financing activities", "cff"), n)
    cf_net = get_vals(find_row(raw, "net cash", "net change in cash"), n)

    latest = n - 1
    prev   = n - 2 if n >= 2 else 0

    sales_growth = (sales[latest]-sales[prev])/sales[prev]*100 if sales[latest] and sales[prev] else None
    np_growth    = (np_[latest]-np_[prev])/np_[prev]*100 if np_[latest] and np_[prev] else None
    cagr         = (pow(sales[latest]/sales[0], 1/(n-1))-1)*100 if sales[0] and sales[latest] and n>1 else None
    total_equity = safe(equity[latest]) + safe(reserves[latest])
    total_debt   = safe(borrowings[latest])
    de_ratio     = total_debt/total_equity if total_equity else None
    ocf_np       = cf_op[latest]/np_[latest] if cf_op[latest] and np_[latest] else None

    return {
        "comp_name": comp_name, "mcap": mcap, "price": price,
        "years": years, "n": n, "latest": latest, "prev": prev,
        "sales": sales, "ebitda": ebitda, "np": np_, "pbt": pbt,
        "depr": depr, "interest": interest, "tax": tax,
        "gp_margin": gp_margin, "ebitda_margin": ebitda_margin, "np_margin": np_margin,
        "gross_profit": gross_profit, "selling": selling, "emp": emp,
        "equity": equity, "reserves": reserves, "borrowings": borrowings,
        "cash": cash, "inventory": inventory, "receivables": receivables, "net_block": net_block,
        "cf_op": cf_op, "cf_inv": cf_inv, "cf_fin": cf_fin, "cf_net": cf_net,
        "sales_growth": sales_growth, "np_growth": np_growth, "cagr": cagr,
        "total_equity": total_equity, "total_debt": total_debt,
        "de_ratio": de_ratio, "ocf_np": ocf_np,
    }, None


def parse_dcf(dcf_raw, price):
    if not dcf_raw:
        return {"found": False}
    wacc_row       = find_dcf_row(dcf_raw, "wacc", "discount rate", "required rate")
    tgr_row        = find_dcf_row(dcf_raw, "terminal growth rate", "terminal growth", "perpetuity growth", "tgr")
    tv_row         = find_dcf_row(dcf_raw, "terminal value")
    pvtv_row       = find_dcf_row(dcf_raw, "pv of terminal", "pv terminal", "present value of terminal")
    pvfcf_row      = find_dcf_row(dcf_raw, "sum of pv of fcf", "sum of pv", "pv of fcf", "pv fcf", "present value of fcf")
    ev_row         = find_dcf_row(dcf_raw, "enterprise value", "ev", "total firm value")
    shares_row     = find_dcf_row(dcf_raw, "shares outstanding", "number of shares", "no. of shares", "equity shares outstanding")
    net_debt_row   = find_dcf_row(dcf_raw, "net debt", "less: net debt")
    intrinsic_row  = find_dcf_row(dcf_raw, "fair value per share", "intrinsic value per share",
                                  "intrinsic value", "fair value", "target price", "per share value", "value per share")
    fcf_row        = find_dcf_row(dcf_raw, "free cash flow to firm", "free cash flow", "fcf", "fcff", "unlevered free")

    wacc       = get_dcf_val(wacc_row)
    term_growth= get_dcf_val(tgr_row)
    if wacc and wacc > 1:       wacc       /= 100
    if term_growth and term_growth > 1: term_growth /= 100

    ev          = find_first_large_val(ev_row)
    pv_fcf      = find_first_large_val(pvfcf_row)
    pv_tv       = find_first_large_val(pvtv_row)
    intrinsic_ps= find_first_large_val(intrinsic_row)
    shares      = get_dcf_val(shares_row)
    net_debt    = get_dcf_val(net_debt_row)
    tv          = find_first_large_val(tv_row)

    # FCF projection values
    fcf_vals  = []
    fcf_years = []
    if fcf_row:
        label_in_b = str(fcf_row[0]).strip() == ""
        start = 2 if label_in_b else 1
        fcf_vals = []
        for v in fcf_row[start:]:
            try:
                fv = float(str(v).replace(",","").strip())
                if not pd.isna(fv):
                    fcf_vals.append(fv)
            except: pass
        fcf_idx = dcf_raw.index(fcf_row)
        if fcf_idx > 0:
            hr = dcf_raw[fcf_idx-1]
            for v in hr[1:]:
                m = re.search(r"(\d{4})", str(v))
                if m:
                    fcf_years.append(int(m.group(1)))

    mos = None
    if intrinsic_ps and price:
        mos = (intrinsic_ps - price) / price * 100

    found = bool(intrinsic_ps or ev or wacc)
    return {
        "found": found,
        "wacc": wacc, "term_growth": term_growth,
        "ev": ev, "pv_fcf": pv_fcf, "pv_tv": pv_tv,
        "intrinsic_ps": intrinsic_ps, "mos": mos,
        "shares": shares, "net_debt": net_debt, "tv": tv,
        "fcf_vals": fcf_vals, "fcf_years": fcf_years,
        "price": price,
    }


def parse_forecast(fcast_raw):
    if not fcast_raw:
        return {"found": False}
    rev_row  = find_row(fcast_raw, "revenue", "sales", "total revenue", "net revenue")
    np_row   = find_row(fcast_raw, "net profit", "pat", "net income")
    ebit_row = find_row(fcast_raw, "ebitda")
    date_row = next((r for r in fcast_raw if any(re.search(r"\d{4}", str(c)) for c in r)), None)
    f_years  = get_years(date_row) if date_row else []
    f_rev    = get_vals(rev_row,  len(f_years)) if rev_row  else []
    f_np     = get_vals(np_row,   len(f_years)) if np_row   else []
    f_ebit   = get_vals(ebit_row, len(f_years)) if ebit_row else []
    if f_years and (any(v is not None for v in f_rev) or any(v is not None for v in f_np)):
        return {"found": True, "years": f_years, "rev": f_rev, "np": f_np, "ebit": f_ebit}
    return {"found": False}


# ─────────────────────────────────────────────
# RENDER INSIGHTS
# ─────────────────────────────────────────────
def render_insights(d, dcf):
    yrs = d["years"]
    lat = d["latest"]
    n   = d["n"]
    insights = []

    # 1. Revenue CAGR
    cagr = d["cagr"]
    cagr_str = f"{cagr:.1f}%" if cagr is not None else "N/A"
    commentary = ("Strong top-line growth trajectory." if (cagr or 0)>15
                  else "Moderate growth pace." if (cagr or 0)>8
                  else "Growth has been tepid; monitor for demand constraints.")
    insights.append(dict(
        icon="📊", bg="rgba(91,143,232,0.1)",
        text=f"<strong>Revenue CAGR of {cagr_str}</strong> over {n-1} years — "
             f"₹{fmt(d['sales'][0])} Cr (FY{yrs[0]}) → ₹{fmt(d['sales'][lat])} Cr (FY{yrs[lat]}). {commentary}"
    ))

    # 2. Profitability
    np_lat = d["np"][lat]
    prof_dir = "improving" if (np_lat or 0) > (d["np"][0] or 0) else "declining"
    npm = d["np_margin"][lat]
    p_comment = ("Losses raise near-term solvency concerns." if (np_lat or 0)<0
                 else "Healthy net margin indicative of pricing power." if (npm or 0)>15
                 else "Margins have room to expand.")
    insights.append(dict(
        icon="✅" if (np_lat or 0)>=0 else "⚠️",
        bg="rgba(45,189,138,0.1)" if (np_lat or 0)>=0 else "rgba(224,82,82,0.1)",
        text=f"Net profitability is <strong>{prof_dir}</strong>. Latest PAT: ₹{fmt(np_lat)} Cr "
             f"(margin: {fmtpct(npm)}). {p_comment}"
    ))

    # 3. EBITDA
    em0  = d["ebitda_margin"][0]
    emL  = d["ebitda_margin"][lat]
    em_dir = "expanded" if (emL or 0)>(em0 or 0) else "contracted"
    em_comment = ("Above-average operational efficiency." if (emL or 0)>20
                  else "Decent operating leverage." if (emL or 0)>12
                  else "Pressure on operational costs — monitor input cost inflation.")
    insights.append(dict(
        icon="🏭",
        bg="rgba(45,189,138,0.1)" if (emL or 0)>(em0 or 0) else "rgba(224,82,82,0.1)",
        text=f"EBITDA margin has <strong>{em_dir}</strong> from {fmtpct(em0)} → {fmtpct(emL)}. {em_comment}"
    ))

    # 4. Leverage
    de = d["de_ratio"]
    de_comment = ("High leverage — refinancing risk elevated." if (de or 0)>3
                  else "Moderate debt load. Watch interest coverage." if (de or 0)>1.5
                  else "Low leverage — strong balance sheet position." if (de or 0)<0.5
                  else "Manageable debt levels.")
    insights.append(dict(
        icon="🔴" if (de or 0)>2 else "🟡",
        bg="rgba(45,189,138,0.1)" if (de or 0)<1 else "rgba(224,148,58,0.1)",
        text=f"Debt/Equity ratio: <strong>{fmtx(de)}</strong>. "
             f"Borrowings: ₹{fmt(d['borrowings'][lat])} Cr. {de_comment}"
    ))

    # 5. Cash Quality
    cfo = d["cf_op"][lat]
    ocf = d["ocf_np"]
    ocf_comment = ("Excellent cash conversion — profits well-backed by cash." if (ocf or 0)>1.2
                   else "Good earnings quality." if (ocf or 0)>0.8
                   else "Negative OCF despite profits is a red flag — check working capital." if (ocf or 0)<0
                   else "Earnings outpacing cash — possible working capital stretch.")
    insights.append(dict(
        icon="💵",
        bg="rgba(45,189,138,0.1)" if (cfo or 0)>0 else "rgba(224,82,82,0.1)",
        text=f"Operating cash flow FY{yrs[lat]}: <strong>₹{fmt(cfo)} Cr</strong>. "
             f"OCF/PAT: <strong>{fmtx(ocf)}</strong>. {ocf_comment}"
    ))

    # 6. DCF
    if dcf.get("found") and dcf.get("intrinsic_ps"):
        mos = dcf["mos"]
        mos_comment = ("Significant discount to intrinsic value — attractive if assumptions are conservative." if (mos or 0)>30
                       else "Moderate upside. Accumulate on dips." if (mos or 0)>10
                       else "Stock trading at a steep premium to intrinsic value. Avoid." if (mos or 0)<-20
                       else "Near fair value — hold if already invested.")
        insights.append(dict(
            icon="⬡",
            bg="rgba(201,168,76,0.1)" if (mos or 0)>0 else "rgba(224,82,82,0.1)",
            text=f"<strong>DCF Valuation:</strong> Intrinsic value ₹{fmt(dcf['intrinsic_ps'],0)}/share vs. "
                 f"CMP ₹{fmt(dcf.get('price'),0)}. Margin of safety: "
                 f"<strong>{abs(mos):.1f}% {'upside' if (mos or 0)>0 else 'downside'}</strong>. {mos_comment}"
        ))
    else:
        insights.append(dict(
            icon="💡", bg="rgba(155,126,232,0.1)",
            text="No DCF sheet found. Add a <strong>DCF</strong> or <strong>Valuation</strong> sheet "
                 "with WACC, FCF projections, terminal growth rate, and intrinsic value per share "
                 "for automated valuation verdicts."
        ))

    # 7. Gross margin
    gpm = d["gp_margin"][lat]
    sell = d["selling"][lat]
    sal  = d["sales"][lat]
    sell_pct = f"{sell/sal*100:.1f}" if (sell and sal) else "?"
    gp_comment = ("Wide gross margins — strong pricing power or low material intensity." if (gpm or 0)>50
                  else "Healthy gross margins." if (gpm or 0)>30
                  else "Thin gross margins — business likely commodity or capital-intensive.")
    insights.append(dict(
        icon="📦", bg="rgba(91,143,232,0.08)",
        text=f"Gross margin (latest FY): <strong>{fmtpct(gpm)}</strong>. "
             f"S&A expense: ₹{fmt(sell)} Cr ({sell_pct}% of revenue). {gp_comment}"
    ))

    html = '<div class="sec-title">Investment-Grade Insights</div>'
    for ins in insights:
        html += f"""
<div class="insight-row">
  <div class="insight-icon" style="background:{ins['bg']}">{ins['icon']}</div>
  <div class="insight-body">{ins['text']}</div>
</div>"""
    return html


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
def main():
    # Brand header
    st.markdown("""
<div class="brand-header">
  <div class="brand-mark">📈</div>
  <div>
    <div class="brand-name">CapitalLens</div>
    <div class="brand-tag">Investment Research</div>
  </div>
</div>""", unsafe_allow_html=True)

    # ── Upload ──
    if "data" not in st.session_state:
        st.markdown("""
<div style="text-align:center;padding:30px 0 10px">
  <div style="font-size:10px;letter-spacing:3px;text-transform:uppercase;color:#C9A84C;margin-bottom:16px">Financial Analysis</div>
  <div style="font-family:'Playfair Display',Georgia,serif;font-size:38px;line-height:1.1;letter-spacing:-0.5px;margin-bottom:14px">
    Analyse Any <em style="font-style:italic;color:#C9A84C">Financial Model</em>
  </div>
  <div style="font-size:13.5px;color:#8A95AA;line-height:1.7;max-width:500px;margin:0 auto 32px">
    Upload your Excel workbook with P&amp;L, Balance Sheet, Cash Flow,
    DCF, and Forecast sheets. Get instant investment-grade insights.
  </div>
</div>""", unsafe_allow_html=True)

        uploaded = st.file_uploader(
            "Drop your Excel file here", type=["xlsx", "xls"],
            label_visibility="collapsed"
        )

        if uploaded:
            with st.spinner("Parsing workbook…"):
                try:
                    xf = pd.ExcelFile(uploaded)
                    all_sheets = xf.sheet_names

                    data_sheet = next((s for s in all_sheets if "data" in s.lower()), None)
                    if not data_sheet:
                        st.error("❌ No 'Data Sheet' tab found. Please check your file.")
                        return

                    wb = pd.ExcelFile(uploaded)
                    raw      = load_sheet_raw(wb, data_sheet)

                    dcf_sheet   = next((s for s in all_sheets if any(k in s.lower() for k in ["dcf","valuat","intrinsic"])), None)
                    fcast_sheet = next((s for s in all_sheets if any(k in s.lower() for k in ["forecast","project","fcast"])
                                       and "dcf" not in s.lower()), None)

                    dcf_raw   = load_sheet_raw(wb, dcf_sheet)   if dcf_sheet   else None
                    fcast_raw = load_sheet_raw(wb, fcast_sheet) if fcast_sheet else None

                    d, err = parse_data(raw)
                    if err:
                        st.error(f"❌ {err}")
                        return

                    dcf      = parse_dcf(dcf_raw, d["price"])
                    forecast = parse_forecast(fcast_raw)

                    st.session_state["data"]      = d
                    st.session_state["dcf"]       = dcf
                    st.session_state["forecast"]  = forecast
                    st.session_state["all_sheets"]= all_sheets
                    st.session_state["file_name"] = uploaded.name
                    st.rerun()
                except Exception as ex:
                    st.error(f"❌ Could not read file: {ex}")
        return

    # ── Dashboard ──
    d        = st.session_state["data"]
    dcf      = st.session_state["dcf"]
    forecast = st.session_state["forecast"]
    yrs      = d["years"]
    n        = d["n"]
    lat      = d["latest"]
    prev     = d["prev"]
    yl       = [f"FY{y}" for y in yrs]

    # Top row: company name + reset
    col_name, col_reset = st.columns([5, 1])
    with col_name:
        st.markdown(f"""
<div style="font-family:'Playfair Display',Georgia,serif;font-size:30px;letter-spacing:-0.4px">{d['comp_name'] or "Company"}</div>
<div style="font-size:11.5px;color:#8A95AA;margin-top:6px">
  FY{yrs[0]} – FY{yrs[lat]} &nbsp;·&nbsp; {n} years &nbsp;·&nbsp; {len(d.get('years',[]))} data years
</div>""", unsafe_allow_html=True)
    with col_reset:
        if st.button("↩ Reset", use_container_width=True):
            for k in ["data","dcf","forecast","all_sheets","file_name"]:
                st.session_state.pop(k, None)
            st.rerun()

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Verdict Banner ──
    if dcf.get("found") and dcf.get("intrinsic_ps") and d["price"]:
        mos = dcf["mos"] or 0
        ip  = dcf["intrinsic_ps"]
        pr  = d["price"]
        if mos > 20:
            vcls, vico, vtit = "undervalued", "🟢", f"Potentially Undervalued — {mos:.1f}% Margin of Safety"
            vsub = f"DCF intrinsic value of ₹{fmt(ip,0)}/share vs. CMP ₹{fmt(pr,0)}/share. The stock appears to trade at a significant discount to its intrinsic worth. Always verify assumptions."
        elif mos < -10:
            vcls, vico, vtit = "overvalued", "🔴", f"Potentially Overvalued — {abs(mos):.1f}% Premium to Intrinsic Value"
            vsub = f"DCF intrinsic value of ₹{fmt(ip,0)}/share vs. CMP ₹{fmt(pr,0)}/share. Stock appears pricey relative to modelled cash flows. Re-check growth assumptions."
        else:
            vcls, vico, vtit = "fair", "🟡", "Fairly Valued — Within ±10% of Intrinsic Value"
            vsub = f"DCF intrinsic value of ₹{fmt(ip,0)}/share vs. CMP ₹{fmt(pr,0)}/share. Stock is approximately priced to its modelled value. Watch margin and growth execution."
        st.markdown(f"""
<div class="verdict verdict-{vcls}">
  <div class="verdict-icon">{vico}</div>
  <div>
    <div class="verdict-title">{vtit}</div>
    <div class="verdict-sub">{vsub}</div>
  </div>
</div>""", unsafe_allow_html=True)

    # ── KPI Row 1: Operating ──
    sg = d["sales_growth"]
    ng = d["np_growth"]
    cg = d["cagr"]
    em = d["ebitda_margin"][lat]

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(kpi_card("Revenue (₹ Cr)", fmt(d["sales"][lat]),
            pct(sg)+" YoY" if sg is not None else f"FY{yrs[lat]}",
            "val-sapp", "foot-pos" if (sg or 0)>=0 else "foot-neg", "sapphire"), unsafe_allow_html=True)
    with k2:
        np_lat = d["np"][lat]
        st.markdown(kpi_card("Net Profit (₹ Cr)", fmt(np_lat),
            pct(ng)+" YoY" if ng is not None else f"FY{yrs[lat]}",
            "val-pos" if (np_lat or 0)>=0 else "val-neg",
            "foot-pos" if (ng or 0)>=0 else "foot-neg",
            "emerald" if (np_lat or 0)>=0 else "crimson"), unsafe_allow_html=True)
    with k3:
        st.markdown(kpi_card("EBITDA Margin", fmtpct(em), f"FY{yrs[lat]}",
            "val-gold", "", "gold"), unsafe_allow_html=True)
    with k4:
        st.markdown(kpi_card("Revenue CAGR", fmtpct(cg), f"{n-1}-yr compound",
            "val-pos", "foot-pos", "emerald"), unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── KPI Row 2: Valuation & Leverage ──
    de = d["de_ratio"]
    ocf = d["ocf_np"]
    k5, k6, k7, k8 = st.columns(4)
    with k5:
        st.markdown(kpi_card("Market Cap (₹ Cr)", fmt(d["mcap"]) if d["mcap"] else "—",
            f"CMP ₹{fmt(d['price'],0)}" if d["price"] else "N/A",
            "val-sapp", "", "sapphire"), unsafe_allow_html=True)
    with k6:
        de_cls = "val-neg" if (de or 0)>2 else "val-pos" if (de or 0)<1 else "val-gold"
        st.markdown(kpi_card("Debt / Equity", fmtx(de),
            f"₹{fmt(d['total_debt'])} Cr debt", de_cls, "", "crimson" if (de or 0)>2 else "gold"), unsafe_allow_html=True)
    with k7:
        ocf_cls = "val-pos" if (ocf or 0)>1 else "val-neg" if (ocf or 0)<0.5 else ""
        st.markdown(kpi_card("OCF / Net Profit", fmtx(ocf), "Earnings quality",
            ocf_cls, "", "emerald" if (ocf or 0)>1 else "amber"), unsafe_allow_html=True)
    with k8:
        if dcf.get("found") and dcf.get("intrinsic_ps"):
            mos = dcf["mos"] or 0
            st.markdown(kpi_card("Intrinsic Value (₹/sh)", f"₹{fmt(dcf['intrinsic_ps'],0)}",
                f"{mos:.1f}% {'upside' if mos>0 else 'downside'}",
                "val-pos" if mos>0 else "val-neg",
                "foot-pos" if mos>0 else "foot-neg",
                "emerald" if mos>0 else "crimson"), unsafe_allow_html=True)
        else:
            st.markdown(kpi_card("Intrinsic Value", "—", "Upload DCF sheet", "val-vio", "", "violet"), unsafe_allow_html=True)

    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

    # ── TABS ──
    tab_labels = ["📊 P&L", "📐 Margins", "🏦 Balance Sheet", "💧 Cash Flow", "⬡ DCF", "🔭 Forecast", "💡 Insights"]
    tabs = st.tabs(tab_labels)

    # ── TAB 0: P&L ──────────────────────────────
    with tabs[0]:
        st.markdown('<div class="sec-title">Revenue, EBITDA & Net Profit (₹ Cr)</div>', unsafe_allow_html=True)
        fig = bar_chart(yl, [
            {"label": "Revenue",    "data": d["sales"],  "color": "rgba(91,143,232,0.70)"},
            {"label": "EBITDA",     "data": d["ebitda"], "color": "rgba(45,189,138,0.70)"},
            {"label": "Net Profit", "data": d["np"],     "color": "rgba(224,82,82,0.75)"},
        ])
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-title" style="margin-top:20px">Annual Growth Table</div>', unsafe_allow_html=True)
        rows = []
        for i, y in enumerate(yrs):
            sg_ = (d["sales"][i]-d["sales"][i-1])/d["sales"][i-1]*100 if i>0 and d["sales"][i] and d["sales"][i-1] else None
            ng_ = (d["np"][i]-d["np"][i-1])/d["np"][i-1]*100 if i>0 and d["np"][i] and d["np"][i-1] else None
            rows.append({"vals": [
                f"FY{y}", fmt(d["sales"][i]), pct(sg_), fmt(d["ebitda"][i]),
                fmtpct(d["ebitda_margin"][i]), fmt(d["np"][i]), pct(ng_)
            ], "cls_2": "td-pos" if (sg_ or 0)>=0 else "td-neg",
               "cls_5": "td-pos" if (d["np"][i] or 0)>=0 else "td-neg",
               "cls_6": "td-pos" if (ng_ or 0)>=0 else "td-neg"})
        st.markdown(styled_table(
            ["Year","Revenue","Rev Growth","EBITDA","EBITDA Margin","Net Profit","NP Growth"], rows
        ), unsafe_allow_html=True)

    # ── TAB 1: Margins ──────────────────────────
    with tabs[1]:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.markdown('<div class="sec-title">Margin Trends (%)</div>', unsafe_allow_html=True)
            fig = line_chart(yl, [
                {"label": "Gross Margin",  "data": d["gp_margin"],     "color": C["sapphire"],
                 "fill": "tozeroy", "fillcolor": "rgba(91,143,232,0.06)"},
                {"label": "EBITDA Margin", "data": d["ebitda_margin"], "color": C["emerald"]},
                {"label": "Net Margin",    "data": d["np_margin"],     "color": C["crimson"]},
            ], is_pct=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with col_b:
            st.markdown('<div class="sec-title">Latest FY Margins</div>', unsafe_allow_html=True)
            for label, val, color in [
                ("Gross Margin", d["gp_margin"][lat], C["sapphire"]),
                ("EBITDA Margin", d["ebitda_margin"][lat], C["emerald"]),
                ("Net Margin", d["np_margin"][lat], C["crimson"]),
            ]:
                w = clamp(val or 0, 0, 100)
                st.markdown(f"""
<div class="margin-bar-wrap">
  <div class="margin-bar-label"><span>{label}</span>
    <span style="color:{color};font-family:'IBM Plex Mono',monospace">{fmtpct(val)}</span></div>
  <div class="margin-bar-track">
    <div class="margin-bar-fill" style="width:{w}%;background:{color}"></div>
  </div>
</div>""", unsafe_allow_html=True)

        st.markdown('<div class="sec-title" style="margin-top:20px">Margin History</div>', unsafe_allow_html=True)
        rows = [{"vals": [f"FY{y}", fmtpct(d["gp_margin"][i]), fmtpct(d["ebitda_margin"][i]), fmtpct(d["np_margin"][i])]}
                for i, y in enumerate(yrs)]
        st.markdown(styled_table(["Year","Gross Margin","EBITDA Margin","Net Margin"], rows), unsafe_allow_html=True)

    # ── TAB 2: Balance Sheet ─────────────────────
    with tabs[2]:
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown('<div class="sec-title">Equity vs Borrowings (₹ Cr)</div>', unsafe_allow_html=True)
            total_eq = [safe(d["equity"][i]) + safe(d["reserves"][i]) for i in range(n)]
            fig = bar_chart(yl, [
                {"label": "Total Equity",  "data": total_eq,        "color": "rgba(45,189,138,0.70)"},
                {"label": "Borrowings",    "data": d["borrowings"], "color": "rgba(224,82,82,0.70)"},
            ])
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with col_b:
            st.markdown('<div class="sec-title">Assets Overview (₹ Cr)</div>', unsafe_allow_html=True)
            fig = bar_chart(yl, [
                {"label": "Net Block",   "data": d["net_block"],   "color": "rgba(91,143,232,0.70)"},
                {"label": "Cash",        "data": d["cash"],        "color": "rgba(201,168,76,0.70)"},
                {"label": "Receivables", "data": d["receivables"], "color": "rgba(155,126,232,0.60)"},
                {"label": "Inventory",   "data": d["inventory"],   "color": "rgba(224,148,58,0.60)"},
            ])
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        st.markdown('<div class="sec-title" style="margin-top:20px">Balance Sheet Summary (₹ Cr)</div>', unsafe_allow_html=True)
        rows = [{"vals": [f"FY{y}",
            fmt(safe(d["equity"][i])+safe(d["reserves"][i])),
            fmt(d["borrowings"][i]), fmt(d["cash"][i]),
            fmt(d["inventory"][i]), fmt(d["receivables"][i]), fmt(d["net_block"][i])]}
                for i, y in enumerate(yrs)]
        st.markdown(styled_table(
            ["Year","Total Equity","Borrowings","Cash","Inventory","Receivables","Net Block"], rows
        ), unsafe_allow_html=True)

    # ── TAB 3: Cash Flow ─────────────────────────
    with tabs[3]:
        col_a, col_b = st.columns([2, 1])
        with col_a:
            st.markdown('<div class="sec-title">Cash Flow Waterfall (₹ Cr)</div>', unsafe_allow_html=True)
            fig = bar_chart(yl, [
                {"label": "CFO (Operating)",  "data": d["cf_op"],  "color": "rgba(45,189,138,0.75)"},
                {"label": "CFI (Investing)",  "data": d["cf_inv"], "color": "rgba(224,82,82,0.70)"},
                {"label": "CFF (Financing)",  "data": d["cf_fin"], "color": "rgba(91,143,232,0.70)"},
            ])
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with col_b:
            st.markdown('<div class="sec-title">Latest FY Cash Flows</div>', unsafe_allow_html=True)
            for label, val, clr in [
                ("Operating (CFO)", d["cf_op"][lat],  C["emerald"]),
                ("Investing (CFI)", d["cf_inv"][lat], C["crimson"]),
                ("Financing (CFF)", d["cf_fin"][lat], C["sapphire"]),
                ("Net Change",      d["cf_net"][lat] if d["cf_net"] else None, C["gold"]),
            ]:
                color = clr if (val or 0)>=0 else C["crimson"]
                st.markdown(f"""
<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid rgba(255,255,255,0.05)">
  <span style="font-size:12px;color:#8A95AA">{label}</span>
  <span style="font-family:'IBM Plex Mono',monospace;font-size:14px;color:{color}">₹{fmt(val)} Cr</span>
</div>""", unsafe_allow_html=True)

        st.markdown('<div class="sec-title" style="margin-top:20px">Cash Flow History (₹ Cr)</div>', unsafe_allow_html=True)
        rows = [{"vals": [f"FY{y}", fmt(d["cf_op"][i]), fmt(d["cf_inv"][i]),
                          fmt(d["cf_fin"][i]), fmt(d["cf_net"][i]) if d["cf_net"] else "—"],
                 "cls_1": "td-pos" if (d["cf_op"][i] or 0)>=0 else "td-neg"}
                for i, y in enumerate(yrs)]
        st.markdown(styled_table(["Year","CFO","CFI","CFF","Net Cash"], rows), unsafe_allow_html=True)

    # ── TAB 4: DCF ───────────────────────────────
    with tabs[4]:
        if not dcf.get("found"):
            st.markdown("""
<div style="text-align:center;padding:60px 20px">
  <div style="font-size:48px;margin-bottom:16px">⬡</div>
  <div style="font-family:'Playfair Display',Georgia,serif;font-size:22px;margin-bottom:10px">No DCF Sheet Detected</div>
  <div style="font-size:13px;color:#8A95AA;line-height:1.7;max-width:420px;margin:0 auto">
    Add a sheet named <strong>DCF</strong> or <strong>Valuation</strong> with your WACC,
    FCF projections, terminal growth rate, and intrinsic value per share for automated
    valuation verdicts and sensitivity analysis.
  </div>
</div>""", unsafe_allow_html=True)
        else:
            # Assumption cards
            assumptions = [
                ("WACC", f"{dcf['wacc']*100:.1f}%" if dcf.get("wacc") else "N/A", "Discount rate applied to FCFs"),
                ("Terminal Growth", f"{dcf['term_growth']*100:.1f}%" if dcf.get("term_growth") else "N/A", "Perpetuity growth rate"),
                ("Enterprise Value", f"₹{fmt(dcf.get('ev'),0)} Cr" if dcf.get("ev") else "N/A", "Total firm value"),
                ("PV of FCFs", f"₹{fmt(dcf.get('pv_fcf'),0)} Cr" if dcf.get("pv_fcf") else "N/A", "Sum of discounted FCFs"),
                ("PV of Terminal Value", f"₹{fmt(dcf.get('pv_tv'),0)} Cr" if dcf.get("pv_tv") else "N/A", "Discounted terminal value"),
                ("Intrinsic Value/Share", f"₹{fmt(dcf.get('intrinsic_ps'),0)}" if dcf.get("intrinsic_ps") else "N/A", "Fair value per equity share"),
            ]
            cols = st.columns(3)
            for idx, (lbl, val, note) in enumerate(assumptions):
                with cols[idx % 3]:
                    st.markdown(f"""
<div class="assumption-card">
  <div class="assumption-label">{lbl}</div>
  <div class="assumption-val">{val}</div>
  <div class="assumption-note">{note}</div>
</div>""", unsafe_allow_html=True)

            # FCF chart
            if dcf.get("fcf_vals"):
                st.markdown('<div class="sec-title" style="margin-top:24px">Free Cash Flow Projections (₹ Cr)</div>', unsafe_allow_html=True)
                fc_yrs = [str(y) for y in dcf["fcf_years"]] if dcf.get("fcf_years") else [f"Y{i+1}" for i in range(len(dcf["fcf_vals"]))]
                fig = waterfall_chart(fc_yrs[:len(dcf["fcf_vals"])], dcf["fcf_vals"])
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            # Sensitivity table
            if dcf.get("wacc") and dcf.get("term_growth") and dcf.get("intrinsic_ps"):
                st.markdown('<div class="sec-title" style="margin-top:24px">Sensitivity: Intrinsic Value vs WACC × Terminal Growth (₹/share)</div>', unsafe_allow_html=True)
                base_wacc = dcf["wacc"]
                base_tg   = dcf["term_growth"]
                base_iv   = dcf["intrinsic_ps"]
                wacc_range = [base_wacc - 0.02, base_wacc - 0.01, base_wacc, base_wacc + 0.01, base_wacc + 0.02]
                tg_range   = [base_tg - 0.01, base_tg, base_tg + 0.01, base_tg + 0.02]
                price      = dcf.get("price") or base_iv

                def approx_iv(w, g):
                    # Scale: if wacc increases, IV decreases proportionally
                    scale = (base_wacc - base_tg) / ((w - g) if w != g else 0.001)
                    return round(base_iv * scale, 0)

                header_html = "<tr><th>WACC \\ TGR</th>" + "".join(f"<th>{g*100:.1f}%</th>" for g in tg_range) + "</tr>"
                rows_html = ""
                for w in wacc_range:
                    rows_html += f"<tr><th>{w*100:.1f}%</th>"
                    for g in tg_range:
                        iv = approx_iv(w, g)
                        is_cur = abs(w - base_wacc) < 0.005 and abs(g - base_tg) < 0.005
                        diff   = (iv - price) / price * 100 if price else 0
                        cls = "sens-current" if is_cur else ("sens-good" if diff > 10 else "sens-bad" if diff < -10 else "sens-mid")
                        rows_html += f'<td class="{cls}">₹{fmt(iv,0)}</td>'
                    rows_html += "</tr>"

                st.markdown(f"""
<table class="sens-table">
<thead>{header_html}</thead>
<tbody>{rows_html}</tbody>
</table>""", unsafe_allow_html=True)

    # ── TAB 5: Forecast ──────────────────────────
    with tabs[5]:
        if not forecast.get("found"):
            st.markdown("""
<div style="text-align:center;padding:60px 20px">
  <div style="font-size:48px;margin-bottom:16px">🔭</div>
  <div style="font-family:'Playfair Display',Georgia,serif;font-size:22px;margin-bottom:10px">No Forecast Sheet Detected</div>
  <div style="font-size:13px;color:#8A95AA;line-height:1.7;max-width:420px;margin:0 auto">
    Add a sheet named <strong>Forecast</strong>, <strong>Revenue Forecast</strong>, or <strong>Projections</strong>
    to see forward estimates alongside historical performance.
  </div>
</div>""", unsafe_allow_html=True)
        else:
            f = forecast
            f_yl = [f"FY{y}" for y in f["years"]]
            combined_labels = [f"FY{y} (H)" for y in yrs] + [f"FY{y} (F)" for y in f["years"]]
            combined_rev = (d["sales"] or []) + (f["rev"] or [])
            combined_np  = (d["np"] or []) + (f["np"] or [])
            split = len(yrs)
            rev_colors = ["rgba(91,143,232,0.65)"]*split + ["rgba(201,168,76,0.75)"]*len(f["years"])

            fig = go.Figure()
            fig.add_trace(go.Bar(name="Revenue", x=combined_labels, y=combined_rev,
                                 marker_color=rev_colors, marker=dict(line=dict(width=0))))
            fig.add_trace(go.Bar(name="Net Profit", x=combined_labels, y=combined_np,
                                 marker_color=["rgba(45,189,138,0.65)"]*split + ["rgba(45,189,138,0.4)"]*len(f["years"]),
                                 marker=dict(line=dict(width=0))))
            fig.update_layout(**PLOTLY_LAYOUT, barmode="group", height=280)
            st.markdown('<div class="sec-title">Revenue Forecast vs Historical (₹ Cr)</div>', unsafe_allow_html=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            st.markdown('<div class="sec-title" style="margin-top:20px">Forecast Summary</div>', unsafe_allow_html=True)
            rows = [{"vals": [
                f"FY{f['years'][i]}", fmt(f["rev"][i]), fmt(f["ebit"][i]),
                fmt(f["np"][i])
            ], "cls_3": "td-pos" if (f["np"][i] or 0)>=0 else "td-neg"}
                    for i in range(len(f["years"]))]
            st.markdown(styled_table(
                ["Year","Revenue (₹Cr)","EBITDA (₹Cr)","Net Profit (₹Cr)"], rows, highlight_last=False
            ), unsafe_allow_html=True)

    # ── TAB 6: Insights ──────────────────────────
    with tabs[6]:
        st.markdown(render_insights(d, dcf), unsafe_allow_html=True)


# ─────────────────────────────────────────────
if __name__ == "__main__":
    main()
