import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Myntra Reconciliation",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

/* ── Global font ── */
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* ── Sidebar background only ── */
section[data-testid="stSidebar"] {
    background: #0f172a !important;
    border-right: 1px solid #1e293b;
}

/* ── Sidebar: only target text nodes we own ── */
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown li,
section[data-testid="stSidebar"] .stMarkdown span,
section[data-testid="stSidebar"] .element-container p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stMultiSelect label,
section[data-testid="stSidebar"] .stDateInput label,
section[data-testid="stSidebar"] .stFileUploader label,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] .upload-hint,
section[data-testid="stSidebar"] .upload-hint b {
    color: #e2e8f0 !important;
}
section[data-testid="stSidebar"] .upload-hint b { color: #7dd3fc !important; }

/* sidebar multiselect tags */
section[data-testid="stSidebar"] [data-baseweb="tag"] {
    background: #1e40af !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"] span {
    color: #ffffff !important;
}

/* ── Main area ── */
.main .block-container { padding-top: 1.5rem; max-width: 1400px; }

/* ── Metric cards ── */
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.1rem 1.25rem;
    box-shadow: 0 2px 6px rgba(0,0,0,.07);
}
.metric-label {
    font-size: .72rem; font-weight: 700; color: #475569;
    text-transform: uppercase; letter-spacing: .07em; margin-bottom: .3rem;
}
.metric-value {
    font-size: 1.5rem; font-weight: 700; color: #0f172a;
    font-family: 'JetBrains Mono', monospace;
}
.metric-value.positive { color: #15803d; }
.metric-value.negative { color: #b91c1c; }
.metric-sub { font-size: .72rem; color: #64748b; margin-top: .2rem; font-weight: 500; }

/* ── Section headers ── */
.section-header {
    display: flex; align-items: center; gap: .6rem;
    font-size: 1rem; font-weight: 700; color: #0f172a;
    border-bottom: 2px solid #e2e8f0; padding-bottom: .5rem; margin-bottom: 1rem;
}

/* ── Info / warn / success boxes ── */
.info-box {
    background: #eff6ff; border-left: 3px solid #3b82f6;
    border-radius: 0 8px 8px 0; padding: .8rem 1rem;
    font-size: .83rem; color: #1e3a5f; font-weight: 500;
}
.info-box b  { color: #1d4ed8; }
.info-box code { background: #dbeafe; color: #1e40af; padding: .1rem .3rem; border-radius: 3px; }
.warn-box {
    background: #fffbeb; border-left: 3px solid #f59e0b;
    border-radius: 0 8px 8px 0; padding: .8rem 1rem;
    font-size: .83rem; color: #78350f; font-weight: 500;
}
.success-box {
    background: #f0fdf4; border-left: 3px solid #22c55e;
    border-radius: 0 8px 8px 0; padding: .8rem 1rem;
    font-size: .83rem; color: #14532d; font-weight: 500;
}

/* ── Upload hint box (inside sidebar) ── */
.upload-hint {
    background: #1e293b; border: 1px solid #334155;
    border-radius: 8px; padding: .85rem 1rem;
    font-size: .78rem; line-height: 1.8;
    color: #e2e8f0;
    margin-top: .5rem;
}

/* ── CSV file pills ── */
.file-pill {
    display: inline-block; background: #1e40af; color: #fff !important;
    border-radius: 999px; padding: .2rem .7rem;
    font-size: .72rem; font-weight: 600; margin: .15rem .1rem;
}

/* ── Page headings ── */
h1 { color: #0f172a !important; font-weight: 800 !important; }
h2 { color: #1e293b !important; font-weight: 700 !important; }

/* ── Streamlit native metric widget ── */
[data-testid="stMetricLabel"] { color: #334155 !important; font-weight: 600 !important; }
[data-testid="stMetricValue"] { color: #0f172a !important; font-weight: 700 !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab"]                     { color: #334155 !important; font-weight: 600; }
.stTabs [data-baseweb="tab"][aria-selected="true"]{ color: #1d4ed8 !important; }
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════
RETURN_CHARGES = 45
MARKETING_PCT  = 0.03
ROYALTY_PCT    = 0.01
GST_RATE       = 0.18

REQUIRED_RATE_COLS = [
    "Brand Name", "Category",
    "Lower Limit Commision", "Upper Limit Commision", "Commision Charge",
    "GT Lower Limit", "GT Upper Limit", "GT Charges",
    "Lower Limit Fixed Fee", "Upper Limit Fixed Fee", "Fix Fee",
]


# ═════════════════════════════════════════════════════════════════════════════
#  FILE LOADERS
# ═════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def load_slab(file_bytes: bytes):
    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    ws   = wb["Rates"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        raise ValueError("'Rates' sheet is empty.")

    header = [c for c in rows[0] if c is not None and str(c).strip()]
    n_cols = len(header)
    data_rows = []
    for r in rows[1:]:
        r = list(r)[:n_cols] + [None] * max(0, n_cols - len(r))
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        data_rows.append(r[:n_cols])

    rates = pd.DataFrame(data_rows, columns=header)
    rates.columns = [str(c).strip() for c in rates.columns]

    missing = [c for c in REQUIRED_RATE_COLS if c not in rates.columns]
    if missing:
        raise ValueError(f"'Rates' sheet missing columns: {', '.join(missing)}")

    rates["Brand Name"] = rates["Brand Name"].astype(str).str.strip()
    rates["Category"]   = rates["Category"].astype(str).str.strip()
    for c in [c for c in REQUIRED_RATE_COLS if c not in ("Brand Name", "Category")]:
        rates[c] = pd.to_numeric(rates[c], errors="coerce")

    # Drop rows where all numeric slab limits are NaN (blank spacer rows)
    slab_num_cols = ["Lower Limit Commision", "Upper Limit Commision",
                     "GT Lower Limit", "GT Upper Limit",
                     "Lower Limit Fixed Fee", "Upper Limit Fixed Fee"]
    rates = rates.dropna(subset=slab_num_cols, how="all").reset_index(drop=True)

    oms_map, yrn_map = {}, {}
    if "Replace Sku" in wb.sheetnames:
        for row in wb["Replace Sku"].iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * max(0, 5 - len(row))
            if row[2] and row[4]:
                oms_map[str(row[2]).strip()] = str(row[4]).strip()
            if row[1] and row[4]:
                yrn_map[str(row[1]).strip()] = str(row[4]).strip()

    pwn_map = {}
    for sn in wb.sheetnames:
        if sn.strip().lower().startswith("price we need"):
            for row in wb[sn].iter_rows(min_row=2, values_only=True):
                row = list(row) + [None] * 3
                if row[1] and row[2] is not None:
                    pwn_map[str(row[1]).strip()] = row[2]
            break

    closed_map = {}
    if "Closed" in wb.sheetnames:
        for row in wb["Closed"].iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * 3
            if row[1] and row[2] is not None:
                closed_map[str(row[1]).strip()] = row[2]

    return rates, oms_map, yrn_map, pwn_map, closed_map


@st.cache_data(show_spinner=False)
def load_fix_rate(file_bytes: bytes) -> pd.DataFrame:
    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return pd.DataFrame()

    header = [str(c).strip() if c is not None else f"col_{i}"
              for i, c in enumerate(rows[0])]
    df = pd.DataFrame(rows[1:], columns=header)

    required = ["Brand Name", "Style Id", "Start Date", "End Date",
                "GT Charge", "Commission", "Fixed fee"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Fix Rate sheet missing columns: {', '.join(missing)}")

    df = df.dropna(how="all").reset_index(drop=True)
    df["Brand Name"] = df["Brand Name"].astype(str).str.strip()
    df["Style Id"]   = pd.to_numeric(df["Style Id"],   errors="coerce")
    df["GT Charge"]  = pd.to_numeric(df["GT Charge"],  errors="coerce").fillna(0)
    df["Commission"] = pd.to_numeric(df["Commission"], errors="coerce").fillna(0)
    df["Fixed fee"]  = pd.to_numeric(df["Fixed fee"],  errors="coerce").fillna(0)

    def _to_date(v):
        if isinstance(v, datetime): return v.date()
        if v is None or (isinstance(v, str) and not v.strip()): return None
        try: return pd.to_datetime(v).date()
        except: return None

    df["Start Date"] = df["Start Date"].apply(_to_date)
    df["End Date"]   = df["End Date"].apply(_to_date)

    # Only keep rows that have a valid Style Id
    df = df[df["Style Id"].notna()].reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False)
def load_csv_files(files_bytes: list, filenames: list) -> pd.DataFrame:
    """
    Load one or more CSV files and concatenate them.
    Adds a '_source_file' column so rows can be traced back to their file.
    """
    frames = []
    for fb, fn in zip(files_bytes, filenames):
        try:
            df = pd.read_csv(BytesIO(fb))
            df["_source_file"] = fn
            frames.append(df)
        except Exception as e:
            st.warning(f"Could not read '{fn}': {e}")
    if not frames:
        raise ValueError("No valid CSV files could be read.")
    combined = pd.concat(frames, ignore_index=True)
    return combined


# ═════════════════════════════════════════════════════════════════════════════
#  FIX RATE LOOKUP
# ═════════════════════════════════════════════════════════════════════════════

def get_fix_rate_row(fix_df, brand: str, style_id, order_date):
    """
    Return matching fix-rate Series or None.
    Match: Brand Name == brand  AND  Style Id == style_id
           AND  Start Date <= order_date <= End Date

    Falls back to undated rows if no dated window covers order_date.
    Returns None → caller uses normal Slab lookup.
    """
    if fix_df is None or fix_df.empty:
        return None
    try:
        sid = int(float(style_id))
    except Exception:
        return None

    sub = fix_df[
        (fix_df["Brand Name"] == brand) &
        (fix_df["Style Id"]   == sid)
    ]
    if sub.empty:
        return None

    if order_date is not None:
        dated = sub[
            sub["Start Date"].notna() & sub["End Date"].notna() &
            (sub["Start Date"] <= order_date) & (sub["End Date"] >= order_date)
        ]
        if not dated.empty:
            return dated.iloc[0]
        undated = sub[sub["Start Date"].isna() | sub["End Date"].isna()]
        if not undated.empty:
            return undated.iloc[0]
        # brand+style exists but no date window matches → fall through to Slab
        return None

    return sub.iloc[0]


# ═════════════════════════════════════════════════════════════════════════════
#  SLAB LOOKUP  (NaN-safe)
# ═════════════════════════════════════════════════════════════════════════════

def _lookup(subset: pd.DataFrame, lo_col: str, hi_col: str, val: float):
    """
    Find slab row where lo_col <= val < hi_col.
    Skips rows where either bound is NaN (blank/spacer rows in the sheet).
    Falls back to lo <= val <= hi for the top slab row.
    """
    valid = subset[subset[lo_col].notna() & subset[hi_col].notna()]
    m = valid[(valid[lo_col] <= val) & (valid[hi_col] > val)]
    if m.empty:
        # Top slab: exact match on upper bound
        m = valid[(valid[lo_col] <= val) & (valid[hi_col] >= val)]
    return m.iloc[0] if not m.empty else None


def get_charges_from_slab(rates: pd.DataFrame, brand: str, cat: str, SP: float):
    """
    Normal slab lookup (used when no Fix Rate match exists).
    Formula:
      1. GT        = slab lookup by SP
      2. V         = SP - GT
      3. comm_rate = slab lookup by V  (not SP!)
      4. comm_amt  = rate × V
      5. fixed_fee = slab lookup by SP
    Returns (GT, V, comm_rate, comm_amt, fixed_fee) or all-None on miss.
    """
    sub = rates[
        (rates["Brand Name"] == brand) &
        (rates["Category"].str.lower() == cat.lower())
    ]
    if sub.empty:
        return None, None, None, None, None

    gt_row  = _lookup(sub, "GT Lower Limit",       "GT Upper Limit",       SP)
    fee_row = _lookup(sub, "Lower Limit Fixed Fee", "Upper Limit Fixed Fee", SP)
    if gt_row is None or fee_row is None:
        return None, None, None, None, None

    GT        = float(gt_row["GT Charges"])
    fixed_fee = float(fee_row["Fix Fee"])
    V         = round(SP - GT, 2)

    comm_row = _lookup(sub, "Lower Limit Commision", "Upper Limit Commision", V)
    if comm_row is None:
        return None, None, None, None, None

    comm_rate = float(comm_row["Commision Charge"])
    comm_amt  = round(comm_rate * V, 2)
    return GT, V, comm_rate, comm_amt, fixed_fee


def get_pwn_price(oms_map, yrn_map, pwn_map, closed_map, seller_sku, myntra_sku=""):
    key     = str(seller_sku).strip()
    mkey    = str(myntra_sku).strip()
    oms_sku = oms_map.get(key, key)
    if oms_sku == key and mkey:
        oms_sku = yrn_map.get(mkey, oms_sku)
    if oms_sku in closed_map:
        return closed_map[oms_sku], oms_sku, "Closed"
    if oms_sku in pwn_map:
        return pwn_map[oms_sku], oms_sku, "PWN"
    return None, oms_sku, None


# ═════════════════════════════════════════════════════════════════════════════
#  CORE RECONCILIATION
# ═════════════════════════════════════════════════════════════════════════════

def reconcile(rates, df, oms_map, yrn_map, pwn_map, closed_map,
              fix_df=None, report_date=None, brand_filter=None):
    """
    Priority per order row:
      1. Fix Rate  (Brand + Style Id + date window)
             → GT, Commission (flat ₹), Fixed fee from Fix Rate sheet
             → Myntra Payable = SP - GT - Commission - Fixed fee  (GST embedded)
      2. Slab      (Brand + Category + SP range)
             → normal formula with separate GST
      3. No match  → flagged unmatched
    """
    df = df.copy()
    df["brand"]        = df["brand"].astype(str).str.strip()
    df["article type"] = df["article type"].astype(str).str.strip()

    if brand_filter:
        df = df[df["brand"].isin(brand_filter)].copy()

    records = []
    for _, row in df.iterrows():
        SP    = float(row["final amount"])
        brand = row["brand"]
        cat   = row["article type"]
        sid   = row.get("style id", None)

        # ── 1. Fix Rate ───────────────────────────────────────────────────
        fix_row = get_fix_rate_row(fix_df, brand, sid, report_date)

        if fix_row is not None:
            GT        = float(fix_row["GT Charge"])
            comm_amt  = float(fix_row["Commission"])   # flat ₹
            fixed_fee = float(fix_row["Fixed fee"])
            V         = round(SP - GT, 2)
            total_ch  = round(comm_amt + GT + fixed_fee, 2)
            myntra_pay= round(SP - total_ch, 2)        # GST already embedded
            gst       = 0.0
            marketing = round(SP * MARKETING_PCT, 2)
            royalty   = round(V  * ROYALTY_PCT, 2)
            rec = dict(
                _GT=GT, _V=V, _comm_rate=None, _comm_amt=comm_amt,
                _fixed_fee=fixed_fee, _total_charges=total_ch,
                _gst=gst, _myntra_payable=myntra_pay,
                _marketing=marketing, _royalty=royalty,
                _slab_ok=True, _rate_source="Fix Rate",
            )

        else:
            # ── 2. Slab ───────────────────────────────────────────────────
            GT, V, comm_rate, comm_amt, fixed_fee = get_charges_from_slab(
                rates, brand, cat, SP)

            if GT is None:
                rec = dict(
                    _GT=None, _V=None, _comm_rate=None, _comm_amt=None,
                    _fixed_fee=None, _total_charges=None, _gst=None,
                    _myntra_payable=None, _marketing=None, _royalty=None,
                    _slab_ok=False, _rate_source="No Match",
                )
            else:
                total_ch  = round(comm_amt + GT + fixed_fee, 2)
                gst       = round((total_ch - GT) * GST_RATE, 2)
                myntra_pay= round(SP - total_ch - gst, 2)
                marketing = round(SP * MARKETING_PCT, 2)
                royalty   = round(V  * ROYALTY_PCT, 2)
                rec = dict(
                    _GT=GT, _V=V, _comm_rate=comm_rate, _comm_amt=comm_amt,
                    _fixed_fee=fixed_fee, _total_charges=total_ch,
                    _gst=gst, _myntra_payable=myntra_pay,
                    _marketing=marketing, _royalty=royalty,
                    _slab_ok=True, _rate_source="Slab",
                )

        pwn_price, oms_sku, pwn_source = get_pwn_price(
            oms_map, yrn_map, pwn_map, closed_map,
            row.get("seller sku code", ""),
            row.get("myntra sku code", ""),
        )
        rec["_pwn"]        = pwn_price
        rec["_oms_sku"]    = oms_sku
        rec["_pwn_source"] = pwn_source
        records.append(rec)

    enrich = pd.DataFrame(records, index=df.index)
    return pd.concat([df.reset_index(drop=True), enrich.reset_index(drop=True)], axis=1)


# ═════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT
# ═════════════════════════════════════════════════════════════════════════════

HEADERS = [
    "order release id", "order line id", "STYLE ID", "Myntra SKU Code", "SKU",
    "Original SKU", "Article Type", "MRP", "PWN+10%",
    "Marketing Charges 3%", "Return Charges", "Royalty", "Total Amount",
    "Commission Amount", "GT Charges", "Fixed Fee",
    "Total Charges", "GST Amount", "Myntra Payble Amount",
    "Rebate", "Diffrence", "Selling Price", "Selling Price -GT Price",
    "Date", "Brand", "Order Status", "Rate Source",
]
COL_W = [18,15,12,24,24,22,14,8,10,18,14,10,13,17,11,10,13,11,18,8,11,13,18,13,12,12,12]

_HF = PatternFill("solid", fgColor="1F4E79")
_YF = PatternFill("solid", fgColor="FCE4D6")
_BF = PatternFill("solid", fgColor="DEEAF1")
_GF = PatternFill("solid", fgColor="E2EFDA")
_AF = PatternFill("solid", fgColor="FFF9E6")
_A1 = PatternFill("solid", fgColor="EBF3FA")
_A2 = PatternFill("solid", fgColor="FFFFFF")
_thin = Side(style="thin", color="B8CCE4")
_TB   = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _cell(ws, r, c, val, fill, font=None, fmt=None, align="center"):
    cell = ws.cell(row=r, column=c, value=val)
    cell.fill      = fill
    cell.border    = _TB
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if font: cell.font          = font
    if fmt:  cell.number_format = fmt
    return cell


def build_excel(result_df: pd.DataFrame, report_date_str: str = "") -> bytes:
    wb         = Workbook()
    all_brands = sorted(result_df["brand"].unique())

    for brand in all_brands:
        bdf   = result_df[result_df["brand"] == brand].reset_index(drop=True)
        ws    = wb.create_sheet(title=brand[:31])

        hfont = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
        nfont = Font(size=9, name="Calibri")
        ofont = Font(size=9, name="Calibri", italic=True, color="C55A11")
        gfont = Font(bold=True, size=9, name="Calibri", color="375623")
        ffont = Font(bold=True, size=9, name="Calibri", color="7B4F00")

        for ci, (h, w) in enumerate(zip(HEADERS, COL_W), 1):
            _cell(ws, 1, ci, h, _HF, hfont)
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 36

        for ri, (_, row) in enumerate(bdf.iterrows(), 2):
            is_fix = row.get("_rate_source") == "Fix Rate"
            alt    = _AF if is_fix else (_A1 if ri % 2 == 0 else _A2)
            SP     = row.get("final amount", 0)
            V      = row.get("_V")
            gst_v  = row.get("_gst")
            rfont  = ffont if is_fix else nfont

            vals = [
                row.get("order release id", ""),
                row.get("order line id", ""),
                row.get("style id", ""),
                row.get("myntra sku code", ""),
                row.get("seller sku code", ""),
                row.get("_oms_sku", ""),
                row.get("article type", ""),
                row.get("total mrp", ""),
                row.get("_pwn"),
                row.get("_marketing"),
                RETURN_CHARGES,
                row.get("_royalty"),
                None,                           # M  Total Amount (formula)
                row.get("_comm_amt"),
                row.get("_GT"),
                row.get("_fixed_fee"),
                row.get("_total_charges"),
                gst_v if gst_v else None,
                row.get("_myntra_payable"),
                0,
                None,                           # U  Difference (formula)
                SP,
                V,
                report_date_str or "—",
                row.get("brand", ""),
                row.get("order status", ""),
                row.get("_rate_source", ""),
            ]

            num_ci = {10,11,12,13,14,15,16,17,18,19,20,21,22,23}

            for ci, val in enumerate(vals, 1):
                if ci == 9:
                    _cell(ws, ri, ci, val,
                          alt if val is not None else _YF,
                          rfont if val is not None else ofont,
                          fmt="#,##0.00" if val is not None else None)
                elif ci == 19:
                    _cell(ws, ri, ci, val, _GF, gfont,
                          fmt="#,##0.00" if val is not None else None)
                elif ci == 21:
                    _cell(ws, ri, ci, None, _BF, nfont)
                elif ci == 13:
                    _cell(ws, ri, ci, None, alt, rfont)
                elif ci == 27:
                    _cell(ws, ri, ci, val,
                          _AF if is_fix else _A2,
                          ffont if is_fix else nfont)
                else:
                    fmt = "#,##0.00" if ci in num_ci else ("#,##0" if ci == 8 else None)
                    _cell(ws, ri, ci, val, alt, rfont, fmt)

            ws.cell(row=ri, column=13).value         = f'=IF(I{ri}="","",K{ri}+J{ri}+I{ri}+L{ri})'
            ws.cell(row=ri, column=13).number_format = "#,##0.00"
            ws.cell(row=ri, column=21).value         = f'=IF(I{ri}="","",S{ri}-M{ri}+T{ri})'
            ws.cell(row=ri, column=21).number_format = "#,##0.00"

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

        nr = len(bdf) + 3
        nc = ws.cell(row=nr, column=1,
            value="🟡 Amber = Fix Rate (GT/Comm/Fee from Fix Rate sheet, GST embedded).  "
                  "⬜ White/Blue = normal Slab lookup.  "
                  "🟠 Orange PWN cells = fill manually.")
        nc.font = Font(bold=True, italic=True, size=9, color="7B4F00", name="Calibri")
        nc.fill = PatternFill("solid", fgColor="FFF9E6")
        ws.merge_cells(f"A{nr}:{get_column_letter(len(HEADERS))}{nr}")

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws_s = wb.create_sheet(title="Summary", index=0)
    sh   = ["Brand","Orders","Fix Rate Orders","Selling Price",
            "GT Charges","Commission","Fixed Fee","Total Charges","GST","Myntra Payable"]
    sw   = [14,8,14,18,14,14,12,14,10,18]
    hf2  = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    for ci,(h,w) in enumerate(zip(sh,sw),1):
        _cell(ws_s, 1, ci, h, _HF, hf2)
        ws_s.column_dimensions[get_column_letter(ci)].width = w
    ws_s.row_dimensions[1].height = 30

    grand = {k:0.0 for k in ["sp","gt","cm","ff","tc","gst","pay","fix_n","n"]}

    for ri, brand in enumerate(all_brands, 2):
        bdf   = result_df[result_df["brand"] == brand]
        ok    = bdf[bdf["_slab_ok"] == True]
        fix_n = int((bdf["_rate_source"] == "Fix Rate").sum())
        n     = len(bdf)
        sp    = round(bdf["final amount"].sum(), 2)
        gt    = round(ok["_GT"].sum(), 2)
        cm    = round(ok["_comm_amt"].sum(), 2)
        ff    = round(ok["_fixed_fee"].sum(), 2)
        tc    = round(ok["_total_charges"].sum(), 2)
        gst   = round(ok["_gst"].sum(), 2)
        pay   = round(ok["_myntra_payable"].sum(), 2)

        alt = _A1 if ri % 2 == 0 else _A2
        nf  = Font(size=9, name="Calibri")
        bf  = Font(bold=True, size=9, name="Calibri")
        for ci, v in enumerate([brand,n,fix_n,sp,gt,cm,ff,tc,gst,pay],1):
            _cell(ws_s, ri, ci, v, alt, bf if ci==1 else nf,
                  fmt="#,##0.00" if ci > 3 else None)
        for k,v in zip(["n","sp","gt","cm","ff","tc","gst","pay","fix_n"],
                        [n,sp,gt,cm,ff,tc,gst,pay,fix_n]):
            grand[k] += v

    gr  = len(all_brands) + 2
    gf2 = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    gv  = ["GRAND TOTAL", grand["n"], grand["fix_n"],
           round(grand["sp"],2), round(grand["gt"],2), round(grand["cm"],2),
           round(grand["ff"],2), round(grand["tc"],2), round(grand["gst"],2),
           round(grand["pay"],2)]
    for ci,v in enumerate(gv,1):
        _cell(ws_s, gr, ci, v, _HF, gf2, fmt="#,##0.00" if ci > 3 else None)

    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🧾 Myntra Recon")
    st.markdown("---")

    st.markdown("**① Slab File**")
    slab_file = st.file_uploader(
        "Slab.xlsx (Rates + Replace Sku)",
        type=["xlsx"], key="slab",
    )

    st.markdown("**② Orders CSV** *(one or more)*")
    csv_files = st.file_uploader(
        "Seller Orders Report (.csv)",
        type=["csv"], key="csv",
        accept_multiple_files=True,
        help="Upload one or more Myntra Seller Orders CSV files — they will be merged automatically.",
    )

    st.markdown("**③ Fix Rate Sheet** *(optional)*")
    fix_file = st.file_uploader(
        "Fix_Rate_sheet.xlsx",
        type=["xlsx"], key="fix",
        help="Brand + Style ID fixed GT/Commission/Fee overrides with date windows.",
    )

    st.markdown("**Report Date**")
    report_date_input = st.date_input(
        "Order date (for Fix Rate date range filter)",
        value=datetime.today().date(),
        help="Used to match Fix Rate rows where Start Date ≤ date ≤ End Date.",
    )

    st.markdown("---")

    brand_filter = []
    if csv_files:
        try:
            frames = []
            for f in csv_files:
                frames.append(pd.read_csv(f))
                f.seek(0)
            _df_tmp = pd.concat(frames, ignore_index=True)
            _df_tmp["brand"] = _df_tmp["brand"].astype(str).str.strip()
            all_b = sorted(_df_tmp["brand"].unique().tolist())
            brand_filter = st.multiselect("Filter by Brand", options=all_b, default=all_b)
        except Exception:
            pass

    st.markdown("---")
    st.markdown("""
<div class="upload-hint">
<b>Priority per order:</b><br>
1. Fix Rate (Brand + Style ID + Date)<br>
2. Slab (Brand + Category + SP range)<br><br>
<b>Fix Rate formula:</b><br>
Payable = SP − GT − Comm − Fee<br>
(GST embedded, no separate deduction)<br><br>
<b>Slab formula:</b><br>
V = SP − GT<br>
Comm = rate(V) × V<br>
GST = (Comm + Fee) × 18%<br>
Payable = SP − TC − GST
</div>
""", unsafe_allow_html=True)


# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("## Myntra Seller Reconciliation")

if not slab_file or not csv_files:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""<div class="info-box"><b>Step 1 – Slab.xlsx</b><br>
Commission / GT / Fixed-fee rate table. Must have a <code>Rates</code> sheet with:<br>
Brand Name · Category · slab limits (commission, GT, fixed fee).</div>""",
            unsafe_allow_html=True)
    with c2:
        st.markdown("""<div class="info-box"><b>Step 2 – Orders CSV (one or more)</b><br>
Myntra Seller Orders export. Key columns needed:<br>
<code>brand</code> · <code>article type</code> · <code>final amount</code> · <code>style id</code><br>
Upload multiple CSVs — they will be merged automatically.</div>""",
            unsafe_allow_html=True)
    st.stop()

# ── Load slab ─────────────────────────────────────────────────────────────────
with st.spinner("Loading slab rates…"):
    try:
        rates, oms_map, yrn_map, pwn_map, closed_map = load_slab(slab_file.read())
    except Exception as e:
        st.error(f"❌ Slab file error: {e}"); st.stop()

# ── Load fix rate ─────────────────────────────────────────────────────────────
fix_df = None
if fix_file:
    with st.spinner("Loading fix rates…"):
        try:
            fix_df    = load_fix_rate(fix_file.read())
            n_combos  = fix_df[["Brand Name","Style Id"]].drop_duplicates().shape[0]
            st.sidebar.success(f"✅ Fix Rate — {len(fix_df)} rows · {n_combos} Brand+Style combos")
        except Exception as e:
            st.warning(f"Fix Rate file could not be loaded: {e}")

# ── Load CSV(s) ───────────────────────────────────────────────────────────────
with st.spinner("Reading orders…"):
    try:
        files_bytes = [f.read() for f in csv_files]
        filenames   = [f.name   for f in csv_files]
        orders_df   = load_csv_files(files_bytes, filenames)

        orders_df["brand"]        = orders_df["brand"].astype(str).str.strip()
        orders_df["article type"] = orders_df["article type"].astype(str).str.strip()
        orders_df["final amount"] = pd.to_numeric(orders_df["final amount"], errors="coerce").fillna(0)
        orders_df["seller price"] = pd.to_numeric(orders_df["seller price"], errors="coerce").fillna(0)
        orders_df["style id"]     = pd.to_numeric(orders_df["style id"],     errors="coerce")
    except Exception as e:
        st.error(f"❌ CSV error: {e}"); st.stop()

# Show which files were loaded
if len(csv_files) > 1:
    pills = " ".join(f'<span class="file-pill">{fn}</span>' for fn in filenames)
    st.markdown(
        f'<div class="info-box">📂 <b>{len(csv_files)} CSV files merged</b> '
        f'— {len(orders_df):,} total rows<br>{pills}</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

# ── Reconcile ─────────────────────────────────────────────────────────────────
with st.spinner("Reconciling…"):
    result = reconcile(
        rates, orders_df, oms_map, yrn_map, pwn_map, closed_map,
        fix_df=fix_df,
        report_date=report_date_input,
        brand_filter=brand_filter if brand_filter else None,
    )

ok_df    = result[result["_slab_ok"] == True]
bad_df   = result[result["_slab_ok"] == False]
fix_rows = int((result["_rate_source"] == "Fix Rate").sum())

# ── KPIs ──────────────────────────────────────────────────────────────────────
total_orders  = len(result)
total_sp      = result["final amount"].sum()
total_payable = ok_df["_myntra_payable"].sum()
total_tc_gst  = ok_df["_total_charges"].sum() + ok_df["_gst"].sum()
unmatched     = len(bad_df)

cols = st.columns(6)
kpis = [
    ("Total Orders",         f"{total_orders:,}",           None,
     f"{len(csv_files)} file(s) merged"),
    ("Total Selling Price",  f"₹{total_sp:,.0f}",          None,        "sum of final amount"),
    ("Total Myntra Payable", f"₹{total_payable:,.0f}",     None,        "after all deductions"),
    ("Total Charges + GST",  f"₹{total_tc_gst:,.0f}",     None,        "platform fees"),
    ("Fix Rate Applied",     f"{fix_rows}/{total_orders}",
     "positive" if fix_rows > 0 else None,                              "style-level overrides"),
    ("Unmatched Rows",       f"{unmatched}",
     "negative" if unmatched > 0 else "positive",                       "no slab / fix rate found"),
]
for col, (label, value, cls, sub) in zip(cols, kpis):
    with col:
        st.markdown(f"""<div class="metric-card">
  <div class="metric-label">{label}</div>
  <div class="metric-value {cls or ''}">{value}</div>
  <div class="metric-sub">{sub}</div>
</div>""", unsafe_allow_html=True)

if fix_rows > 0:
    st.markdown(f"""<br><div class="warn-box">
🔒 <b>{fix_rows} orders</b> used <b>Fix Rate</b> (Brand + Style ID + Date match).
GT / Commission / Fixed Fee come from the Fix Rate sheet — GST is already embedded,
no separate GST is deducted. All other orders use normal Slab lookup.
</div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Brand summary ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">📊 Brand Summary</div>', unsafe_allow_html=True)

summary_rows = []
for brand in sorted(result["brand"].unique()):
    bdf  = result[result["brand"] == brand]
    bok  = bdf[bdf["_slab_ok"] == True]
    fx_n = int((bdf["_rate_source"] == "Fix Rate").sum())
    summary_rows.append({
        "Brand":               brand,
        "Orders":              len(bdf),
        "Fix Rate Orders":     fx_n,
        "Selling Price (₹)":   round(bdf["final amount"].sum(), 2),
        "GT Charges (₹)":      round(bok["_GT"].sum(), 2),
        "Commission (₹)":      round(bok["_comm_amt"].sum(), 2),
        "Fixed Fee (₹)":       round(bok["_fixed_fee"].sum(), 2),
        "Total Charges (₹)":   round(bok["_total_charges"].sum(), 2),
        "GST (₹)":             round(bok["_gst"].sum(), 2),
        "Myntra Payable (₹)":  round(bok["_myntra_payable"].sum(), 2),
    })

sum_df = pd.DataFrame(summary_rows)
st.dataframe(
    sum_df.style
        .format({c: "{:,.2f}" for c in sum_df.columns if "₹" in c}, na_rep="")
        .set_properties(**{"text-align": "center"}),
    use_container_width=True, hide_index=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# ── Order-level tabs ──────────────────────────────────────────────────────────
brands_in_result = sorted(result["brand"].unique().tolist())

if brands_in_result:
    st.markdown('<div class="section-header">📋 Order-Level Detail</div>', unsafe_allow_html=True)
    tabs = st.tabs([f"  {b}  " for b in brands_in_result] + ["⚠️ Unmatched"])

    DISPLAY_COLS = {
        "order release id": "Order ID",
        "order line id":    "Line ID",
        "style id":         "Style ID",
        "seller sku code":  "SKU",
        "article type":     "Category",
        "order status":     "Status",
        "_source_file":     "Source File",
        "_rate_source":     "Rate Source",
        "final amount":     "Selling Price ₹",
        "_V":               "SP−GT ₹",
        "_comm_amt":        "Commission ₹",
        "_GT":              "GT Charges ₹",
        "_fixed_fee":       "Fixed Fee ₹",
        "_total_charges":   "Total Charges ₹",
        "_gst":             "GST ₹",
        "_myntra_payable":  "Myntra Payable ₹",
        "_pwn":             "PWN+10% ₹",
        "_marketing":       "Marketing 3% ₹",
        "_royalty":         "Royalty 1% ₹",
    }
    NUM_KEYS = {"final amount","_V","_comm_amt","_GT","_fixed_fee",
                "_total_charges","_gst","_myntra_payable","_pwn","_marketing","_royalty"}

    for tab, brand in zip(tabs[:-1], brands_in_result):
        with tab:
            bdf = result[result["brand"] == brand].copy()
            ok  = bdf[bdf["_slab_ok"] == True]
            mis = bdf[bdf["_slab_ok"] == False]
            fx  = int((bdf["_rate_source"] == "Fix Rate").sum())

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Orders",          len(bdf))
            c2.metric("Selling Price",   f"₹{bdf['final amount'].sum():,.0f}")
            c3.metric("Myntra Payable",  f"₹{ok['_myntra_payable'].sum():,.0f}")
            c4.metric("Fix Rate Rows",   fx)
            c5.metric("Unmatched",       len(mis),
                      delta=str(len(mis)) if len(mis) else None,
                      delta_color="inverse")

            # Only show Source File column when multiple files loaded
            disp_cols = {k: v for k, v in DISPLAY_COLS.items()
                         if k != "_source_file" or len(csv_files) > 1}

            display_df = bdf[[c for c in disp_cols if c in bdf.columns]].rename(columns=disp_cols)
            num_dcols  = [disp_cols[k] for k in NUM_KEYS if k in disp_cols and k in bdf.columns]
            for c in num_dcols:
                if c in display_df.columns:
                    display_df[c] = pd.to_numeric(display_df[c], errors="coerce")

            def _highlight(row):
                if row.get("Rate Source") == "Fix Rate":
                    return ["background-color:#fffbeb"] * len(row)
                return [""] * len(row)

            st.dataframe(
                display_df.style
                    .format({c: "{:,.2f}" for c in num_dcols if c in display_df.columns},
                            na_rep="")
                    .apply(_highlight, axis=1),
                use_container_width=True,
                hide_index=True,
                height=min(420, 45 + 35 * len(display_df)),
            )

            if len(mis) > 0:
                st.markdown(f"""<div class="warn-box">
⚠️ {len(mis)} rows had no matching slab or fix rate.
Check brand name spelling and that the category exists in the Slab file.</div>""",
                    unsafe_allow_html=True)
                st.dataframe(
                    mis[["seller sku code","style id","article type","final amount","order status"]],
                    use_container_width=True, hide_index=True,
                )

    with tabs[-1]:
        if bad_df.empty:
            st.markdown('<div class="success-box">✅ All rows matched — no unmatched orders.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="warn-box">⚠️ {len(bad_df)} unmatched orders.</div>',
                        unsafe_allow_html=True)
            st.dataframe(
                bad_df[["brand","seller sku code","style id","article type",
                        "final amount","order status"]],
                use_container_width=True, hide_index=True,
            )

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">⬇️ Export</div>', unsafe_allow_html=True)

col_dl, col_info = st.columns([1, 3])
with col_dl:
    with st.spinner("Preparing Excel…"):
        xlsx_bytes = build_excel(result, report_date_str=str(report_date_input))
    st.download_button(
        label="📥 Download Reconciliation (.xlsx)",
        data=xlsx_bytes,
        file_name="Myntra_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_info:
    st.markdown("""<div class="info-box">
One sheet per brand + <b>Summary</b> sheet.<br>
🟡 <b>Amber rows</b> = Fix Rate applied (style-level override, GST embedded).<br>
⬜ White rows = normal Slab lookup (GST calculated separately).<br>
🟠 <b>Orange PWN cells</b> = SKU not resolved — fill manually.<br>
PWN+10% auto-filled via Replace Sku → Price We Need / Closed lookup.
</div>""", unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)
