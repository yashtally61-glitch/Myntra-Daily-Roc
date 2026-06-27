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

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

section[data-testid="stSidebar"] {
    background: #0f172a !important;
    border-right: 1px solid #1e293b;
}
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
section[data-testid="stSidebar"] .upload-hint {
    color: #e2e8f0 !important;
}
section[data-testid="stSidebar"] .upload-hint b { color: #7dd3fc !important; }
section[data-testid="stSidebar"] [data-baseweb="tag"] { background: #1e40af !important; }
section[data-testid="stSidebar"] [data-baseweb="tag"] span { color: #ffffff !important; }

.main .block-container { padding-top: 1.5rem; max-width: 1400px; }

.metric-card {
    background: #ffffff; border: 1px solid #e2e8f0;
    border-radius: 12px; padding: 1.1rem 1.25rem;
    box-shadow: 0 2px 6px rgba(0,0,0,.07);
}
.metric-label { font-size:.72rem; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:.07em; margin-bottom:.3rem; }
.metric-value { font-size:1.5rem; font-weight:700; color:#0f172a; font-family:'JetBrains Mono',monospace; }
.metric-value.positive { color:#15803d; }
.metric-value.negative { color:#b91c1c; }
.metric-sub { font-size:.72rem; color:#64748b; margin-top:.2rem; font-weight:500; }

.section-header {
    display:flex; align-items:center; gap:.6rem;
    font-size:1rem; font-weight:700; color:#0f172a;
    border-bottom:2px solid #e2e8f0; padding-bottom:.5rem; margin-bottom:1rem;
}

.info-box { background:#eff6ff; border-left:3px solid #3b82f6; border-radius:0 8px 8px 0; padding:.8rem 1rem; font-size:.83rem; color:#1e3a5f; font-weight:500; }
.info-box b { color:#1d4ed8; }
.info-box code { background:#dbeafe; color:#1e40af; padding:.1rem .3rem; border-radius:3px; }
.warn-box { background:#fffbeb; border-left:3px solid #f59e0b; border-radius:0 8px 8px 0; padding:.8rem 1rem; font-size:.83rem; color:#78350f; font-weight:500; }
.success-box { background:#f0fdf4; border-left:3px solid #22c55e; border-radius:0 8px 8px 0; padding:.8rem 1rem; font-size:.83rem; color:#14532d; font-weight:500; }

.upload-hint { background:#1e293b; border:1px solid #334155; border-radius:8px; padding:.85rem 1rem; font-size:.78rem; line-height:1.8; color:#e2e8f0; margin-top:.5rem; }

.file-pill { display:inline-block; background:#1e40af; color:#fff !important; border-radius:999px; padding:.2rem .7rem; font-size:.72rem; font-weight:600; margin:.15rem .1rem; }

/* Brand charges config table */
.charges-table { width:100%; border-collapse:collapse; font-size:.83rem; }
.charges-table th { background:#1e3a5f; color:#fff; padding:.5rem .75rem; text-align:center; font-weight:700; }
.charges-table td { padding:.35rem .75rem; text-align:center; border-bottom:1px solid #e2e8f0; }
.charges-table tr:nth-child(even) td { background:#f8fafc; }

h1 { color:#0f172a !important; font-weight:800 !important; }
h2 { color:#1e293b !important; font-weight:700 !important; }
[data-testid="stMetricLabel"] { color:#334155 !important; font-weight:600 !important; }
[data-testid="stMetricValue"] { color:#0f172a !important; font-weight:700 !important; }
.stTabs [data-baseweb="tab"] { color:#334155 !important; font-weight:600; }
.stTabs [data-baseweb="tab"][aria-selected="true"] { color:#1d4ed8 !important; }
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
#  DEFAULTS
# ═════════════════════════════════════════════════════════════════════════════
DEFAULT_MARKETING_TYPE  = "%"
DEFAULT_MARKETING_VALUE = 3.0       # 3% of SP
DEFAULT_RETURN_CHARGES  = 45.0      # flat ₹
DEFAULT_ROYALTY_TYPE    = "%"
DEFAULT_ROYALTY_VALUE   = 1.0       # 1% of V (SP-GT) — or flat ₹ if type="Flat ₹"
GST_RATE                = 0.18

REQUIRED_RATE_COLS = [
    "Brand Name", "Category",
    "Lower Limit Commision", "Upper Limit Commision", "Commision Charge",
    "GT Lower Limit", "GT Upper Limit", "GT Charges",
    "Lower Limit Fixed Fee", "Upper Limit Fixed Fee", "Fix Fee",
]


# ═════════════════════════════════════════════════════════════════════════════
#  BRAND CHARGES HELPER
# ═════════════════════════════════════════════════════════════════════════════

def compute_brand_charges(brand: str, SP: float, V: float, brand_cfg: dict) -> tuple:
    """
    Returns (marketing, return_ch, royalty) for a given brand using
    the per-brand config dict. Falls back to defaults if brand not configured.

    brand_cfg structure:
      { "IKRASS": { "mkt_type": "%", "mkt_val": 3.0,
                    "return":   45.0,
                    "roy_type": "%", "roy_val": 1.0 }, ... }

    Marketing:
      type="%"      → marketing = SP × (val/100)
      type="Flat ₹" → marketing = val (fixed per order)

    Return Charges: always flat ₹

    Royalty:
      type="%"      → royalty = V × (val/100)   (V = SP − GT)
      type="Flat ₹" → royalty = val (fixed per order)
    """
    cfg = brand_cfg.get(brand, {})
    mkt_type = cfg.get("mkt_type",  DEFAULT_MARKETING_TYPE)
    mkt_val  = cfg.get("mkt_val",   DEFAULT_MARKETING_VALUE)
    ret_ch   = cfg.get("return",    DEFAULT_RETURN_CHARGES)
    roy_type = cfg.get("roy_type",  DEFAULT_ROYALTY_TYPE)
    roy_val  = cfg.get("roy_val",   cfg.get("roy_pct", DEFAULT_ROYALTY_VALUE))

    if mkt_type == "%":
        marketing = round(SP * (mkt_val / 100), 2)
    else:
        marketing = round(float(mkt_val), 2)

    return_ch = round(float(ret_ch), 2)

    if roy_type == "%":
        royalty = round(V * (roy_val / 100), 2)
    else:
        royalty = round(float(roy_val), 2)

    return marketing, return_ch, royalty


# ═════════════════════════════════════════════════════════════════════════════
#  SLAB VALIDATION  (NEW)
# ═════════════════════════════════════════════════════════════════════════════

def validate_slab_ranges(rates: pd.DataFrame) -> list:
    """
    Scans every (Brand Name, Category) group for the three pair-columns
    (Commission, GT, Fixed Fee) and flags:
      - INVALID  : a row where Lower Limit > Upper Limit
      - GAP      : a value range between consecutive brackets that no row covers
      - CONFLICT : two overlapping brackets that disagree on the charge value
                   (a silent wrong-answer risk — the app picks whichever row
                   happens to come first)
    Returns a list of dicts: {brand, category, kind, column, detail}
    Harmless overlaps (identical charge value on both sides) are NOT reported.
    """
    issues = []
    pairs = [
        ("Lower Limit Commision", "Upper Limit Commision", "Commision Charge", "Commission"),
        ("GT Lower Limit",        "GT Upper Limit",        "GT Charges",       "GT"),
        ("Lower Limit Fixed Fee", "Upper Limit Fixed Fee", "Fix Fee",          "Fixed Fee"),
    ]
    for (brand, cat), g in rates.groupby(["Brand Name", "Category"]):
        for lo_col, hi_col, val_col, label in pairs:
            gg = g.dropna(subset=[lo_col, hi_col]).sort_values(lo_col).reset_index()
            if gg.empty:
                continue
            for i in range(len(gg)):
                cur = gg.iloc[i]
                if cur[hi_col] < cur[lo_col]:
                    issues.append(dict(
                        brand=brand, category=cat, kind="INVALID", column=label,
                        detail=f"row has {lo_col}={cur[lo_col]} > {hi_col}={cur[hi_col]}",
                    ))
                if i + 1 < len(gg):
                    nxt = gg.iloc[i + 1]
                    if nxt[lo_col] > cur[hi_col]:
                        issues.append(dict(
                            brand=brand, category=cat, kind="GAP", column=label,
                            detail=f"no bracket covers {cur[hi_col]} – {nxt[lo_col]}",
                        ))
                    elif nxt[lo_col] < cur[hi_col]:
                        v1, v2 = cur[val_col], nxt[val_col]
                        if v1 != v2:
                            issues.append(dict(
                                brand=brand, category=cat, kind="CONFLICT", column=label,
                                detail=(f"[{cur[lo_col]}-{cur[hi_col]}) and "
                                        f"[{nxt[lo_col]}-{nxt[hi_col]}) overlap with "
                                        f"different values ({v1} vs {v2}) — wrong "
                                        f"amount may be picked"),
                            ))
    return issues


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
    for c in [c for c in REQUIRED_RATE_COLS if c not in ("Brand Name","Category")]:
        rates[c] = pd.to_numeric(rates[c], errors="coerce")

    slab_num_cols = ["Lower Limit Commision","Upper Limit Commision",
                     "GT Lower Limit","GT Upper Limit",
                     "Lower Limit Fixed Fee","Upper Limit Fixed Fee"]
    rates = rates.dropna(subset=slab_num_cols, how="all").reset_index(drop=True)

    # NEW: validate slab ranges and surface any data problems to the user.
    slab_issues = validate_slab_ranges(rates)

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

    return rates, oms_map, yrn_map, pwn_map, closed_map, slab_issues


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

    required = ["Brand Name","Style Id","Start Date","End Date",
                "GT Charge","Commission","Fixed fee"]
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
    df = df[df["Style Id"].notna()].reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False)
def load_csv_files(files_bytes: list, filenames: list) -> pd.DataFrame:
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
    return pd.concat(frames, ignore_index=True)


# ═════════════════════════════════════════════════════════════════════════════
#  FIX RATE LOOKUP
# ═════════════════════════════════════════════════════════════════════════════

def get_fix_rate_row(fix_df, brand, style_id, order_date):
    if fix_df is None or fix_df.empty:
        return None
    try:
        sid = int(float(style_id))
    except Exception:
        return None
    sub = fix_df[(fix_df["Brand Name"]==brand) & (fix_df["Style Id"]==sid)]
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
        return None
    return sub.iloc[0]


# ═════════════════════════════════════════════════════════════════════════════
#  SLAB LOOKUP
# ═════════════════════════════════════════════════════════════════════════════

def _lookup(subset, lo_col, hi_col, val):
    valid = subset[subset[lo_col].notna() & subset[hi_col].notna()]
    m = valid[(valid[lo_col] <= val) & (valid[hi_col] > val)]
    if m.empty:
        m = valid[(valid[lo_col] <= val) & (valid[hi_col] >= val)]
    return m.iloc[0] if not m.empty else None


def get_charges_from_slab(rates, brand, cat, SP):
    try:
        brand_str = str(brand).strip()
        cat_str   = str(cat).strip().lower()
    except Exception:
        return None, None, None, None, None
    if not brand_str or brand_str.lower() in ("nan","none",""):
        return None, None, None, None, None
    if not cat_str or cat_str in ("nan","none",""):
        return None, None, None, None, None
    try:
        SP = float(SP)
    except Exception:
        return None, None, None, None, None

    sub = rates[
        (rates["Brand Name"] == brand_str) &
        (rates["Category"].str.lower() == cat_str)
    ]
    if sub.empty:
        return None, None, None, None, None

    gt_row = _lookup(sub, "GT Lower Limit", "GT Upper Limit", SP)
    if gt_row is None:
        return None, None, None, None, None

    GT = float(gt_row["GT Charges"])
    V  = round(SP - GT, 2)

    fee_limits_present = (
        sub["Lower Limit Fixed Fee"].notna().any() and
        sub["Upper Limit Fixed Fee"].notna().any()
    )
    if fee_limits_present:
        fee_row = _lookup(sub, "Lower Limit Fixed Fee", "Upper Limit Fixed Fee", V)
        if fee_row is None:
            return None, None, None, None, None
        fixed_fee = float(fee_row["Fix Fee"])
    else:
        import math
        raw_fee = gt_row.get("Fix Fee", None)
        if raw_fee is None or (isinstance(raw_fee, float) and math.isnan(raw_fee)):
            fee_vals = sub["Fix Fee"].dropna()
            fixed_fee = float(fee_vals.iloc[0]) if not fee_vals.empty else 0.0
        else:
            fixed_fee = float(raw_fee)

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
              fix_df=None, report_date=None, brand_filter=None, brand_cfg=None):
    df = df.copy()
    df["brand"]        = df["brand"].astype(str).str.strip()
    df["article type"] = df["article type"].astype(str).str.strip()
    if brand_filter:
        df = df[df["brand"].isin(brand_filter)].copy()
    if brand_cfg is None:
        brand_cfg = {}

    records = []
    for _, row in df.iterrows():
        try:
            SP = float(row["final amount"])
        except Exception:
            SP = 0.0
        brand = str(row.get("brand","") or "").strip()
        cat   = str(row.get("article type","") or "").strip()
        sid   = row.get("style id", None)

        # ── 1. Fix Rate ───────────────────────────────────────────────────
        fix_row = get_fix_rate_row(fix_df, brand, sid, report_date)

        if fix_row is not None:
            GT        = float(fix_row["GT Charge"])
            comm_amt  = float(fix_row["Commission"])
            fixed_fee = float(fix_row["Fixed fee"])
            V         = round(SP - GT, 2)
            total_ch  = round(comm_amt + GT + fixed_fee, 2)
            myntra_pay= round(SP - total_ch, 2)
            gst       = 0.0
            marketing, return_ch, royalty = compute_brand_charges(brand, SP, V, brand_cfg)
            rec = dict(
                _GT=GT, _V=V, _comm_rate=None, _comm_amt=comm_amt,
                _fixed_fee=fixed_fee, _total_charges=total_ch,
                _gst=gst, _myntra_payable=myntra_pay,
                _marketing=marketing, _return_ch=return_ch, _royalty=royalty,
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
                    _myntra_payable=None, _marketing=None,
                    _return_ch=None, _royalty=None,
                    _slab_ok=False, _rate_source="No Match",
                )
            else:
                total_ch   = round(comm_amt + GT + fixed_fee, 2)
                gst        = round((total_ch - GT) * GST_RATE, 2)
                myntra_pay = round(SP - total_ch - gst, 2)
                marketing, return_ch, royalty = compute_brand_charges(brand, SP, V, brand_cfg)
                rec = dict(
                    _GT=GT, _V=V, _comm_rate=comm_rate, _comm_amt=comm_amt,
                    _fixed_fee=fixed_fee, _total_charges=total_ch,
                    _gst=gst, _myntra_payable=myntra_pay,
                    _marketing=marketing, _return_ch=return_ch, _royalty=royalty,
                    _slab_ok=True, _rate_source="Slab",
                )

        pwn_price, oms_sku, pwn_source = get_pwn_price(
            oms_map, yrn_map, pwn_map, closed_map,
            row.get("seller sku code",""),
            row.get("myntra sku code",""),
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
    "Marketing Charges", "Return Charges", "Royalty", "Total Amount",
    "Commission Amount", "GT Charges", "Fixed Fee",
    "Total Charges", "GST Amount", "Myntra Payble Amount",
    "Rebate", "Diffrence", "Selling Price", "Selling Price -GT Price",
    "Date", "Brand", "Order Status", "Rate Source",
]
COL_W = [18,15,12,24,24,22,14,8,10,16,14,10,13,17,11,10,13,11,18,8,11,13,18,13,12,12,12]

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
    cell.fill = fill; cell.border = _TB
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if font: cell.font = font
    if fmt:  cell.number_format = fmt
    return cell


def build_excel(result_df: pd.DataFrame, report_date_str: str = "",
                brand_cfg: dict = None) -> bytes:
    if brand_cfg is None:
        brand_cfg = {}
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

        # Build dynamic header labels showing actual rates for this brand
        cfg = brand_cfg.get(brand, {})
        mkt_type = cfg.get("mkt_type", DEFAULT_MARKETING_TYPE)
        mkt_val  = cfg.get("mkt_val",  DEFAULT_MARKETING_VALUE)
        roy_type = cfg.get("roy_type", DEFAULT_ROYALTY_TYPE)
        roy_val  = cfg.get("roy_val",  cfg.get("roy_pct", DEFAULT_ROYALTY_VALUE))
        ret_ch   = cfg.get("return",   DEFAULT_RETURN_CHARGES)

        mkt_label = f"Marketing ({mkt_val}{'%' if mkt_type=='%' else '₹ flat'})"
        roy_label = f"Royalty ({roy_val}{'%' if roy_type=='%' else '₹ flat'})"
        ret_label = f"Return Charges (₹{ret_ch:.0f})"

        dynamic_headers = list(HEADERS)
        dynamic_headers[9]  = mkt_label
        dynamic_headers[10] = ret_label
        dynamic_headers[11] = roy_label

        for ci, (h, w) in enumerate(zip(dynamic_headers, COL_W), 1):
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
                row.get("order release id",""),
                row.get("order line id",""),
                row.get("style id",""),
                row.get("myntra sku code",""),
                row.get("seller sku code",""),
                row.get("_oms_sku",""),
                row.get("article type",""),
                row.get("total mrp",""),
                row.get("_pwn"),               # I  PWN+10%
                row.get("_marketing"),          # J  Marketing
                row.get("_return_ch"),          # K  Return Charges
                row.get("_royalty"),            # L  Royalty
                None,                           # M  Total Amount (formula)
                row.get("_comm_amt"),           # N  Commission
                row.get("_GT"),                 # O  GT Charges
                row.get("_fixed_fee"),          # P  Fixed Fee
                row.get("_total_charges"),      # Q  Total Charges
                gst_v if gst_v else None,       # R  GST
                row.get("_myntra_payable"),     # S  Myntra Payable
                0,                              # T  Rebate
                None,                           # U  Difference (formula)
                SP,                             # V  Selling Price
                V,                              # W  SP-GT
                report_date_str or "—",         # X  Date
                row.get("brand",""),            # Y  Brand
                row.get("order status",""),     # Z  Order Status
                row.get("_rate_source",""),     # AA Rate Source
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
                    _cell(ws, ri, ci, val, _AF if is_fix else _A2,
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
                  "⬜ White/Blue = normal Slab.  🟠 Orange PWN = fill manually.")
        nc.font = Font(bold=True, italic=True, size=9, color="7B4F00", name="Calibri")
        nc.fill = PatternFill("solid", fgColor="FFF9E6")
        ws.merge_cells(f"A{nr}:{get_column_letter(len(HEADERS))}{nr}")

    # ── Summary ───────────────────────────────────────────────────────────────
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
        n=len(bdf); sp=round(bdf["final amount"].sum(),2)
        gt=round(ok["_GT"].sum(),2); cm=round(ok["_comm_amt"].sum(),2)
        ff=round(ok["_fixed_fee"].sum(),2); tc=round(ok["_total_charges"].sum(),2)
        gst=round(ok["_gst"].sum(),2); pay=round(ok["_myntra_payable"].sum(),2)
        alt = _A1 if ri%2==0 else _A2
        nf=Font(size=9,name="Calibri"); bf=Font(bold=True,size=9,name="Calibri")
        for ci,v in enumerate([brand,n,fix_n,sp,gt,cm,ff,tc,gst,pay],1):
            _cell(ws_s, ri, ci, v, alt, bf if ci==1 else nf,
                  fmt="#,##0.00" if ci>3 else None)
        for k,v in zip(["n","sp","gt","cm","ff","tc","gst","pay","fix_n"],
                        [n,sp,gt,cm,ff,tc,gst,pay,fix_n]):
            grand[k] += v

    gr=len(all_brands)+2
    gf2=Font(bold=True,color="FFFFFF",size=10,name="Calibri")
    gv=["GRAND TOTAL",grand["n"],grand["fix_n"],
        round(grand["sp"],2),round(grand["gt"],2),round(grand["cm"],2),
        round(grand["ff"],2),round(grand["tc"],2),round(grand["gst"],2),
        round(grand["pay"],2)]
    for ci,v in enumerate(gv,1):
        _cell(ws_s, gr, ci, v, _HF, gf2, fmt="#,##0.00" if ci>3 else None)

    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    buf = BytesIO(); wb.save(buf)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🧾 Myntra Recon")
    st.markdown("---")

    st.markdown("**① Slab File**")
    slab_file = st.file_uploader("Slab.xlsx (Rates + Replace Sku)",
                                  type=["xlsx"], key="slab")

    st.markdown("**② Orders CSV** *(one or more)*")
    csv_files = st.file_uploader("Seller Orders Report (.csv)",
                                  type=["csv"], key="csv",
                                  accept_multiple_files=True)

    st.markdown("**③ Fix Rate Sheet** *(optional)*")
    fix_file = st.file_uploader("Fix_Rate_sheet.xlsx",
                                 type=["xlsx"], key="fix")

    st.markdown("**Report Date**")
    report_date_input = st.date_input(
        "Order date (for Fix Rate date range filter)",
        value=datetime.today().date())

    st.markdown("---")
    brand_filter = []
    if csv_files:
        try:
            frames = []
            for f in csv_files:
                frames.append(pd.read_csv(f)); f.seek(0)
            _df_tmp = pd.concat(frames, ignore_index=True)
            _df_tmp["brand"] = _df_tmp["brand"].astype(str).str.strip()
            all_b = sorted(_df_tmp["brand"].unique().tolist())
            brand_filter = st.multiselect("Filter by Brand", options=all_b, default=all_b)
        except Exception:
            pass

    st.markdown("---")
    st.markdown("""<div class="upload-hint">
<b>Priority per order:</b><br>
1. Fix Rate (Brand + Style ID + Date)<br>
2. Slab (Brand + Category + SP range)<br><br>
<b>Fix Rate:</b> Payable = SP − GT − Comm − Fee (GST embedded)<br>
<b>Slab:</b> V=SP−GT · Comm=rate(V)×V · GST=(Comm+Fee)×18%
</div>""", unsafe_allow_html=True)


# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("## Myntra Seller Reconciliation")

if not slab_file or not csv_files:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""<div class="info-box"><b>Step 1 – Slab.xlsx</b><br>
Rate table with a <code>Rates</code> sheet — Brand Name, Category, slab limits.</div>""",
            unsafe_allow_html=True)
    with c2:
        st.markdown("""<div class="info-box"><b>Step 2 – Orders CSV (one or more)</b><br>
Myntra Seller Orders export — <code>brand</code>, <code>article type</code>,
<code>final amount</code>, <code>style id</code>. Multiple files are merged.</div>""",
            unsafe_allow_html=True)
    st.stop()

# ── Load files ────────────────────────────────────────────────────────────────
with st.spinner("Loading slab rates…"):
    try:
        rates, oms_map, yrn_map, pwn_map, closed_map, slab_issues = load_slab(slab_file.read())
    except Exception as e:
        st.error(f"❌ Slab error: {e}"); st.stop()

# NEW: surface slab data-quality problems right away, before reconciling.
if slab_issues:
    n_conflict = sum(1 for i in slab_issues if i["kind"] == "CONFLICT")
    n_gap      = sum(1 for i in slab_issues if i["kind"] in ("GAP", "INVALID"))
    with st.expander(
        f"⚠️ Slab data issues found — {len(slab_issues)} "
        f"({n_conflict} conflicting value{'s' if n_conflict!=1 else ''}, "
        f"{n_gap} gap/invalid range{'s' if n_gap!=1 else ''}) — click to review",
        expanded=True,
    ):
        st.markdown("""<div class="warn-box">
These problems live in the <b>Rates</b> sheet of your Slab.xlsx, not in this app.
<b>CONFLICT</b> = two overlapping brackets disagree on the charge — the app silently
picks whichever one appears first, which can give the wrong Fixed Fee / Commission / GT.
<b>GAP / INVALID</b> = a price range with no matching bracket at all — those orders will
show up as "Unmatched". Fix the Rates sheet and re-upload to clear these.</div>""",
            unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(slab_issues), use_container_width=True, hide_index=True)

fix_df = None
if fix_file:
    with st.spinner("Loading fix rates…"):
        try:
            fix_df   = load_fix_rate(fix_file.read())
            n_combos = fix_df[["Brand Name","Style Id"]].drop_duplicates().shape[0]
            st.sidebar.success(f"✅ Fix Rate — {len(fix_df)} rows · {n_combos} combos")
        except Exception as e:
            st.warning(f"Fix Rate not loaded: {e}")

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

if len(csv_files) > 1:
    pills = " ".join(f'<span class="file-pill">{fn}</span>' for fn in filenames)
    st.markdown(f'<div class="info-box">📂 <b>{len(csv_files)} CSV files merged</b> '
                f'— {len(orders_df):,} total rows<br>{pills}</div>',
                unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

# ── Brand Charges Configuration ───────────────────────────────────────────────
all_brands_in_data = sorted(orders_df["brand"].dropna().unique().tolist())
all_brands_in_data = [b for b in all_brands_in_data
                      if b and b.lower() not in ("nan","none","")]

with st.expander("⚙️ Brand Charges Configuration", expanded=False):
    st.markdown("""<div class="info-box" style="margin-bottom:1rem;">
Set <b>Marketing</b>, <b>Return Charges</b> and <b>Royalty</b> per brand.
Marketing and Royalty can each be a <b>%</b> or a <b>Flat ₹</b> amount.<br>
Pre-filled with defaults (Marketing 3%, Return ₹45, Royalty 1%).<br>
Only change what's different for a particular brand.
</div>""", unsafe_allow_html=True)

    brand_cfg = {}

    # Header row
    hcols = st.columns([1.8, 1.4, 1.3, 1.5, 1.4, 1.3])
    hcols[0].markdown("**Brand**")
    hcols[1].markdown("**Marketing Type**")
    hcols[2].markdown("**Marketing Value**")
    hcols[3].markdown("**Return Charges (₹)**")
    hcols[4].markdown("**Royalty Type**")
    hcols[5].markdown("**Royalty Value**")

    st.markdown("<hr style='margin:.3rem 0 .6rem 0;border-color:#e2e8f0'>",
                unsafe_allow_html=True)

    for brand in all_brands_in_data:
        cols = st.columns([1.8, 1.4, 1.3, 1.5, 1.4, 1.3])

        with cols[0]:
            st.markdown(f"<div style='padding:.45rem 0;font-weight:600;color:#1e293b'>"
                        f"{brand}</div>", unsafe_allow_html=True)

        with cols[1]:
            mkt_type = st.selectbox(
                "mkt_type", ["%" , "Flat ₹"],
                key=f"mkt_type_{brand}", label_visibility="collapsed"
            )

        with cols[2]:
            mkt_val = st.number_input(
                "mkt_val", min_value=0.0, value=DEFAULT_MARKETING_VALUE,
                step=0.5, format="%.2f",
                key=f"mkt_val_{brand}", label_visibility="collapsed",
                help="Enter % value (e.g. 3 = 3%) or flat ₹ amount depending on type"
            )

        with cols[3]:
            ret_ch = st.number_input(
                "return", min_value=0.0, value=DEFAULT_RETURN_CHARGES,
                step=1.0, format="%.2f",
                key=f"ret_{brand}", label_visibility="collapsed"
            )

        with cols[4]:
            roy_type = st.selectbox(
                "roy_type", ["%", "Flat ₹"],
                key=f"roy_type_{brand}", label_visibility="collapsed"
            )

        with cols[5]:
            roy_val = st.number_input(
                "roy_val", min_value=0.0, value=DEFAULT_ROYALTY_VALUE,
                step=0.1, format="%.2f",
                key=f"roy_val_{brand}", label_visibility="collapsed",
                help="% of V (SP − GT) if type is %, otherwise a flat ₹ amount per order"
            )

        brand_cfg[brand] = {
            "mkt_type": mkt_type,
            "mkt_val":  mkt_val,
            "return":   ret_ch,
            "roy_type": roy_type,
            "roy_val":  roy_val,
        }

    st.markdown("<br>", unsafe_allow_html=True)

    # Preview table
    preview_rows = []
    for b, cfg in brand_cfg.items():
        mt = cfg["mkt_type"]
        mv = cfg["mkt_val"]
        rt = cfg["roy_type"]
        rv = cfg["roy_val"]
        preview_rows.append({
            "Brand":      b,
            "Marketing":  f"{mv}% of SP" if mt=="%" else f"₹{mv:.2f} flat",
            "Return (₹)": f"₹{cfg['return']:.2f}",
            "Royalty":    f"{rv}% of V" if rt=="%" else f"₹{rv:.2f} flat",
        })
    if preview_rows:
        st.markdown("**Current configuration:**")
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True,
                     hide_index=True)


# ── Reconcile ─────────────────────────────────────────────────────────────────
with st.spinner("Reconciling…"):
    result = reconcile(
        rates, orders_df, oms_map, yrn_map, pwn_map, closed_map,
        fix_df=fix_df,
        report_date=report_date_input,
        brand_filter=brand_filter if brand_filter else None,
        brand_cfg=brand_cfg,
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
    ("Total Orders",        f"{total_orders:,}",        None,
     f"{len(csv_files)} file(s) merged"),
    ("Total Selling Price", f"₹{total_sp:,.0f}",       None, "sum of final amount"),
    ("Myntra Payable",      f"₹{total_payable:,.0f}",  None, "after all deductions"),
    ("Total Charges+GST",   f"₹{total_tc_gst:,.0f}",  None, "platform fees"),
    ("Fix Rate Applied",    f"{fix_rows}/{total_orders}",
     "positive" if fix_rows>0 else None, "style-level overrides"),
    ("Unmatched Rows",      f"{unmatched}",
     "negative" if unmatched>0 else "positive", "no slab/fix rate"),
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
🔒 <b>{fix_rows} orders</b> used Fix Rate (Brand + Style ID + Date).
GT/Commission/Fixed Fee from Fix Rate sheet — GST embedded, no separate deduction.
</div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Brand summary ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">📊 Brand Summary</div>', unsafe_allow_html=True)

summary_rows = []
for brand in sorted(result["brand"].unique()):
    bdf  = result[result["brand"] == brand]
    bok  = bdf[bdf["_slab_ok"] == True]
    fx_n = int((bdf["_rate_source"] == "Fix Rate").sum())
    cfg  = brand_cfg.get(brand, {})
    mt   = cfg.get("mkt_type", DEFAULT_MARKETING_TYPE)
    mv   = cfg.get("mkt_val",  DEFAULT_MARKETING_VALUE)
    rt   = cfg.get("roy_type", DEFAULT_ROYALTY_TYPE)
    rv   = cfg.get("roy_val",  DEFAULT_ROYALTY_VALUE)
    summary_rows.append({
        "Brand":               brand,
        "Orders":              len(bdf),
        "Fix Rate Orders":     fx_n,
        "Marketing Rate":      f"{mv}% of SP" if mt=="%" else f"₹{mv:.2f} flat",
        "Return Charges (₹)":  cfg.get("return", DEFAULT_RETURN_CHARGES),
        "Royalty":             f"{rv}% of V" if rt=="%" else f"₹{rv:.2f} flat",
        "Selling Price (₹)":   round(bdf["final amount"].sum(), 2),
        "GT Charges (₹)":      round(bok["_GT"].sum(), 2),
        "Commission (₹)":      round(bok["_comm_amt"].sum(), 2),
        "Total Charges (₹)":   round(bok["_total_charges"].sum(), 2),
        "GST (₹)":             round(bok["_gst"].sum(), 2),
        "Myntra Payable (₹)":  round(bok["_myntra_payable"].sum(), 2),
    })

sum_df = pd.DataFrame(summary_rows)
st.dataframe(
    sum_df.style
        .format({c: "{:,.2f}" for c in sum_df.columns
                 if "₹" in c and c not in ("Return Charges (₹)","Royalty")},
                na_rep="")
        .set_properties(**{"text-align":"center"}),
    use_container_width=True, hide_index=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# ── Order-level tabs ──────────────────────────────────────────────────────────
brands_in_result = sorted(result["brand"].unique().tolist())

if brands_in_result:
    st.markdown('<div class="section-header">📋 Order-Level Detail</div>',
                unsafe_allow_html=True)
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
        "_marketing":       "Marketing ₹",
        "_return_ch":       "Return Charges ₹",
        "_royalty":         "Royalty ₹",
        "_pwn":             "PWN+10% ₹",
    }
    NUM_KEYS = {"final amount","_V","_comm_amt","_GT","_fixed_fee",
                "_total_charges","_gst","_myntra_payable",
                "_marketing","_return_ch","_royalty","_pwn"}

    for tab, brand in zip(tabs[:-1], brands_in_result):
        with tab:
            bdf = result[result["brand"] == brand].copy()
            ok  = bdf[bdf["_slab_ok"] == True]
            mis = bdf[bdf["_slab_ok"] == False]
            fx  = int((bdf["_rate_source"] == "Fix Rate").sum())

            # Show the configured charges for this brand
            cfg = brand_cfg.get(brand, {})
            mt  = cfg.get("mkt_type", DEFAULT_MARKETING_TYPE)
            mv  = cfg.get("mkt_val",  DEFAULT_MARKETING_VALUE)
            rc  = cfg.get("return",   DEFAULT_RETURN_CHARGES)
            rt  = cfg.get("roy_type", DEFAULT_ROYALTY_TYPE)
            rv  = cfg.get("roy_val",  DEFAULT_ROYALTY_VALUE)
            mkt_str = f"{mv}% of SP" if mt=="%" else f"₹{mv:.2f} flat"
            roy_str = f"{rv}% of V" if rt=="%" else f"₹{rv:.2f} flat"
            st.markdown(
                f'<div class="info-box" style="margin-bottom:.75rem;">'
                f'<b>Charges applied:</b> '
                f'Marketing = {mkt_str} &nbsp;·&nbsp; '
                f'Return = ₹{rc:.0f} &nbsp;·&nbsp; '
                f'Royalty = {roy_str}</div>',
                unsafe_allow_html=True)

            c1,c2,c3,c4,c5 = st.columns(5)
            c1.metric("Orders",         len(bdf))
            c2.metric("Selling Price",  f"₹{bdf['final amount'].sum():,.0f}")
            c3.metric("Myntra Payable", f"₹{ok['_myntra_payable'].sum():,.0f}")
            c4.metric("Fix Rate Rows",  fx)
            c5.metric("Unmatched",      len(mis),
                      delta=str(len(mis)) if len(mis) else None,
                      delta_color="inverse")

            disp_cols = {k:v for k,v in DISPLAY_COLS.items()
                         if k != "_source_file" or len(csv_files) > 1}

            display_df = bdf[[c for c in disp_cols if c in bdf.columns]].rename(columns=disp_cols)
            num_dcols  = [disp_cols[k] for k in NUM_KEYS
                          if k in disp_cols and k in bdf.columns]
            for c in num_dcols:
                if c in display_df.columns:
                    display_df[c] = pd.to_numeric(display_df[c], errors="coerce")

            def _highlight(row):
                if row.get("Rate Source") == "Fix Rate":
                    return ["background-color:#fffbeb"] * len(row)
                return [""] * len(row)

            st.dataframe(
                display_df.style
                    .format({c:"{:,.2f}" for c in num_dcols if c in display_df.columns},
                            na_rep="")
                    .apply(_highlight, axis=1),
                use_container_width=True, hide_index=True,
                height=min(420, 45+35*len(display_df)),
            )

            if len(mis) > 0:
                st.markdown(f"""<div class="warn-box">
⚠️ {len(mis)} rows had no matching slab or fix rate.</div>""",
                    unsafe_allow_html=True)
                st.dataframe(
                    mis[["seller sku code","style id","article type",
                         "final amount","order status"]],
                    use_container_width=True, hide_index=True)

    with tabs[-1]:
        if bad_df.empty:
            st.markdown('<div class="success-box">✅ All rows matched.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="warn-box">⚠️ {len(bad_df)} unmatched orders.</div>',
                        unsafe_allow_html=True)
            st.dataframe(
                bad_df[["brand","seller sku code","style id","article type",
                        "final amount","order status"]],
                use_container_width=True, hide_index=True)

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">⬇️ Export</div>', unsafe_allow_html=True)

col_dl, col_info = st.columns([1,3])
with col_dl:
    with st.spinner("Preparing Excel…"):
        xlsx_bytes = build_excel(result,
                                  report_date_str=str(report_date_input),
                                  brand_cfg=brand_cfg)
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
Column headers show the <b>actual rate</b> applied for each brand
(e.g. "Marketing (2.5% of SP)" or "Return Charges (₹50)").<br>
🟡 Amber rows = Fix Rate. ⬜ White = Slab. 🟠 Orange PWN = fill manually.
</div>""", unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)
