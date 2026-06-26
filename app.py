import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page config
st.set_page_config(
    page_title="Myntra Reconciliation",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #0f172a;
    border-right: 1px solid #1e293b;
}
section[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
section[data-testid="stSidebar"] .stFileUploader label { color: #cbd5e1 !important; font-size: 0.78rem; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #ffffff !important; }
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div { color: #e2e8f0 !important; }
section[data-testid="stSidebar"] .stMarkdown { color: #e2e8f0 !important; }

/* Multiselect in sidebar */
section[data-testid="stSidebar"] .stMultiSelect label { color: #cbd5e1 !important; }
section[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="tag"] { background: #1e40af !important; }
section[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="tag"] span { color: #ffffff !important; }

/* Main background */
.main .block-container { padding-top: 1.5rem; max-width: 1400px; }

/* Metric cards */
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.1rem 1.25rem;
    box-shadow: 0 2px 6px rgba(0,0,0,.07);
}
.metric-label {
    font-size: 0.72rem; font-weight: 700; color: #475569;
    text-transform: uppercase; letter-spacing: .07em; margin-bottom: .3rem;
}
.metric-value {
    font-size: 1.5rem; font-weight: 700; color: #0f172a;
    font-family: 'JetBrains Mono', monospace;
}
.metric-value.positive { color: #15803d; }
.metric-value.negative { color: #b91c1c; }
.metric-sub { font-size: 0.72rem; color: #64748b; margin-top: .2rem; font-weight: 500; }

/* Section header */
.section-header {
    display: flex; align-items: center; gap: .6rem;
    font-size: 1rem; font-weight: 700; color: #0f172a;
    border-bottom: 2px solid #e2e8f0; padding-bottom: .5rem; margin-bottom: 1rem;
}

/* Upload hint */
.upload-hint {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: .85rem 1rem;
    font-size: .78rem;
    color: #e2e8f0 !important;
    margin-top: .5rem;
    line-height: 1.8;
}
.upload-hint b { color: #7dd3fc !important; }

/* Info box */
.info-box {
    background: #eff6ff;
    border-left: 3px solid #3b82f6;
    border-radius: 0 8px 8px 0;
    padding: .8rem 1rem;
    font-size: .83rem;
    color: #1e3a5f;
    font-weight: 500;
}
.info-box b { color: #1d4ed8; }
.info-box code { background: #dbeafe; color: #1e40af; padding: .1rem .3rem; border-radius: 3px; }

/* Warn box */
.warn-box {
    background: #fffbeb;
    border-left: 3px solid #f59e0b;
    border-radius: 0 8px 8px 0;
    padding: .8rem 1rem;
    font-size: .83rem;
    color: #78350f;
    font-weight: 500;
}

/* Success box */
.success-box {
    background: #f0fdf4;
    border-left: 3px solid #22c55e;
    border-radius: 0 8px 8px 0;
    padding: .8rem 1rem;
    font-size: .83rem;
    color: #14532d;
    font-weight: 500;
}

/* Page title */
h1 { color: #0f172a !important; font-weight: 800 !important; }
h2 { color: #1e293b !important; font-weight: 700 !important; }
h3 { color: #1e293b !important; }

/* Streamlit metric widget text */
[data-testid="stMetricLabel"] { color: #334155 !important; font-weight: 600 !important; }
[data-testid="stMetricValue"] { color: #0f172a !important; font-weight: 700 !important; }

/* Tab text */
.stTabs [data-baseweb="tab"] { color: #334155 !important; font-weight: 600; }
.stTabs [data-baseweb="tab"][aria-selected="true"] { color: #1d4ed8 !important; }

/* Dataframe */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
#  CORE RECONCILIATION LOGIC
# =============================================================================

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


@st.cache_data(show_spinner=False)
def load_slab(file_bytes: bytes):
    wb = __import__("openpyxl").load_workbook(BytesIO(file_bytes), data_only=True)

    # ---- Rates sheet --------------------------------------------------------
    ws = wb["Rates"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        raise ValueError("'Rates' sheet is empty.")

    header = list(rows[0])
    while header and (header[-1] is None or str(header[-1]).strip() == ""):
        header.pop()

    n_cols = len(header)
    if n_cols == 0:
        raise ValueError("Could not find a valid header row in the 'Rates' sheet.")

    data_rows = []
    for r in rows[1:]:
        r = list(r)
        if len(r) < n_cols:
            r = r + [None] * (n_cols - len(r))
        elif len(r) > n_cols:
            r = r[:n_cols]
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        data_rows.append(r)

    rates = pd.DataFrame(data_rows, columns=header)
    rates.columns = [str(c).strip() for c in rates.columns]

    missing = [c for c in REQUIRED_RATE_COLS if c not in rates.columns]
    if missing:
        raise ValueError(
            f"'Rates' sheet is missing required column(s): {', '.join(missing)}. "
            f"Found columns: {list(rates.columns)}"
        )

    rates["Brand Name"] = rates["Brand Name"].astype(str).str.strip()
    rates["Category"]   = rates["Category"].astype(str).str.strip()

    numeric_cols = [
        "Lower Limit Commision", "Upper Limit Commision", "Commision Charge",
        "GT Lower Limit", "GT Upper Limit", "GT Charges",
        "Lower Limit Fixed Fee", "Upper Limit Fixed Fee", "Fix Fee",
    ]
    for c in numeric_cols:
        rates[c] = pd.to_numeric(rates[c], errors="coerce")

    # ---- Replace Sku sheet --------------------------------------------------
    # col[1] = YRN NUMBER, col[2] = MYNTRA SKU CODE, col[4] = OMS SKU CODE
    oms_map = {}   # MYNTRA SKU CODE → OMS SKU CODE
    yrn_map = {}   # YRN NUMBER      → OMS SKU CODE
    if "Replace Sku" in wb.sheetnames:
        ws2 = wb["Replace Sku"]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * max(0, 5 - len(row))
            if row[2] and row[4]:
                oms_map[str(row[2]).strip()] = str(row[4]).strip()
            if row[1] and row[4]:
                yrn_map[str(row[1]).strip()] = str(row[4]).strip()

    # ---- Price We Need Excel sheet — OMS Child SKU → PWN+10% ---------------
    pwn_map = {}
    pwn_sheet_name = None
    for sn in wb.sheetnames:
        if sn.strip().lower().startswith("price we need"):
            pwn_sheet_name = sn
            break
    if pwn_sheet_name:
        ws3 = wb[pwn_sheet_name]
        for row in ws3.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * max(0, 3 - len(row))
            if row[1] and row[2] is not None:
                pwn_map[str(row[1]).strip()] = row[2]

    # ---- Closed sheet — OMS Child SKU → Closed Price ------------------------
    closed_map = {}
    if "Closed" in wb.sheetnames:
        ws4 = wb["Closed"]
        for row in ws4.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * max(0, 3 - len(row))
            if row[1] and row[2] is not None:
                closed_map[str(row[1]).strip()] = row[2]

    return rates, oms_map, yrn_map, pwn_map, closed_map


def get_pwn_price(oms_map: dict, yrn_map: dict, pwn_map: dict, closed_map: dict,
                  seller_sku_code: str, myntra_sku_code: str = ""):
    key  = str(seller_sku_code).strip()
    mkey = str(myntra_sku_code).strip()

    # Step 1: Resolve via MYNTRA SKU CODE column in Replace Sku
    oms_sku = oms_map.get(key, key)

    # Step 2: If not resolved, try Myntra SKU Code via YRN NUMBER column
    if oms_sku == key and mkey:
        oms_sku = yrn_map.get(mkey, oms_sku)

    # Step 3: Check Closed sheet first (using resolved OMS SKU)
    if oms_sku in closed_map:
        return closed_map[oms_sku], oms_sku, "Closed"

    # Step 4: Check Price We Need Excel
    if oms_sku in pwn_map:
        return pwn_map[oms_sku], oms_sku, "PWN"

    return None, oms_sku, None


def _lookup(subset: pd.DataFrame, lo_col: str, hi_col: str, val: float):
    m = subset[(subset[lo_col] <= val) & (subset[hi_col] > val)]
    if m.empty:
        m = subset[(subset[lo_col] <= val) & (subset[hi_col] >= val)]
    return m.iloc[0] if not m.empty else None


def get_charges(rates: pd.DataFrame, brand: str, cat: str, SP: float):
    """
    Returns (GT, V, comm_rate, comm_amt, fixed_fee) or all None on miss.
    1. GT      = slab lookup by SP
    2. V       = SP - GT
    3. comm    = rate(V) * V
    4. fixed   = slab lookup by SP
    """
    sub = rates[
        (rates["Brand Name"] == brand) &
        (rates["Category"].str.lower() == cat.lower())
    ]
    if sub.empty:
        return None, None, None, None, None

    gt_row  = _lookup(sub, "GT Lower Limit",        "GT Upper Limit",        SP)
    fee_row = _lookup(sub, "Lower Limit Fixed Fee",  "Upper Limit Fixed Fee", SP)
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


def reconcile(rates: pd.DataFrame, df: pd.DataFrame, oms_map: dict, yrn_map: dict,
              pwn_map: dict, closed_map: dict, brand_filter=None):
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

        GT, V, comm_rate, comm_amt, fixed_fee = get_charges(rates, brand, cat, SP)
        pwn_price, oms_sku, pwn_source = get_pwn_price(
            oms_map, yrn_map, pwn_map, closed_map,
            row.get("seller sku code", ""),
            row.get("myntra sku code", "")
        )

        if GT is None:
            rec = dict(
                _GT=None, _V=None, _comm_rate=None, _comm_amt=None,
                _fixed_fee=None, _total_charges=None, _gst=None,
                _myntra_payable=None, _marketing=None, _royalty=None,
                _slab_ok=False,
            )
        else:
            total_ch   = round(comm_amt + GT + fixed_fee, 2)
            gst        = round((total_ch - GT) * GST_RATE, 2)
            myntra_pay = round(SP - total_ch - gst, 2)
            marketing  = round(SP * MARKETING_PCT, 2)
            royalty    = round(V  * ROYALTY_PCT, 2)
            rec = dict(
                _GT=GT, _V=V, _comm_rate=comm_rate, _comm_amt=comm_amt,
                _fixed_fee=fixed_fee, _total_charges=total_ch, _gst=gst,
                _myntra_payable=myntra_pay, _marketing=marketing,
                _royalty=royalty, _slab_ok=True,
            )
        rec["_pwn"]        = pwn_price
        rec["_oms_sku"]    = oms_sku
        rec["_pwn_source"] = pwn_source
        records.append(rec)

    enrich = pd.DataFrame(records, index=df.index)
    return pd.concat([df.reset_index(drop=True), enrich.reset_index(drop=True)], axis=1)


# =============================================================================
#  EXCEL EXPORT
# =============================================================================

HEADERS = [
    "order release id", "order line id", "STYLE ID", "Myntra SKU Code", "SKU",
    "Original SKU", "Article Type", "MRP", "PWN+10%", "Marketing Charges 3%",
    "Return Charges", "Royalty", "Total Amount", "Commission Amount",
    "GT Charges", "Fixed Fee", "Total Charges", "GST Amount",
    "Myntra Payble Amount", "Rebate", "Diffrence", "Selling Price",
    "Selling Price -GT Price", "Date", "Brand", "Order Status", "Price Source",
]

COL_W = [18,15,12,24,24,22,14,8,10,18,14,10,13,17,11,10,13,11,18,8,11,13,18,13,12,12,12]

_H  = PatternFill("solid", fgColor="1F4E79")
_Y  = PatternFill("solid", fgColor="FCE4D6")
_B  = PatternFill("solid", fgColor="DEEAF1")
_G  = PatternFill("solid", fgColor="E2EFDA")
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


def build_excel(result_df: pd.DataFrame) -> bytes:
    wb         = Workbook()
    all_brands = sorted(result_df["brand"].unique())

    for brand in all_brands:
        bdf = result_df[result_df["brand"] == brand].reset_index(drop=True)
        ws  = wb.create_sheet(title=brand[:31])

        hfont = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
        nfont = Font(size=9, name="Calibri")
        ofont = Font(size=9, name="Calibri", italic=True, color="C55A11")
        gfont = Font(bold=True, size=9, name="Calibri", color="375623")

        for ci, (h, w) in enumerate(zip(HEADERS, COL_W), 1):
            _cell(ws, 1, ci, h, _H, hfont)
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 36

        for ri, (_, row) in enumerate(bdf.iterrows(), 2):
            alt = _A1 if ri % 2 == 0 else _A2
            SP  = row.get("final amount", 0)
            V   = row.get("_V")

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
                None,                       # M: Total Amount (formula)
                row.get("_comm_amt"),
                row.get("_GT"),
                row.get("_fixed_fee"),
                row.get("_total_charges"),
                row.get("_gst"),
                row.get("_myntra_payable"),
                0,                          # T: Rebate
                None,                       # U: Difference (formula)
                SP,
                V,
                "25-Jun-2026",
                row.get("brand", ""),
                row.get("order status", ""),
                row.get("_pwn_source") or "Manual",
            ]

            num_cols = {10,11,12,13,14,15,16,17,18,19,20,21,22,23}
            for ci, val in enumerate(vals, 1):
                if ci == 9:
                    fill_h = alt if val is not None else _Y
                    font_h = nfont if val is not None else ofont
                    _cell(ws, ri, ci, val, fill_h, font_h,
                          fmt="#,##0.00" if val is not None else None)
                elif ci == 19:
                    _cell(ws, ri, ci, val, _G, gfont,
                          fmt="#,##0.00" if val is not None else None)
                elif ci == 21:
                    _cell(ws, ri, ci, None, _B, nfont)
                elif ci == 13:
                    _cell(ws, ri, ci, None, alt, nfont)
                else:
                    fmt = "#,##0.00" if ci in num_cols else (
                          "#,##0"    if ci == 8 else None)
                    _cell(ws, ri, ci, val, alt, nfont, fmt)

            m = ws.cell(row=ri, column=13)
            m.value = f'=IF(I{ri}="","",K{ri}+J{ri}+I{ri}+L{ri})'
            m.number_format = "#,##0.00"

            u = ws.cell(row=ri, column=21)
            u.value = f'=IF(I{ri}="","",S{ri}-M{ri}+T{ri})'
            u.number_format = "#,##0.00"

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

        nr = len(bdf) + 3
        note = ws.cell(row=nr, column=1,
            value="PWN+10% (Column I) auto-filled via Replace Sku → Price We Need / Closed lookup. "
                  "Orange cells = no SKU/price match found — fill manually.")
        note.font  = Font(bold=True, italic=True, size=9, color="C55A11", name="Calibri")
        note.fill  = PatternFill("solid", fgColor="FFF2CC")
        ws.merge_cells(f"A{nr}:{get_column_letter(len(HEADERS))}{nr}")

    # Summary sheet
    ws_sum = wb.create_sheet(title="Summary", index=0)
    sum_headers = ["Brand", "Orders", "Total Selling Price",
                   "Total GT Charges", "Total Commission",
                   "Total Fixed Fee", "Total Charges",
                   "Total GST", "Total Myntra Payable"]
    sum_widths  = [14, 8, 20, 16, 16, 14, 14, 12, 20]

    hfont = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    for ci, (h, w) in enumerate(zip(sum_headers, sum_widths), 1):
        _cell(ws_sum, 1, ci, h, _H, hfont)
        ws_sum.column_dimensions[get_column_letter(ci)].width = w
    ws_sum.row_dimensions[1].height = 30

    grand = {k: 0.0 for k in ["sp","gt","comm","ff","tc","gst","pay"]}
    grand["n"] = 0

    for ri, brand in enumerate(all_brands, 2):
        bdf = result_df[result_df["brand"] == brand]
        ok  = bdf[bdf["_slab_ok"] == True]
        n   = len(bdf)
        sp  = round(bdf["final amount"].sum(), 2)
        gt  = round(ok["_GT"].sum(), 2)
        cm  = round(ok["_comm_amt"].sum(), 2)
        ff  = round(ok["_fixed_fee"].sum(), 2)
        tc  = round(ok["_total_charges"].sum(), 2)
        gst = round(ok["_gst"].sum(), 2)
        pay = round(ok["_myntra_payable"].sum(), 2)

        alt = _A1 if ri % 2 == 0 else _A2
        nf  = Font(size=9, name="Calibri")
        bf  = Font(bold=True, size=9, name="Calibri")
        for ci, v in enumerate([brand,n,sp,gt,cm,ff,tc,gst,pay], 1):
            fmt = "#,##0.00" if ci > 2 else None
            _cell(ws_sum, ri, ci, v, alt, bf if ci == 1 else nf, fmt)

        for k, v in zip(["n","sp","gt","comm","ff","tc","gst","pay"],
                         [n, sp, gt, cm, ff, tc, gst, pay]):
            grand[k] += v

    gr = len(all_brands) + 2
    gf = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    gv = [
        "GRAND TOTAL", grand["n"],
        round(grand["sp"],2), round(grand["gt"],2), round(grand["comm"],2),
        round(grand["ff"],2), round(grand["tc"],2), round(grand["gst"],2),
        round(grand["pay"],2),
    ]
    for ci, v in enumerate(gv, 1):
        fmt = "#,##0.00" if ci > 2 else None
        _cell(ws_sum, gr, ci, v, _H, gf, fmt)

    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
#  STREAMLIT UI
# =============================================================================

with st.sidebar:
    st.markdown("## Myntra Recon")
    st.markdown("---")

    st.markdown("**Step 1 - Upload Slab File**")
    slab_file = st.file_uploader(
        "Slab.xlsx (Rates + Replace Sku sheets)",
        type=["xlsx"], key="slab",
        help="Must contain a 'Rates' sheet with Brand Name, Category, slab limits."
    )

    st.markdown("**Step 2 - Upload Orders CSV**")
    csv_file = st.file_uploader(
        "Seller Orders Report (.csv)",
        type=["csv"], key="csv",
        help="Myntra Seller Orders export with brand, article type, final amount, seller price."
    )

    st.markdown("---")

    brand_filter = []
    if csv_file:
        try:
            _df_tmp = pd.read_csv(csv_file)
            csv_file.seek(0)
            _df_tmp["brand"] = _df_tmp["brand"].astype(str).str.strip()
            all_b = sorted(_df_tmp["brand"].unique().tolist())
            brand_filter = st.multiselect(
                "Filter by Brand",
                options=all_b,
                default=all_b,
                help="Select which brands to reconcile."
            )
        except Exception:
            pass

    st.markdown("---")
    st.markdown("""
<div class="upload-hint">
<b>Formula used:</b><br>
- GT = slab lookup by SP<br>
- V = SP - GT<br>
- Commission = rate(V) * V<br>
- Fixed Fee = slab lookup by SP<br>
- Total Charges = Comm + GT + Fee<br>
- GST = (Total - GT) * 18%<br>
- Myntra Payable = SP - TC - GST<br>
- Marketing = SP * 3%<br>
- Royalty = V * 1%
</div>
""", unsafe_allow_html=True)


# Main area
st.markdown("## Myntra Seller Reconciliation")
st.markdown("Upload your **Slab** and **Orders CSV** in the sidebar to begin.")

if not slab_file or not csv_file:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
<div class="info-box">
<b>Step 1 - Slab.xlsx</b><br>
Your commission/GT/fixed-fee rate table. Must have a <code>Rates</code> sheet with:<br>
Brand Name - Category - slab limits (commission, GT, fixed fee)
</div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
<div class="info-box">
<b>Step 2 - Orders CSV</b><br>
Myntra Seller Orders export. Key columns needed:<br>
<code>brand</code> - <code>article type</code> - <code>final amount</code> - <code>seller price</code>
</div>""", unsafe_allow_html=True)
    st.stop()


# Load data
with st.spinner("Loading slab rates..."):
    try:
        rates, oms_map, yrn_map, pwn_map, closed_map = load_slab(slab_file.read())
    except Exception as e:
        st.error(f"Could not read Slab file: {e}")
        st.stop()

with st.spinner("Reading orders..."):
    try:
        orders_df = pd.read_csv(csv_file)
        orders_df["brand"]        = orders_df["brand"].astype(str).str.strip()
        orders_df["article type"] = orders_df["article type"].astype(str).str.strip()
        orders_df["final amount"] = pd.to_numeric(orders_df["final amount"], errors="coerce").fillna(0)
        orders_df["seller price"] = pd.to_numeric(orders_df["seller price"], errors="coerce").fillna(0)
    except Exception as e:
        st.error(f"Could not read CSV: {e}")
        st.stop()


# Run reconciliation
with st.spinner("Reconciling..."):
    result = reconcile(rates, orders_df, oms_map, yrn_map, pwn_map, closed_map,
                       brand_filter if brand_filter else None)

ok_df  = result[result["_slab_ok"] == True]
bad_df = result[result["_slab_ok"] == False]


# KPI row
total_orders  = len(result)
total_sp      = result["final amount"].sum()
total_payable = ok_df["_myntra_payable"].sum()
total_charges = ok_df["_total_charges"].sum()
total_gst     = ok_df["_gst"].sum()
unmatched     = len(bad_df)
pwn_found     = result["_pwn"].notna().sum()
pwn_missing   = total_orders - pwn_found

cols = st.columns(6)
kpi_data = [
    ("Total Orders",         f"{total_orders:,}",                None,       "orders in report"),
    ("Total Selling Price",  f"Rs {total_sp:,.0f}",              None,       "sum of final amount"),
    ("Total Myntra Payable", f"Rs {total_payable:,.0f}",         None,       "after all deductions"),
    ("Total Charges + GST",  f"Rs {total_charges+total_gst:,.0f}", None,     "platform fees"),
    ("PWN Price Matched",    f"{pwn_found}/{total_orders}",
     "positive" if pwn_missing == 0 else "negative",
     "via Replace Sku → Price We Need / Closed"),
    ("Unmatched Rows",       f"{unmatched}",
     "negative" if unmatched > 0 else "positive",
     "no slab found"),
]
for col, (label, value, cls, sub) in zip(cols, kpi_data):
    with col:
        st.markdown(f"""
<div class="metric-card">
  <div class="metric-label">{label}</div>
  <div class="metric-value {cls or ''}">{value}</div>
  <div class="metric-sub">{sub}</div>
</div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# Brand summary table
st.markdown('<div class="section-header">Brand Summary</div>', unsafe_allow_html=True)

summary_rows = []
for brand in sorted(result["brand"].unique()):
    bdf = result[result["brand"] == brand]
    bok = bdf[bdf["_slab_ok"] == True]
    summary_rows.append({
        "Brand":              brand,
        "Orders":             len(bdf),
        "Selling Price (Rs)": round(bdf["final amount"].sum(), 2),
        "GT Charges (Rs)":    round(bok["_GT"].sum(), 2),
        "Commission (Rs)":    round(bok["_comm_amt"].sum(), 2),
        "Fixed Fee (Rs)":     round(bok["_fixed_fee"].sum(), 2),
        "Total Charges (Rs)": round(bok["_total_charges"].sum(), 2),
        "GST (Rs)":           round(bok["_gst"].sum(), 2),
        "Myntra Payable (Rs)":round(bok["_myntra_payable"].sum(), 2),
    })

sum_df = pd.DataFrame(summary_rows)
st.dataframe(
    sum_df.style
        .format({c: "{:,.2f}" for c in sum_df.columns if "(Rs)" in c}, na_rep="")
        .set_properties(**{"text-align": "center"}),
    use_container_width=True, hide_index=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# Per-brand detail tabs
brands_in_result = sorted(result["brand"].unique().tolist())

if brands_in_result:
    st.markdown('<div class="section-header">Order-Level Detail</div>', unsafe_allow_html=True)
    tabs = st.tabs([f"  {b}  " for b in brands_in_result] + ["Unmatched"])

    DISPLAY_COLS = {
        "order release id": "Order ID",
        "order line id":    "Line ID",
        "seller sku code":  "SKU",
        "_oms_sku":         "Original SKU",
        "article type":     "Category",
        "order status":     "Status",
        "final amount":     "Selling Price",
        "_pwn":             "PWN+10%",
        "_pwn_source":      "Price Source",
        "_V":               "SP-GT",
        "_comm_amt":        "Commission",
        "_GT":              "GT Charges",
        "_fixed_fee":       "Fixed Fee",
        "_total_charges":   "Total Charges",
        "_gst":             "GST",
        "_myntra_payable":  "Myntra Payable",
        "_marketing":       "Marketing 3%",
        "_royalty":         "Royalty 1%",
    }

    NUM_KEYS = {
        "final amount", "_pwn", "_V", "_comm_amt", "_GT",
        "_fixed_fee", "_total_charges", "_gst", "_myntra_payable",
        "_marketing", "_royalty",
    }

    for tab, brand in zip(tabs[:-1], brands_in_result):
        with tab:
            bdf = result[result["brand"] == brand].copy()
            ok  = bdf[bdf["_slab_ok"] == True]
            mis = bdf[bdf["_slab_ok"] == False]

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Orders", len(bdf))
            with c2:
                st.metric("Selling Price", f"Rs {bdf['final amount'].sum():,.0f}")
            with c3:
                st.metric("Myntra Payable", f"Rs {ok['_myntra_payable'].sum():,.0f}")
            with c4:
                st.metric("Unmatched", len(mis),
                          delta=f"{len(mis)} rows" if len(mis) else None,
                          delta_color="inverse")

            display_df = ok[list(DISPLAY_COLS.keys())].rename(columns=DISPLAY_COLS)

            # Coerce numeric columns
            num_display_cols = [DISPLAY_COLS[k] for k in NUM_KEYS if k in DISPLAY_COLS]
            for _c in num_display_cols:
                if _c in display_df.columns:
                    display_df[_c] = pd.to_numeric(display_df[_c], errors="coerce")

            fmt_map = {c: "{:,.2f}" for c in num_display_cols if c in display_df.columns}

            st.dataframe(
                display_df.style
                    .format(fmt_map, na_rep="")
                    .map(
                        lambda v: "background-color:#f0fdf4;color:#14532d;font-weight:600"
                        if isinstance(v, (int, float)) and v > 0 else "",
                        subset=["Myntra Payable"] if "Myntra Payable" in display_df.columns else []
                    ),
                use_container_width=True,
                hide_index=True,
                height=min(400, 45 + 35 * len(display_df)),
            )

            if len(mis) > 0:
                st.markdown(f"""
<div class="warn-box">
{len(mis)} rows had no matching slab (brand+category not found in Rates sheet).
Check brand name spelling and that the category exists in the Slab file.
</div>""", unsafe_allow_html=True)
                st.dataframe(
                    mis[["seller sku code", "article type", "final amount", "order status"]],
                    use_container_width=True, hide_index=True,
                )

    with tabs[-1]:
        if bad_df.empty:
            st.markdown("""
<div class="success-box">All rows matched a slab - no unmatched orders.</div>
""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
<div class="warn-box">{len(bad_df)} orders could not be matched to a slab.</div>
""", unsafe_allow_html=True)
            st.dataframe(
                bad_df[["brand", "seller sku code", "article type",
                        "final amount", "order status"]],
                use_container_width=True, hide_index=True,
            )


# Download
st.markdown("---")
st.markdown('<div class="section-header">Export</div>', unsafe_allow_html=True)

col_dl, col_info = st.columns([1, 3])
with col_dl:
    with st.spinner("Preparing Excel..."):
        xlsx_bytes = build_excel(result)
    st.download_button(
        label="Download Reconciliation (.xlsx)",
        data=xlsx_bytes,
        file_name="Myntra_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_info:
    st.markdown("""
<div class="info-box">
The Excel file has <b>one sheet per brand</b> + a <b>Summary</b> sheet.<br>
Column layout matches your reference workbook exactly.<br>
<b>PWN+10% (col I)</b> is auto-filled via Replace Sku → Price We Need / Closed lookup.<br>
Orange cells = SKU not found in any sheet — fill manually.
</div>""", unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)
