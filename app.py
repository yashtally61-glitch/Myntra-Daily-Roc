import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Myntra Reconciliation",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #0f172a;
    border-right: 1px solid #1e293b;
}
section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
section[data-testid="stSidebar"] .stFileUploader label { color: #94a3b8 !important; font-size: 0.78rem; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #f8fafc !important; }

/* Main background */
.main .block-container { padding-top: 1.5rem; max-width: 1400px; }

/* Metric cards */
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 1rem 1.25rem;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
}
.metric-label { font-size: 0.72rem; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: .06em; margin-bottom: .2rem; }
.metric-value { font-size: 1.55rem; font-weight: 700; color: #0f172a; font-family: 'JetBrains Mono', monospace; }
.metric-value.positive { color: #16a34a; }
.metric-value.negative { color: #dc2626; }
.metric-sub { font-size: 0.72rem; color: #94a3b8; margin-top: .15rem; }

/* Section header */
.section-header {
    display: flex; align-items: center; gap: .6rem;
    font-size: 1rem; font-weight: 700; color: #0f172a;
    border-bottom: 2px solid #e2e8f0; padding-bottom: .5rem; margin-bottom: 1rem;
}

/* Brand pill */
.brand-pill {
    display: inline-block; padding: .15rem .55rem;
    border-radius: 999px; font-size: 0.72rem; font-weight: 600;
    background: #dbeafe; color: #1d4ed8;
}

/* Status badge */
.status-SH  { background:#dcfce7; color:#15803d; padding:.1rem .4rem; border-radius:4px; font-size:.7rem; font-weight:700; }
.status-WP  { background:#fef9c3; color:#92400e; padding:.1rem .4rem; border-radius:4px; font-size:.7rem; font-weight:700; }
.status-PK  { background:#e0f2fe; color:#0369a1; padding:.1rem .4rem; border-radius:4px; font-size:.7rem; font-weight:700; }
.status-F   { background:#f3e8ff; color:#7e22ce; padding:.1rem .4rem; border-radius:4px; font-size:.7rem; font-weight:700; }
.status-C   { background:#fee2e2; color:#b91c1c; padding:.1rem .4rem; border-radius:4px; font-size:.7rem; font-weight:700; }

/* Upload zone */
.upload-hint { background:#1e293b; border-radius:8px; padding:.75rem 1rem; font-size:.78rem; color:#94a3b8; margin-top:.5rem; line-height:1.6; }

/* Info box */
.info-box { background:#eff6ff; border-left:3px solid #3b82f6; border-radius:0 8px 8px 0; padding:.7rem 1rem; font-size:.82rem; color:#1e40af; }
.warn-box  { background:#fffbeb; border-left:3px solid #f59e0b; border-radius:0 8px 8px 0; padding:.7rem 1rem; font-size:.82rem; color:#92400e; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  CORE RECONCILIATION LOGIC
# ══════════════════════════════════════════════════════════════════════════════

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
    """
    Parse Slab.xlsx → rates DataFrame + sku_map dict.

    FIX: the previous version did
        rates = pd.DataFrame(rows[1:], columns=rows[0])
    which throws
        "Length mismatch: Expected axis has N elements, new values have M elements"
    whenever the header row and a data row don't have the exact same
    number of cells (e.g. a trailing blank column read by openpyxl, or a
    stray value past the last header). We now build the header safely,
    drop fully-empty trailing columns, and pad/truncate every data row to
    match the header length instead of letting pandas raise.
    """
    wb = __import__("openpyxl").load_workbook(BytesIO(file_bytes), data_only=True)

    # ---- Rates sheet -----------------------------------------------------
    ws = wb["Rates"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        raise ValueError("'Rates' sheet is empty.")

    header = list(rows[0])

    # Drop trailing columns whose header is blank/None — these are usually
    # stray formatting artifacts from Excel and are the #1 cause of the
    # "Length mismatch" error.
    while header and (header[-1] is None or str(header[-1]).strip() == ""):
        header.pop()

    n_cols = len(header)
    if n_cols == 0:
        raise ValueError("Could not find a valid header row in the 'Rates' sheet.")

    data_rows = []
    for r in rows[1:]:
        r = list(r)
        if len(r) < n_cols:
            r = r + [None] * (n_cols - len(r))   # pad short rows
        elif len(r) > n_cols:
            r = r[:n_cols]                        # truncate long rows
        # skip fully blank rows
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        data_rows.append(r)

    rates = pd.DataFrame(data_rows, columns=header)
    rates.columns = [str(c).strip() for c in rates.columns]

    missing = [c for c in REQUIRED_RATE_COLS if c not in rates.columns]
    if missing:
        raise ValueError(
            "'Rates' sheet is missing required column(s): "
            f"{', '.join(missing)}. Found columns: {list(rates.columns)}"
        )

    rates["Brand Name"] = rates["Brand Name"].astype(str).str.strip()
    rates["Category"]   = rates["Category"].astype(str).str.strip()

    # numeric coercion so a stray text cell can't silently break the lookups
    numeric_cols = [
        "Lower Limit Commision", "Upper Limit Commision", "Commision Charge",
        "GT Lower Limit", "GT Upper Limit", "GT Charges",
        "Lower Limit Fixed Fee", "Upper Limit Fixed Fee", "Fix Fee",
    ]
    for c in numeric_cols:
        rates[c] = pd.to_numeric(rates[c], errors="coerce")

    # ---- Replace Sku sheet — MYNTRA SKU CODE → OMS SKU CODE --------------
    # Columns: DATE, YRN NUMBER, MYNTRA SKU CODE, STYLE ID, OMS SKU CODE, BRAND
    # NOTE: the orders CSV's "seller sku code" column is what matches
    # "MYNTRA SKU CODE" here (NOT the CSV's "myntra sku code" column,
    # which instead matches "YRN NUMBER" — verified against real data).
    oms_map = {}
    if "Replace Sku" in wb.sheetnames:
        ws2 = wb["Replace Sku"]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * max(0, 5 - len(row))  # defensive pad
            if row[2] and row[4]:                     # MYNTRA SKU CODE → OMS SKU CODE
                oms_map[str(row[2]).strip()] = str(row[4]).strip()

    # ---- "Price We Need Excel" sheet — OMS Child SKU → PWN+10% ----------
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
            if row[1] and row[2] is not None:          # OMS Child SKU → PWN+10%
                pwn_map[str(row[1]).strip()] = row[2]

    # ---- "Closed" sheet — OMS Child SKU → Closed Price (fallback) -------
    closed_map = {}
    if "Closed" in wb.sheetnames:
        ws4 = wb["Closed"]
        for row in ws4.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * max(0, 3 - len(row))
            if row[1] and row[2] is not None:           # OMS Child SKU → Closed Price
                closed_map[str(row[1]).strip()] = row[2]

    return rates, oms_map, pwn_map, closed_map


def get_pwn_price(oms_map: dict, pwn_map: dict, closed_map: dict, seller_sku_code: str):
    """
    Resolve the PWN+10% price for an order row, following the chain:
      seller sku code  --[Replace Sku: MYNTRA SKU CODE→OMS SKU CODE]-->  oms_sku
      oms_sku --[Closed: OMS Child SKU→Closed Price]--> price            (preferred —
          a SKU listed in Closed means it's no longer active, so its Closed
          Price is the correct/current one even if it still also appears,
          possibly stale, in Price We Need Excel)
      oms_sku --[Price We Need Excel: OMS Child SKU→PWN+10%]--> price    (fallback)
    Returns (price_or_None, oms_sku_or_None, source) where source is
    "PWN", "Closed", or None.
    """
    key = str(seller_sku_code).strip()
    oms_sku = oms_map.get(key)
    if not oms_sku:
        return None, None, None

    if oms_sku in closed_map:
        return closed_map[oms_sku], oms_sku, "Closed"
    if oms_sku in pwn_map:
        return pwn_map[oms_sku], oms_sku, "PWN"
    return None, oms_sku, None


def _lookup(subset: pd.DataFrame, lo_col: str, hi_col: str, val: float):
    """Find the slab row where lo_col <= val < hi_col."""
    m = subset[(subset[lo_col] <= val) & (subset[hi_col] > val)]
    if m.empty:
        m = subset[(subset[lo_col] <= val) & (subset[hi_col] >= val)]
    return m.iloc[0] if not m.empty else None


def get_charges(rates: pd.DataFrame, brand: str, cat: str, SP: float):
    """
    Returns (GT, V, comm_rate, comm_amt, fixed_fee) or all None on miss.

    Formula (matches reference workbook exactly):
      1. GT      = slab lookup by SP
      2. V       = SP − GT
      3. comm    = rate(V) × V     ← rate looked up by V, not SP
      4. fixed   = slab lookup by SP
    """
    sub = rates[
        (rates["Brand Name"] == brand) &
        (rates["Category"].str.lower() == cat.lower())
    ]
    if sub.empty:
        return None, None, None, None, None

    gt_row  = _lookup(sub, "GT Lower Limit",           "GT Upper Limit",           SP)
    fee_row = _lookup(sub, "Lower Limit Fixed Fee",    "Upper Limit Fixed Fee",    SP)
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


def reconcile(rates: pd.DataFrame, df: pd.DataFrame, oms_map: dict, pwn_map: dict,
              closed_map: dict, brand_filter=None):
    """
    Run reconciliation on a Myntra seller-orders CSV DataFrame.
    Returns enriched DataFrame with all reconciliation columns, including
    the resolved PWN+10% price (_pwn), the OMS SKU it was resolved through
    (_oms_sku), and where the price came from (_pwn_source: "PWN"/"Closed"/None).
    """
    df = df.copy()
    df["brand"]        = df["brand"].astype(str).str.strip()
    df["article type"] = df["article type"].astype(str).str.strip()

    if brand_filter:
        df = df[df["brand"].isin(brand_filter)].copy()

    records = []
    for _, row in df.iterrows():
        SP   = float(row["final amount"])
        brand = row["brand"]
        cat   = row["article type"]

        GT, V, comm_rate, comm_amt, fixed_fee = get_charges(rates, brand, cat, SP)
        pwn_price, oms_sku, pwn_source = get_pwn_price(
            oms_map, pwn_map, closed_map, row.get("seller sku code", "")
        )

        if GT is None:
            rec = dict(
                _GT=None, _V=None, _comm_rate=None, _comm_amt=None,
                _fixed_fee=None, _total_charges=None, _gst=None,
                _myntra_payable=None, _marketing=None, _royalty=None,
                _slab_ok=False,
            )
        else:
            total_ch  = round(comm_amt + GT + fixed_fee, 2)
            gst       = round((total_ch - GT) * GST_RATE, 2)
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


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT  (exact reference-workbook column layout)
# ══════════════════════════════════════════════════════════════════════════════

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
_Y  = PatternFill("solid", fgColor="FCE4D6")   # orange = user fills
_B  = PatternFill("solid", fgColor="DEEAF1")
_G  = PatternFill("solid", fgColor="E2EFDA")
_A1 = PatternFill("solid", fgColor="EBF3FA")
_A2 = PatternFill("solid", fgColor="FFFFFF")
_thin = Side(style="thin", color="B8CCE4")
_TB = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _cell(ws, r, c, val, fill, font=None, fmt=None, align="center"):
    cell = ws.cell(row=r, column=c, value=val)
    cell.fill   = fill
    cell.border = _TB
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if font:  cell.font   = font
    if fmt:   cell.number_format = fmt
    return cell


def build_excel(result_df: pd.DataFrame) -> bytes:
    wb  = Workbook()
    all_brands = sorted(result_df["brand"].unique())

    # ── one sheet per brand ──────────────────────────────────────────────────
    for brand in all_brands:
        bdf = result_df[result_df["brand"] == brand].reset_index(drop=True)
        ws  = wb.create_sheet(title=brand[:31])

        hfont = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
        nfont = Font(size=9, name="Calibri")
        bfont = Font(bold=True, size=9, name="Calibri")
        ofont = Font(size=9, name="Calibri", italic=True, color="C55A11")
        gfont = Font(bold=True, size=9, name="Calibri", color="375623")

        # Header
        for ci, (h, w) in enumerate(zip(HEADERS, COL_W), 1):
            _cell(ws, 1, ci, h, _H, hfont)
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 36

        # Data rows
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
                row.get("_oms_sku", ""),           # F: Original SKU (resolved OMS SKU)
                row.get("article type", ""),
                row.get("total mrp", ""),
                row.get("_pwn"),                    # I: PWN+10% – auto-filled from Slab lookups
                row.get("_marketing"),             # J
                RETURN_CHARGES,                    # K
                row.get("_royalty"),               # L
                None,                              # M: Total Amount (formula)
                row.get("_comm_amt"),              # N
                row.get("_GT"),                    # O
                row.get("_fixed_fee"),             # P
                row.get("_total_charges"),         # Q
                row.get("_gst"),                   # R
                row.get("_myntra_payable"),        # S
                0,                                 # T: Rebate
                None,                              # U: Difference (formula)
                SP,                                # V
                V,                                 # W
                "25-Jun-2026",                     # X
                row.get("brand", ""),              # Y
                row.get("order status", ""),       # Z
                row.get("_pwn_source") or "Manual", # AA: Price Source
            ]

            num_cols = {10,11,12,13,14,15,16,17,18,19,20,21,22,23}
            for ci, val in enumerate(vals, 1):
                if ci == 9:                        # PWN+10% – orange if missing, normal if auto-filled
                    fill_h = alt if val is not None else _Y
                    font_h = nfont if val is not None else ofont
                    _cell(ws, ri, ci, val, fill_h, font_h,
                          fmt="#,##0.00" if val is not None else None)
                elif ci == 19:                     # Myntra Payable – green
                    _cell(ws, ri, ci, val, _G, gfont,
                          fmt="#,##0.00" if val is not None else None)
                elif ci == 21:                     # Difference (formula)
                    _cell(ws, ri, ci, None, _B, nfont)
                elif ci == 13:                     # Total Amount (formula)
                    _cell(ws, ri, ci, None, alt, nfont)
                else:
                    fmt = "#,##0.00" if ci in num_cols else (
                          "#,##0"    if ci == 8 else None)
                    _cell(ws, ri, ci, val, alt, nfont, fmt)

            # Excel formulas for Total Amount (M) and Difference (U)
            m = ws.cell(row=ri, column=13)
            m.value = f'=IF(I{ri}="","",K{ri}+J{ri}+I{ri}+L{ri})'
            m.number_format = "#,##0.00"

            u = ws.cell(row=ri, column=21)
            u.value = f'=IF(I{ri}="","",S{ri}-M{ri}+T{ri})'
            u.number_format = "#,##0.00"

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

        # Note row
        nr = len(bdf) + 3
        note = ws.cell(row=nr, column=1,
            value="📌 PWN+10% (Column H) auto-filled via Replace Sku → Price We Need / Closed lookup. "
                  "Orange cells = no SKU/price match found — fill manually.")
        note.font  = Font(bold=True, italic=True, size=9, color="C55A11", name="Calibri")
        note.fill  = PatternFill("solid", fgColor="FFF2CC")
        ws.merge_cells(f"A{nr}:{get_column_letter(len(HEADERS))}{nr}")

    # ── Summary sheet ────────────────────────────────────────────────────────
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
            _cell(ws_sum, ri, ci, v, alt,
                  bf if ci == 1 else nf, fmt)

        for k, v in zip(["n","sp","gt","comm","ff","tc","gst","pay"],
                         [n, sp, gt, cm, ff, tc, gst, pay]):
            grand[k] += v

    # Grand total
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

    # Remove default sheet if present
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🧾 Myntra Recon")
    st.markdown("---")

    st.markdown("**① Upload Slab File**")
    slab_file = st.file_uploader(
        "Slab.xlsx  (Rates + Replace Sku sheets)",
        type=["xlsx"], key="slab",
        help="Must contain a 'Rates' sheet with Brand Name, Category, slab limits."
    )

    st.markdown("**② Upload Orders CSV**")
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
• GT = slab lookup by SP<br>
• V = SP − GT<br>
• Commission = rate(V) × V<br>
• Fixed Fee = slab lookup by SP<br>
• Total Charges = Comm + GT + Fee<br>
• GST = (Total − GT) × 18%<br>
• Myntra Payable = SP − TC − GST<br>
• Marketing = SP × 3%<br>
• Royalty = V × 1%
</div>
""", unsafe_allow_html=True)


# ── Main area ─────────────────────────────────────────────────────────────────
st.markdown("## Myntra Seller Reconciliation")
st.markdown("Upload your **Slab** and **Orders CSV** in the sidebar to begin.")

if not slab_file or not csv_file:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
<div class="info-box">
📂 <b>Step 1 — Slab.xlsx</b><br>
Your commission/GT/fixed-fee rate table. Must have a <code>Rates</code> sheet with:<br>
Brand Name · Category · slab limits (commission, GT, fixed fee)
</div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
<div class="info-box">
📄 <b>Step 2 — Orders CSV</b><br>
Myntra Seller Orders export. Key columns needed:<br>
<code>brand</code> · <code>article type</code> · <code>final amount</code> · <code>seller price</code>
</div>""", unsafe_allow_html=True)
    st.stop()


# ── Load data ─────────────────────────────────────────────────────────────────
with st.spinner("Loading slab rates…"):
    try:
        rates, oms_map, pwn_map, closed_map = load_slab(slab_file.read())
    except Exception as e:
        st.error(f"❌ Could not read Slab file: {e}")
        st.stop()

with st.spinner("Reading orders…"):
    try:
        orders_df = pd.read_csv(csv_file)
        orders_df["brand"]        = orders_df["brand"].astype(str).str.strip()
        orders_df["article type"] = orders_df["article type"].astype(str).str.strip()
        orders_df["final amount"] = pd.to_numeric(orders_df["final amount"], errors="coerce").fillna(0)
        orders_df["seller price"] = pd.to_numeric(orders_df["seller price"], errors="coerce").fillna(0)
    except Exception as e:
        st.error(f"❌ Could not read CSV: {e}")
        st.stop()


# ── Run reconciliation ────────────────────────────────────────────────────────
with st.spinner("Reconciling…"):
    result = reconcile(rates, orders_df, oms_map, pwn_map, closed_map,
                        brand_filter if brand_filter else None)

ok_df  = result[result["_slab_ok"] == True]
bad_df = result[result["_slab_ok"] == False]


# ── KPI row ───────────────────────────────────────────────────────────────────
total_orders   = len(result)
total_sp       = result["final amount"].sum()
total_payable  = ok_df["_myntra_payable"].sum()
total_charges  = ok_df["_total_charges"].sum()
total_gst      = ok_df["_gst"].sum()
unmatched      = len(bad_df)
pwn_found      = result["_pwn"].notna().sum()
pwn_missing    = total_orders - pwn_found

cols = st.columns(6)
kpi_data = [
    ("Total Orders",         f"{total_orders:,}",          None,       "orders in report"),
    ("Total Selling Price",  f"₹{total_sp:,.0f}",          None,       "sum of final amount"),
    ("Total Myntra Payable", f"₹{total_payable:,.0f}",     None,       "after all deductions"),
    ("Total Charges + GST",  f"₹{total_charges+total_gst:,.0f}", None, "platform fees"),
    ("PWN Price Matched",    f"{pwn_found}/{total_orders}",
     "positive" if pwn_missing == 0 else "negative",       "via Replace Sku → Price We Need/Closed"),
    ("Unmatched Rows",       f"{unmatched}",
     "negative" if unmatched > 0 else "positive",          "no slab found"),
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


# ── Brand summary table ───────────────────────────────────────────────────────
st.markdown('<div class="section-header">📊 Brand Summary</div>', unsafe_allow_html=True)

summary_rows = []
for brand in sorted(result["brand"].unique()):
    bdf  = result[result["brand"] == brand]
    bok  = bdf[bdf["_slab_ok"] == True]
    summary_rows.append({
        "Brand":                brand,
        "Orders":               len(bdf),
        "Selling Price (₹)":    round(bdf["final amount"].sum(), 2),
        "GT Charges (₹)":       round(bok["_GT"].sum(), 2),
        "Commission (₹)":       round(bok["_comm_amt"].sum(), 2),
        "Fixed Fee (₹)":        round(bok["_fixed_fee"].sum(), 2),
        "Total Charges (₹)":    round(bok["_total_charges"].sum(), 2),
        "GST (₹)":              round(bok["_gst"].sum(), 2),
        "Myntra Payable (₹)":   round(bok["_myntra_payable"].sum(), 2),
    })

sum_df = pd.DataFrame(summary_rows)
st.dataframe(
    sum_df.style
        .format({c: "₹{:,.2f}" for c in sum_df.columns if "₹" in c})
        .set_properties(**{"text-align": "center"}),
    use_container_width=True, hide_index=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# ── Per-brand detail tabs ─────────────────────────────────────────────────────
brands_in_result = sorted(result["brand"].unique().tolist())

if brands_in_result:
    st.markdown('<div class="section-header">📋 Order-Level Detail</div>', unsafe_allow_html=True)
    tabs = st.tabs([f"  {b}  " for b in brands_in_result] + ["⚠️ Unmatched"])

    DISPLAY_COLS = {
        "order release id":  "Order ID",
        "order line id":     "Line ID",
        "seller sku code":   "SKU",
        "_oms_sku":          "Original SKU",
        "article type":      "Category",
        "order status":      "Status",
        "final amount":      "Selling Price ₹",
        "_pwn":              "PWN+10% ₹",
        "_pwn_source":       "Price Source",
        "_V":                "SP − GT ₹",
        "_comm_amt":         "Commission ₹",
        "_GT":               "GT Charges ₹",
        "_fixed_fee":        "Fixed Fee ₹",
        "_total_charges":    "Total Charges ₹",
        "_gst":              "GST ₹",
        "_myntra_payable":   "Myntra Payable ₹",
        "_marketing":        "Marketing 3% ₹",
        "_royalty":          "Royalty 1% ₹",
    }

    for tab, brand in zip(tabs[:-1], brands_in_result):
        with tab:
            bdf = result[result["brand"] == brand].copy()
            ok  = bdf[bdf["_slab_ok"] == True]
            mis = bdf[bdf["_slab_ok"] == False]

            # mini KPIs
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Orders", len(bdf))
            with c2:
                st.metric("Selling Price", f"₹{bdf['final amount'].sum():,.0f}")
            with c3:
                st.metric("Myntra Payable", f"₹{ok['_myntra_payable'].sum():,.0f}")
            with c4:
                st.metric("Unmatched", len(mis),
                          delta=f"{len(mis)} rows" if len(mis) else None,
                          delta_color="inverse")

            display_df = ok[list(DISPLAY_COLS.keys())].rename(columns=DISPLAY_COLS)
            num_cols_disp = [v for k, v in DISPLAY_COLS.items()
                             if k.startswith("_") or k in ("final amount",)]

            st.dataframe(
    display_df.style
        .format({c: "₹{:,.2f}" for c in num_cols_disp if c in display_df.columns}, na_rep="")
        .map(
                        lambda v: "background-color:#f0fdf4;color:#166534;font-weight:600"
                        if isinstance(v, (int, float)) and v > 0 else "",
                        subset=["Myntra Payable ₹"] if "Myntra Payable ₹" in display_df.columns else []
                    ),
                use_container_width=True,
                hide_index=True,
                height=min(400, 45 + 35 * len(display_df)),
            )

            if len(mis) > 0:
                st.markdown(f"""
<div class="warn-box">⚠️ {len(mis)} rows had no matching slab (brand+category not found in Rates sheet).
Check brand name spelling and that the category exists in the Slab file.</div>""",
                    unsafe_allow_html=True)
                st.dataframe(
                    mis[["seller sku code", "article type", "final amount", "order status"]],
                    use_container_width=True, hide_index=True,
                )

    with tabs[-1]:
        if bad_df.empty:
            st.success("✅ All rows matched a slab — no unmatched orders.")
        else:
            st.warning(f"{len(bad_df)} orders could not be matched to a slab.")
            st.dataframe(
                bad_df[["brand", "seller sku code", "article type",
                        "final amount", "order status"]],
                use_container_width=True, hide_index=True,
            )


# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">⬇️ Export</div>', unsafe_allow_html=True)

col_dl, col_info = st.columns([1, 3])
with col_dl:
    with st.spinner("Preparing Excel…"):
        xlsx_bytes = build_excel(result)
    st.download_button(
        label="📥 Download Reconciliation (.xlsx)",
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
<b>PWN+10% (col H)</b> is left blank — fill it to auto-calculate Total Amount & Difference.
</div>""", unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)
