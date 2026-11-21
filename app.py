import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from io import BytesIO
import re

# --- CẤU HÌNH GIAO DIỆN WEB (TRẮNG - XANH NAVY - VÀNG GOLD) ---
st.set_page_config(page_title="Dashboard Báo Cáo Tuần TMĐT", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main { background-color: #FFFFFF; }
    h1, h2, h3 { color: #0F2C59; font-family: 'Segoe UI', sans-serif; }
    .kpi-card {
        background-color: #F8FAFC;
        border-left: 5px solid #0F2C59;
        border-top: 1px solid #E2E8F0;
        border-right: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
        padding: 13px 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        margin-bottom: 10px;
    }
    .kpi-card-shopee {
        background-color: #FFFDF7;
        border-left: 5px solid #D4AF37;
        border-top: 1px solid #FDE68A;
        border-right: 1px solid #FDE68A;
        border-bottom: 1px solid #FDE68A;
        padding: 13px 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        margin-bottom: 10px;
    }
    .kpi-card-tiktok {
        background-color: #F4F7FB;
        border-left: 5px solid #1E3A8A;
        border-top: 1px solid #DBEAFE;
        border-right: 1px solid #DBEAFE;
        border-bottom: 1px solid #DBEAFE;
        padding: 13px 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        margin-bottom: 10px;
    }
    .kpi-label { color: #475569; font-weight: 600; font-size: 0.85rem; margin-bottom: 4px; }
    .kpi-value { color: #0F2C59; font-weight: 700; font-size: 1.40rem; line-height: 1.25; margin-bottom: 6px; }
    .kpi-sub { font-size: 0.78rem; color: #64748B; display: flex; align-items: center; flex-wrap: wrap; gap: 6px; }
    .badge-up { background-color: #DCFCE7; color: #15803D; padding: 1px 7px; border-radius: 12px; font-weight: 600; font-size: 0.74rem; }
    .badge-down { background-color: #FEE2E2; color: #B91C1C; padding: 1px 7px; border-radius: 12px; font-weight: 600; font-size: 0.74rem; }
    .badge-neutral { background-color: #E2E8F0; color: #475569; padding: 1px 7px; border-radius: 12px; font-weight: 600; font-size: 0.74rem; }
</style>
""", unsafe_allow_html=True)

def render_kpi_card(label, curr_str, prev_str=None, pct_change=None, inverse_color=False, fallback_sub="", card_style="kpi-card"):
    sub_html = ""
    if prev_str is not None:
        badge_html = ""
        if pct_change is not None:
            is_positive = pct_change > 0
            if inverse_color:
                badge_cls = "badge-down" if is_positive else ("badge-up" if pct_change < 0 else "badge-neutral")
            else:
                badge_cls = "badge-up" if is_positive else ("badge-down" if pct_change < 0 else "badge-neutral")
            arrow = "↑" if pct_change > 0 else ("↓" if pct_change < 0 else "→")
            badge_html = f'<span class="{badge_cls}">{arrow} {pct_change:+.1f}%</span>'
        sub_html = f'<div class="kpi-sub"><span>Tuần trước: <b>{prev_str}</b></span> {badge_html}</div>'
    elif fallback_sub:
        sub_html = f'<div class="kpi-sub"><span class="badge-neutral">{fallback_sub}</span></div>'
    else:
        sub_html = '<div class="kpi-sub">&nbsp;</div>'

    st.markdown(f"""
    <div class="{card_style}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{curr_str}</div>
        {sub_html}
    </div>
    """, unsafe_allow_html=True)

# --- BỘ ĐỌC FILE ĐA NĂNG ---
HEADER_KEYWORDS = [
    'mã đơn hàng', 'order id', 'id đơn hàng',
    'trạng thái đơn hàng', 'order status',
    'tên sản phẩm', 'product name', 'số lượng', 'quantity'
]

def clean_col_name(c):
    if c is None:
        return ""
    return re.sub(r'\s+', ' ', str(c)).strip()

def detect_header_and_clean(df_raw):
    curr_cols_lower = [clean_col_name(c).lower() for c in df_raw.columns]
    if any(k in curr_cols_lower for k in HEADER_KEYWORDS):
        df_raw.columns = [clean_col_name(c) for c in df_raw.columns]
        return df_raw

    for idx in range(min(15, len(df_raw))):
        row_vals = [clean_col_name(v).lower() for v in df_raw.iloc[idx].tolist()]
        if sum(1 for k in HEADER_KEYWORDS if k in row_vals) >= 2:
            new_cols = [clean_col_name(v) for v in df_raw.iloc[idx].tolist()]
            df_clean = df_raw.iloc[idx + 1:].copy()
            df_clean.columns = new_cols
            return df_clean.reset_index(drop=True)

    df_raw.columns = [clean_col_name(c) for c in df_raw.columns]
    return df_raw

def read_excel_full_cells(uploaded_file):
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.getvalue()), data_only=True)
    best_df = None
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.values)
        if not rows:
            continue
        max_col = 0
        for r in rows[:20]:
            for idx, val in enumerate(r):
                if val is not None and str(val).strip() != "":
                    max_col = max(max_col, idx + 1)
        if max_col == 0:
            continue
        trimmed_rows = [list(r[:max_col]) for r in rows]
        df_temp = pd.DataFrame(trimmed_rows[1:], columns=[clean_col_name(c) for c in trimmed_rows[0]])
        df_checked = detect_header_and_clean(df_temp)
        cols_lower = [c.lower() for c in df_checked.columns]
        if any(k in cols_lower for k in HEADER_KEYWORDS):
            best_df = df_checked
            break
        if best_df is None:
            best_df = df_checked
    wb.close()
    return best_df if best_df is not None else pd.DataFrame()

def smart_read_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.csv') or name.endswith('.txt'):
        raw_bytes = uploaded_file.getvalue()
        for enc in ['utf-8-sig', 'utf-8', 'utf-16', 'cp1258', 'latin1']:
            try:
                text = raw_bytes.decode(enc)
                first_line = text.splitlines()[0] if text.splitlines() else ""
                sep = '\t' if '\t' in first_line else (';' if first_line.count(';') > first_line.count(',') else ',')
                df = pd.read_csv(BytesIO(raw_bytes), encoding=enc, sep=sep, dtype=str)
                return detect_header_and_clean(df)
            except Exception:
                continue
        raise ValueError(f"Không thể giải mã file CSV: {uploaded_file.name}")
    elif name.endswith('.xlsx'):
        return read_excel_full_cells(uploaded_file)
    else:
        xls = pd.ExcelFile(uploaded_file)
        df_first = pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=str)
        return detect_header_and_clean(df_first)

def find_col(df, exact_case_candidates, exact_insensitive_candidates=None, partial_candidates=None):
    for cand in exact_case_candidates:
        if cand in df.columns:
            return cand

    all_cands = exact_case_candidates + (exact_insensitive_candidates or [])
    for cand in all_cands:
        for col in df.columns:
            if col.lower() == cand.lower():
                return col

    if partial_candidates:
        for p in partial_candidates:
            for col in df.columns:
                if p.lower() in col.lower():
                    return col
    return None

def parse_number_universal(series):
    def _conv(val):
        if val is None or pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip()
        s = re.sub(r'(?i)(vnd|vnđ|đ|₫|\s)', '', s)
        if s in ['', '-', 'nan', 'none', 'null']:
            return 0.0
        if ',' in s and '.' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        elif ',' in s:
            parts = s.split(',')
            if len(parts[-1]) == 3 and len(parts) > 1:
                s = s.replace(',', '')
            else:
                s = s.replace(',', '.')
        elif '.' in s:
            if s.count('.') > 1:
                s = s.replace('.', '')
        try:
            return float(s)
        except ValueError:
            return 0.0

    return series.apply(_conv)

# --- QUY TẮC PHÂN LOẠI 5 NHÓM TRẠNG THÁI ---
def classify_order_status(platform, raw_status, sub_status="", return_type="", ret_qty=0):
    st_val = str(raw_status).strip().lower()
    sub_val = str(sub_status).strip().lower()
    ret_val = str(return_type).strip().lower()

    if platform == 'TikTok Shop':
        if any(k in ret_val for k in ['return', 'refund', 'hoàn', 'trả']) or ret_qty > 0:
            return 'Trả hàng / Hoàn tiền'
        if 'cancel' in ret_val or any(k in st_val for k in ['hủy', 'cancel']):
            return 'Đơn hủy'
        if any(k in st_val for k in ['đã hoàn tất', 'completed']) or any(k in sub_val for k in ['đã giao', 'delivered']):
            return 'Giao thành công'
        if any(k in st_val for k in ['cần vận chuyển', 'chưa thanh toán', 'to ship', 'unpaid']) or any(k in sub_val for k in ['chờ', 'awaiting']):
            return 'Đang chờ giao hàng'
        if any(k in st_val for k in ['đã vận chuyển', 'shipped']) or any(k in sub_val for k in ['trung chuyển', 'transit']):
            return 'Đang vận chuyển'
        return 'Đang vận chuyển'

    else:
        if (ret_val not in ['', 'nan', 'none', 'null'] and any(k in ret_val for k in [
            'đã chấp thuận', 'đã duyệt', 'yêu cầu chờ xử lý', 'đang xử lý',
            'đã hoàn tiền', 'hoàn tiền thành công', 'refunded', 'returned'
        ])) or ret_qty > 0:
            return 'Trả hàng / Hoàn tiền'
        if any(k in st_val for k in ['hủy', 'cancel', 'thất bại', 'failed', 'không thành công', 'đã đóng']):
            return 'Đơn hủy'
        if any(k in st_val for k in ['hoàn thành', 'đã giao', 'đã nhận được hàng', 'completed', 'delivered']):
            return 'Giao thành công'
        if any(k in st_val for k in ['trả hàng', 'hoàn tiền', 'return', 'refund', 'khiếu nại']):
            return 'Trả hàng / Hoàn tiền'
        if any(k in st_val for k in ['chờ lấy', 'chờ giao', 'chờ xác nhận', 'awaiting', 'to ship', 'unpaid', 'đang chuẩn bị', 'chờ thanh toán']):
            return 'Đang chờ giao hàng'
        if any(k in st_val for k in ['đang giao', 'vận chuyển', 'transit', 'shipped', 'đã bàn giao']):
            return 'Đang vận chuyển'
        return 'Đang vận chuyển'

# --- HÀM CHUẨN HÓA DỮ LIỆU CHUNG ---
def parse_universal_orders(uploaded_file, hint_platform):
    df = smart_read_file(uploaded_file)
    cols_lower = [c.lower() for c in df.columns]

    if any(k in cols_lower for k in ['mã kiện hàng', 'được shopee trợ giá', 'trạng thái trả hàng/hoàn tiền']):
        platform = 'Shopee'
    elif any(k in cols_lower for k in [
        'sku subtotal after discount', 'tổng phụ sku sau giảm giá',
        'cancelation/return type', 'loại hủy/trả hàng',
        'order substatus', 'trạng thái phụ', 'seller sku', 'sku người bán'
    ]):
        platform = 'TikTok Shop'
    else:
        platform = hint_platform

    col_order = find_col(df, ['Mã đơn hàng', 'Order ID', 'ID đơn hàng'], partial_candidates=['mã đơn', 'order id', 'id đơn'])
    col_status = find_col(df, ['Trạng Thái Đơn Hàng', 'Trạng thái đơn hàng', 'Order Status'], partial_candidates=['trạng thái đơn', 'order status'])
    col_substatus = find_col(df, ['Order Substatus', 'Trạng thái phụ'], partial_candidates=['substatus', 'trạng thái phụ'])
    col_return_type = find_col(df, [
        'Trạng thái Trả hàng/Hoàn tiền', 'Trạng thái trả hàng',
        'Cancelation/Return Type', 'Loại hủy/trả hàng'
    ], partial_candidates=['trả hàng/hoàn tiền', 'return type', 'hủy/trả'])
    col_ret_qty = find_col(df, ['Số lượng sản phẩm được hoàn trả', 'Sku Quantity of return', 'Số lượng SKU hoàn trả'])
    col_prod = find_col(df, ['Tên sản phẩm', 'Tên Sản Phẩm', 'Product Name'], partial_candidates=['tên sản phẩm', 'product name'])
    col_sku = find_col(df, ['Tên phân loại hàng', 'SKU phân loại hàng', 'Variation', 'Biến thể', 'Seller SKU', 'SKU người bán'], partial_candidates=['phân loại', 'variation', 'sku'])
    col_qty = find_col(df, ['Số lượng', 'Quantity'], partial_candidates=['số lượng', 'quantity'])

    col_reason = find_col(df, [
        'Lý do hủy', 'Lý do Hủy', 'Cancel Reason', 'Return Reason',
        'Lý do trả hàng', 'Lý do Trả hàng/Hoàn tiền', 'Cancel By', 'Người hủy'
    ], partial_candidates=['lý do', 'reason'])

    if col_order:
        df = df[df[col_order].notna()].copy()
        s_ord = df[col_order].astype(str).str.strip()
        df = df[
            (s_ord != '') &
            (~s_ord.str.lower().isin(['nan', 'none', 'null', 'tổng', 'total'])) &
            (~s_ord.str.lower().str.contains('order id|mã đơn|id đơn|platform unique', na=False))
        ].copy()

    out = pd.DataFrame()
    out['order_id'] = df[col_order].astype(str).str.strip() if col_order else df.index.astype(str)
    out['raw_status'] = df[col_status].fillna('').astype(str).str.strip() if col_status else 'Hoàn thành'
    sub_series = df[col_substatus].fillna('').astype(str).str.strip() if col_substatus else pd.Series([''] * len(df), index=df.index)
    ret_series = df[col_return_type].fillna('').astype(str).str.strip() if col_return_type else pd.Series([''] * len(df), index=df.index)
    ret_qty_series = parse_number_universal(df[col_ret_qty]) if col_ret_qty else pd.Series([0.0] * len(df), index=df.index)

    out['status_group'] = [
        classify_order_status(platform, st_v, sub_v, ret_v, rq_v)
        for st_v, sub_v, ret_v, rq_v in zip(out['raw_status'], sub_series, ret_series, ret_qty_series)
    ]
    out['product_name'] = df[col_prod].fillna(f'Sản phẩm {platform}').astype(str).str.strip() if col_prod else f'Sản phẩm {platform}'
    out['sku'] = df[col_sku].fillna('Mặc định').astype(str).str.strip() if col_sku else 'Mặc định'
    out['qty'] = parse_number_universal(df[col_qty]) if col_qty else 1

    if platform == 'Shopee':
        col_order_val = find_col(df, ['Tổng giá trị đơn hàng (VND)'], partial_candidates=['tổng giá trị đơn hàng'])
        col_item_val = find_col(df, ['Tổng số tiền Người mua thanh toán'], exact_insensitive_candidates=['Giá ưu đãi'])
        
        if col_order_val and col_item_val:
            order_val_s = parse_number_universal(df[col_order_val])
            item_val_s = parse_number_universal(df[col_item_val])
            out['_order_val'] = order_val_s
            out['_item_val'] = item_val_s
            order_item_sum = out.groupby('order_id')['_item_val'].transform('sum')
            order_row_count = out.groupby('order_id')['_item_val'].transform('count')
            out['revenue'] = [
                (ov * (iv / ois)) if ois > 0 else (ov / orc)
                for ov, iv, ois, orc in zip(out['_order_val'], out['_item_val'], order_item_sum, order_row_count)
            ]
            out.drop(columns=['_order_val', '_item_val'], inplace=True)
            col_rev_used = col_order_val
        elif col_item_val:
            out['revenue'] = parse_number_universal(df[col_item_val])
            col_rev_used = col_item_val
        else:
            out['revenue'] = 0.0
            col_rev_used = 'Không tìm thấy'
    else:
        col_rev = find_col(
            df,
            exact_case_candidates=['SKU Subtotal After Discount', 'Tổng phụ SKU sau giảm giá'],
            exact_insensitive_candidates=['SKU Subtotal Before Discount', 'Tổng phụ SKU trước giảm giá', 'Order Amount', 'Tổng số tiền đơn hàng'],
            partial_candidates=['sku subtotal after', 'tổng phụ sku']
        )
        out['revenue'] = parse_number_universal(df[col_rev]) if col_rev else 0.0
        col_rev_used = col_rev or 'Không tìm thấy'

    raw_reason = df[col_reason].replace(['', 'nan', 'None', 'null'], pd.NA) if col_reason else pd.Series([pd.NA] * len(df), index=df.index)
    out['reason'] = [
        str(r).strip() if pd.notna(r) else (f"Hoàn/Trả: {rt}" if rt not in ['', 'nan', 'None', 'null'] else 'Không rõ lý do')
        for r, rt in zip(raw_reason, ret_series)
    ]
    out['platform'] = platform
    out['is_valid'] = ~out['status_group'].isin(['Đơn hủy', 'Trả hàng / Hoàn tiền'])

    meta = {
        'file_name': uploaded_file.name,
        'platform': platform,
        'rows': len(out),
        'orders': out['order_id'].nunique(),
        'col_rev': col_rev_used,
        'col_status': col_status or 'Không tìm thấy'
    }
    return out, meta

def load_week_data(shopee_f, tiktok_f):
    dfs, metas = [], []
    for f, hint in [(shopee_f, 'Shopee'), (tiktok_f, 'TikTok Shop')]:
        if f is not None:
            df_parsed, meta = parse_universal_orders(f, hint)
            dfs.append(df_parsed)
            metas.append(meta)
    return (pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()), metas

# --- TÍNH TOÁN CHỈ SỐ TỔNG HỢP & TÁCH RIÊNG SHOPEE / TIKTOK ---
def calc_platform_kpi(df, ad_spend=0):
    if df.empty:
        return {'gmv': 0.0, 'tot_orders': 0, 'ad_spend': ad_spend, 'roi': 0.0, 'cpo': 0.0}
    gmv = df['revenue'].sum()
    tot_orders = df['order_id'].nunique()
    roi = (gmv / ad_spend) if ad_spend > 0 else 0.0
    cpo = (ad_spend / tot_orders) if (ad_spend > 0 and tot_orders > 0) else 0.0
    return {'gmv': gmv, 'tot_orders': tot_orders, 'ad_spend': ad_spend, 'roi': roi, 'cpo': cpo}

def calc_metrics(df, ads_sp=0, ads_tt=0):
    if df.empty:
        return None
    
    total_ads = ads_sp + ads_tt
    df_sp = df[df['platform'] == 'Shopee']
    df_tt = df[df['platform'] == 'TikTok Shop']
    
    valid = df[df['is_valid'] == True]
    invalid = df[df['is_valid'] == False]
    
    return {
        'all': calc_platform_kpi(df, total_ads),
        'Shopee': calc_platform_kpi(df_sp, ads_sp),
        'TikTok Shop': calc_platform_kpi(df_tt, ads_tt),
        'valid_df': valid,
        'invalid_df': invalid,
        'full_df': df
    }

def calc_pct(curr_val, prev_val):
    if prev_val == 0:
        return None
    return ((curr_val - prev_val) / prev_val) * 100

def render_4_kpi_row(curr_kpi, prev_kpi=None, card_style="kpi-card", prefix=""):
    c1, c2, c3, c4 = st.columns(4)
    has_prev = prev_kpi is not None and prev_kpi['tot_orders'] > 0
    with c1:
        render_kpi_card(
            label=f"{prefix}Tổng doanh thu ({curr_kpi['tot_orders']:,} đơn)",
            curr_str=f"{curr_kpi['gmv']:,.0f} đ",
            prev_str=f"{prev_kpi['gmv']:,.0f} đ ({prev_kpi['tot_orders']:,} đơn)" if has_prev else None,
            pct_change=calc_pct(curr_kpi['gmv'], prev_kpi['gmv']) if has_prev else None,
            fallback_sub=f"Tổng {curr_kpi['tot_orders']:,} đơn đặt",
            card_style=card_style
        )
    with c2:
        render_kpi_card(
            label=f"{prefix}Số tiền Ads",
            curr_str=f"{curr_kpi['ad_spend']:,.0f} đ",
            prev_str=f"{prev_kpi['ad_spend']:,.0f} đ" if prev_kpi is not None else None,
            pct_change=calc_pct(curr_kpi['ad_spend'], prev_kpi['ad_spend']) if (prev_kpi is not None and prev_kpi['ad_spend'] > 0) else None,
            inverse_color=True,
            fallback_sub="Nhập tiền Ads ở cột trái" if curr_kpi['ad_spend'] == 0 else "Ngân sách quảng cáo",
            card_style=card_style
        )
    with c3:
        render_kpi_card(
            label=f"{prefix}ROI (Doanh thu / Ads)",
            curr_str=f"{curr_kpi['roi']:.2f}" if curr_kpi['ad_spend'] > 0 else "Chưa nhập Ads",
            prev_str=f"{prev_kpi['roi']:.2f}" if (prev_kpi is not None and prev_kpi['ad_spend'] > 0) else ("Chưa nhập Ads" if prev_kpi is not None else None),
            pct_change=calc_pct(curr_kpi['roi'], prev_kpi['roi']) if (prev_kpi is not None and curr_kpi['ad_spend'] > 0 and prev_kpi['ad_spend'] > 0) else None,
            fallback_sub="Tỷ suất doanh thu trên Ads",
            card_style=card_style
        )
    with c4:
        render_kpi_card(
            label=f"{prefix}Chi phí cho 1 đơn hàng",
            curr_str=f"{curr_kpi['cpo']:,.0f} đ" if curr_kpi['ad_spend'] > 0 else "0 đ",
            prev_str=f"{prev_kpi['cpo']:,.0f} đ" if (prev_kpi is not None and prev_kpi['ad_spend'] > 0) else ("0 đ" if prev_kpi is not None else None),
            pct_change=calc_pct(curr_kpi['cpo'], prev_kpi['cpo']) if (prev_kpi is not None and curr_kpi['ad_spend'] > 0 and prev_kpi['ad_spend'] > 0) else None,
            inverse_color=True,
            fallback_sub="Tiền Ads / Tổng số đơn",
            card_style=card_style
        )

# --- HÀM XUẤT EXCEL DASHBOARD CHUYÊN NGHIỆP (TRẮNG - XANH NAVY - VÀNG GOLD) ---
def style_table_sheet(ws, df, title, subtitle, currency_cols=None, pct_cols=None, int_cols=None, growth_col=None):
    currency_cols = currency_cols or []
    pct_cols = pct_cols or []
    int_cols = int_cols or []

    ws.views.sheetView[0].showGridLines = True
    navy_fill = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    red_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    thin_side = Side(border_style="thin", color="CBD5E1")
    gold_bottom = Side(border_style="medium", color="D4AF37")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ncols = max(len(df.columns), 4)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c_title = ws.cell(1, 1, title)
    c_title.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    c_title.fill = navy_fill
    c_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c_sub = ws.cell(2, 1, subtitle)
    c_sub.font = Font(name="Segoe UI", size=9.5, italic=True, color="0F2C59")
    c_sub.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    c_sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22

    header_row = 4
    ws.row_dimensions[header_row].height = 26
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(header_row, col_idx, col_name)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=gold_bottom)

    for row_offset, (_, row) in enumerate(df.iterrows()):
        r_idx = header_row + 1 + row_offset
        ws.row_dimensions[r_idx].height = 22
        row_fill = zebra_fill if r_idx % 2 == 1 else white_fill
        for c_idx, col_name in enumerate(df.columns, start=1):
            val = row[col_name]
            cell = ws.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10, color="1E293B")
            cell.fill = row_fill
            cell.border = cell_border

            if col_name in currency_cols:
                cell.number_format = '#,##0 "đ"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in pct_cols:
                cell.number_format = '0.0"%"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in int_cols:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            if growth_col and col_name == growth_col and isinstance(val, (int, float)):
                cell.number_format = '+0.0"%";-0.0"%";0.0"%"'
                if val > 0:
                    cell.fill = green_fill
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="15803D")
                elif val < 0:
                    cell.fill = red_fill
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="B91C1C")

    for col_idx, col_name in enumerate(df.columns, start=1):
        max_len = len(str(col_name))
        for r_idx in range(header_row + 1, header_row + 1 + len(df)):
            val = ws.cell(r_idx, col_idx).value
            if val is not None:
                max_len = max(max_len, min(len(str(val)), 48))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(max_len + 5, 14)

def build_executive_excel(m_curr, m_prev, status_df, split_st, sku_merged, cancel_summary, wow_table):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "1_Tong_Quan_Dashboard"
    ws.views.sheetView[0].showGridLines = True

    navy_fill = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    sp_header_fill = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
    tt_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    
    card_bg_all = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    card_bg_sp = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    card_bg_tt = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    
    thin_gray = Side(border_style="thin", color="CBD5E1")
    thick_navy = Side(border_style="medium", color="0F2C59")
    thick_gold = Side(border_style="medium", color="D4AF37")
    thick_blue = Side(border_style="medium", color="1E3A8A")

    ws.merge_cells("A1:H1")
    c1 = ws["A1"]
    c1.value = "📊 BÁO CÁO HIỆU SUẤT KINH DOANH TUẦN — SHOPEE & TIKTOK SHOP"
    c1.font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    c1.fill = navy_fill
    c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:H2")
    c2 = ws["A2"]
    c2.value = "Tổng hợp doanh thu, ngân sách quảng cáo, tỷ suất ROI & chi phí trên mỗi đơn hàng"
    c2.font = Font(name="Segoe UI", size=10, italic=True, color="0F2C59")
    c2.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22

    def write_kpi_block(start_row, section_title, header_fill, bg_fill, accent_side, kpi_curr, kpi_prev):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=8)
        sec = ws.cell(start_row, 1, section_title)
        sec.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        sec.fill = header_fill
        sec.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[start_row].height = 25

        r_lbl = start_row + 1
        r_val = start_row + 2
        r_sub = start_row + 3
        ws.row_dimensions[r_lbl].height = 20
        ws.row_dimensions[r_val].height = 28
        ws.row_dimensions[r_sub].height = 19

        has_p = kpi_prev is not None and kpi_prev['tot_orders'] > 0
        cards = [
            (
                f"Tổng doanh thu ({kpi_curr['tot_orders']:,} đơn)",
                kpi_curr['gmv'],
                '#,##0 "đ"',
                f"Tuần trước: {kpi_prev['gmv']:,.0f} đ ({kpi_prev['tot_orders']:,} đơn)" if has_p else f"Tổng {kpi_curr['tot_orders']:,} đơn"
            ),
            (
                "Số tiền Ads",
                kpi_curr['ad_spend'],
                '#,##0 "đ"',
                f"Tuần trước: {kpi_prev['ad_spend']:,.0f} đ" if kpi_prev is not None else "Ngân sách quảng cáo"
            ),
            (
                "ROI (Doanh thu / Ads)",
                round(kpi_curr['roi'], 2) if kpi_curr['ad_spend'] > 0 else "Chưa nhập Ads",
                '0.00' if kpi_curr['ad_spend'] > 0 else '@',
                f"Tuần trước: {kpi_prev['roi']:.2f}" if (kpi_prev is not None and kpi_prev['ad_spend'] > 0) else "Tỷ suất Doanh thu / Ads"
            ),
            (
                "Chi phí cho 1 đơn hàng",
                round(kpi_curr['cpo'], 0) if kpi_curr['ad_spend'] > 0 else 0,
                '#,##0 "đ"',
                f"Tuần trước: {kpi_prev['cpo']:,.0f} đ" if (kpi_prev is not None and kpi_prev['ad_spend'] > 0) else "Tiền Ads / Tổng số đơn"
            )
        ]

        for idx, (lbl, val, num_fmt, sub_txt) in enumerate(cards):
            c_start = idx * 2 + 1
            c_end = c_start + 1
            for rr in (r_lbl, r_val, r_sub):
                ws.merge_cells(start_row=rr, start_column=c_start, end_row=rr, end_column=c_end)
                for cc in (c_start, c_end):
                    cell = ws.cell(rr, cc)
                    cell.fill = bg_fill
                    cell.border = Border(
                        left=accent_side if cc == c_start else None,
                        right=thin_gray if cc == c_end else None,
                        top=thin_gray if rr == r_lbl else None,
                        bottom=thin_gray if rr == r_sub else None
                    )

            c_l = ws.cell(r_lbl, c_start, lbl)
            c_l.font = Font(name="Segoe UI", size=9.5, bold=True, color="475569")
            c_l.alignment = Alignment(horizontal="left", vertical="center", indent=1)

            c_v = ws.cell(r_val, c_start, val)
            c_v.font = Font(name="Segoe UI", size=14, bold=True, color="0F2C59")
            c_v.number_format = num_fmt
            c_v.alignment = Alignment(horizontal="left", vertical="center", indent=1)

            c_s = ws.cell(r_sub, c_start, sub_txt)
            c_s.font = Font(name="Segoe UI", size=8.5, italic=True, color="64748B")
            c_s.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    curr_row = 4
    write_kpi_block(curr_row, "🌐 TỔNG HỢP TOÀN SHOP (SHOPEE + TIKTOK SHOP)", navy_fill, card_bg_all, thick_navy, m_curr['all'], m_prev['all'] if m_prev else None)
    curr_row += 5

    if m_curr['Shopee']['tot_orders'] > 0 or m_curr['Shopee']['ad_spend'] > 0:
        write_kpi_block(curr_row, "🟠 SÀN SHOPEE", sp_header_fill, card_bg_sp, thick_gold, m_curr['Shopee'], m_prev['Shopee'] if m_prev else None)
        curr_row += 5

    if m_curr['TikTok Shop']['tot_orders'] > 0 or m_curr['TikTok Shop']['ad_spend'] > 0:
        write_kpi_block(curr_row, "⚫ SÀN TIKTOK SHOP", tt_header_fill, card_bg_tt, thick_blue, m_curr['TikTok Shop'], m_prev['TikTok Shop'] if m_prev else None)
        curr_row += 5

    # Bảng 5 nhóm trạng thái đơn hàng + Biểu đồ cột ngay trên Sheet 1
    st_start = curr_row
    ws.merge_cells(start_row=st_start, start_column=1, end_row=st_start, end_column=6)
    sec_st = ws.cell(st_start, 1, "📦 PHÂN BỔ 5 NHÓM TRẠNG THÁI ĐƠN HÀNG (TOÀN SHOP)")
    sec_st.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    sec_st.fill = navy_fill
    sec_st.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[st_start].height = 25

    st_headers = ["Trạng thái đơn", "Số đơn", "SL sản phẩm", "Tổng doanh thu", "% Đơn", "% Doanh thu"]
    r_hdr = st_start + 1
    ws.row_dimensions[r_hdr].height = 24
    for c_i, h_txt in enumerate(st_headers, start=1):
        c = ws.cell(r_hdr, c_i, h_txt)
        c.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F2C59")
        c.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thick_gold)

    for row_offset, (_, r_data) in enumerate(status_df.iterrows()):
        idx = r_hdr + 1 + row_offset
        ws.row_dimensions[idx].height = 21
        vals = [
            r_data['status_group'],
            float(r_data['So_Don']),
            float(r_data['So_Luong_SP']),
            float(r_data['Doanh_Thu']),
            float(r_data['Ty_Trong_Don (%)']),
            float(r_data['Ty_Trong_DT (%)'])
        ]
        for c_i, v in enumerate(vals, start=1):
            c = ws.cell(idx, c_i, v)
            c.font = Font(name="Segoe UI", size=9.5, bold=(c_i == 1), color="1E293B")
            c.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
            if c_i in (2, 3):
                c.number_format = '#,##0'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif c_i == 4:
                c.number_format = '#,##0 "đ"'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif c_i in (5, 6):
                c.number_format = '0.0"%"'
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Doanh thu theo trạng thái đơn"
    chart.y_axis.title = "VNĐ"
    chart.legend = None
    data_ref = Reference(ws, min_col=4, min_row=r_hdr, max_row=r_hdr + len(status_df))
    cats_ref = Reference(ws, min_col=1, min_row=r_hdr + 1, max_row=r_hdr + len(status_df))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 14
    chart.height = 7.5
    ws.add_chart(chart, f"G{st_start}")

    col_widths_s1 = {1: 24, 2: 14, 3: 15, 4: 20, 5: 13, 6: 15, 7: 22, 8: 22}
    for c_idx, w in col_widths_s1.items():
        ws.column_dimensions[get_column_letter(c_idx)].width = w

    ws2 = wb.create_sheet("2_Trang_Thai_Theo_San")
    style_table_sheet(
        ws2, split_st,
        title="PHÂN BỔ 5 TRẠNG THÁI ĐƠN HÀNG THEO TỪNG SÀN",
        subtitle="Đối chiếu chi tiết số đơn và doanh thu giữa Shopee và TikTok Shop",
        currency_cols=["Doanh thu Shopee", "Doanh thu TikTok Shop"],
        int_cols=["Số đơn Shopee", "Số đơn TikTok Shop"]
    )

    if not wow_table.empty:
        ws3 = wb.create_sheet("3_So_Sanh_2_Tuan")
        style_table_sheet(
            ws3, wow_table,
            title="SO SÁNH TĂNG TRƯỞNG 2 TUẦN LIÊN TIẾP (WoW)",
            subtitle="Đối chiếu Toàn Shop, Shopee và TikTok Shop giữa Tuần trước và Tuần này",
            growth_col="Tăng trưởng (%)"
        )

    ws4 = wb.create_sheet("4_Hieu_Suat_San_Pham")
    sku_export = sku_merged.rename(columns={
        'platform': 'Sàn TMĐT',
        'product_name': 'Tên sản phẩm',
        'Tong_Doanh_Thu': 'Tổng doanh thu',
        'Tong_Don': 'Tổng số đơn',
        'SL_Ban': 'SL sản phẩm'
    })
    style_table_sheet(
        ws4, sku_export,
        title="BÁO CÁO HIỆU SUẤT DOANH THU THEO SẢN PHẨM",
        subtitle="Sắp xếp theo thứ tự doanh thu từ cao xuống thấp",
        currency_cols=["Tổng doanh thu"],
        int_cols=["Tổng số đơn", "SL sản phẩm"]
    )

    if not cancel_summary.empty:
        ws5 = wb.create_sheet("5_Phan_Tich_Hoan_Huy")
        cancel_export = cancel_summary.rename(columns={
            'platform': 'Sàn TMĐT',
            'status_group': 'Nhóm trạng thái',
            'product_name': 'Tên sản phẩm',
            'reason': 'Lý do hủy / hoàn',
            'So_Don_Loi': 'Số đơn lỗi',
            'Gia_Tri_That_Thoat': 'Giá trị đơn lỗi'
        })
        style_table_sheet(
            ws5, cancel_export,
            title="PHÂN TÍCH CHI TIẾT ĐƠN HỦY & TRẢ HÀNG / HOÀN TIỀN",
            subtitle="Thống kê nguyên nhân và sản phẩm có tỷ lệ hoàn/hủy cao",
            currency_cols=["Giá trị đơn lỗi"],
            int_cols=["Số đơn lỗi"]
        )

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- GIAO DIỆN CHÍNH ---
st.title("📊 HỆ THỐNG BÁO CÁO HIỆU SUẤT BÁN HÀNG TUẦN")
st.caption("Báo cáo Tổng hợp & Tách riêng Shopee / TikTok Shop | Doanh thu, Ads, ROI, Chi phí 1 đơn | So sánh WoW")

with st.sidebar:
    st.header("📅 Dữ liệu TUẦN NÀY (Bắt buộc)")
    sp_curr = st.file_uploader("File Shopee Tuần này (.xlsx, .csv)", type=['xlsx', 'xls', 'csv'], key="sp_curr")
    ads_sp_curr = st.number_input("Chi phí Ads Shopee Tuần này (VNĐ)", min_value=0, value=0, step=100000, key="ads_sp_curr")
    
    tt_curr = st.file_uploader("File TikTok Tuần này (.xlsx, .csv)", type=['xlsx', 'xls', 'csv'], key="tt_curr")
    ads_tt_curr = st.number_input("Chi phí Ads TikTok Tuần này (VNĐ)", min_value=0, value=0, step=100000, key="ads_tt_curr")

    st.divider()
    st.header("⏮️ Dữ liệu TUẦN TRƯỚC (Tùy chọn)")
    st.caption("Tải thêm file và Ads tuần trước để tự động đối chiếu:")
    sp_prev = st.file_uploader("File Shopee Tuần trước (.xlsx, .csv)", type=['xlsx', 'xls', 'csv'], key="sp_prev")
    ads_sp_prev = st.number_input("Chi phí Ads Shopee Tuần trước (VNĐ)", min_value=0, value=0, step=100000, key="ads_sp_prev")
    
    tt_prev = st.file_uploader("File TikTok Tuần trước (.xlsx, .csv)", type=['xlsx', 'xls', 'csv'], key="tt_prev")
    ads_tt_prev = st.number_input("Chi phí Ads TikTok Tuần trước (VNĐ)", min_value=0, value=0, step=100000, key="ads_tt_prev")

df_curr_raw, metas_curr = load_week_data(sp_curr, tt_curr)
df_prev_raw, metas_prev = load_week_data(sp_prev, tt_prev)

with st.sidebar:
    if metas_curr or metas_prev:
        st.divider()
        with st.expander("🛠️ Kiểm tra nhận diện cột tự động"):
            for m in (metas_curr + metas_prev):
                st.caption(f"**{m['file_name']}** ({m['platform']}): {m['orders']:,} đơn | Cột tiền: `{m['col_rev']}`")

if not df_curr_raw.empty:
    m_curr = calc_metrics(df_curr_raw, ads_sp_curr, ads_tt_curr)
    m_prev = calc_metrics(df_prev_raw, ads_sp_prev, ads_tt_prev) if not df_prev_raw.empty else None

    # --- PHẦN 1: TỔNG QUAN CHỈ SỐ TUẦN NÀY ---
    st.subheader("1. Tổng quan chỉ số tuần này")
    
    st.markdown("**🌐 Tổng hợp Toàn Shop (Shopee + TikTok Shop)**")
    render_4_kpi_row(
        curr_kpi=m_curr['all'],
        prev_kpi=m_prev['all'] if m_prev else None,
        card_style="kpi-card"
    )

    if m_curr['Shopee']['tot_orders'] > 0 or ads_sp_curr > 0:
        st.markdown("**🟠 Sàn Shopee**")
        render_4_kpi_row(
            curr_kpi=m_curr['Shopee'],
            prev_kpi=m_prev['Shopee'] if m_prev else None,
            card_style="kpi-card-shopee"
        )

    if m_curr['TikTok Shop']['tot_orders'] > 0 or ads_tt_curr > 0:
        st.markdown("**⚫ Sàn TikTok Shop**")
        render_4_kpi_row(
            curr_kpi=m_curr['TikTok Shop'],
            prev_kpi=m_prev['TikTok Shop'] if m_prev else None,
            card_style="kpi-card-tiktok"
        )

    # --- PHẦN 2: BẢNG & BIỂU ĐỒ PHÂN BỔ 5 NHÓM TRẠNG THÁI ---
    st.divider()
    st.subheader("2. Bảng tổng hợp 5 nhóm trạng thái đơn hàng")
    
    status_order = ['Giao thành công', 'Đang vận chuyển', 'Đang chờ giao hàng', 'Đơn hủy', 'Trả hàng / Hoàn tiền']
    
    status_df = (
        m_curr['full_df'].groupby('status_group', as_index=False)
        .agg(So_Don=('order_id', 'nunique'), So_Luong_SP=('qty', 'sum'), Doanh_Thu=('revenue', 'sum'))
    )
    base_status = pd.DataFrame({'status_group': status_order})
    status_df = pd.merge(base_status, status_df, on='status_group', how='left').fillna(0)
    tot_all_orders = m_curr['all']['tot_orders']
    tot_all_gmv = m_curr['all']['gmv']
    status_df['Ty_Trong_Don (%)'] = (status_df['So_Don'] / tot_all_orders * 100).round(1) if tot_all_orders > 0 else 0
    status_df['Ty_Trong_DT (%)'] = (status_df['Doanh_Thu'] / tot_all_gmv * 100).round(1) if tot_all_gmv > 0 else 0

    col_st1, col_st2 = st.columns([5, 5])
    with col_st1:
        st.markdown("**Tổng hợp chung 2 sàn**")
        st.dataframe(
            status_df.rename(columns={
                'status_group': 'Trạng thái đơn',
                'So_Don': 'Số đơn',
                'So_Luong_SP': 'SL sản phẩm',
                'Doanh_Thu': 'Tổng doanh thu',
                'Ty_Trong_Don (%)': '% Đơn',
                'Ty_Trong_DT (%)': '% Doanh thu'
            }).style.format({
                'Số đơn': '{:,.0f}',
                'SL sản phẩm': '{:,.0f}',
                'Tổng doanh thu': '{:,.0f} đ',
                '% Đơn': '{:.1f}%',
                '% Doanh thu': '{:.1f}%'
            }),
            use_container_width=True,
            hide_index=True
        )

    with col_st2:
        st.markdown("**Biểu đồ doanh thu theo trạng thái & sàn**")
        st_plat_df = (
            m_curr['full_df'].groupby(['status_group', 'platform'], as_index=False)
            .agg(Doanh_Thu=('revenue', 'sum'), So_Don=('order_id', 'nunique'))
        )
        fig_st = px.bar(
            st_plat_df, x='status_group', y='Doanh_Thu', color='platform', barmode='group',
            text_auto='.3s',
            category_orders={'status_group': status_order},
            color_discrete_map={'Shopee': '#D4AF37', 'TikTok Shop': '#0F2C59'},
            labels={'status_group': 'Trạng thái', 'Doanh_Thu': 'Doanh thu (VNĐ)', 'platform': 'Sàn'}
        )
        fig_st.update_traces(textposition='outside')
        fig_st.update_layout(plot_bgcolor='white', height=280, margin=dict(t=20, b=20))
        st.plotly_chart(fig_st, use_container_width=True)

    piv_orders = m_curr['full_df'].pivot_table(
        index='status_group', columns='platform', values='order_id', aggfunc='nunique', fill_value=0
    )
    piv_rev = m_curr['full_df'].pivot_table(
        index='status_group', columns='platform', values='revenue', aggfunc='sum', fill_value=0
    )
    split_st = pd.DataFrame({'Trạng thái đơn': status_order})
    for p in ['Shopee', 'TikTok Shop']:
        split_st[f'Số đơn {p}'] = split_st['Trạng thái đơn'].map(piv_orders[p]).fillna(0) if p in piv_orders else 0
        split_st[f'Doanh thu {p}'] = split_st['Trạng thái đơn'].map(piv_rev[p]).fillna(0) if p in piv_rev else 0

    with st.expander("📋 Xem bảng chi tiết 5 trạng thái tách riêng Shopee và TikTok Shop"):
        st.dataframe(
            split_st.style.format({
                'Số đơn Shopee': '{:,.0f}',
                'Doanh thu Shopee': '{:,.0f} đ',
                'Số đơn TikTok Shop': '{:,.0f}',
                'Doanh thu TikTok Shop': '{:,.0f} đ'
            }),
            use_container_width=True,
            hide_index=True
        )

    # --- CÁC TAB PHÂN TÍCH CHUYÊN SÂU ---
    tab1, tab2, tab3 = st.tabs([
        "📈 Hiệu suất Sản phẩm & Sàn",
        "🔄 So sánh Tuần này vs Tuần trước (Tách sàn)",
        "⚠️ Phân tích Đơn Hoàn / Hủy"
    ])

    with tab1:
        col1, col2 = st.columns([6, 4])
        with col1:
            st.markdown("**Top 5 Sản phẩm chủ lực (Theo Tổng doanh thu)**")
            top_prod = (
                m_curr['full_df'].groupby('product_name', as_index=False)
                .agg({'revenue': 'sum', 'qty': 'sum', 'order_id': 'nunique'})
                .sort_values(by='revenue', ascending=False).head(5)
            )
            fig_bar = px.bar(
                top_prod, x='product_name', y='revenue', text_auto='.3s',
                color_discrete_sequence=['#0F2C59'],
                labels={'product_name': 'Sản phẩm', 'revenue': 'Doanh thu (VNĐ)'}
            )
            fig_bar.update_traces(marker_line_color='#D4AF37', marker_line_width=1.5, textposition='outside')
            fig_bar.update_layout(plot_bgcolor='white', xaxis_tickangle=-15, height=380)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col2:
            st.markdown("**Tỷ trọng doanh thu theo sàn**")
            plat_df = m_curr['full_df'].groupby('platform', as_index=False)['revenue'].sum()
            fig_pie = px.pie(
                plat_df, values='revenue', names='platform', hole=0.45,
                color='platform', color_discrete_map={'Shopee': '#D4AF37', 'TikTok Shop': '#0F2C59'}
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=380)
            st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("**Bảng hiệu suất chi tiết theo từng Sản phẩm**")
        sku_merged = (
            m_curr['full_df'].groupby(['platform', 'product_name'], as_index=False)
            .agg(
                Tong_Doanh_Thu=('revenue', 'sum'),
                Tong_Don=('order_id', 'nunique'),
                SL_Ban=('qty', 'sum')
            )
            .sort_values(by='Tong_Doanh_Thu', ascending=False)
        )

        st.dataframe(
            sku_merged.style.format({
                'Tong_Doanh_Thu': '{:,.0f} đ',
                'Tong_Don': '{:,.0f}',
                'SL_Ban': '{:,.0f}'
            }),
            use_container_width=True
        )

    wow_table = pd.DataFrame()
    with tab2:
        if m_prev is not None:
            st.markdown("**Bảng so sánh tăng trưởng 2 tuần liên tiếp (Tổng hợp & Tách riêng từng sàn)**")
            wow_items = [
                ("TỔNG SHOP", "Tổng doanh thu", m_prev['all']['gmv'], m_curr['all']['gmv']),
                ("TỔNG SHOP", "Tổng số đơn", m_prev['all']['tot_orders'], m_curr['all']['tot_orders']),
                ("TỔNG SHOP", "Tổng tiền Ads", m_prev['all']['ad_spend'], m_curr['all']['ad_spend']),
                ("TỔNG SHOP", "ROI Tổng", round(m_prev['all']['roi'], 2), round(m_curr['all']['roi'], 2)),
                ("TỔNG SHOP", "Chi phí / 1 đơn hàng", round(m_prev['all']['cpo'], 0), round(m_curr['all']['cpo'], 0)),
                ("Shopee", "Doanh thu Shopee", m_prev['Shopee']['gmv'], m_curr['Shopee']['gmv']),
                ("Shopee", "Số đơn Shopee", m_prev['Shopee']['tot_orders'], m_curr['Shopee']['tot_orders']),
                ("Shopee", "Tiền Ads Shopee", m_prev['Shopee']['ad_spend'], m_curr['Shopee']['ad_spend']),
                ("Shopee", "ROI Shopee", round(m_prev['Shopee']['roi'], 2), round(m_curr['Shopee']['roi'], 2)),
                ("Shopee", "Chi phí / 1 đơn Shopee", round(m_prev['Shopee']['cpo'], 0), round(m_curr['Shopee']['cpo'], 0)),
                ("TikTok Shop", "Doanh thu TikTok", m_prev['TikTok Shop']['gmv'], m_curr['TikTok Shop']['gmv']),
                ("TikTok Shop", "Số đơn TikTok", m_prev['TikTok Shop']['tot_orders'], m_curr['TikTok Shop']['tot_orders']),
                ("TikTok Shop", "Tiền Ads TikTok", m_prev['TikTok Shop']['ad_spend'], m_curr['TikTok Shop']['ad_spend']),
                ("TikTok Shop", "ROI TikTok", round(m_prev['TikTok Shop']['roi'], 2), round(m_curr['TikTok Shop']['roi'], 2)),
                ("TikTok Shop", "Chi phí / 1 đơn TikTok", round(m_prev['TikTok Shop']['cpo'], 0), round(m_curr['TikTok Shop']['cpo'], 0)),
            ]
            wow_rows = []
            for grp_name, name, v_prev, v_curr in wow_items:
                diff = v_curr - v_prev
                pct = ((diff / v_prev) * 100) if v_prev != 0 else 0
                wow_rows.append({
                    'Phân loại': grp_name,
                    'Chỉ số': name,
                    'Tuần trước': v_prev,
                    'Tuần này': v_curr,
                    'Chênh lệch': diff,
                    'Tăng trưởng (%)': round(pct, 1)
                })
            wow_table = pd.DataFrame(wow_rows)
            st.dataframe(
                wow_table.style.format({
                    'Tuần trước': '{:,.2f}',
                    'Tuần này': '{:,.2f}',
                    'Chênh lệch': '{:+,.2f}',
                    'Tăng trưởng (%)': '{:+.1f}%'
                }),
                use_container_width=True,
                hide_index=True
            )

            st.markdown("**So sánh Tổng doanh thu theo Sàn (Tuần trước vs Tuần này)**")
            p_prev = m_prev['full_df'].groupby('platform', as_index=False)['revenue'].sum()
            p_prev['Kỳ báo cáo'] = 'Tuần trước'
            p_curr = m_curr['full_df'].groupby('platform', as_index=False)['revenue'].sum()
            p_curr['Kỳ báo cáo'] = 'Tuần này'
            p_comp = pd.concat([p_prev, p_curr], ignore_index=True)

            fig_wow = px.bar(
                p_comp, x='platform', y='revenue', color='Kỳ báo cáo', barmode='group',
                text_auto='.3s',
                color_discrete_map={'Tuần trước': '#94A3B8', 'Tuần này': '#0F2C59'},
                labels={'platform': 'Sàn TMĐT', 'revenue': 'Tổng doanh thu (VNĐ)'}
            )
            fig_wow.update_layout(plot_bgcolor='white', height=380)
            st.plotly_chart(fig_wow, use_container_width=True)
        else:
            st.info("👈 Tải thêm file dữ liệu **Tuần trước** ở thanh công cụ bên trái để hệ thống tự động lập bảng so sánh tăng trưởng 2 tuần.")

    cancel_summary = pd.DataFrame()
    with tab3:
        inv_df = m_curr['invalid_df']
        if not inv_df.empty:
            c_left, c_right = st.columns([5, 5])
            with c_left:
                st.markdown("**Phân loại Trả hàng/Hoàn tiền vs Đơn hủy**")
                type_df = inv_df.groupby('status_group', as_index=False).agg(
                    So_Don=('order_id', 'nunique'),
                    Gia_Tri=('revenue', 'sum')
                )
                fig_ctype = px.pie(
                    type_df, values='Gia_Tri', names='status_group', hole=0.4,
                    color_discrete_sequence=['#DC2626', '#94A3B8']
                )
                fig_ctype.update_traces(textposition='inside', textinfo='percent+label')
                fig_ctype.update_layout(height=350)
                st.plotly_chart(fig_ctype, use_container_width=True)

            with c_right:
                st.markdown("**Top Lý do Hủy / Trả hàng phổ biến**")
                reason_df = (
                    inv_df.groupby(['platform', 'reason'], as_index=False)
                    .agg(So_Don=('order_id', 'nunique'), Gia_Tri=('revenue', 'sum'))
                    .sort_values(by='So_Don', ascending=False).head(8)
                )
                fig_reason = px.bar(
                    reason_df, x='So_Don', y='reason', color='platform', orientation='h',
                    color_discrete_map={'Shopee': '#D4AF37', 'TikTok Shop': '#0F2C59'},
                    labels={'So_Don': 'Số đơn', 'reason': 'Lý do'}
                )
                fig_reason.update_layout(plot_bgcolor='white', yaxis={'categoryorder': 'total ascending'}, height=350)
                st.plotly_chart(fig_reason, use_container_width=True)

            st.markdown("**Danh sách chi tiết Sản phẩm bị Hoàn / Hủy nhiều nhất**")
            cancel_summary = (
                inv_df.groupby(['platform', 'status_group', 'product_name', 'reason'], as_index=False)
                .agg(So_Don_Loi=('order_id', 'nunique'), Gia_Tri_That_Thoat=('revenue', 'sum'))
                .sort_values(by='Gia_Tri_That_Thoat', ascending=False)
            )
            st.dataframe(
                cancel_summary.style.format({'Gia_Tri_That_Thoat': '{:,.0f} đ'}),
                use_container_width=True
            )
        else:
            st.success("🎉 Tuần này không ghi nhận đơn hàng nào bị hủy hoặc hoàn trả.")

    # --- NÚT TẢI FILE EXCEL DASHBOARD CHUYÊN NGHIỆP ---
    st.divider()
    excel_bytes = build_executive_excel(m_curr, m_prev, status_df, split_st, sku_merged, cancel_summary, wow_table)
    st.download_button(
        label="📥 Tải xuống Báo Cáo Tuần Đầy Đủ (.xlsx)",
        data=excel_bytes,
        file_name="Bao_Cao_Dashboard_Tuan_TMDT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Vui lòng tải lên ít nhất 1 file tất cả đơn hàng **Tuần này** (Shopee hoặc TikTok Shop) ở thanh bên trái để bắt đầu phân tích.")
