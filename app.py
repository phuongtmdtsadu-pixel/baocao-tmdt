import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from io import BytesIO
import re

# --- CẤU HÌNH GIAO DIỆN (TRẮNG - XANH NAVY - VÀNG GOLD) ---
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
        padding: 14px 16px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        margin-bottom: 10px;
    }
    .kpi-label { color: #475569; font-weight: 600; font-size: 0.88rem; margin-bottom: 4px; }
    .kpi-value { color: #0F2C59; font-weight: 700; font-size: 1.45rem; line-height: 1.25; margin-bottom: 6px; }
    .kpi-sub { font-size: 0.80rem; color: #64748B; display: flex; align-items: center; flex-wrap: wrap; gap: 6px; }
    .badge-up { background-color: #DCFCE7; color: #15803D; padding: 1px 7px; border-radius: 12px; font-weight: 600; font-size: 0.76rem; }
    .badge-down { background-color: #FEE2E2; color: #B91C1C; padding: 1px 7px; border-radius: 12px; font-weight: 600; font-size: 0.76rem; }
    .badge-neutral { background-color: #E2E8F0; color: #475569; padding: 1px 7px; border-radius: 12px; font-weight: 600; font-size: 0.76rem; }
</style>
""", unsafe_allow_html=True)

def render_kpi_card(label, curr_str, prev_str=None, pct_change=None, inverse_color=False, fallback_sub=""):
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
    <div class="kpi-card">
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

# --- TÍNH TOÁN CHỈ SỐ TỔNG QUAN (TỔNG DOANH THU, ADS, ROI, CHI PHÍ / 1 ĐƠN) ---
def calc_metrics(df, ad_spend=0):
    if df.empty:
        return None
    
    valid = df[df['is_valid'] == True]
    invalid = df[df['is_valid'] == False]
    
    gmv = df['revenue'].sum()                          # Tổng doanh thu
    tot_orders = df['order_id'].nunique()              # Tổng số đơn
    
    roi = (gmv / ad_spend) if ad_spend > 0 else 0      # ROI = Tổng doanh thu / Tiền Ads
    cpo = (ad_spend / tot_orders) if (ad_spend > 0 and tot_orders > 0) else 0  # Chi phí 1 đơn = Tổng tiền Ads / Tổng số đơn
    
    return {
        'gmv': gmv,
        'tot_orders': tot_orders,
        'ad_spend': ad_spend,
        'roi': roi,
        'cpo': cpo,
        'valid_df': valid,
        'invalid_df': invalid,
        'full_df': df
    }

def calc_pct(curr_val, prev_val):
    if prev_val == 0:
        return None
    return ((curr_val - prev_val) / prev_val) * 100

def to_excel(
