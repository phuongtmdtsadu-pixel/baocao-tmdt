import streamlit as st
import pandas as pd
import plotly.express as px
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
    .kpi-value { color: #0F2C59; font-weight: 700; font-size: 1.55rem; line-height: 1.25; margin-bottom: 6px; }
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

# --- BỘ ĐỌC FILE ĐA NĂNG (CHỐNG LỖI ENCODING, NHIỀU SHEET, LỆCH DÒNG TIÊU ĐỀ) ---
HEADER_KEYWORDS = [
    'mã đơn hàng', 'order id', 'trạng thái đơn hàng', 'order status',
    'tên sản phẩm', 'product name', 'số lượng', 'quantity'
]

def clean_col_name(c):
    return re.sub(r'\s+', ' ', str(c)).strip()

def detect_header_and_clean(df_raw):
    # Kiểm tra xem tên cột hiện tại đã chứa từ khóa chuẩn chưa
    curr_cols_lower = [clean_col_name(c).lower() for c in df_raw.columns]
    if any(k in curr_cols_lower for k in HEADER_KEYWORDS):
        df_raw.columns = [clean_col_name(c) for c in df_raw.columns]
        return df_raw

    # Quét 15 dòng đầu tiên để tìm đúng dòng chứa tiêu đề bảng
    for idx in range(min(15, len(df_raw))):
        row_vals = [clean_col_name(v).lower() for v in df_raw.iloc[idx].tolist()]
        if sum(1 for k in HEADER_KEYWORDS if k in row_vals) >= 2:
            new_cols = [clean_col_name(v) for v in df_raw.iloc[idx].tolist()]
            df_clean = df_raw.iloc[idx + 1:].copy()
            df_clean.columns = new_cols
            return df_clean.reset_index(drop=True)

    df_raw.columns = [clean_col_name(c) for c in df_raw.columns]
    return df_raw

def smart_read_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.csv') or name.endswith('.txt'):
        raw_bytes = uploaded_file.getvalue()
        for enc in ['utf-8-sig', 'utf-8', 'utf-16', 'cp1258', 'latin1']:
            try:
                text = raw_bytes.decode(enc)
                # Tự nhận diện dấu phân cách (, hoặc ; hoặc tab)
                first_line = text.splitlines()[0] if text.splitlines() else ""
                sep = '\t' if '\t' in first_line else (';' if first_line.count(';') > first_line.count(',') else ',')
                df = pd.read_csv(BytesIO(raw_bytes), encoding=enc, sep=sep, dtype=str)
                return detect_header_and_clean(df)
            except Exception:
                continue
        raise ValueError(f"Không thể giải mã file CSV: {uploaded_file.name}")
    else:
        xls = pd.ExcelFile(uploaded_file)
        # Tìm sheet chứa dữ liệu đơn hàng (nếu file có nhiều sheet)
        for sheet in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=sheet, dtype=str)
            df_checked = detect_header_and_clean(df_temp)
            cols_lower = [c.lower() for c in df_checked.columns]
            if any(k in cols_lower for k in HEADER_KEYWORDS):
                return df_checked
        # Mặc định lấy sheet đầu tiên nếu không khớp từ khóa
        df_first = pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=str)
        return detect_header_and_clean(df_first)

def find_col(df, exact_candidates, partial_candidates=None):
    cols_map = {c.lower(): c for c in df.columns}
    # Ưu tiên khớp chính xác 100%
    for cand in exact_candidates:
        if cand.lower() in cols_map:
            return cols_map[cand.lower()]
    # Khớp gần đúng nếu sàn đổi nhẹ tên cột
    if partial_candidates:
        for p in partial_candidates:
            for col_low, col_orig in cols_map.items():
                if p.lower() in col_low:
                    return col_orig
    return None

def parse_number_universal(series):
    # Xử lý chuẩn cả số thập phân quốc tế (250000.00), dấu phẩy hàng nghìn (250,000) và kiểu VN (250.000)
    def _conv(val):
        if pd.isna(val):
            return 0.0
        s = str(val).strip()
        s = re.sub(r'(?i)(vnd|vnđ|đ|₫|\s)', '', s)
        if s in ['', '-', 'nan', 'none', 'null']:
            return 0.0
        # Nếu có cả dấu chấm và dấu phẩy (vd: 1,250,000.00 hoặc 1.250.000,00)
        if ',' in s and '.' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        elif ',' in s:
            # Nếu sau dấu phẩy có đúng 3 chữ số -> dấu phẩy ngăn cách hàng nghìn (vd: 250,000)
            parts = s.split(',')
            if len(parts[-1]) == 3 and len(parts) > 1:
                s = s.replace(',', '')
            else:
                s = s.replace(',', '.')
        elif '.' in s:
            # Nếu có nhiều dấu chấm (vd: 1.250.000) -> ngăn cách hàng nghìn
            if s.count('.') > 1:
                s = s.replace('.', '')
        try:
            return float(s)
        except ValueError:
            return 0.0

    return series.apply(_conv)

# --- BỘ PHÂN LOẠI 5 TRẠNG THÁI CHUẨN MỌI PHIÊN BẢN SHOPEE & TIKTOK ---
def classify_order_status(raw_status, extra_status=""):
    s = str(raw_status).strip().lower()
    e = str(extra_status).strip().lower()

    # 1. Trả hàng / Hoàn tiền (Kiểm tra trạng thái hoàn thực tế ở cột phụ)
    if e not in ['', 'nan', 'none', 'null'] and any(k in e for k in [
        'đã chấp thuận', 'đã duyệt', 'yêu cầu chờ xử lý', 'đang xử lý',
        'đã hoàn tiền', 'hoàn tiền thành công', 'refunded', 'returned', 'return/refund'
    ]):
        return 'Trả hàng / Hoàn tiền'

    # 2. Đơn hủy
    if any(k in s for k in ['hủy', 'cancel', 'thất bại', 'failed', 'không thành công', 'đã đóng']):
        return 'Đơn hủy'

    # 3. Giao thành công (Bao gồm cả câu xác nhận đã nhận hàng của Shopee & Completed/Delivered của TikTok)
    if any(k in s for k in ['hoàn thành', 'đã giao', 'đã nhận được hàng', 'completed', 'delivered']):
        return 'Giao thành công'

    # 4. Trả hàng / Hoàn tiền ở cột trạng thái chính
    if any(k in s for k in ['trả hàng', 'hoàn tiền', 'return', 'refund', 'khiếu nại']):
        return 'Trả hàng / Hoàn tiền'

    # 5. Đang chờ giao hàng
    if any(k in s for k in [
        'chờ lấy', 'chờ giao', 'chờ xác nhận', 'awaiting', 'to ship',
        'unpaid', 'on hold', 'đang chuẩn bị', 'chờ thanh toán', 'sẵn sàng giao'
    ]):
        return 'Đang chờ giao hàng'

    # 6. Đang vận chuyển
    if any(k in s for k in [
        'đang giao', 'vận chuyển', 'transit', 'shipped', 'đã bàn giao', 'shipping', 'in transit'
    ]):
        return 'Đang vận chuyển'

    return 'Đang vận chuyển'

# --- HÀM CHUẨN HÓA CHUNG CHO MỌI FILE BÁO CÁO ĐƠN HÀNG ---
def parse_universal_orders(uploaded_file, hint_platform):
    df = smart_read_file(uploaded_file)
    cols_lower = [c.lower() for c in df.columns]

    # Tự động nhận diện đúng sàn dựa trên cột đặc trưng (kể cả khi người dùng upload nhầm ô)
    if any(k in cols_lower for k in ['mã kiện hàng', 'được shopee trợ giá', 'trạng thái trả hàng/hoàn tiền']):
        platform = 'Shopee'
    elif any(k in cols_lower for k in ['sku subtotal after discount', 'cancelation/return type', 'order substatus', 'seller sku']):
        platform = 'TikTok Shop'
    else:
        platform = hint_platform

    col_order = find_col(df, ['Mã đơn hàng', 'Order ID'], ['mã đơn', 'order id'])
    col_status = find_col(df, ['Trạng Thái Đơn Hàng', 'Trạng thái đơn hàng', 'Order Status'], ['trạng thái đơn', 'order status'])
    col_extra = find_col(df, [
        'Trạng thái Trả hàng/Hoàn tiền', 'Trạng thái trả hàng',
        'Cancelation/Return Type', 'Loại hủy/trả hàng', 'Order Substatus'
    ], ['trả hàng/hoàn tiền', 'return type'])
    col_prod = find_col(df, ['Tên sản phẩm', 'Tên Sản Phẩm', 'Product Name'], ['tên sản phẩm', 'product name'])
    col_sku = find_col(df, ['Tên phân loại hàng', 'SKU phân loại hàng', 'Variation', 'Seller SKU'], ['phân loại', 'variation', 'sku'])
    col_qty = find_col(df, ['Số lượng', 'Quantity'], ['số lượng', 'quantity'])
    col_rev = find_col(df, [
        'Tổng số tiền Người mua thanh toán',
        'SKU Subtotal After Discount',
        'Tổng phụ SKU sau giảm giá',
        'Giá ưu đãi',
        'Tổng giá bán (sản phẩm)',
        'SKU Subtotal Before Discount',
        'Order Amount',
        'Tổng số tiền đơn hàng'
    ], ['người mua thanh toán', 'sku subtotal after', 'tổng phụ sku', 'giá ưu đãi'])
    col_reason = find_col(df, [
        'Lý do hủy', 'Lý do Hủy', 'Cancel Reason', 'Return Reason',
        'Lý do trả hàng', 'Lý do Trả hàng/Hoàn tiền', 'Cancel By', 'Người hủy'
    ], ['lý do', 'reason'])

    # Loại bỏ dòng trống, dòng tổng cộng ở cuối file hoặc dòng giải thích cột của TikTok
    if col_order:
        df = df[df[col_order].notna()].copy()
        s_ord = df[col_order].astype(str).str.strip()
        df = df[
            (s_ord != '') &
            (~s_ord.str.lower().isin(['nan', 'none', 'null', 'tổng', 'total'])) &
            (~s_ord.str.lower().str.contains('order id|mã đơn|platform unique', na=False))
        ].copy()

    out = pd.DataFrame()
    out['order_id'] = df[col_order].astype(str).str.strip() if col_order else df.index.astype(str)
    out['raw_status'] = df[col_status].fillna('').astype(str).str.strip() if col_status else 'Hoàn thành'
    extra_series = df[col_extra].fillna('').astype(str).str.strip() if col_extra else pd.Series([''] * len(df), index=df.index)

    out['status_group'] = [classify_order_status(st_v, ex_v) for st_v, ex_v in zip(out['raw_status'], extra_series)]
    out['product_name'] = df[col_prod].fillna(f'Sản phẩm {platform}').astype(str).str.strip() if col_prod else f'Sản phẩm {platform}'
    out['sku'] = df[col_sku].fillna('Mặc định').astype(str).str.strip() if col_sku else 'Mặc định'
    out['qty'] = parse_number_universal(df[col_qty]) if col_qty else 1
    out['revenue'] = parse_number_universal(df[col_rev]) if col_rev else 0

    raw_reason = df[col_reason].replace(['', 'nan', 'None', 'null'], pd.NA) if col_reason else pd.Series([pd.NA] * len(df), index=df.index)
    out['reason'] = [
        str(r).strip() if pd.notna(r) else (f"Hoàn/Trả: {ex}" if ex not in ['', 'nan', 'None', 'null'] else 'Không rõ lý do')
        for r, ex in zip(raw_reason, extra_series)
    ]
    out['platform'] = platform
    out['is_valid'] = ~out['status_group'].isin(['Đơn hủy', 'Trả hàng / Hoàn tiền'])

    meta = {
        'file_name': uploaded_file.name,
        'platform': platform,
        'rows': len(out),
        'orders': out['order_id'].nunique(),
        'col_rev': col_rev or 'Không tìm thấy',
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

def calc_metrics(df, ad_spend=0):
    if df.empty:
        return None
    valid = df[df['is_valid'] == True]
    invalid = df[df['is_valid'] == False]
    delivered = df[df['status_group'] == 'Giao thành công']
    
    gmv = df['revenue'].sum()
    net_rev = valid['revenue'].sum()
    delivered_rev = delivered['revenue'].sum()
    lost_rev = invalid['revenue'].sum()
    
    tot_orders = df['order_id'].nunique()
    val_orders = valid['order_id'].nunique()
    inv_orders = invalid['order_id'].nunique()
    cancel_rate = (inv_orders / tot_orders * 100) if tot_orders > 0 else 0
    
    aov = (net_rev / val_orders) if val_orders > 0 else 0
    roi = (net_rev / ad_spend) if ad_spend > 0 else 0
    
    return {
        'gmv': gmv, 'net_rev': net_rev, 'delivered_rev': delivered_rev, 'lost_rev': lost_rev,
        'tot_orders': tot_orders, 'val_orders': val_orders, 'inv_orders': inv_orders,
        'cancel_rate': cancel_rate, 'aov': aov, 'roi': roi,
        'valid_df': valid, 'invalid_df': invalid, 'full_df': df
    }

def calc_pct(curr_val, prev_val):
    if prev_val == 0:
        return None
    return ((curr_val - prev_val) / prev_val) * 100

def to_excel(df_summary, df_status, df_sku, df_cancel, df_wow):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_summary.to_excel(writer, index=False, sheet_name='Tong_Quan_Tuan')
        df_status.to_excel(writer, index=False, sheet_name='Phan_Bo_5_Trang_Thai')
        if not df_wow.empty:
            df_wow.to_excel(writer, index=False, sheet_name='So_Sanh_2_Tuan')
        df_sku.to_excel(writer, index=False, sheet_name='Hieu_Suat_San_Pham')
        if not df_cancel.empty:
            df_cancel.to_excel(writer, index=False, sheet_name='Phan_Tich_Hoan_Huy')
    return output.getvalue()

# --- GIAO DIỆN CHÍNH ---
st.title("📊 HỆ THỐNG BÁO CÁO HIỆU SUẤT BÁN HÀNG TUẦN")
st.caption("Tự động nhận diện mọi định dạng Excel/CSV (Shopee & TikTok Shop) | 5 trạng thái đơn hàng | So sánh WoW")

with st.sidebar:
    st.header("📅 Dữ liệu TUẦN NÀY (Bắt buộc)")
    sp_curr = st.file_uploader("File Shopee Tuần này (.xlsx, .csv)", type=['xlsx', 'xls', 'csv'], key="sp_curr")
    tt_curr = st.file_uploader("File TikTok Tuần này (.csv, .xlsx)", type=['csv', 'xlsx', 'xls'], key="tt_curr")
    ads_curr = st.number_input("Chi phí Ads Tuần này (VNĐ)", min_value=0, value=0, step=100000)

    st.divider()
    st.header("⏮️ Dữ liệu TUẦN TRƯỚC (Tùy chọn)")
    st.caption("Tải thêm file tuần trước để tự động đối chiếu bên dưới từng chỉ số:")
    sp_prev = st.file_uploader("File Shopee Tuần trước (.xlsx, .csv)", type=['xlsx', 'xls', 'csv'], key="sp_prev")
    tt_prev = st.file_uploader("File TikTok Tuần trước (.csv, .xlsx)", type=['csv', 'xlsx', 'xls'], key="tt_prev")
    ads_prev = st.number_input("Chi phí Ads Tuần trước (VNĐ)", min_value=0, value=0, step=100000)

df_curr_raw, metas_curr = load_week_data(sp_curr, tt_curr)
df_prev_raw, metas_prev = load_week_data(sp_prev, tt_prev)

with st.sidebar:
    if metas_curr or metas_prev:
        st.divider()
        with st.expander("🛠️ Kiểm tra nhận diện cột tự động"):
            for m in (metas_curr + metas_prev):
                st.caption(f"**{m['file_name']}** ({m['platform']}): {m['orders']:,} đơn | Cột tiền: `{m['col_rev']}`")

if not df_curr_raw.empty:
    m_curr = calc_metrics(df_curr_raw, ads_curr)
    m_prev = calc_metrics(df_prev_raw, ads_prev) if not df_prev_raw.empty else None

    # --- PHẦN 1: THẺ KPI TỔNG QUAN ---
    st.subheader("1. Tổng quan chỉ số tuần này")
    
    r1_c1, r1_c2, r1_c3 = st.columns(3)
    with r1_c1:
        render_kpi_card(
            label="Tổng doanh thu (GMV)",
            curr_str=f"{m_curr['gmv']:,.0f} đ",
            prev_str=f"{m_prev['gmv']:,.0f} đ" if m_prev else None,
            pct_change=calc_pct(m_curr['gmv'], m_prev['gmv']) if m_prev else None,
            fallback_sub=f"Tổng {m_curr['tot_orders']:,} đơn đặt"
        )
    with r1_c2:
        render_kpi_card(
            label="Doanh thu trừ Hoàn/Hủy",
            curr_str=f"{m_curr['net_rev']:,.0f} đ",
            prev_str=f"{m_prev['net_rev']:,.0f} đ" if m_prev else None,
            pct_change=calc_pct(m_curr['net_rev'], m_prev['net_rev']) if m_prev else None,
            fallback_sub=f"{m_curr['val_orders']:,} đơn hợp lệ"
        )
    with r1_c3:
        render_kpi_card(
            label="Đã giao thành công",
            curr_str=f"{m_curr['delivered_rev']:,.0f} đ",
            prev_str=f"{m_prev['delivered_rev']:,.0f} đ" if m_prev else None,
            pct_change=calc_pct(m_curr['delivered_rev'], m_prev['delivered_rev']) if m_prev else None,
            fallback_sub="Tiền hàng đã giao xong"
        )

    r2_c1, r2_c2, r2_c3 = st.columns(3)
    with r2_c1:
        render_kpi_card(
            label=f"Thất thoát Hoàn / Hủy ({m_curr['inv_orders']:,} đơn - {m_curr['cancel_rate']:.1f}%)",
            curr_str=f"-{m_curr['lost_rev']:,.0f} đ",
            prev_str=f"-{m_prev['lost_rev']:,.0f} đ ({m_prev['inv_orders']:,} đơn - {m_prev['cancel_rate']:.1f}%)" if m_prev else None,
            pct_change=calc_pct(m_curr['lost_rev'], m_prev['lost_rev']) if m_prev else None,
            inverse_color=True,
            fallback_sub=f"{m_curr['inv_orders']:,} đơn hoàn/hủy"
        )
    with r2_c2:
        render_kpi_card(
            label="Giá trị TB đơn (AOV)",
            curr_str=f"{m_curr['aov']:,.0f} đ",
            prev_str=f"{m_prev['aov']:,.0f} đ" if m_prev else None,
            pct_change=calc_pct(m_curr['aov'], m_prev['aov']) if m_prev else None,
            fallback_sub=f"Tính trên {m_curr['val_orders']:,} đơn"
        )
    with r2_c3:
        curr_roi_str = f"{m_curr['roi']:.2f}" if ads_curr > 0 else "Chưa nhập Ads"
        prev_roi_str = f"{m_prev['roi']:.2f}" if (m_prev and ads_prev > 0) else ("Chưa nhập Ads" if m_prev else None)
        roi_pct = calc_pct(m_curr['roi'], m_prev['roi']) if (m_prev and ads_curr > 0 and ads_prev > 0) else None
        render_kpi_card(
            label="ROI Thực Tế (Net/Ads)",
            curr_str=curr_roi_str,
            prev_str=prev_roi_str,
            pct_change=roi_pct,
            fallback_sub=f"Ads: {ads_curr:,.0f} đ" if ads_curr > 0 else "Nhập chi phí Ads ở cột trái"
        )

    # --- PHẦN 2: BẢNG & BIỂU ĐỒ PHÂN BỔ 5 NHÓM TRẠNG THÁI ---
    st.divider()
    st.subheader("2. Phân bổ 5 nhóm trạng thái đơn hàng (Shopee & TikTok)")
    
    status_order = ['Giao thành công', 'Đang vận chuyển', 'Đang chờ giao hàng', 'Đơn hủy', 'Trả hàng / Hoàn tiền']
    status_df = (
        m_curr['full_df'].groupby('status_group', as_index=False)
        .agg(So_Don=('order_id', 'nunique'), So_Luong_SP=('qty', 'sum'), Doanh_Thu=('revenue', 'sum'))
    )
    base_status = pd.DataFrame({'status_group': status_order})
    status_df = pd.merge(base_status, status_df, on='status_group', how='left').fillna(0)
    status_df['Ty_Trong_Don (%)'] = (status_df['So_Don'] / m_curr['tot_orders'] * 100).round(1) if m_curr['tot_orders'] > 0 else 0
    status_df['Ty_Trong_DT (%)'] = (status_df['Doanh_Thu'] / m_curr['gmv'] * 100).round(1) if m_curr['gmv'] > 0 else 0

    col_st1, col_st2 = st.columns([5, 5])
    with col_st1:
        st.dataframe(
            status_df.rename(columns={
                'status_group': 'Trạng thái đơn',
                'So_Don': 'Số đơn',
                'So_Luong_SP': 'SL sản phẩm',
                'Doanh_Thu': 'Giá trị đơn hàng',
                'Ty_Trong_Don (%)': '% Đơn',
                'Ty_Trong_DT (%)': '% Doanh thu'
            }).style.format({
                'Số đơn': '{:,.0f}',
                'SL sản phẩm': '{:,.0f}',
                'Giá trị đơn hàng': '{:,.0f} đ',
                '% Đơn': '{:.1f}%',
                '% Doanh thu': '{:.1f}%'
            }),
            use_container_width=True,
            hide_index=True
        )

    with col_st2:
        color_map_status = {
            'Giao thành công': '#0F2C59',
            'Đang vận chuyển': '#2563EB',
            'Đang chờ giao hàng': '#D4AF37',
            'Đơn hủy': '#94A3B8',
            'Trả hàng / Hoàn tiền': '#DC2626'
        }
        fig_st = px.bar(
            status_df, x='status_group', y='Doanh_Thu', text_auto='.3s',
            color='status_group', color_discrete_map=color_map_status,
            labels={'status_group': 'Trạng thái', 'Doanh_Thu': 'Giá trị (VNĐ)'}
        )
        fig_st.update_traces(textposition='outside')
        fig_st.update_layout(plot_bgcolor='white', showlegend=False, height=280, margin=dict(t=20, b=20))
        st.plotly_chart(fig_st, use_container_width=True)

    # --- CÁC TAB PHÂN TÍCH CHUYÊN SÂU ---
    tab1, tab2, tab3 = st.tabs([
        "📈 Hiệu suất Sản phẩm & Sàn",
        "🔄 So sánh Tuần này vs Tuần trước",
        "⚠️ Phân tích Đơn Hoàn / Hủy"
    ])

    with tab1:
        col1, col2 = st.columns([6, 4])
        with col1:
            st.markdown("**Top 5 Sản phẩm chủ lực (Doanh thu trừ hoàn/hủy)**")
            top_prod = (
                m_curr['valid_df'].groupby('product_name', as_index=False)
                .agg({'revenue': 'sum', 'qty': 'sum', 'order_id': 'nunique'})
                .sort_values(by='revenue', ascending=False).head(5)
            )
            fig_bar = px.bar(
                top_prod, x='product_name', y='revenue', text_auto='.3s',
                color_discrete_sequence=['#0F2C59'],
                labels={'product_name': 'Sản phẩm', 'revenue': 'Doanh thu thực (VNĐ)'}
            )
            fig_bar.update_traces(marker_line_color='#D4AF37', marker_line_width=1.5, textposition='outside')
            fig_bar.update_layout(plot_bgcolor='white', xaxis_tickangle=-15, height=380)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col2:
            st.markdown("**Tỷ trọng doanh thu thực theo sàn**")
            plat_df = m_curr['valid_df'].groupby('platform', as_index=False)['revenue'].sum()
            fig_pie = px.pie(
                plat_df, values='revenue', names='platform', hole=0.45,
                color='platform', color_discrete_map={'Shopee': '#D4AF37', 'TikTok Shop': '#0F2C59'}
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=380)
            st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("**Bảng đối chiếu Doanh thu theo từng Sản phẩm**")
        sku_all = m_curr['full_df'].groupby(['platform', 'product_name'], as_index=False).agg(
            Tong_Dat=('revenue', 'sum'),
            Tong_Don=('order_id', 'nunique')
        )
        sku_val = m_curr['valid_df'].groupby(['platform', 'product_name'], as_index=False).agg(
            Doanh_Thu_Thuc=('revenue', 'sum'),
            Don_Hop_Le=('order_id', 'nunique'),
            SL_Ban=('qty', 'sum')
        )
        sku_merged = pd.merge(sku_all, sku_val, on=['platform', 'product_name'], how='left').fillna(0)
        sku_merged['Tien_Hoan_Huy'] = sku_merged['Tong_Dat'] - sku_merged['Doanh_Thu_Thuc']
        sku_merged['Ty_Le_Hoan_Huy (%)'] = ((sku_merged['Tong_Don'] - sku_merged['Don_Hop_Le']) / sku_merged['Tong_Don'] * 100).round(1)
        sku_merged = sku_merged.sort_values(by='Doanh_Thu_Thuc', ascending=False)

        st.dataframe(
            sku_merged.style.format({
                'Tong_Dat': '{:,.0f} đ',
                'Tien_Hoan_Huy': '{:,.0f} đ',
                'Doanh_Thu_Thuc': '{:,.0f} đ',
                'SL_Ban': '{:,.0f}',
                'Ty_Le_Hoan_Huy (%)': '{:.1f}%'
            }),
            use_container_width=True
        )

    wow_table = pd.DataFrame()
    with tab2:
        if m_prev is not None:
            st.markdown("**Bảng so sánh tăng trưởng 2 tuần liên tiếp (WoW)**")
            metrics_list = [
                ("Tổng doanh thu (GMV)", m_prev['gmv'], m_curr['gmv']),
                ("Doanh thu trừ hoàn/hủy", m_prev['net_rev'], m_curr['net_rev']),
                ("Doanh thu đã giao thành công", m_prev['delivered_rev'], m_curr['delivered_rev']),
                ("Thất thoát do hoàn/hủy", m_prev['lost_rev'], m_curr['lost_rev']),
                ("Số đơn hợp lệ (không hoàn/hủy)", m_prev['val_orders'], m_curr['val_orders']),
                ("Tỷ lệ hoàn/hủy (%)", round(m_prev['cancel_rate'], 2), round(m_curr['cancel_rate'], 2)),
                ("Giá trị trung bình đơn (AOV)", m_prev['aov'], m_curr['aov']),
                ("Chi phí Quảng cáo", ads_prev, ads_curr),
                ("ROI Thực Tế", round(m_prev['roi'], 2), round(m_curr['roi'], 2))
            ]
            wow_rows = []
            for name, v_prev, v_curr in metrics_list:
                diff = v_curr - v_prev
                pct = ((diff / v_prev) * 100) if v_prev != 0 else 0
                wow_rows.append({
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
                use_container_width=True
            )

            st.markdown("**So sánh Doanh thu trừ hoàn/hủy theo Sàn (Tuần trước vs Tuần này)**")
            p_prev = m_prev['valid_df'].groupby('platform', as_index=False)['revenue'].sum()
            p_prev['Kỳ báo cáo'] = 'Tuần trước'
            p_curr = m_curr['valid_df'].groupby('platform', as_index=False)['revenue'].sum()
            p_curr['Kỳ báo cáo'] = 'Tuần này'
            p_comp = pd.concat([p_prev, p_curr], ignore_index=True)

            fig_wow = px.bar(
                p_comp, x='platform', y='revenue', color='Kỳ báo cáo', barmode='group',
                text_auto='.3s',
                color_discrete_map={'Tuần trước': '#94A3B8', 'Tuần này': '#0F2C59'},
                labels={'platform': 'Sàn TMĐT', 'revenue': 'Doanh thu thực (VNĐ)'}
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

    # --- NÚT TẢI FILE EXCEL TỔNG HỢP ---
    st.divider()
    summary_export = pd.DataFrame([{
        'Tong_Doanh_Thu_GMV': m_curr['gmv'],
        'Doanh_Thu_Tru_Hoan_Huy': m_curr['net_rev'],
        'Da_Giao_Thanh_Cong': m_curr['delivered_rev'],
        'Tien_Hoan_Huy': m_curr['lost_rev'],
        'Tong_Don_Dat': m_curr['tot_orders'],
        'Don_Hop_Le': m_curr['val_orders'],
        'Don_Hoan_Huy': m_curr['inv_orders'],
        'Ty_Le_Hoan_Huy (%)': round(m_curr['cancel_rate'], 2),
        'AOV': round(m_curr['aov'], 0),
        'Chi_Phi_Ads': ads_curr,
        'ROI_Thuc_Te': round(m_curr['roi'], 2)
    }])
    excel_bytes = to_excel(summary_export, status_df, sku_merged, cancel_summary, wow_table)
    st.download_button(
        label="📥 Tải xuống Báo Cáo Tuần Đầy Đủ (.xlsx)",
        data=excel_bytes,
        file_name="Bao_Cao_Tuan_TMDT_Universal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Vui lòng tải lên ít nhất 1 file tất cả đơn hàng **Tuần này** (Shopee hoặc TikTok Shop) ở thanh bên trái để bắt đầu phân tích.")
