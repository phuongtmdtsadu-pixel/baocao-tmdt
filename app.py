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
    div[data-testid="metric-container"] {
        background-color: #F8FAFC;
        border-left: 5px solid #0F2C59;
        border-top: 1px solid #E2E8F0;
        border-right: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
        padding: 14px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }
    div[data-testid="metric-container"] label { color: #475569; font-weight: 600; font-size: 0.88rem; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #0F2C59; font-size: 1.45rem; }
</style>
""", unsafe_allow_html=True)

# --- HÀM CHUẨN HÓA TÊN CỘT (XÓA DẤU XUỐNG DÒNG \n & KHOẢNG TRẮNG THỪA CHO CẢ 2 SÀN) ---
def normalize_columns(df):
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    return df

def find_col(df, candidates):
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

# --- HÀM PHÂN LOẠI 5 NHÓM TRẠNG THÁI THỐNG NHẤT CHO CẢ SHOPEE & TIKTOK ---
def classify_order_status(raw_status, extra_status=""):
    s = str(raw_status).strip().lower()
    e = str(extra_status).strip().lower()
    combined = f"{s} {e}"
    
    # 1. Trả hàng / Hoàn tiền (Kiểm tra cả trạng thái chính và trạng thái phụ/hoàn tiền của 2 sàn)
    if any(k in combined for k in [
        'trả hàng', 'hoàn tiền', 'return', 'refund', 'khiếu nại',
        'đã duyệt', 'chấp nhận hoàn', 'đang hoàn'
    ]):
        return 'Trả hàng / Hoàn tiền'
    
    # 2. Đơn hủy (Bao gồm giao thất bại / hủy bởi người mua, người bán, hệ thống)
    if any(k in combined for k in [
        'hủy', 'cancel', 'thất bại', 'failed', 'không thành công', 'đã đóng'
    ]):
        return 'Đơn hủy'
    
    # 3. Giao thành công
    if any(k in combined for k in [
        'hoàn thành', 'đã giao', 'completed', 'delivered'
    ]):
        return 'Giao thành công'
    
    # 4. Đang chờ giao hàng (Chưa bàn giao cho ĐVVC / Chờ lấy hàng / Chờ thanh toán)
    if any(k in combined for k in [
        'chờ lấy', 'chờ giao', 'chờ xác nhận', 'awaiting', 'to ship',
        'unpaid', 'on hold', 'đang chuẩn bị', 'chờ thanh toán', 'sẵn sàng giao'
    ]):
        return 'Đang chờ giao hàng'
    
    # 5. Đang vận chuyển (Đã bàn giao ĐVVC, đang trên đường giao)
    if any(k in combined for k in [
        'đang giao', 'vận chuyển', 'transit', 'shipped', 'đã bàn giao', 'shipping', 'in transit'
    ]):
        return 'Đang vận chuyển'
    
    return 'Đang vận chuyển'

def clean_currency(series):
    # Xử lý chuẩn cả định dạng số lẫn chuỗi có chữ VND/đ hoặc dấu phẩy ngăn cách hàng nghìn
    s = series.astype(str).str.replace('VND', '', case=False, regex=False)
    s = s.str.replace('đ', '', case=False, regex=False)
    s = s.str.replace(',', '', regex=False).str.strip()
    return pd.to_numeric(s, errors='coerce').fillna(0)

# --- BỘ ĐỌC DỮ LIỆU SHOPEE ---
def parse_shopee(file):
    df = pd.read_excel(file)
    df = normalize_columns(df)
    
    col_order = find_col(df, ['Mã đơn hàng', 'Order ID'])
    col_status = find_col(df, ['Trạng Thái Đơn Hàng', 'Trạng thái đơn hàng', 'Order Status'])
    col_return = find_col(df, ['Trạng thái Trả hàng/Hoàn tiền', 'Trạng thái trả hàng'])
    col_prod = find_col(df, ['Tên sản phẩm', 'Tên Sản Phẩm', 'Product Name'])
    col_sku = find_col(df, ['Tên phân loại hàng', 'SKU phân loại hàng', 'Variation'])
    col_qty = find_col(df, ['Số lượng', 'Quantity'])
    # Dùng giá theo từng dòng sản phẩm để không bị cộng trùng khi 1 đơn có nhiều món
    col_rev = find_col(df, ['Tổng số tiền Người mua thanh toán', 'Giá ưu đãi', 'Tổng giá bán (sản phẩm)'])
    col_reason = find_col(df, ['Lý do hủy', 'Lý do Hủy', 'Lý do Trả hàng/Hoàn tiền', 'Ghi chú'])

    # Loại bỏ dòng trống hoặc dòng giải thích nếu có
    if col_order:
        df = df[df[col_order].notna() & (df[col_order].astype(str).str.strip() != '')].copy()

    out = pd.DataFrame()
    out['order_id'] = df[col_order].astype(str).str.strip() if col_order else df.index.astype(str)
    out['raw_status'] = df[col_status].astype(str).str.strip() if col_status else 'Hoàn thành'
    ret_series = df[col_return].fillna('').astype(str) if col_return else pd.Series([''] * len(df), index=df.index)
    
    out['status_group'] = [classify_order_status(st_val, rt_val) for st_val, rt_val in zip(out['raw_status'], ret_series)]
    out['product_name'] = df[col_prod].fillna('Sản phẩm Shopee').astype(str).str.strip() if col_prod else 'Sản phẩm Shopee'
    out['sku'] = df[col_sku].fillna('Mặc định').astype(str).str.strip() if col_sku else 'Mặc định'
    out['qty'] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0) if col_qty else 1
    out['revenue'] = clean_currency(df[col_rev]) if col_rev else 0
    out['reason'] = df[col_reason].replace(['', 'nan', 'None'], pd.NA).fillna('Không rõ lý do').astype(str).str.strip() if col_reason else 'Không có dữ liệu lý do'
    out['platform'] = 'Shopee'
    out['is_valid'] = ~out['status_group'].isin(['Đơn hủy', 'Trả hàng / Hoàn tiền'])
    return out

# --- BỘ ĐỌC DỮ LIỆU TIKTOK SHOP (ĐỒNG BỘ QUY TẮC VỚI SHOPEE) ---
def parse_tiktok(file):
    df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
    df = normalize_columns(df)
    
    col_order = find_col(df, ['Order ID', 'Mã đơn hàng'])
    col_status = find_col(df, ['Order Status', 'Trạng thái đơn hàng', 'Trạng Thái Đơn Hàng'])
    col_substatus = find_col(df, ['Order Substatus', 'Cancelation/Return Type', 'Trạng thái phụ', 'Loại hủy/trả hàng'])
    col_prod = find_col(df, ['Product Name', 'Tên sản phẩm', 'Tên Sản Phẩm'])
    col_sku = find_col(df, ['Variation', 'Seller SKU', 'Tên phân loại hàng', 'Phân loại'])
    col_qty = find_col(df, ['Quantity', 'Số lượng'])
    # Dùng giá theo từng dòng SKU để chống nhân đôi doanh thu khi 1 đơn TikTok có nhiều SKU
    col_rev = find_col(df, [
        'SKU Subtotal After Discount',
        'Tổng phụ SKU sau giảm giá',
        'SKU Subtotal Before Discount',
        'Order Amount',
        'Tổng số tiền đơn hàng'
    ])
    col_reason = find_col(df, ['Cancel Reason', 'Return Reason', 'Lý do hủy', 'Lý do trả hàng', 'Cancel By', 'Người hủy'])

    # Loại bỏ dòng giải thích phụ (nếu file TikTok có dòng mô tả cột ngay dưới tiêu đề)
    if col_order:
        df = df[df[col_order].notna()].copy()
        df = df[~df[col_order].astype(str).str.lower().str.contains('order id|mã đơn|platform unique', na=False)].copy()

    out = pd.DataFrame()
    out['order_id'] = df[col_order].astype(str).str.strip() if col_order else df.index.astype(str)
    out['raw_status'] = df[col_status].astype(str).str.strip() if col_status else 'Completed'
    sub_series = df[col_substatus].fillna('').astype(str) if col_substatus else pd.Series([''] * len(df), index=df.index)
    
    out['status_group'] = [classify_order_status(st_val, sb_val) for st_val, sb_val in zip(out['raw_status'], sub_series)]
    out['product_name'] = df[col_prod].fillna('Sản phẩm TikTok').astype(str).str.strip() if col_prod else 'Sản phẩm TikTok'
    out['sku'] = df[col_sku].fillna('Mặc định').astype(str).str.strip() if col_sku else 'Mặc định'
    out['qty'] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0) if col_qty else 1
    out['revenue'] = clean_currency(df[col_rev]) if col_rev else 0
    out['reason'] = df[col_reason].replace(['', 'nan', 'None'], pd.NA).fillna('Không rõ lý do').astype(str).str.strip() if col_reason else 'Không có dữ liệu lý do'
    out['platform'] = 'TikTok Shop'
    out['is_valid'] = ~out['status_group'].isin(['Đơn hủy', 'Trả hàng / Hoàn tiền'])
    return out

def load_week_data(shopee_f, tiktok_f):
    dfs = []
    if shopee_f:
        dfs.append(parse_shopee(shopee_f))
    if tiktok_f:
        dfs.append(parse_tiktok(tiktok_f))
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

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
st.caption("Đồng bộ quy tắc Shopee & TikTok Shop | 5 trạng thái đơn hàng | Doanh thu trừ hoàn/hủy | So sánh WoW")

with st.sidebar:
    st.header("📅 Dữ liệu TUẦN NÀY (Bắt buộc)")
    sp_curr = st.file_uploader("Shopee Tuần này (.xlsx)", type=['xlsx'], key="sp_curr")
    tt_curr = st.file_uploader("TikTok Tuần này (.csv, .xlsx)", type=['csv', 'xlsx'], key="tt_curr")
    ads_curr = st.number_input("Chi phí Ads Tuần này (VNĐ)", min_value=0, value=0, step=100000)

    st.divider()
    st.header("⏮️ Dữ liệu TUẦN TRƯỚC (Tùy chọn)")
    st.caption("Tải thêm file tuần trước để tự động tính % tăng trưởng:")
    sp_prev = st.file_uploader("Shopee Tuần trước (.xlsx)", type=['xlsx'], key="sp_prev")
    tt_prev = st.file_uploader("TikTok Tuần trước (.csv, .xlsx)", type=['csv', 'xlsx'], key="tt_prev")
    ads_prev = st.number_input("Chi phí Ads Tuần trước (VNĐ)", min_value=0, value=0, step=100000)

df_curr_raw = load_week_data(sp_curr, tt_curr)
df_prev_raw = load_week_data(sp_prev, tt_prev)

if not df_curr_raw.empty:
    m_curr = calc_metrics(df_curr_raw, ads_curr)
    m_prev = calc_metrics(df_prev_raw, ads_prev) if not df_prev_raw.empty else None

    def fmt_delta(curr_val, prev_val):
        if m_prev is None or prev_val == 0:
            return None
        pct = ((curr_val - prev_val) / prev_val) * 100
        return f"{pct:+.1f}% so với tuần trước"

    # --- PHẦN 1: THẺ KPI TỔNG QUAN ---
    st.subheader("1. Tổng quan chỉ số tuần này")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(
        "Doanh thu trừ Hoàn/Hủy",
        f"{m_curr['net_rev']:,.0f} đ",
        fmt_delta(m_curr['net_rev'], m_prev['net_rev']) if m_prev else f"Tổng đặt: {m_curr['gmv']:,.0f} đ"
    )
    c2.metric(
        "Đã giao thành công",
        f"{m_curr['delivered_rev']:,.0f} đ",
        fmt_delta(m_curr['delivered_rev'], m_prev['delivered_rev']) if m_prev else "Tiền hàng đã giao xong"
    )
    c3.metric(
        "Thất thoát Hoàn / Hủy",
        f"-{m_curr['lost_rev']:,.0f} đ",
        f"{m_curr['inv_orders']:,} đơn ({m_curr['cancel_rate']:.1f}%)",
        delta_color="inverse"
    )
    c4.metric(
        "Giá trị TB đơn (AOV)",
        f"{m_curr['aov']:,.0f} đ",
        fmt_delta(m_curr['aov'], m_prev['aov']) if m_prev else f"{m_curr['val_orders']:,} đơn hợp lệ"
    )
    c5.metric(
        "ROI Thực Tế (Net/Ads)",
        f"{m_curr['roi']:.2f}" if ads_curr > 0 else "Chưa nhập Ads",
        fmt_delta(m_curr['roi'], m_prev['roi']) if (m_prev and ads_curr > 0 and ads_prev > 0) else None
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
                ("Tổng giá trị đặt (GMV)", m_prev['gmv'], m_curr['gmv']),
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
        'Tong_Dat_Hang_GMV': m_curr['gmv'],
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
        file_name="Bao_Cao_Tuan_TMDT_Dong_Bo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Vui lòng tải lên ít nhất 1 file tất cả đơn hàng **Tuần này** (Shopee hoặc TikTok Shop) ở thanh bên trái để bắt đầu phân tích.")
