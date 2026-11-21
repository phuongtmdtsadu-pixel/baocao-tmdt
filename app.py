import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# --- CẤU HÌNH TRANG & GIAO DIỆN (TRẮNG - XANH NAVY - VÀNG GOLD) ---
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
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }
    div[data-testid="metric-container"] label { color: #475569; font-weight: 600; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #0F2C59; }
</style>
""", unsafe_allow_html=True)

# --- HÀM CHUẨN HÓA DỮ LIỆU SHOPEE & TIKTOK ---
def parse_shopee(file):
    df = pd.read_excel(file)
    # Map các cột chuẩn của file xuất từ Shopee Seller Center
    col_map = {
        'Mã đơn hàng': 'order_id',
        'Trạng thái Đơn Hàng': 'status',
        'Tên Sản Phẩm': 'product_name',
        'Tên phân loại hàng': 'sku',
        'Số lượng': 'qty',
        'Tổng giá bán (sản phẩm)': 'revenue'
    }
    # Lọc các cột tồn tại
    available_cols = {k: v for k, v in col_map.items() if k in df.columns}
    df = df.rename(columns=available_cols)
    df['platform'] = 'Shopee'
    
    # Chuẩn hóa trạng thái: Loại bỏ đơn Hủy và Trả hàng/Hoàn tiền
    cancel_keywords = ['hủy', 'trả hàng', 'hoàn tiền', 'cancel', 'return']
    df['is_valid'] = ~df['status'].astype(str).str.lower().apply(
        lambda x: any(k in x for k in cancel_keywords)
    )
    df['revenue'] = pd.to_numeric(df['revenue'], errors='coerce').fillna(0)
    df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
    return df

def parse_tiktok(file):
    df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
    # Map các cột chuẩn của file xuất từ TikTok Shop Seller Center
    col_map = {
        'Order ID': 'order_id',
        'Order Status': 'status',
        'Product Name': 'product_name',
        'Variation': 'sku',
        'Quantity': 'qty',
        'SKU Subtotal After Discount': 'revenue'
    }
    available_cols = {k: v for k, v in col_map.items() if k in df.columns}
    df = df.rename(columns=available_cols)
    df['platform'] = 'TikTok Shop'
    
    # Chuẩn hóa trạng thái: Loại bỏ đơn Canceled / Returned
    cancel_keywords = ['canceled', 'cancelled', 'return', 'refund', 'hủy', 'hoàn']
    df['is_valid'] = ~df['status'].astype(str).str.lower().apply(
        lambda x: any(k in x for k in cancel_keywords)
    )
    # Làm sạch định dạng tiền tệ (nếu có chữ VND hoặc dấu phẩy)
    df['revenue'] = df['revenue'].astype(str).str.replace(r'[^\d.]', '', regex=True)
    df['revenue'] = pd.to_numeric(df['revenue'], errors='coerce').fillna(0)
    df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
    return df

def to_excel(df_summary, df_top_sku, df_raw):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_summary.to_excel(writer, index=False, sheet_name='Tong_Quan_Tuan')
        df_top_sku.to_excel(writer, index=False, sheet_name='Top_San_Pham')
        df_raw.to_excel(writer, index=False, sheet_name='Du_Lieu_Chi_Tiet')
    return output.getvalue()

# --- GIAO DIỆN CHÍNH ---
st.title("📊 HỆ THỐNG BÁO CÁO HIỆU SUẤT BÁN HÀNG TUẦN")
st.caption("Tự động tổng hợp dữ liệu đa sàn (Shopee & TikTok Shop) | Loại trừ đơn hủy & hoàn trả")

# Sidebar Upload
with st.sidebar:
    st.header("📁 Tải dữ liệu đầu vào")
    st.info("Tải file xuất đơn hàng từ Seller Center của từng sàn:")
    shopee_file = st.file_uploader("1. File đơn hàng Shopee (.xlsx)", type=['xlsx'])
    tiktok_file = st.file_uploader("2. File đơn hàng TikTok Shop (.csv, .xlsx)", type=['csv', 'xlsx'])
    
    st.divider()
    ad_spend = st.number_input("Chi phí Quảng cáo tuần (VNĐ)", min_value=0, value=0, step=100000, help="Nhập tổng chi phí Ads để tính ROI thực tế")

# Xử lý dữ liệu khi có file tải lên
dfs = []
if shopee_file:
    dfs.append(parse_shopee(shopee_file))
if tiktok_file:
    dfs.append(parse_tiktok(tiktok_file))

if dfs:
    full_df = pd.concat(dfs, ignore_index=True)
    valid_df = full_df[full_df['is_valid'] == True]
    
    # --- TÍNH TOÁN CHỈ SỐ KPI ---
    total_gmv = full_df['revenue'].sum()
    net_revenue = valid_df['revenue'].sum()
    canceled_revenue = total_gmv - net_revenue
    
    total_orders = full_df['order_id'].nunique()
    valid_orders = valid_df['order_id'].nunique()
    cancel_rate = ((total_orders - valid_orders) / total_orders * 100) if total_orders > 0 else 0
    
    aov = (net_revenue / valid_orders) if valid_orders > 0 else 0
    roi = (net_revenue / ad_spend) if ad_spend > 0 else 0

    # --- HIỂN THỊ THẺ KPI ---
    st.subheader("1. Tổng quan chỉ số tuần")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Doanh thu thực (Net)", f"{net_revenue:,.0f} đ", f"-{canceled_revenue:,.0f} đ hoàn/hủy", delta_color="inverse")
    c2.metric("Đơn hàng thành công", f"{valid_orders:,} đơn", f"Tổng đặt: {total_orders:,}")
    c3.metric("Tỷ lệ Hủy / Hoàn", f"{cancel_rate:.1f}%")
    c4.metric("Giá trị TB đơn (AOV)", f"{aov:,.0f} đ")
    c5.metric("ROI Quảng cáo", f"{roi:.2f}" if ad_spend > 0 else "Chưa nhập Ads")

    st.divider()

    # --- BIỂU ĐỒ TRỰC QUAN ---
    col_chart1, col_chart2 = st.columns([6, 4])

    with col_chart1:
        st.subheader("2. Top 5 Sản phẩm chủ lực (Doanh thu thực)")
        top_products = (
            valid_df.groupby('product_name', as_index=False)
            .agg({'revenue': 'sum', 'qty': 'sum', 'order_id': 'nunique'})
            .sort_values(by='revenue', ascending=False)
            .head(5)
        )
        fig_bar = px.bar(
            top_products,
            x='product_name',
            y='revenue',
            text_auto='.3s',
            color_discrete_sequence=['#0F2C59'],
            labels={'product_name': 'Sản phẩm', 'revenue': 'Doanh thu thực (VNĐ)'}
        )
        fig_bar.update_traces(marker_line_color='#D4AF37', marker_line_width=1.5, textposition='outside')
        fig_bar.update_layout(plot_bgcolor='white', xaxis_tickangle=-15, height=400)
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_chart2:
        st.subheader("3. Tỷ trọng doanh thu theo sàn")
        platform_df = valid_df.groupby('platform', as_index=False)['revenue'].sum()
        fig_pie = px.pie(
            platform_df,
            values='revenue',
            names='platform',
            hole=0.45,
            color='platform',
            color_discrete_map={'Shopee': '#D4AF37', 'TikTok Shop': '#0F2C59'}
        )
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie.update_layout(height=400)
        st.plotly_chart(fig_pie, use_container_width=True)

    # --- BẢNG CHI TIẾT TOP SKU & XUẤT FILE ---
    st.subheader("4. Bảng chi tiết hiệu suất theo Sản phẩm")
    sku_table = (
        valid_df.groupby(['platform', 'product_name'], as_index=False)
        .agg(
            So_Don=('order_id', 'nunique'),
            So_Luong_Ban=('qty', 'sum'),
            Doanh_Thu_Thuc=('revenue', 'sum')
        )
        .sort_values(by='Doanh_Thu_Thuc', ascending=False)
    )
    st.dataframe(
        sku_table.style.format({'Doanh_Thu_Thuc': '{:,.0f} đ', 'So_Luong_Ban': '{:,.0f}'}),
        use_container_width=True
    )

    # Nút xuất Excel
    summary_export = pd.DataFrame([{
        'Tong_GMV': total_gmv,
        'Doanh_Thu_Thuc': net_revenue,
        'Doanh_Thu_Huy_Hoan': canceled_revenue,
        'Don_Thanh_Cong': valid_orders,
        'Ty_Le_Huy_Hoan (%)': round(cancel_rate, 2),
        'AOV': round(aov, 0),
        'Chi_Phi_Ads': ad_spend,
        'ROI': round(roi, 2)
    }])
    
    excel_data = to_excel(summary_export, sku_table, valid_df)
    st.download_button(
        label="📥 Tải xuống Báo Cáo Tuần (.xlsx)",
        data=excel_data,
        file_name="Bao_Cao_Tuan_TMDT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👈 Vui lòng tải lên ít nhất 1 file báo cáo từ **Shopee** hoặc **TikTok Shop** ở thanh công cụ bên trái để hệ thống tự động phân tích.")