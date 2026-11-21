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
    .kpi-value { color: #0F2C59; font-weight: 700; font-size: 1.50rem; line-height: 1.25; margin-bottom: 6px; }
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

# --- BỘ ĐỌC FILE ĐA NĂNG (XỬ LÝ MỌI ĐỊNH DẠNG EXCEL/CSV CỦA SHOPEE & TIKTOK) ---
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

# --- QUY TẮC PHÂN LOẠI 5 NHÓM TRẠNG THÁI THỐNG NHẤT CHO SHOPEE & TIKTOK ---
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

    # Tự động nhận diện sàn kể cả khi TikTok xuất giao diện Tiếng Việt
    if any(k in cols_lower for k in ['m
