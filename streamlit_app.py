import datetime
import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. CẤU HÌNH TRANG & GIỮ NGUYÊN GIAO DIỆN (INTER FONT)
# ==========================================
st.set_page_config(layout="wide", page_title="MID Furniture - Quản Lý Tiến Độ")

# Đường dẫn URL chứa Logo cố định của công ty
LOGO_URL = "https://i.postimg.cc/d0ynyKDz/MID-FB.jpg"

# Inject CSS để đồng bộ phông chữ toàn cục bên ngoài thành Inter
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;500;600;700;800&display=swap');
* { font-family: 'Inter', sans-serif; }
.report-title { font-weight: 900; text-transform: uppercase; letter-spacing: -0.05em; color: #0f172a; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. KẾT NỐI & TỰ ĐỘNG LẤY DỮ LIỆU TỪ GOOGLE SHEETS (GIỮ NGUYÊN)
# ==========================================
@st.cache_data(ttl=60) 
def load_data_from_sheets():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read()
        return df
    except Exception as e:
        st.error(f"Lỗi kết nối tới Google Sheet Database: {str(e)}")
        return pd.DataFrame()

db_raw = load_data_from_sheets()

if db_raw.empty:
    st.warning("Đang chờ cấu hình kết nối hoặc Database trên Google Sheets hiện đang trống.")
    db = pd.DataFrame(columns=["project", "pic", "contractDate", "leadtime", "loadingDate", "status", "resolvedIssues", "newIssues", "week", "month", "year"])
else:
    db = db_raw.copy()

# ==========================================
# 3. HÀM TÍNH TOÁN THỜI GIAN AN TOÀN
# ==========================================
def get_week_range_str(week_num, year):
    try:
        w = int(week_num)
        y = int(year)
        first_day_of_year = datetime.date(y, 1, 1)
        if first_day_of_year.weekday() > 3:
            first_monday = first_day_of_year + datetime.timedelta(days=(7 - first_day_of_year.weekday()))
        else:
            first_monday = first_day_of_year - datetime.timedelta(days=first_day_of_year.weekday())
        start_date = first_monday + datetime.timedelta(weeks=w - 1)
        end_date = start_date + datetime.timedelta(days=6)
        return f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"
    except:
        return "01/06 - 07/06"

# ==========================================
# 4. GIAO DIỆN KHỐI TRÊN (TIÊU ĐỀ & HÌNH ẢNH LOGO GỐC)
# ==========================================
col_title, col_logo_zone = st.columns([3, 1])

with col_title:
    st.markdown(f"""
        <div style="background-color: white; padding: 24px; border-radius: 16px; border: 1px solid #f1f5f9; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05); height: 125px; display: flex; flex-direction: column; justify-content: center;">
            <span style="align-self: flex-start; padding: 2px 10px; font-size: 11px; font-weight: 600; background-color: #ecfdf5; color: #047857; border-radius: 9999px; border: 1px solid #d1fae5; text-transform: uppercase; letter-spacing: 0.05em; line-height: 1.2;">HỆ THỐNG BÁO CÁO</span>
            <h1 class="report-title" style="font-size: 23px; margin-top: 6px; margin-bottom: 0px; line-height: 1.1;">QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
            <p style="font-size: 13px; color: #64748b; margin-top: 4px; margin-bottom: 0px; line-height: 1.2;">MID Furniture – Report System</p>
        </div>
    """, unsafe_allow_html=True)

with col_logo_zone:
    st.markdown(f"""
        <div style="display: flex; align-items: center; justify-content: flex-end; height: 125px; box-sizing: border-box; padding-right: 10px;">
            <img src="{LOGO_URL}" style="max-height: 105px; width: auto; object-fit: contain;" alt="MID Logo">
        </div>
    """, unsafe_allow_html=True)

try:
    sheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
except:
    sheet_url = "https://docs.google.com"

# Chế độ xem đồng bộ dữ liệu
st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
view_mode = st.radio("Chọn chế độ tổng hợp dữ liệu:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)

# ==========================================
# 5. KHU VỰC BỘ LỌC THỜI GIAN VÀ PHÂN LOẠI
# ==========================================
db['year'] = pd.to_numeric(db['year'], errors='coerce').fillna(datetime.date.today().year).astype(int)
db['week'] = pd.to_numeric(db['week'], errors='coerce').fillna(1).astype(int)
db['month'] = pd.to_numeric(db['month'], errors='coerce').fillna(1).astype(int)

available_years = sorted(db['year'].unique()) if not db.empty else [2026]
col_f1, col_f2, col_f3, col_f4 = st.columns(4)

with col_f1:
    selected_year = st.selectbox("Chọn Năm", available_years, index=0)
    db_filtered = db[db['year'] == selected_year] if not db.empty else db

if view_mode == "Xem báo cáo theo Tuần":
    with col_f2:
        available_weeks = sorted(db_filtered['week'].unique()) if not db_filtered.empty else [1]
        selected_week = st.selectbox("Chọn Tuần", available_weeks, index=0)
        db_filtered = db_filtered[db_filtered['week'] == selected_week] if not db_filtered.empty else db_filtered
    week_range = get_week_range_str(selected_week, selected_year)
    header_title = f"Tuần {selected_week} ({selected_year}) | {week_range}"
else:
    with col_f2:
        available_months = sorted(db_filtered['month'].unique()) if not db_filtered.empty else [1]
        selected_month = st.selectbox("Chọn Tháng", available_months, index=0)
        db_filtered = db_filtered[db_filtered['month'] == selected_month] if not db_filtered.empty else db_filtered
    header_title = f"Tháng {selected_month}/{selected_year}"

with col_f3:
    all_pics = ["Tất cả Người phụ trách"] + sorted(list(db_filtered['pic'].dropna().unique())) if not db_filtered.empty else ["Tất cả Người phụ trách"]
    selected_pic = st.selectbox("Lọc nhanh Người phụ trách:", all_pics)
    if selected_pic != "Tất cả Người phụ trách" and not db_filtered.empty:
        db_filtered = db_filtered[db_filtered['pic'] == selected_pic]

with col_f4:
    all_status = ["Tất cả Trạng thái"] + sorted(list(db_filtered['status'].dropna().unique())) if not db_filtered.empty else ["Tất cả Trạng thái"]
    selected_status = st.selectbox("Lọc nhanh Trạng thái:", all_status)
    if selected_status != "Tất cả Trạng thái" and not db_filtered.empty:
        db_filtered = db_filtered[db_filtered['status'] == selected_status]

st.markdown(f"<p style='font-size: 14px; color: #64748b; margin-top:10px;'>Đang hiển thị <b>{len(db_filtered)}</b> đơn hàng lấy trực tiếp theo thời gian thực từ Google Sheets.</p>", unsafe_allow_html=True)

# ==========================================
# 6. XÂY DỰNG GIAO DIỆN BẢNG (TÁCH TIÊU ĐỀ & TABLE BODY)
# ==========================================

# 1. Phần tiêu đề (Header) - Sẽ hiển thị cố định phía trên bảng
header_section = f"""
<div style="background-color: white; padding: 20px; border-radius: 16px 16px 0 0; border: 1px solid #e2e8f0; border-bottom: none; display: flex; align-items: center; justify-content: space-between;">
    <div style="display: flex; align-items: center; gap: 16px;">
        <h2 style="font-size: 20px; font-weight: 900; color: #0f172a; text-transform: uppercase; margin: 0;">Báo Cáo Tiến Độ Đơn Hàng</h2>
        <span style="background-color: #1e1b4b; color: white; padding: 4px 12px; border-radius: 6px; font-size: 11px; font-weight: 700;">{header_title}</span>
    </div>
    <a href="{sheet_url}" target="_blank" style="text-decoration: none; border: 1px solid #cbd5e1; padding: 8px 16px; border-radius: 8px; color: #334155; font-size: 13px; font-weight: 500; background: white;">
        ⚙️ Data Management
    </a>
</div>
"""

# 2. Phần bảng dữ liệu (Table Body) - Phần này sẽ có thanh cuộn
table_content = f"""
<div style="max-height: 500px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 0 0 16px 16px; background: white;">
    <table style="width: 100%; border-collapse: collapse; font-size: 12px; table-layout: fixed;">
        <thead style="position: sticky; top: 0; background: #0f172a; color: white;">
            <tr>
                <th style="padding: 12px 6px; text-align: left; width: 14%;">Đơn hàng</th>
                <th style="padding: 12px 6px; text-align: center; width: 8%;">Phụ trách</th>
                <th style="padding: 12px 6px; text-align: center; width: 8%;">Ký HĐ</th>
                <th style="padding: 12px 6px; text-align: center; width: 8%;">Leadtime</th>
                <th style="padding: 12px 6px; text-align: center; width: 10%;">Loading DK</th>
                <th style="padding: 12px 6px; text-align: center; width: 12%;">Trạng thái</th>
                <th style="padding: 12px 6px; text-align: left; width: 20%;">Vấn đề đã có giải pháp</th>
                <th style="padding: 12px 6px; text-align: left; width: 20%;">Vấn đề mới cần giải quyết</th>
            </tr>
        </thead>
        <tbody>
"""

# Vòng lặp thêm dữ liệu (giữ nguyên logic của bạn)
if not db_filtered.empty:
    for _, row in db_filtered.iterrows():
        # ... (giữ nguyên đoạn xử lý resolved_txt, new_txt, st_class như cũ) ...
        table_content += f"""<tr>...</tr>"""
else:
    table_content += """<tr><td colspan="8" style="text-align: center; padding: 30px;">Không có dữ liệu</td></tr>"""

table_content += "</tbody></table></div>"

# ==========================================
# 7. RENDER RA GIAO DIỆN
# ==========================================
st.markdown(header_section, unsafe_allow_html=True)
st.markdown(table_content, unsafe_allow_html=True)

# ==========================================
# 8. DÒNG CREDIT ĐƯỢC CĂN GIỮA TUYỆT ĐỐI Ở CUỐI TRANG WEB
# ==========================================
st.markdown("""
<div style="text-align: center; font-size: 11px; color: #94a3b8; font-weight: 500; letter-spacing: 0.02em; padding-top: 25px; padding-bottom: 15px;">
    System developed by April &copy; 2026 MID Furniture Report System
</div>
""", unsafe_allow_html=True)
