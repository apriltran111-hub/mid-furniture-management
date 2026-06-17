import datetime
import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. CẤU HÌNH TRANG
# ==========================================
st.set_page_config(layout="wide", page_title="MID Furniture - Quản Lý Tiến Độ")
LOGO_URL = "https://i.postimg.cc/d0ynyKDz/MID-FB.jpg"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;500;600;700;800&display=swap');
* { font-family: 'Inter', sans-serif; }
.report-title { font-weight: 900; text-transform: uppercase; color: #0f172a; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. KẾT NỐI DỮ LIỆU
# ==========================================
@st.cache_data(ttl=60)
def load_data_from_sheets():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        return conn.read()
    except: return pd.DataFrame()

db_raw = load_data_from_sheets()
db = db_raw.copy() if not db_raw.empty else pd.DataFrame(columns=["project", "pic", "contractDate", "leadtime", "loadingDate", "status", "resolvedIssues", "newIssues", "week", "month", "year"])

# ==========================================
# 3. HÀM TÍNH TOÁN
# ==========================================
def get_week_range_str(week_num, year):
    try:
        start_date = datetime.date(int(year), 1, 1) + datetime.timedelta(weeks=int(week_num) - 1)
        start_date -= datetime.timedelta(days=start_date.weekday())
        end_date = start_date + datetime.timedelta(days=6)
        return f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"
    except: return "01/06 - 07/06"

# ==========================================
# 4. GIAO DIỆN PHÍA TRÊN
# ==========================================
col_title, col_logo_zone = st.columns([3, 1])
with col_title:
    st.markdown("""
    <div style="padding: 10px; border-radius: 12px; height: 110px;">
        <h1 class="report-title" style="font-size: 20px;">QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 5. BỘ LỌC DỮ LIỆU
# ==========================================
view_mode = st.radio("Chọn chế độ:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)
db_filtered = db.copy()
col_f1, col_f2, col_f3, col_f4 = st.columns(4)

with col_f1: selected_year = st.selectbox("Chọn Năm", sorted(db['year'].unique()))
db_filtered = db_filtered[db_filtered['year'] == selected_year]

if view_mode == "Xem báo cáo theo Tuần":
    with col_f2: selected_week = st.selectbox("Chọn Tuần", sorted(db_filtered['week'].unique()))
    db_filtered = db_filtered[db_filtered['week'] == selected_week]
    header_title = f"Tuần {selected_week} ({selected_year}) | {get_week_range_str(selected_week, selected_year)}"
else:
    with col_f2: selected_month = st.selectbox("Chọn Tháng", sorted(db_filtered['month'].unique()))
    db_filtered = db_filtered[db_filtered['month'] == selected_month]
    header_title = f"Tháng {selected_month}/{selected_year}"

# ==========================================
# 6. RENDER GIAO DIỆN TÁCH RỜI (HEADER - TH - BODY)
# ==========================================
try: sheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
except: sheet_url = "#"

# Khối Header (Đưa ra ngoài bảng)
st.markdown(f"""
<div style="background: white; padding: 15px 20px; border: 1px solid #e2e8f0; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
    <h2 style="font-size: 16px; margin: 0; text-transform: uppercase;">Báo Cáo Tiến Độ Đơn Hàng | <b>{header_title}</b></h2>
    <a href="{sheet_url}" target="_blank" style="padding: 6px 12px; background: #f1f5f9; border-radius: 6px; text-decoration: none; color: #334155; font-size: 12px;">⚙️ Data Management</a>
</div>
""", unsafe_allow_html=True)

# Khối Table Head (Đưa ra ngoài bảng - Fixed)
st.markdown("""
<div style="background: #0f172a; color: white; padding: 10px 20px; font-size: 12px; font-weight: 700; text-transform: uppercase; border-left: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; display: grid; grid-template-columns: 14% 8% 8% 8% 10% 12% 20% 20%;">
    <div>Đơn hàng</div><div>Phụ trách</div><div>Ký HĐ</div><div>Leadtime</div><div>Loading DK</div><div>Trạng thái</div><div>Giải pháp</div><div>Vấn đề mới</div>
</div>
""", unsafe_allow_html=True)

# Khối Bảng dữ liệu (Chỉ chứa body, có scroll)
rows_html = ""
for _, row in db_filtered.iterrows():
    rows_html += f"<tr><td style='width: 14%;'>{row.get('project', '-')}</td><td style='width: 8%;'>{row.get('pic', '-')}</td><td style='width: 8%;'>{row.get('contractDate', '-')}</td><td style='width: 8%;'>{row.get('leadtime', '-')}</td><td style='width: 10%;'>{row.get('loadingDate', '-')}</td><td style='width: 12%;'>{row.get('status', '-')}</td><td style='width: 20%;'>{row.get('resolvedIssues', '-')}</td><td style='width: 20%;'>{row.get('newIssues', '-')}</td></tr>"

components.html(f"""
<div style="max-height: 500px; overflow-y: auto; border: 1px solid #e2e8f0; border-top: none; background: white;">
    <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
        <tbody>{rows_html}</tbody>
    </table>
</div>
""", height=505)

# ==========================================
# 7. CREDIT
# ==========================================
st.markdown("<div style='text-align: center; font-size: 10px; color: #94a3b8; padding: 10px;'>System developed by April &copy; 2026 MID Furniture Report System</div>", unsafe_allow_html=True)

# ==========================================
# 8. DÒNG CREDIT ĐƯỢC CĂN GIỮA TUYỆT ĐỐI Ở CUỐI TRANG WEB
# ==========================================
st.markdown("""
<div style="text-align: center; font-size: 11px; color: #94a3b8; font-weight: 500; letter-spacing: 0.02em; padding-top: 25px; padding-bottom: 15px;">
    System developed by April &copy; 2026 MID Furniture Report System
</div>
""", unsafe_allow_html=True)
