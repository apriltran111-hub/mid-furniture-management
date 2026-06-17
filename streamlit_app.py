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
    except:
        return pd.DataFrame()

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
# 4. GIAO DIỆN KHỐI TRÊN
# ==========================================
col_title, col_logo_zone = st.columns([3, 1])
with col_title:
    st.markdown(f"""
    <div style="padding: 10px; border-radius: 12px; border: 1px solid #f1f5f9; height: 110px;">
        <span style="font-size: 10px; font-weight: 600; background-color: #ecfdf5; color: #047857; padding: 2px 8px; border-radius: 999px;">HỆ THỐNG BÁO CÁO</span>
        <h1 class="report-title" style="font-size: 20px; margin: 5px 0;">QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
        <p style="font-size: 12px; color: #64748b;">MID Furniture – Report System</p>
    </div>
    """, unsafe_allow_html=True)
with col_logo_zone:
    st.markdown(f'<div style="text-align: right;"><img src="{LOGO_URL}" style="max-height: 90px;"></div>', unsafe_allow_html=True)

# ==========================================
# 5. BỘ LỌC DỮ LIỆU
# ==========================================
view_mode = st.radio("Chọn chế độ:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)
db['year'] = pd.to_numeric(db['year'], errors='coerce').fillna(2026).astype(int)
db['week'] = pd.to_numeric(db['week'], errors='coerce').fillna(1).astype(int)
db['month'] = pd.to_numeric(db['month'], errors='coerce').fillna(1).astype(int)

col_f1, col_f2, col_f3, col_f4 = st.columns(4)
with col_f1: selected_year = st.selectbox("Chọn Năm", sorted(db['year'].unique()))
db_filtered = db[db['year'] == selected_year]

if view_mode == "Xem báo cáo theo Tuần":
    with col_f2: selected_week = st.selectbox("Chọn Tuần", sorted(db_filtered['week'].unique()))
    db_filtered = db_filtered[db_filtered['week'] == selected_week]
    header_title = f"Tuần {selected_week} ({selected_year}) | {get_week_range_str(selected_week, selected_year)}"
else:
    with col_f2: selected_month = st.selectbox("Chọn Tháng", sorted(db_filtered['month'].unique()))
    db_filtered = db_filtered[db_filtered['month'] == selected_month]
    header_title = f"Tháng {selected_month}/{selected_year}"

with col_f3: selected_pic = st.selectbox("Lọc Người phụ trách:", ["Tất cả"] + sorted(db_filtered['pic'].dropna().unique().tolist()))
if selected_pic != "Tất cả": db_filtered = db_filtered[db_filtered['pic'] == selected_pic]

with col_f4: selected_status = st.selectbox("Lọc Trạng thái:", ["Tất cả"] + sorted(db_filtered['status'].dropna().unique().tolist()))
if selected_status != "Tất cả": db_filtered = db_filtered[db_filtered['status'] == selected_status]

st.markdown(f"<p style='font-size: 13px; color: #64748b;'>Đang hiển thị <b>{len(db_filtered)}</b> đơn hàng.</p>", unsafe_allow_html=True)

# ==========================================
# 6 & 7. GIAO DIỆN BẢNG TỐI ƯU (STICKY HEADER)
# ==========================================
try: sheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
except: sheet_url = "https://docs.google.com"

# Render phần Header cố định
st.markdown(f"""
<div style="background-color: white; padding: 10px 20px; border: 1px solid #e2e8f0; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
    <h2 style="font-size: 16px; font-weight: 900; margin: 0; text-transform: uppercase;">Báo Cáo Tiến Độ Đơn Hàng | <span style="color: #4338ca;">{header_title}</span></h2>
    <a href="{sheet_url}" target="_blank" style="padding: 5px 12px; font-size: 12px; background: #f1f5f9; border-radius: 6px; text-decoration: none; color: #334155;">⚙️ Data Management</a>
</div>
""", unsafe_allow_html=True)

# Render phần bảng với thanh cuộn
status_map = {'Complete': 'status-complete', 'Đang sản xuất': 'status-sx', 'Chờ lệnh sản xuất': 'status-cho-sx', 'Chờ hàng sang': 'status-sang', 'Chờ phản hồi Quality': 'status-quality', 'Chờ lắp đặt': 'status-lapdat', 'Pending': 'status-pending', 'Tiến hành bản vẽ': 'status-banve'}
rows_html = ""
for _, row in db_filtered.iterrows():
    st_label = str(row.get('status', 'Pending'))
    rows_html += f"<tr><td>{row.get('project', '-')}</td><td>{row.get('pic', '-')}</td><td>{row.get('contractDate', '-')}</td><td>{row.get('leadtime', '-')}</td><td>{row.get('loadingDate', '-')}</td><td><span style='padding: 2px 8px; border-radius: 10px; font-size: 10px; border: 1px solid #ccc;'>{st_label}</span></td><td>{str(row.get('resolvedIssues', '-'))}</td><td>{str(row.get('newIssues', '-'))}</td></tr>"

st.markdown(f"""
<div style="max-height: 500px; overflow-y: auto; border: 1px solid #e2e8f0; border-top: none; background: white;">
    <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
        <thead style="position: sticky; top: 0; background: #0f172a; color: white;">
            <tr><th style="padding: 10px;">Đơn hàng</th><th>Phụ trách</th><th>Ký HĐ</th><th>Leadtime</th><th>Loading DK</th><th>Trạng thái</th><th>Giải pháp</th><th>Vấn đề mới</th></tr>
        </thead>
        <tbody>{rows_html}</tbody>
    </table>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 8. CREDIT
# ==========================================
st.markdown("<div style='text-align: center; font-size: 10px; color: #94a3b8; padding: 10px;'>System developed by April &copy; 2026 MID Furniture Report System</div>", unsafe_allow_html=True)
