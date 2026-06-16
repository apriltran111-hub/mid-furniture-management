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
# 2. KẾT NỐI & TỰ ĐỘNG LẤY DỮ LIỆU TỪ GOOGLE SHEETS
# ==========================================
@st.cache_data(ttl=60) # Tự động làm mới dữ liệu sau mỗi 60 giây nếu có thay đổi trên Sheet
def load_data_from_sheets():
    try:
        # Khởi tạo kết nối Google Sheets dựa trên link cấu hình trong secrets.toml
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read()
        return df
    except Exception as e:
        st.error(f"Lỗi kết nối tới Google Sheet Database: {str(e)}")
        return pd.DataFrame()

db_raw = load_data_from_sheets()

# Trường hợp không kết nối được sheet hoặc sheet trống, hệ thống không sập mà hiển thị bảng trống
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
# 4. GIAO DIỆN KHỐI TRÊN (TIÊU ĐỀ & DATA MANAGEMENT & LOGO)
# ==========================================
col_title, col_actions = st.columns([2.5, 1.5])

with col_title:
    st.markdown(f"""
        <div style="background-color: white; padding: 24px; border-radius: 16px; border: 1px solid #f1f5f9; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05); height: 125px; display: flex; flex-direction: column; justify-content: center;">
            <span style="align-self: flex-start; padding: 2px 10px; font-size: 11px; font-weight: 600; background-color: #ecfdf5; color: #047857; border-radius: 9999px; border: 1px solid #d1fae5; text-transform: uppercase; letter-spacing: 0.05em; line-height: 1.2;">Hệ thống quản trị dữ liệu đám mây</span>
            <h1 class="report-title" style="font-size: 23px; margin-top: 6px; margin-bottom: 0px; line-height: 1.1;">HỆ THỐNG QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
            <p style="font-size: 13px; color: #64748b; margin-top: 4px; margin-bottom: 0px; line-height: 1.2;">MID Furniture System (Live Google Sheets Data)</p>
        </div>
    """, unsafe_allow_html=True)

with col_actions:
    col_btn, col_logo = st.columns([1.2, 1])
    with col_btn:
        st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)
        # Nút bấm lấy link gốc cấu hình trong secrets để người dùng click mở tab mới chỉnh sửa Sheet
        try:
            sheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        except:
            sheet_url = "https://docs.google.com"
        st.link_button("⚙️ Data Management", sheet_url, use_container_width=True)
        
    with col_logo:
        st.markdown(f"""
            <div style="background-color: white; border-radius: 16px; border: 1px solid #f1f5f9; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05); display: flex; align-items: center; justify-content: center; height: 125px; box-sizing: border-box; padding: 10px;">
                <img src="{LOGO_URL}" style="max-height: 90px; width: auto; border-radius: 8px; object-fit: contain;" alt="MID Logo">
            </div>
        """, unsafe_allow_html=True)

# Chế độ xem đồng bộ dữ liệu
st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
view_mode = st.radio("Chọn chế độ tổng hợp dữ liệu:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)

# ==========================================
# 5. KHU VỰC BỘ LỌC THỜI GIAN VÀ PHÂN LOẠI
# ==========================================
# Chuẩn hóa dữ liệu số thời gian từ Google Sheets
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
# 6. XÂY DỰNG GIAO DIỆN BẢNG HTML CHUẨN ĐẸP
# ==========================================
html_content = f"""
<!DOCTYPE html>
<html>
<head>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
body {{ font-family: 'Inter', sans-serif; background-color: transparent; margin: 0; padding: 0; }}
.table-container {{ background-color: white; padding: 30px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
.table-header {{ border-bottom: 2px solid #0f172a; padding-bottom: 16px; margin-bottom: 24px; display: flex; align-items: center; justify-content: space-between; }}
.table-title-group {{ display: flex; align-items: center; gap: 16px; }}
.table-title {{ font-size: 22px; font-weight: 900; color: #0f172a; text-transform: uppercase; margin: 0; letter-spacing: -0.04em; }}
.table-badge {{ background-color: #1e1b4b; color: white; padding: 6px 14px; border-radius: 6px; font-size: 11px; font-weight: 700; letter-spacing: 0.02em; text-transform: uppercase; }}
.brand-title {{ font-size: 22px; font-weight: 900; color: #172554; letter-spacing: -0.05em; }}
table.custom-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
table.custom-table th {{ background-color: #0f172a; color: white; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; padding: 12px; border: 1px solid #cbd5e1; text-align: center; }}
table.custom-table td {{ padding: 12px; border: 1px solid #e2e8f0; color: #334155; vertical-align: top; line-height: 1.5; }}
table.custom-table tr:nth-child(even) {{ background-color: #f8fafc; }}
.project-name {{ font-weight: 700; color: #0f172a; font-size: 13px; }}
.issue-new {{ background-color: rgba(254, 243, 199, 0.25); }}
.badge {{ padding: 4px 10px; border-radius: 9999px; font-size: 11px; font-weight: 700; border: 1px solid; display: inline-block; text-align: center; white-space: nowrap; }}
.status-complete {{ background-color: #f0fdf4; color: #15803d; border-color: #bbf7d0; }}
.status-sx {{ background-color: #eff6ff; color: #1d4ed8; border-color: #bfdbfe; }}
.status-cho-sx {{ background-color: #eef2ff; color: #4338ca; border-color: #c7d2fe; }}
.status-sang {{ background-color: #fffbeb; color: #b45309; border-color: #fde68a; }}
.status-quality {{ background-color: #fff1f2; color: #be123c; border-color: #fecdd3; }}
.status-lapdat {{ background-color: #faf5ff; color: #6b21a8; border-color: #e9d5ff; }}
.status-pending {{ background-color: #f8fafc; color: #334155; border-color: #e2e8f0; }}
.status-banve {{ background-color: #ecfeff; color: #0e7490; border-color: #c5f6fa; }}
</style>
</head>
<body>
<div class="table-container">
    <div class="table-header">
        <div class="table-title-group">
            <h2 class="table-title">Báo Cáo Tiến Độ Đơn Hàng</h2>
            <span class="table-badge">{header_title}</span>
            <span style="color:#64748b; font-weight:600;">• Live Database</span>
        </div>
        <div class="brand-title">MID FURNITURE</div>
    </div>
    <table class="custom-table">
        <thead>
            <tr>
                <th style="width: 14%; text-align: left;">Đơn hàng</th>
                <th style="width: 8%;">Phụ trách</th>
                <th style="width: 8%;">Ký HĐ</th>
                <th style="width: 8%;">Leadtime</th>
                <th style="width: 10%;">Loading DK</th>
                <th style="width: 12%;">Trạng thái</th>
                <th style="width: 20%; text-align: left;">Vấn đề đã có giải pháp</th>
                <th style="width: 20%; text-align: left;">Vấn đề mới cần giải quyết</th>
            </tr>
        </thead>
        <tbody>
"""

status_map = {
    'Complete': 'status-complete', 'Đang sản xuất': 'status-sx',
    'Chờ lệnh sản xuất': 'status-cho-sx', 'Chờ hàng sang': 'status-sang',
    'Chờ phản hồi Quality': 'status-quality', 'Chờ lắp đặt': 'status-lapdat',
    'Pending': 'status-pending', 'Tiến hành bản vẽ': 'status-banve'
}

if not db_filtered.empty:
    for _, row in db_filtered.iterrows():
        # Đảm bảo xử lý ký tự xuống dòng từ ô dữ liệu Google Sheet lên bảng HTML chuẩn
        resolved_txt = str(row.get('resolvedIssues', '-')).replace('\n', '<br>').replace('nan', '-')
        new_txt = str(row.get('newIssues', '-')).replace('\n', '<br>').replace('nan', '-')
        st_label = str(row.get('status', 'Pending')).strip()
        st_class = status_map.get(st_label, 'status-pending')
        
        html_content += f"""
                <tr>
                    <td class="project-name">{row.get('project', '-')}</td>
                    <td style="text-align: center; font-weight: 600;">{row.get('pic', '-')}</td>
                    <td style="text-align: center; color: #64748b;">{row.get('contractDate', '-')}</td>
                    <td style="text-align: center; color: #64748b;">{row.get('leadtime', '-')}</td>
                    <td style="text-align: center; font-weight: 600;">{row.get('loadingDate', '-')}</td>
                    <td style="text-align: center;"><span class="badge {st_class}">{st_label}</span></td>
                    <td>{resolved_txt}</td>
                    <td class="issue-new">{new_txt}</td>
                </tr>
        """
else:
    html_content += """
            <tr>
                <td colspan="8" style="text-align: center; padding: 30px; color: #94a3b8; font-weight: 500;">
                    Không tìm thấy dữ liệu phù hợp với bộ lọc hiện tại.
                </td>
            </tr>
    """

html_content += """
        </tbody>
    </table>
</div>
</body>
</html>
"""

# ==========================================
# 7. RENDER BẢNG HTML CHUẨN KHÔNG LỖI TEXT THÔ
# ==========================================
components.html(html_content, height=1000, scrolling=True)
