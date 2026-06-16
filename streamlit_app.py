import datetime
import re
import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 1. CẤU HÌNH TRANG NGƯỜI DÙNG
st.set_page_config(layout="wide", page_title="MID Furniture - Quản Lý Tiến Độ")

COLUMNS = [
    "project", "pic", "contractDate", "leadtime", "loadingDate", 
    "status", "resolvedIssues", "newIssues", "week", "month", "year"
]

# 2. ĐỌC GOOGLE SHEETS TRỰC TIẾP QUA PHƯƠNG THỨC EXPORT CSV
database_df = pd.DataFrame(columns=COLUMNS)
raw_url = ""
try:
    # Lấy link từ Secrets
    raw_url = st.secrets["connections"]["my_gsheets"]["spreadsheet"]
    # Chuyển đổi sang link xuất CSV trực tiếp để Pandas đọc ngầm
    csv_url = raw_url.replace("/edit?usp=sharing", "/export?format=csv&gid=0")
    
    # Đọc dữ liệu trực tiếp thời gian thực từ Google Sheets
    database_df = pd.read_csv(csv_url)
    database_df = database_df.dropna(how="all")
    
    # Chuẩn hóa các cột
    for col in COLUMNS:
        if col not in database_df.columns:
            database_df[col] = "-"
    database_df = database_df[COLUMNS]
except Exception as e:
    st.warning("⚠️ Hệ thống đang kết nối đến Google Sheets, vui lòng đợi trong giây lát...")

# --- GIAO DIỆN CHÍNH ---

# Chia khu vực tiêu đề làm 2 cột: Cột trái chứa Title, Cột phải chứa nút truy cập Google Sheets
col_title, col_btn = st.columns([3, 1])

with col_title:
    st.markdown("""
        <div style="background-color: white; padding: 24px; border-radius: 16px; border: 1px solid #f1f5f9; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05); font-family: 'Segoe UI', Roboto, sans-serif;">
            <span style="padding: 4px 12px; font-size: 12px; font-weight: 600; background-color: #ecfdf5; color: #047857; border-radius: 9999px; border: 1px solid #d1fae5; text-transform: uppercase; letter-spacing: 0.05em;">Hệ thống báo cáo</span>
            <h1 style="font-weight: 900; text-transform: uppercase; letter-spacing: -0.05em; color: #0f172a; font-size: 24px; margin-top: 8px; margin-bottom: 0px;">QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
            <p style="font-size: 14px; color: #64748b; margin-top: 4px; margin-bottom: 0px;">MID Furniture – Report System</p>
        </div>
    """, unsafe_allow_html=True)

with col_btn:
    # Tạo khoảng trống phía trên để nút bấm căn lề đẹp mắt với tiêu đề
    st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)
    
    # Tạo nút bấm dạng Link mở trực tiếp link gốc Google Sheets trong Secrets sang một tab mới
    if raw_url:
        st.link_button(
            label="⚙️ Data Management", 
            url=raw_url, 
            use_container_width=True,
            help="Bấm vào đây để mở trang quản lý Google Sheets cập nhật và chỉnh sửa dữ liệu"
        )
    else:
        st.button("⚙️ Data Management (Chưa cấu hình link)", disabled=True, use_container_width=True)

# --- KHU VỰC BỘ LỌC THỜI GIAN VÀ CHẾ ĐỘ TỔNG HỢP ---
st.markdown("<br>", unsafe_allow_html=True)
view_mode = st.radio("Chọn chế độ tổng hợp dữ liệu:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)

db = database_df.copy()

def get_week_range_str(week_num, year):
    try:
        first_day_of_year = datetime.date(year, 1, 1)
        if first_day_of_year.weekday() > 3:
            first_monday = first_day_of_year + datetime.timedelta(days=(7 - first_day_of_year.weekday()))
        else:
            first_monday = first_day_of_year - datetime.timedelta(days=first_day_of_year.weekday())
        start_date = first_monday + datetime.timedelta(weeks=int(week_num) - 1)
        end_date = start_date + datetime.timedelta(days=6)
        return f"{start_date.strftime('%d/%m')}-{end_date.strftime('%d/%m')}"
    except:
        return "01/01-07/01"

def get_status_badge_style(status):
    styles = {
        'Complete': 'background-color: #f0fdf4; color: #15803d; border-color: #bbf7d0;',
        'Đang sản xuất': 'background-color: #eff6ff; color: #1d4ed8; border-color: #bfdbfe;',
        'Chờ lệnh sản xuất': 'background-color: #eef2ff; color: #4338ca; border-color: #c7d2fe;',
        'Chờ hàng sang': 'background-color: #fffbeb; color: #b45309; border-color: #fde68a;',
        'Chờ phản hồi Quality': 'background-color: #fff1f2; color: #be123c; border-color: #fecdd3;',
        'Chờ lắp đặt': 'background-color: #faf5ff; color: #6b21a8; border-color: #e9d5ff;',
        'Pending': 'background-color: #f8fafc; color: #334155; border-color: #e2e8f0;',
        'Tiến hành bản vẽ': 'background-color: #ecfeff; color: #0e7490; border-color: #c5f6fa;'
    }
    return styles.get(status, 'background-color: #f8fafc; color: #334155; border-color: #e2e8f0;')

if db.empty:
    st.info("Trang tính Google Sheets hiện chưa có dữ liệu hoặc đường link chưa chính xác. Vui lòng kiểm tra lại.")
else:
    db['year'] = pd.to_numeric(db['year'], errors='coerce').fillna(0).astype(int)
    db['week'] = pd.to_numeric(db['week'], errors='coerce').fillna(0).astype(int)
    db['month'] = pd.to_numeric(db['month'], errors='coerce').fillna(0).astype(int)

    available_years = sorted(db['year'].unique())
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        selected_year = st.selectbox("Chọn Năm", available_years, index=len(available_years)-1)
        db_filtered = db[db['year'] == selected_year]
        
    if view_mode == "Xem báo cáo theo Tuần":
        with col_f2:
            available_weeks = sorted(db_filtered['week'].unique())
            selected_week = st.selectbox("Chọn Tuần", available_weeks, index=len(available_weeks)-1)
            db_filtered = db_filtered[db_filtered['week'] == selected_week]
        
        week_range = get_week_range_str(selected_week, selected_year)
        header_title = f"Tuần {selected_week} ({selected_year}) | {week_range}"
    else:
        with col_f2:
            available_months = sorted(db_filtered['month'].unique())
            selected_month = st.selectbox("Chọn Tháng", available_months, index=len(available_months)-1)
            db_filtered = db_filtered[db_filtered['month'] == selected_month]
        header_title = f"TỔNG HỢP BÁO CÁO TIẾN ĐỘ - THÁNG {selected_month}/{selected_year}"

    with col_f3:
        all_pics = ["Tất cả Người phụ trách"] + sorted(list(db_filtered['pic'].dropna().unique()))
        selected_pic = st.selectbox("Lọc nhanh Người phụ trách:", all_pics)
        if selected_pic != "Tất cả Người phụ trách":
            db_filtered = db_filtered[db_filtered['pic'] == selected_pic]

    with col_f4:
        all_status = ["Tất cả Trạng thái"] + sorted(list(db_filtered['status'].dropna().unique()))
        selected_status = st.selectbox("Lọc nhanh Trạng thái:", all_status)
        if selected_status != "Tất cả Trạng thái":
            db_filtered = db_filtered[db_filtered['status'] == selected_status]

    st.markdown(f"<p style='font-size: 14px; color: #64748b; margin-top:10px;'>Đang hiển thị <b>{len(db_filtered)}</b> dòng dữ liệu từ Google Sheets.</p>", unsafe_allow_html=True)

    # --- TẠO CHUỖI HTML BẢNG THUẦN ---
    html_body = ""
    for idx, row in db_filtered.reset_index().iterrows():
        resolved_formatted = str(row['resolvedIssues']).replace('\n', '<br>')
        new_formatted = str(row['newIssues']).replace('\n', '<br>')
        badge_css = get_status_badge_style(row['status'])
        row_bg = "#f8fafc" if idx % 2 != 0 else "#ffffff"
        
        html_body += f"""
            <tr style="background-color: {row_bg};">
                <td style="padding: 12px; border: 1px solid #e2e8f0; font-weight: 700; color: #0f172a; font-size: 13px;">{row['project']}</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align: center; font-weight: 600; color: #475569; font-size: 12px;">{row['pic']}</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align: center; color: #64748b; font-size: 12px;">{row['contractDate']}</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align: center; color: #64748b; font-size: 12px;">{row['leadtime']}</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align: center; font-weight: 600; font-size: 12px;">{row['loadingDate']}</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align: center;">
                    <span style="padding: 4px 10px; border-radius: 9999px; font-size: 11px; font-weight: 700; border: 1px solid; display: inline-block; {badge_css}">{row['status']}</span>
                </td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color: #475569; font-size: 12px;">{resolved_formatted}</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color: #475569; font-size: 12px; background-color: rgba(254, 243, 199, 0.2);">{new_formatted}</td>
            </tr>
        """

    full_table_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght=300;400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Inter', sans-serif; background-color: transparent; margin: 0; padding: 0; }
            th { background-color: #0f172a; color: white; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; padding: 12px; border: 1px solid #cbd5e1; font-size: 11px; }
        </style>
    </head>
    <body>
    <div style="background-color: white; padding: 30px; border-radius: 16px; border: 1px solid #e2e8f0; overflow-x: auto;">
        <div style="display: flex; justify-content: space-between; border-bottom: 2px solid #0f172a; padding-bottom: 24px; margin-bottom: 32px; align-items: center;">
            <div>
                <h2 style="font-size: 24px; font-weight: 900; color: #0f172a; text-transform: uppercase; margin: 0; letter-spacing: -0.05em;">Báo Cáo Tiến Độ Đơn Hàng</h2>
                <div style="display: flex; align-items: center; gap: 12px; margin-top: 8px; font-size: 14px; color: #64748b; font-weight: 600;">
                    <span style="background-color: #1e1b4b; color: white; padding: 2px 12px; border-radius: 6px; font-size: 12px; letter-spacing: 0.05em;">""" + header_title + """</span>
                </div>
            </div>
            <div style="text-align: right;">
                <span style="font-size: 24px; font-weight: 900; color: #172554; letter-spacing: -0.05em;">MID FURNITURE</span>
            </div>
        </div>

        <table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px;">
            <thead>
                <tr>
                    <th style="text-align: left; width: 12%;">Đơn hàng</th>
                    <th style="text-align: center; width: 7%;">Phụ trách</th>
                    <th style="text-align: center; width: 7%;">Ký HĐ</th>
                    <th style="text-align: center; width: 7%;">Leadtime</th>
                    <th style="text-align: center; width: 9%;">Loading DK</th>
                    <th style="text-align: center; width: 11%;">Trạng thái</th>
                    <th style="text-align: left; width: 22%;">Vấn đề đã có giải pháp</th>
                    <th style="text-align: left; width: 24%;">Vấn đề mới cần giải quyết</th>
                </tr>
            </thead>
            <tbody>
                """ + html_body + """
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """
    
    dynamic_height = 200 + (len(db_filtered) * 75)
    if dynamic_height < 500: dynamic_height = 500
    
    components.html(full_table_html, height=dynamic_height, scrolling=True)

# --- DÒNG CREDIT PHÍA CUỐI TRANG (FOOTER) ---
st.markdown("<br><hr>", unsafe_allow_html=True)
st.markdown("""
    <div style="text-align: center; color: #94a3b8; font-size: 13px; font-family: 'Segoe UI', Roboto, sans-serif; padding-bottom: 20px;">
        System developed by April • © 2026 MID Furniture Report System
    </div>
""", unsafe_allow_html=True)
