import datetime
import re
import streamlit as st
import pandas as pd
from docx import Document

# 1. CẤU HÌNH TRANG & GIỮ NGUYÊN GIAO DIỆN GỐC (TAILWIND / INTER FONT)
st.set_page_config(layout="wide", page_title="MID Furniture - Quản Lý Tiến Độ")

# Inject CSS để giữ nguyên 100% giao diện, bảng màu Slate/Indigo/Amber và hiệu ứng bảng từ HTML gốc
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;500;600;700;800&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    .report-title { font-weight: 900; text-transform: uppercase; letter-spacing: -0.05em; color: #0f172a; }
    .badge { padding: 4px 10px; border-radius: 9999px; font-size: 11px; font-weight: 700; border: 1px solid; display: inline-block; }
    
    /* Màu sắc trạng thái chuẩn hóa theo file thiết kế gốc */
    .status-complete { background-color: #f0fdf4; color: #15803d; border-color: #bbf7d0; }
    .status-sx { background-color: #eff6ff; color: #1d4ed8; border-color: #bfdbfe; }
    .status-cho-sx { background-color: #eef2ff; color: #4338ca; border-color: #c7d2fe; }
    .status-sang { background-color: #fffbeb; color: #b45309; border-color: #fde68a; }
    .status-quality { background-color: #fff1f2; color: #be123c; border-color: #fecdd3; }
    .status-lapdat { background-color: #faf5ff; color: #6b21a8; border-color: #e9d5ff; }
    .status-pending { background-color: #f8fafc; color: #334155; border-color: #e2e8f0; }
    .status-banve { background-color: #ecfeff; color: #0e7490; border-color: #c5f6fa; }
    
    /* Định dạng cấu trúc Table chuẩn */
    table.custom-table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
    table.custom-table th { background-color: #0f172a; color: white; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; padding: 12px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;}
    table.custom-table td { padding: 12px; border: 1px solid #e2e8f0; color: #334155; }
    table.custom-table tr:nth-child(even) { background-color: #f8fafc; }
    table.custom-table tr:nth-child(odd) { background-color: #ffffff; }
    .project-name { font-weight: 700; color: #0f172a; font-size: 13px; }
    .issue-new { background-color: rgba(254, 243, 199, 0.2); }
    </style>
""", unsafe_allow_html=True)

# 2. KHỞI TẠO CƠ SỞ DỮ LIỆU TRỐNG (KHÔNG CHỨA THÔNG TIN TUẦN 23)
if 'database' not in st.session_state:
    columns = [
        "project", "pic", "contractDate", "leadtime", "loadingDate", 
        "status", "resolvedIssues", "newIssues", "week", "month", "year"
    ]
    st.session_state.database = pd.DataFrame(columns=columns)

# 3. HÀM TỰ ĐỘNG TÍNH TOÁN NGÀY BẮT ĐẦU VÀ KẾT THÚC CỦA TUẦN (TIÊU CHUẨN ISO WEEK)
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

# 4. HÀM ĐỔI MÀU BADGE TRẠNG THÁI THEO THIẾT KẾ GỐC
def get_status_badge(status):
    status_classes = {
        'Complete': 'status-complete',
        'Đang sản xuất': 'status-sx',
        'Chờ lệnh sản xuất': 'status-cho-sx',
        'Chờ hàng sang': 'status-sang',
        'Chờ phản hồi Quality': 'status-quality',
        'Chờ lắp đặt': 'status-lapdat',
        'Pending': 'status-pending',
        'Tiến hành bản vẽ': 'status-banve'
    }
    cls = status_classes.get(status, 'status-pending')
    return f'<span class="badge {cls}">{status}</span>'

# --- GIAO DIỆN CHÍNH ---

# Khối Tiêu đề Hệ thống & Khu vực Nạp File Word độc lập
col_h1, col_h2 = st.columns([2, 1])
with col_h1:
    st.markdown("""
        <div style="background-color: white; padding: 24px; border-radius: 16px; border: 1px solid #f1f5f9; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05);">
            <span style="padding: 4px 12px; font-size: 12px; font-weight: 600; background-color: #ecfdf5; color: #047857; border-radius: 9999px; border: 1px solid #d1fae5; text-transform: uppercase; letter-spacing: 0.05em;">Hệ thống báo cáo</span>
            <h1 class="report-title" style="font-size: 24px; margin-top: 8px; margin-bottom: 0px;">HỆ THỐNG QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
            <p style="font-size: 14px; color: #64748b; margin-top: 4px; margin-bottom: 0px;">MID Furniture System</p>
        </div>
    """, unsafe_allow_html=True)

with col_h2:
    # TÍNH NĂNG 1: Nạp báo cáo hàng tuần bằng file Word (.docx)
    uploaded_file = st.file_uploader("📥 Nạp báo cáo tuần mới (.docx)", type=["docx"])
    if uploaded_file is not None:
        try:
            doc = Document(uploaded_file)
            if len(doc.tables) == 0:
                st.error("Không tìm thấy bảng dữ liệu tiến độ trong file Word!")
            else:
                table = doc.tables[0]
                new_rows = []
                
                # Tự động bóc tách số Tuần và Năm qua tên file (Ví dụ: "Bao_cao_tuan_24_2026.docx")
                file_name = uploaded_file.name
                week_match = re.search(r'tuan_(\d+)', file_name.lower())
                year_match = re.search(r'202\d', file_name)
                
                target_week = int(week_match.group(1)) if week_match else 24
                target_year = int(year_match.group(0)) if year_match else datetime.date.today().year
                
                # Tính toán tự động Tháng tương đối từ số tuần để phục vụ tính năng Tổng hợp theo tháng
                # Thuật toán lấy ngày giữa tuần để xác định tháng chính xác nhất
                first_day = datetime.date(target_year, 1, 1)
                estimated_date = first_day + datetime.timedelta(weeks=target_week - 1) + datetime.timedelta(days=3)
                target_month = estimated_date.month
                
                # Đọc từng hàng trong bảng file Word
                for i, row in enumerate(table.rows):
                    if i == 0: continue # Bỏ qua hàng tiêu đề đầu tiên
                    text_cells = [cell.text.strip() for cell in row.cells]
                    if len(text_cells) >= 8:
                        new_rows.append({
                            "project": text_cells[0], "pic": text_cells[1], "contractDate": text_cells[2],
                            "leadtime": text_cells[3], "loadingDate": text_cells[4], "status": text_cells[5],
                            "resolvedIssues": text_cells[6], "newIssues": text_cells[7],
                            "week": target_week, "month": target_month, "year": target_year
                        })
                
                if new_rows:
                    df_new = pd.DataFrame(new_rows)
                    # Ghi đè dữ liệu mới: Xóa dữ liệu cũ trùng Tuần + Năm trước khi gộp vào Database
                    if not st.session_state.database.empty:
                        st.session_state.database = st.session_state.database[
                            ~((st.session_state.database['week'] == target_week) & (st.session_state.database['year'] == target_year))
                        ]
                    st.session_state.database = pd.concat([st.session_state.database, df_new], ignore_index=True)
                    st.success(f"Nạp thành công dữ liệu {header_week_title if 'header_week_title' in locals() else f'Tuần {target_week}'} ({len(new_rows)} đơn hàng) vào hệ thống!")
        except Exception as e:
            st.error(f"Lỗi khi đọc file Word: {str(e)}. Vui lòng kiểm tra lại định dạng bảng.")

# --- KHU VỰC BỘ LỌC THỜI GIAN VÀ CHẾ ĐỘ TỔNG HỢP ---
st.markdown("<br>", unsafe_allow_html=True)
view_mode = st.radio("Chọn chế độ tổng hợp dữ liệu:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)

db = st.session_state.database.copy()

if db.empty:
    # Hiển thị giao diện hướng dẫn ban đầu khi chưa có dữ liệu
    st.info("Hệ thống hiện tại chưa có dữ liệu. Vui lòng nạp file báo cáo Word (.docx) ở góc phải để bắt đầu làm việc.")
else:
    available_years = sorted(db['year'].unique())
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        selected_year = st.selectbox("Chọn Năm", available_years, index=len(available_years)-1)
        db_filtered = db[db['year'] == selected_year]
        
    if view_mode == "Xem báo cáo theo Tuần":
        with col_f2:
            # TÍNH NĂNG 3: Chọn tuần muốn xem báo cáo
            available_weeks = sorted(db_filtered['week'].unique())
            selected_week = st.selectbox("Chọn Tuần", available_weeks, index=len(available_weeks)-1)
            db_filtered = db_filtered[db_filtered['week'] == selected_week]
        
        week_range = get_week_range_str(selected_week, selected_year)
        # TÍNH NĂNG ĐÒI HỎI: Ghi cụ thể ngày bao nhiêu trên tiêu đề tuần
        header_title = f"Tuần {selected_week} ({selected_year}) | {week_range}"
    else:
        with col_f2:
            # TÍNH NĂNG 2 & 4: Tổng hợp dữ liệu và Chọn xem báo cáo theo tháng
            available_months = sorted(db_filtered['month'].unique())
            selected_month = st.selectbox("Chọn Tháng", available_months, index=len(available_months)-1)
            db_filtered = db_filtered[db_filtered['month'] == selected_month]
        header_title = f"TỔNG HỢP BÁO CÁO TIẾN ĐỘ - THÁNG {selected_month}/{selected_year}"

    # Bộ lọc Người phụ trách & Trạng thái hoạt động động dựa trên dữ liệu tải lên
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

    st.markdown(f"<p style='font-size: 14px; color: #64748b; margin-top:10px;'>Đang hiển thị <b>{len(db_filtered)}</b> đơn hàng.</p>", unsafe_allow_html=True)

    # --- KHU VỰC HIỂN THỊ BẢNG (BẢO LƯU NGUYÊN BẢN GIAO DIỆN GỐC) ---
    html_table = f"""
    <div style="background-color: white; padding: 40px; border-radius: 16px; border: 1px solid #e2e8f0; min-width: 1400px; overflow-x: auto;">
        <div style="display: flex; justify-content: space-between; border-bottom: 2px solid #0f172a; padding-bottom: 24px; margin-bottom: 32px;">
            <div>
                <h2 style="font-size: 24px; font-weight: 900; color: #0f172a; text-transform: uppercase; margin: 0;">Báo Cáo Tiến Độ Đơn Hàng</h2>
                <div style="display: flex; align-items: center; gap: 12px; margin-top: 8px; font-size: 14px; color: #64748b; font-weight: 600;">
                    <span style="background-color: #1e1b4b; color: white; padding: 2px 12px; border-radius: 6px; font-size: 12px; letter-spacing: 0.05em;">{header_title}</span>
                    <span>•</span>
                    <span>MID Furniture System</span>
                </div>
            </div>
            <div style="text-align: right;">
                <span style="font-size: 24px; font-weight: 900; color: #172554; letter-spacing: -0.05em;">MID FURNITURE</span>
            </div>
        </div>

        <table class="custom-table">
            <thead>
                <tr>
                    <th style="width: 12%; text-align: left;">Đơn hàng</th>
                    <th style="width: 7%;">Phụ trách</th>
                    <th style="width: 7%;">Ký HĐ</th>
                    <th style="width: 7%;">Leadtime</th>
                    <th style="width: 9%;">Loading DK</th>
                    <th style="width: 11%;">Trạng thái</th>
                    <th style="width: 22%; text-align: left;">Vấn đề đã có giải pháp</th>
                    <th style="width: 24%; text-align: left;">Vấn đề mới cần giải quyết</th>
                </tr>
            </thead>
            <tbody>
    """

    # Vòng lặp xây dựng các hàng dữ liệu từ file Word đã nạp
    for _, row in db_filtered.iterrows():
        resolved_formatted = str(row['resolvedIssues']).replace('\n', '<br>')
        new_formatted = str(row['newIssues']).replace('\n', '<br>')
        badge_html = get_status_badge(row['status'])
        
        html_table += f"""
            <tr>
                <td class="project-name">{row['project']}</td>
                <td style="text-align: center; font-weight: 600; color: #475569;">{row['pic']}</td>
                <td style="text-align: center; color: #64748b;">{row['contractDate']}</td>
                <td style="text-align: center; color: #64748b;">{row['leadtime']}</td>
                <td style="text-align: center; font-weight: 600;">{row['loadingDate']}</td>
                <td style="text-align: center;">{badge_html}</td>
                <td style="color: #475569;">{resolved_formatted}</td>
                <td class="issue-new" style="color: #475569;">{new_formatted}</td>
            </tr>
        """

    html_table += """
            </tbody>
        </table>
    </div>
    """
    # Xuất bản bảng HTML tùy chỉnh lên giao diện ứng dụng trực tuyến
    st.markdown(html_table, unsafe_allow_html=True)
