import datetime
import re
import streamlit as st
import pandas as pd
from docx import Document

# 1. CẤU HÌNH TRANG & GIỮ NGUYÊN GIAO DIỆN (TAILWIND/INTER FONT)
st.set_page_config(layout="wide", page_title="MID Furniture - Quản Lý Tiến Độ")

# Inject CSS để đồng bộ giao diện như file HTML của bạn
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    .report-title { font-weight: 900; text-transform: uppercase; letter-spacing: -0.05em; color: #0f172a; }
    .badge { padding: 4px 10px; border-radius: 9999px; font-size: 11px; font-weight: 700; border: 1px solid; display: inline-block; }
    
    /* Màu sắc trạng thái đồng bộ với file HTML gốc */
    .status-complete { background-color: #f0fdf4; color: #15803d; border-color: #bbf7d0; }
    .status-sx { background-color: #eff6ff; color: #1d4ed8; border-color: #bfdbfe; }
    .status-cho-sx { background-color: #eef2ff; color: #4338ca; border-color: #c7d2fe; }
    .status-sang { background-color: #fffbeb; color: #b45309; border-color: #fde68a; }
    .status-quality { background-color: #fff1f2; color: #be123c; border-color: #fecdd3; }
    .status-lapdat { background-color: #faf5ff; color: #6b21a8; border-color: #e9d5ff; }
    .status-pending { background-color: #f8fafc; color: #334155; border-color: #e2e8f0; }
    .status-banve { background-color: #ecfeff; color: #0e7490; border-color: #c5f6fa; }
    
    /* Định dạng bảng giống hệt bản gốc */
    table.custom-table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
    table.custom-table th { background-color: #0f172a; color: white; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; padding: 12px; border: 1px solid #cbd5e1; font-size: 11px; text-align: center;}
    table.custom-table td { padding: 12px; border: 1px solid #e2e8f0; color: #334155; }
    table.custom-table tr:nth-child(even) { background-color: #f8fafc; }
    table.custom-table tr:nth-child(odd) { background-color: #ffffff; }
    .project-name { font-weight: 700; color: #0f172a; font-size: 13px; }
    .issue-new { background-color: rgba(254, 243, 199, 0.2); }
    </style>
""", unsafe_allow_html=True)

# 2. HÀM MÔ PHỎNG / KHỞI TẠO DỮ LIỆU GỐC TỪ FILE HTML (TUẦN 23 - NĂM 2026)
if 'database' not in st.session_state:
    # Khởi tạo dữ liệu mẫu tuần 23 ban đầu
    initial_data = [
        {"project": "CT06 + CT07", "pic": "Điệp", "contractDate": "-", "leadtime": "-", "loadingDate": "-", "status": "Đang sản xuất", "resolvedIssues": "Đã tiến hành cắt hàng mã CT6; hoàn thiện đơn trong tháng 5.", "newIssues": "Dự kiến ready vào 13/6.\nLưu ý: CT06+07 Inox Gold LED vàng, Inox Silver LED trắng. Cần thử 100% đèn trước khi đóng gói.", "week": 23, "month": 6, "year": 2026},
        {"project": "EUTEK 02", "pic": "Điệp", "contractDate": "-", "leadtime": "25/05/2026", "loadingDate": "25/06/2026", "status": "Đang sản xuất", "resolvedIssues": "Đã chốt mẫu đá Yate bạch ngọc; tài liệu SX ngày 29/04.", "newIssues": "Cập nhật hàng ready: 20/6. Lưu ý: (1) Mặt đá rời để FNB tự lắp; (2) Mặt đá đơn FNB 1 bị vỡ; (3) Thí nghiệm sơn cạnh đá thành công.", "week": 23, "month": 6, "year": 2026},
        {"project": "Signature Nail 02 (DAP)", "pic": "Tùng", "contractDate": "20/05/2026", "leadtime": "-", "loadingDate": "03/07/2026", "status": "Đang sản xuất", "resolvedIssues": "Đã chốt đơn và cọc 21/05/2026. Tiến hành cắt ván từ 29/05.", "newIssues": "Dự kiến loading 3/7. Bổ sung 1 cụm mặt cho đơn 1 bị vỡ mặt đá.", "week": 23, "month": 6, "year": 2026},
        {"project": "NT IRELAND 02", "pic": "Điệp", "contractDate": "14/05/2026", "leadtime": "-", "loadingDate": "10/07/2026", "status": "Đang sản xuất", "resolvedIssues": "LSX ngày 18/05; chốt bản vẽ cắt ván ngày 25/05.", "newIssues": "Dự kiến loading 1/7. Rủi ro tiến độ do khách hàng chậm chuyển tiền cọc.", "week": 23, "month": 6, "year": 2026},
        {"project": "Anh Thư", "pic": "Tùng", "contractDate": "26/05/2026", "leadtime": "-", "loadingDate": "30/07/2026", "status": "Chờ lệnh sản xuất", "resolvedIssues": "Đã chuyển cọc ngày 30/05.", "newIssues": "Ngày 3/6 đã chính thức nhận cọc vào tài khoản VND.", "week": 23, "month": 6, "year": 2026},
        {"project": "Ty Dang Khau", "pic": "Tùng", "contractDate": "-", "leadtime": "-", "loadingDate": "-", "status": "Chờ lệnh sản xuất", "resolvedIssues": "Khách chuyển cọc 03/06.", "newIssues": "Đang tiến hành thực hiện bản vẽ; chờ khách chốt hồ sơ kỹ thuật và mã ván.", "week": 23, "month": 6, "year": 2026},
        {"project": "NASHVILLE 1-2 (Đơn 2)", "pic": "Tùng", "contractDate": "08/04/2026", "leadtime": "-", "loadingDate": "26/04/2026", "status": "Chờ hàng sang", "resolvedIssues": "Thông quan, sailing 13/05. Sailing Thâm Quyến 27/5.", "newIssues": "ETA LAX 9/6, ETA Nashville 15/6. Cần khẩn trương giục khách thanh toán đơn hàng.", "week": 23, "month": 6, "year": 2026},
        {"project": "CT04.2 + CT05.1", "pic": "Điệp", "contractDate": "-", "leadtime": "28/03/2026", "loadingDate": "30/03/2026", "status": "Chờ hàng sang", "resolvedIssues": "Sailing 26/04; thanh toán LCC. Thiếu móc treo bàn CT05.", "newIssues": "3/6 gửi BL release. 7/6 hạ cảng Houston, chờ giao cont cho khách.", "week": 23, "month": 6, "year": 2026},
        {"project": "Lumie", "pic": "Điệp", "contractDate": "06/02/2026", "leadtime": "06/04/2026", "loadingDate": "16/04/2026", "status": "Chờ hàng sang", "resolvedIssues": "Sailing 27/04.", "newIssues": "ETA Savannah 25/6, ETA Atlanta 29/6. Lưu ý gửi PKL và hướng dẫn lắp đặt.", "week": 23, "month": 6, "year": 2026},
        {"project": "CT05+06+07 & Lẻ 01", "pic": "Điệp", "contractDate": "-", "leadtime": "09/04/2026", "loadingDate": "13/04/2026", "status": "Chờ hàng sang", "resolvedIssues": "Đã bổ sung 12 pedicart; sailing 26/05.", "newIssues": "4/6 sailing từ Thượng Hải, dự kiến ETA 5/7. Lưu ý gửi PKL chi tiết.", "week": 23, "month": 6, "year": 2026},
        {"project": "Ngân Cao 02", "pic": "Tùng", "contractDate": "20/04/2026", "leadtime": "27/05/2026", "loadingDate": "30/05/2026", "status": "Chờ hàng sang", "resolvedIssues": "Đã cắt ván và lắp ráp tại xưởng.", "newIssues": "4/6 hàng hạ cảng, thông quan. ETD chính thức 6/6.", "week": 23, "month": 6, "year": 2026},
        {"project": "HM3 (Spain)", "pic": "TA", "contractDate": "15/01/2026", "leadtime": "-", "loadingDate": "20/01/2026", "status": "Chờ phản hồi Quality", "resolvedIssues": "Giao hàng 08/05. Tấm khay dây điện đã thiết kế lại.", "newIssues": "Đã xác nhận lỗi claim và đề xuất đền bù 4,000 EUR; khách chưa xác nhận.", "week": 23, "month": 6, "year": 2026},
        {"project": "Sim Trần", "pic": "Hiếu", "contractDate": "-", "leadtime": "-", "loadingDate": "17/12/2025", "status": "Complete", "resolvedIssues": "Khách thay quầy đảo; LED chờ lắp đặt.", "newIssues": "Chưa thay nguồn LED tủ kệ do tủ nặng và thiếu người hỗ trợ.", "week": 23, "month": 6, "year": 2026},
        {"project": "KD Tùng 1.1 (UK)", "pic": "TA", "contractDate": "25/01/2026", "leadtime": "-", "loadingDate": "30/01/2026", "status": "Chờ phản hồi Quality", "resolvedIssues": "Giao hàng 16/04. Giao giắc LED 18/05.", "newIssues": "Hoàn thiện 120 ốp mặt và 180 ốp chân để chuyển Air sang Anh dự kiến 13/06.", "week": 23, "month": 6, "year": 2026},
        {"project": "CT03.1", "pic": "Điệp", "contractDate": "12/01/2026", "leadtime": "-", "loadingDate": "10/02/2026", "status": "Chờ phản hồi Quality", "resolvedIssues": "Gửi bù 12 vương miện inox và 20 túi vật tư.", "newIssues": "1/6 gửi thùng hàng bù đi HCM. Đang kiểm tra nhà xe.", "week": 23, "month": 6, "year": 2026},
        {"project": "Pink Cactus Nail", "pic": "Hiếu", "contractDate": "-", "leadtime": "-", "loadingDate": "-", "status": "Chờ lắp đặt", "resolvedIssues": "Đã thay nguồn 1 kệ; 1 kệ còn lại hỏng LED.", "newIssues": "Đang tìm LED thay thế tại Mỹ. Nếu không có sẽ gửi từ VN.", "week": 23, "month": 6, "year": 2026},
        {"project": "HM4-Bỉ", "pic": "Điệp", "contractDate": "-", "leadtime": "-", "loadingDate": "-", "status": "Pending", "resolvedIssues": "Đang tính giá.", "newIssues": "Chờ xử lý dứt điểm khiếu nại HM3.", "week": 23, "month": 6, "year": 2026},
        {"project": "Olivia Anh Dang", "pic": "Điệp", "contractDate": "-", "leadtime": "-", "loadingDate": "-", "status": "Tiến hành bản vẽ", "resolvedIssues": "Đã vẽ và gửi báo giá.", "newIssues": "Bổ sung WC và mặt tiền. KTS đang thực hiện.", "week": 23, "month": 6, "year": 2026},
        {"project": "Kim Thi Liên", "pic": "Tùng", "contractDate": "-", "leadtime": "-", "loadingDate": "-", "status": "Tiến hành bản vẽ", "resolvedIssues": "Đã tính giá theo bản vẽ Mr Lân.", "newIssues": "Chờ sửa lại bản vẽ 3D.", "week": 23, "month": 6, "year": 2026},
        {"project": "Jennifer Phạm", "pic": "Tùng", "contractDate": "14/04/2026", "leadtime": "-", "loadingDate": "-", "status": "Tiến hành bản vẽ", "resolvedIssues": "Đã vẽ và báo giá.", "newIssues": "Đang điều chỉnh chi tiết nhỏ trước khi chốt đơn.", "week": 23, "month": 6, "year": 2026}
    ]
    st.session_state.database = pd.DataFrame(initial_data)

# 3. HÀM TÍNH NGÀY ĐẦU VÀ CUỐI CỦA TUẦN (ISO WEEK)
def get_week_range_str(week_num, year):
    # Tìm ngày đầu tiên của năm, xác định ngày thuộc tuần yêu cầu
    first_day_of_year = datetime.date(year, 1, 1)
    if first_day_of_year.weekday() > 3:
        first_monday = first_day_of_year + datetime.timedelta(days=(7 - first_day_of_year.weekday()))
    else:
        first_monday = first_day_of_year - datetime.timedelta(days=first_day_of_year.weekday())
    
    start_date = first_monday + datetime.timedelta(weeks=week_num - 1)
    end_date = start_date + datetime.timedelta(days=6)
    return f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"

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

# Khối tiêu đề Hệ thống báo cáo (Bên trái) & Upload File Word (Bên phải)
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
    # Tính năng 1: Nạp báo cáo hàng tuần bằng file Word (.docx)
    uploaded_file = st.file_uploader("📥 Nạp báo cáo tuần mới (.docx)", type=["docx"])
    if uploaded_file is not None:
        try:
            doc = Document(uploaded_file)
            # Giả định file Word chứa 1 bảng có cấu trúc tương đương cấu trúc dữ liệu chính
            table = doc.tables[0]
            new_rows = []
            
            # Đọc tên file để nhận diện tuần (Ví dụ: "Bao_cao_tuan_24_2026.docx")
            file_name = uploaded_file.name
            week_match = re.search(r'tuan_(\d+)', file_name.lower())
            year_match = re.search(r'202\d', file_name)
            
            target_week = int(week_match.group(1)) if week_match else 24
            target_year = int(year_match.group(0)) if year_match else 2026
            target_month = 6 # Tự động map hoặc gán tính toán từ tuần
            
            for i, row in enumerate(table.rows):
                if i == 0: continue # Bỏ qua header của bảng trong file Word
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
                # Xóa dữ liệu cũ của tuần này nếu nạp đè, sau đó append vào Database chung
                st.session_state.database = st.session_state.database[
                    ~((st.session_state.database['week'] == target_week) & (st.session_state.database['year'] == target_year))
                ]
                st.session_state.database = pd.concat([st.session_state.database, df_new], ignore_index=True)
                st.success(f"Nạp thành công dữ liệu Tuần {target_week} ({len(new_rows)} đơn hàng)!")
        except Exception as e:
            st.error(f"Lỗi cấu trúc File Word: {str(e)}")

# --- KHU VỰC CHỌN CHẾ ĐỘ XEM & BỘ LỌC THỜI GIAN ---
st.markdown("<br>", unsafe_allow_html=True)
view_mode = st.radio("Chọn chế độ tổng hợp dữ liệu:", ["Xem báo cáo theo Tuần", "Tổng hợp báo cáo theo Tháng"], horizontal=True)

# Lấy các danh mục duy nhất để làm cấu trúc lọc thời gian
db = st.session_state.database.copy()
available_years = sorted(db['year'].unique())

col_f1, col_f2, col_f3, col_f4 = st.columns(4)

with col_f1:
    selected_year = st.selectbox("Chọn Năm", available_years, index=0)
    db_filtered = db[db['year'] == selected_year]

if view_mode == "Xem báo cáo theo Tuần":
    with col_f2:
        # Tính năng 3: Chọn tuần muốn xem báo cáo
        available_weeks = sorted(db_filtered['week'].unique())
        selected_week = st.selectbox("Chọn Tuần", available_weeks, index=0)
        db_filtered = db_filtered[db_filtered['week'] == selected_week]
    
    # Định dạng tiêu đề tuần yêu cầu cụ thể ngày bao nhiêu: “Tuần X (Năm) | DD/MM-DD/MM”
    week_range = get_week_range_str(selected_week, selected_year)
    header_title = f"TUẦN {selected_week} ({selected_year}) | {week_range}"
else:
    with col_f2:
        # Tính năng 4: Chọn tháng muốn xem báo cáo tổng hợp
        available_months = sorted(db_filtered['month'].unique())
        selected_month = st.selectbox("Chọn Tháng", available_months, index=0)
        db_filtered = db_filtered[db_filtered['month'] == selected_month]
    header_title = f"TỔNG HỢP TIẾN ĐỘ ĐƠN HÀNG - THÁNG {selected_month}/{selected_year}"

# Các bộ lọc nhanh theo Người phụ trách và Trạng thái (Giữ nguyên tính năng lọc nhanh từ HTML gốc)
with col_f3:
    all_pics = ["Tất cả Người phụ trách"] + list(db_filtered['pic'].unique())
    selected_pic = st.selectbox("Lọc nhanh Người phụ trách:", all_pics)
    if selected_pic != "Tất cả Người phụ trách":
        db_filtered = db_filtered[db_filtered['pic'] == selected_pic]

with col_f4:
    all_status = ["Tất cả Trạng thái"] + list(db_filtered['status'].unique())
    selected_status = st.selectbox("Lọc nhanh Trạng thái:", all_status)
    if selected_status != "Tất cả Trạng thái":
        db_filtered = db_filtered[db_filtered['status'] == selected_status]

# Hiển thị số lượng bản ghi đang lọc giống HTML gốc
st.markdown(f"<p style='font-size: 14px; color: #64748b; margin-top:10px;'>Đang hiển thị <b>{len(db_filtered)}</b>/{len(db_filtered)} đơn hàng tương ứng bộ lọc.</p>", unsafe_allow_html=True)

# --- KHU VỰC HIỂN THỊ BẢNG (BẢO LƯU GIAO DIỆN GỐC) ---
# Header Area bên trong bảng
html_table = f"""
<div style="background-color: white; padding: 40px; border-radius: 16px; border: 1px solid #e2e8f0; min-width: 1200px; overflow-x: auto;">
    <div style="display: flex; justify-content: space-between; border-bottom: 2px solid #0f172a; padding-bottom: 24px; margin-bottom: 32px;">
        <div>
            <h2 style="font-size: 24px; font-weight: 900; color: #0f172a; text-transform: uppercase; margin: 0;">Báo Cáo Tiến Độ Đơn Hàng</h2>
            <div style="display: flex; align-items: center; gap: 12px; margin-top: 8px; font-size: 14px; color: #64748b; font-weight: 600;">
                <span style="background-color: #1e1b4b; color: white; padding: 2px 12px; border-radius: 6px; text-transform: uppercase; font-size: 12px; letter-spacing: 0.05em;">{header_title}</span>
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

# Tạo các dòng cho bảng
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

# Render bảng HTML tùy chỉnh lên Streamlit công khai
st.markdown(html_table, unsafe_allow_html=True)
