import datetime
import re
import streamlit as st
import pandas as pd
from docx import Document
import streamlit.components.v1 as components
from streamlit_gsheets import GSheetsConnection

# 1. CẤU HÌNH TRANG NGƯỜI DÙNG
st.set_page_config(layout="wide", page_title="MID Furniture - Quản Lý Tiến Độ")

COLUMNS = [
    "project", "pic", "contractDate", "leadtime", "loadingDate", 
    "status", "resolvedIssues", "newIssues", "week", "month", "year"
]

# 2. KHỞI TẠO KẾT NỐI ĐỌC GOOGLE SHEETS
database_df = pd.DataFrame(columns=COLUMNS)
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    database_df = conn.read(worksheet="database", ttl="0")
    database_df = database_df.dropna(how="all")
    for col in COLUMNS:
        if col not in database_df.columns:
            database_df[col] = "-"
    database_df = database_df[COLUMNS]
except Exception as e:
    st.warning("⚠️ Đang đợi cấu hình hoặc đồng bộ dữ liệu ban đầu từ Google Sheets...")

# --- GIAO DIỆN CHÍNH ---

col_h1, col_h2 = st.columns([2, 1])
with col_h1:
    st.markdown("""
        <div style="background-color: white; padding: 24px; border-radius: 16px; border: 1px solid #f1f5f9; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05); font-family: 'Segoe UI', Roboto, sans-serif;">
            <span style="padding: 4px 12px; font-size: 12px; font-weight: 600; background-color: #ecfdf5; color: #047857; border-radius: 9999px; border: 1px solid #d1fae5; text-transform: uppercase; letter-spacing: 0.05em;">Hệ thống báo cáo</span>
            <h1 style="font-weight: 900; text-transform: uppercase; letter-spacing: -0.05em; color: #0f172a; font-size: 24px; margin-top: 8px; margin-bottom: 0px;">HỆ THỐNG QUẢN LÝ TIẾN ĐỘ ĐƠN HÀNG</h1>
            <p style="font-size: 14px; color: #64748b; margin-top: 4px; margin-bottom: 0px;">MID Furniture System (Cloud Database)</p>
        </div>
    """, unsafe_allow_html=True)

with col_h2:
    uploaded_file = st.file_uploader("📥 Nạp báo cáo tuần mới (.docx)", type=["docx"])
    if uploaded_file is not None:
        try:
            doc = Document(uploaded_file)
            if len(doc.tables) == 0:
                st.error("Không tìm thấy bảng dữ liệu tiến độ trong file Word!")
            else:
                table = doc.tables[0]
                new_rows = []
                
                file_name = uploaded_file.name
                week_match = re.search(r'tuan_(\d+)', file_name.lower())
                year_match = re.search(r'202\d', file_name)
                
                target_week = int(week_match.group(1)) if week_match else 23
                target_year = int(year_match.group(0)) if year_match else datetime.date.today().year
                
                first_day = datetime.date(target_year, 1, 1)
                estimated_date = first_day + datetime.timedelta(weeks=target_week - 1) + datetime.timedelta(days=3)
                target_month = estimated_date.month
                
                for i, row in enumerate(table.rows):
                    if i == 0: continue 
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
                    if not database_df.empty:
                        database_df['week'] = pd.to_numeric(database_df['week'], errors='coerce').fillna(0).astype(int)
                        database_df['year'] = pd.to_numeric(database_df['year'], errors='coerce').fillna(0).astype(int)
                        updated_df = database_df[
                            ~((database_df['week'] == target_week) & (database_df['year'] == target_year))
                        ]
                    else:
                        updated_df = pd.DataFrame(columns=COLUMNS)
                        
                    final_df = pd.concat([updated_df, df_new], ignore_index=True)
                    
                    st.success(f"🎉 Đã gộp thành công dữ liệu Tuần {target_week}!")
                    
                    csv_data = final_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="💾 Bấm vào đây để tải File cập nhật dữ liệu tổng (.csv)",
                        data=csv_data,
                        file_name="database.csv",
                        mime="text/csv"
                    )
                    st.info("💡 **Hành động tiếp theo:** Tải file .csv ở trên về -> Vào Google Sheets của bạn -> Chọn Tệp -> Nhập -> Tải lên file này và chọn 'Thay thế trang tính hiện tại'.")
        except Exception as e:
            st.error(f"Lỗi xử lý file Word: {str(e)}")

# --- KHU VỰC BỘ LỌC THỜI GIAN VÀ CHẾ ĐỘ T
