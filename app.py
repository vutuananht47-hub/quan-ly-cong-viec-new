import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import xlsxwriter

# --- KẾT NỐI HỆ THỐNG ---
def connect_gsheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        # Đọc từ Secrets (Fix lỗi Base64 và JWT)
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_id = "1B5NE0ULV9LFGw6qHNtog4jgjxtA4x2JLYgCXQ6M1P-M" 
        return client.open_by_key(spreadsheet_id).get_worksheet(0)
    except Exception as e:
        st.error(f"Lỗi kết nối hệ thống: {e}")
        return None

# --- HÀM XUẤT EXCEL CHUẨN (Fix lỗi file hỏng) ---
def export_excel_styled(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='BaoCao')
        workbook  = writer.book
        worksheet = writer.sheets['BaoCao']
        header_fmt = workbook.add_format({'bold': True, 'fg_color': '#2980b9', 'font_color': 'white', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
    return output.getvalue()

# --- PHẦN GIAO DIỆN CHÍNH ---
st.set_page_config(layout="wide", page_title="QUẢN LÝ CÔNG VIỆC")
sheet = connect_gsheet()

if sheet:
    # (Toàn bộ logic xử lý dữ liệu, lọc, thêm, sửa, xóa giữ nguyên như bản trước)
    st.success("Hệ thống đã sẵn sàng!")
    # ... dán tiếp phần code hiển thị và form nhập liệu vào đây ...
