import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import xlsxwriter
import plotly.express as px

# --- KẾT NỐI HỆ THỐNG ---
def connect_gsheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_id = "1B5NE0ULV9LFGw6qHNtog4JgjxtA4x2JLYgCXQ6M1P-M" 
        return client.open_by_key(spreadsheet_id).get_worksheet(0)
    except Exception as e:
        st.error(f"Lỗi kết nối: {e}")
        return None

# --- HÀM XUẤT EXCEL LINH HOẠT ---
def export_excel_flexible(df, is_calendar=False):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if is_calendar:
            cols = ['stt', 'team', 'date_time', 'content', 'location', 'host', 'participants', 'note']
            labels = ['STT', 'TỔ', 'THỨ, NGÀY', 'NỘI DUNG ĐĂNG KÝ', 'THỜI GIAN, ĐỊA ĐIỂM', 'CHỦ TRÌ/CHỈ ĐẠO', 'THÀNH PHẦN', 'GHI CHÚ']
        else:
            cols = ['stt', 'team', 'staff', 'content', 'leader', 'progress', 'status', 'product']
            labels = ['STT', 'ĐƠN VỊ/TỔ', 'HỌ VÀ TÊN', 'NỘI DUNG CÔNG VIỆC', 'LÃNH ĐẠO CHỈ ĐẠO', 'TIẾN ĐỘ/THỜI GIAN', 'TRẠNG THÁI', 'SẢN PHẨM']
        
        for c in cols:
            if c not in df.columns: df[c] = ""
        
        df_export = df[cols].copy()
        df_export.columns = labels
        df_export.to_excel(writer, index=False, sheet_name='Data')
        workbook  = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#2980b9', 'font_color': 'white', 'border': 1})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column('A:Z', 20, cell_fmt)
    return output.getvalue()

# --- GIAO DIỆN ---
st.set_page_config(layout="wide", page_title="HỆ THỐNG QUẢN LÝ TỔNG HỢP")
sheet = connect_gsheet()

if sheet:
    raw_data = sheet.get_all_records()
    all_data = pd.DataFrame(raw_data) if raw_data else pd.DataFrame()

    # SIDEBAR
    st.sidebar.header("🔍 BỘ LỌC")
    sel_team = st.sidebar.selectbox("Đơn vị/Tổ:", ["Tổ 1", "Tổ 2", "Tổ 3", "Văn phòng"])
    sel_week = st.sidebar.selectbox("Tuần:", [f"Tuần {str(i).zfill(2)}" for i in range(1, 53)], index=datetime.now().isocalendar()[1]-1)
    sel_type = st.sidebar.selectbox("Loại hình:", ["Đăng ký công việc", "Báo cáo công việc", "Đăng ký lịch tuần"])
    
    staff_list = ["Văn Đức Giao", "Nguyễn Xuân Khánh", "Lê Nguyễn Hạnh Nhi", "Kiều Quang Phương", "Phan Văn Long", "Trần Hoàng Anh", "Trần Hồng Nhung", "Vũ Tuấn Anh", "Bùi Thành Tâm", "Trương Bình Minh", "Hoàng Thị Sinh", "Nguyễn Ngọc Thắng", "Đỗ Hoài Nam", "Lê Tĩnh", "Trương Thị Ngọc Linh", "Tạ Ngọc Thành", "Phùng Hữu Thọ", "Võ Xuân Quý"]
    sel_staff = st.sidebar.selectbox("Cán bộ/Người đăng ký:", staff_list)

    # Lọc dữ liệu cho bảng
    filtered_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
    if sel_type != "Đăng ký lịch tuần":
        filtered_df = filtered_df[filtered_df['staff'] == sel_staff]

    st.header(f"📋 {sel_type}")

    # Gợi ý STT và Form nhập liệu (Giữ nguyên logic cũ bạn đã duyệt)
    # ... (Phần code Form nhập liệu giữ nguyên như bản trước để đảm bảo tính ổn định) ...
    # [TÔI SẼ BỎ QUA ĐOẠN FORM TRONG PHẢN HỒI NÀY ĐỂ TẬP TRUNG VÀO BIỂU ĐỒ - BẠN CỨ GIỮ NGUYÊN FORM CŨ TRONG FILE CỦA BẠN]

    # --- PHẦN KHÔI PHỤC BIỂU ĐỒ ĐÁNH GIÁ HIỆU SUẤT ---
    st.divider()
    st.header("📊 PHÂN TÍCH HIỆU SUẤT CÔNG VIỆC")

    # Chỉ phân tích cho loại hình Báo cáo công việc
    report_data = all_data[(all_data['type'] == "Báo cáo công việc") & (all_data['week'] == sel_week)]

    if not report_data.empty:
        col_chart1, col_chart2 = st.columns(2)

        with col_chart1:
            st.subheader(f"🎯 Hiệu suất theo Tổ ({sel_week})")
            team_stats = report_data.groupby(['team', 'status']).size().reset_index(name='counts')
            fig_team = px.bar(team_stats, x='team', y='counts', color='status', 
                             title="Trạng thái công việc theo từng Tổ",
                             color_discrete_map={"🔵 Mới": "#3498db", "🟢 Hoàn thành": "#2ecc71", "🔴 Trễ hạn": "#e74c3c", "🟡 Đang làm": "#f1c40f"},
                             barmode='group')
            st.plotly_chart(fig_team, use_container_width=True)

        with col_chart2:
            st.subheader(f"👤 Hiệu suất cá nhân: {sel_staff}")
            ind_data = report_data[report_data['staff'] == sel_staff]
            if not ind_data.empty:
                ind_stats = ind_data['status'].value_counts().reset_index()
                ind_stats.columns = ['status', 'count']
                fig_ind = px.pie(ind_stats, values='count', names='status', 
                                title=f"Tỷ lệ hoàn thành của {sel_staff}",
                                color='status',
                                color_discrete_map={"🔵 Mới": "#3498db", "🟢 Hoàn thành": "#2ecc71", "🔴 Trễ hạn": "#e74c3c", "🟡 Đang làm": "#f1c40f"})
                st.plotly_chart(fig_ind, use_container_width=True)
            else:
                st.info(f"Chưa có dữ liệu báo cáo tuần này cho {sel_staff}")

        # BIỂU ĐỒ TỔNG HỢP TOÀN ĐƠN VỊ
        st.subheader("📈 Tổng hợp trạng thái công việc toàn đơn vị")
        total_stats = report_data['status'].value_counts().reset_index()
        total_stats.columns = ['status', 'count']
        fig_total = px.bar(total_stats, x='status', y='count', color='status',
                          text_auto=True,
                          color_discrete_map={"🔵 Mới": "#3498db", "🟢 Hoàn thành": "#2ecc71", "🔴 Trễ hạn": "#e74c3c", "🟡 Đang làm": "#f1c40f"})
        st.plotly_chart(fig_total, use_container_width=True)
    else:
        st.warning("Không có dữ liệu báo cáo trong tuần này để hiển thị biểu đồ.")

    # (Phần xuất Excel giữ nguyên bên dưới)
