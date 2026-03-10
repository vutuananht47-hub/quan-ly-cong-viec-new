import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import xlsxwriter

# --- CẤU HÌNH TRANG ---
st.set_page_config(layout="wide", page_title="QUẢN LÝ CÔNG VIỆC WEB")

# --- KẾT NỐI GOOGLE SHEETS ---
def connect_gsheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        # Đọc trực tiếp từ Secrets của Streamlit
        creds_dict = st.secrets["gcp_service_account"] 
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_id = "1B5NE0ULV9LFGw6qHNtog4jgjxtA4x2JLYgCXQ6M1P-M" 
        return client.open_by_key(spreadsheet_id).get_worksheet(0)
    except Exception as e:
        st.error(f"Lỗi kết nối hệ thống: {e}")
        return None

sheet = connect_gsheet()

# --- HÀM XUẤT EXCEL CÓ MÀU & KẺ BẢNG ---
def export_excel_styled(df):
    cols_order = ['stt', 'team', 'staff', 'content', 'leader', 'progress', 'status']
    for c in cols_order:
        if c not in df.columns: df[c] = ""
    
    df_export = df[cols_order].copy()
    df_export.columns = ['STT', 'ĐƠN VỊ/TỔ', 'HỌ VÀ TÊN', 'NỘI DUNG CÔNG VIỆC', 
                         'LÃNH ĐẠO CHỈ ĐẠO', 'TIẾN ĐỘ/THỜI GIAN', 'TRẠNG THÁI']

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='BaoCao')
        workbook  = writer.book
        worksheet = writer.sheets['BaoCao']

        # Định dạng Header
        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#2980b9', 'font_color': 'white', 'border': 1
        })
        # Định dạng Nội dung (Kẻ bảng)
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})

        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        worksheet.set_column('A:A', 6, cell_fmt)
        worksheet.set_column('B:C', 20, cell_fmt)
        worksheet.set_column('D:D', 55, cell_fmt)
        worksheet.set_column('E:G', 20, cell_fmt)

    return output.getvalue()

# --- XỬ LÝ DỮ LIỆU ---
if sheet:
    raw_data = sheet.get_all_records()
    all_data = pd.DataFrame(raw_data) if raw_data else pd.DataFrame(columns=['team', 'type', 'week', 'staff', 'stt', 'content', 'leader', 'progress', 'status'])

    # SIDEBAR
    st.sidebar.header("🔍 BỘ LỌC HỆ THỐNG")
    sel_team = st.sidebar.selectbox("Đơn vị/Tổ:", ["Tổ 1", "Tổ 2", "Tổ 3", "Văn phòng"])
    sel_week = st.sidebar.selectbox("Tuần:", [f"Tuần {str(i).zfill(2)}" for i in range(1, 53)], 
                                     index=datetime.now().isocalendar()[1]-1)
    sel_type = st.sidebar.selectbox("Loại hình:", ["Đăng ký công việc", "Báo cáo công việc"])
    
    staff_list = ["Văn Đức Giao", "Nguyễn Xuân Khánh", "Lê Nguyễn Hạnh Nhi", "Kiều Quang Phương", "Phan Văn Long", "Trần Hoàng Anh", "Trần Hồng Nhung", "Vũ Tuấn Anh", "Bùi Thành Tâm", "Trương Bình Minh", "Hoàng Thị Sinh", "Nguyễn Ngọc Thắng", "Đỗ Hoài Nam", "Lê Tĩnh", "Trương Thị Ngọc Linh", "Tạ Ngọc Thành", "Phùng Hữu Thọ", "Võ Xuân Quý"]
    sel_staff = st.sidebar.selectbox("Tên cán bộ:", staff_list)

    # Lọc dữ liệu hiển thị
    filtered_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & 
                           (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff)]

    # --- NHẬP / SỬA / XÓA ---
    st.header(f"📋 Cán bộ: {sel_staff}")
    
    options = ["-- Thêm mới --"] + filtered_df['stt'].astype(str).tolist()
    selected_stt = st.selectbox("Chọn STT để Sửa hoặc Xóa:", options)

    with st.form("input_form"):
        c1, c2, c3 = st.columns([1, 2, 1])
        
        if selected_stt != "-- Thêm mới --":
            row = filtered_df[filtered_df['stt'].astype(str) == selected_stt].iloc[0]
            v_stt, v_leader, v_status = str(row['stt']), row['leader'], row['status']
            v_content, v_progress = row['content'], row['progress']
        else:
            v_stt, v_leader, v_status, v_content, v_progress = "", "", "🔵 Mới", "", ""

        stt = c1.text_input("STT", value=v_stt)
        leader = c2.text_input("Lãnh đạo chỉ đạo", value=v_leader)
        status = c3.selectbox("Trạng thái", ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"], 
                              index=["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"].index(v_status))
        content = st.text_area("Nội dung công việc", value=v_content)
        progress = c1.text_input("Tiến độ/Thời gian", value=v_progress)
        
        save = st.form_submit_button("💾 LƯU DỮ LIỆU")
        delete = st.form_submit_button("🗑️ XÓA DÒNG NÀY")

        if save:
            new_data = [sel_team, sel_type, sel_week, sel_staff, stt, content, leader, progress, status]
            if selected_stt == "-- Thêm mới --":
                sheet.append_row(new_data)
            else:
                idx = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & 
                               (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff) & 
                               (all_data['stt'].astype(str) == selected_stt)].index[0]
                sheet.update(f"A{idx+2}:I{idx+2}", [new_data])
            st.success("Thành công!")
            st.rerun()

        if delete and selected_stt != "-- Thêm mới --":
            idx = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & 
                           (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff) & 
                           (all_data['stt'].astype(str) == selected_stt)].index[0]
            sheet.delete_rows(idx + 2)
            st.warning("Đã xóa!")
            st.rerun()

    st.subheader("Dữ liệu hiện tại")
    st.dataframe(filtered_df, use_container_width=True)

    # --- XUẤT EXCEL VỚI TÊN FILE THÔNG MINH ---
    st.divider()
    st.subheader("📥 XUẤT BÁO CÁO EXCEL")
    
    # Xử lý tên file: Loại bỏ dấu tiếng Việt để tránh lỗi trình duyệt (tùy chọn)
    type_name = "DangKy" if "Đăng ký" in sel_type else "BaoCao"
    
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        team_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not team_df.empty:
            st.download_button(f"📥 Tải Excel {sel_team}", 
                               data=export_excel_styled(team_df), 
                               file_name=f"{type_name}_{sel_team.replace(' ','')}_{sel_week.replace(' ','')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    with col_e2:
        unit_df = all_data[(all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not unit_df.empty:
            st.download_button("📥 Tải Excel Toàn Đơn Vị", 
                               data=export_excel_styled(unit_df), 
                               file_name=f"{type_name}_ToanDonVi_{sel_week.replace(' ','')}.xlsx",

                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
