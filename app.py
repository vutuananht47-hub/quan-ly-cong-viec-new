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
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_id = "1B5NE0ULV9LFGw6qHNtog4JgjxtA4x2JLYgCXQ6M1P-M" 
        return client.open_by_key(spreadsheet_id).get_worksheet(0)
    except Exception as e:
        st.error(f"Lỗi kết nối: {e}")
        return None

# --- HÀM XUẤT EXCEL ---
def export_excel_styled(df):
    cols_order = ['stt', 'team', 'staff', 'content', 'leader', 'progress', 'status', 'product']
    for c in cols_order:
        if c not in df.columns: df[c] = ""
    
    df_export = df[cols_order].copy()
    df_export.columns = ['STT', 'ĐƠN VỊ/TỔ', 'HỌ VÀ TÊN', 'NỘI DUNG CÔNG VIỆC', 
                         'LÃNH ĐẠO CHỈ ĐẠO', 'TIẾN ĐỘ/THỜI GIAN', 'TRẠNG THÁI', 'SẢN PHẨM']

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='BaoCao')
        workbook  = writer.book
        worksheet = writer.sheets['BaoCao']
        header_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'fg_color': '#2980b9', 'font_color': 'white', 'border': 1})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column('A:A', 6, cell_fmt)
        worksheet.set_column('B:C', 18, cell_fmt)
        worksheet.set_column('D:D', 45, cell_fmt)
        worksheet.set_column('E:G', 18, cell_fmt)
        worksheet.set_column('H:H', 40, cell_fmt)
    return output.getvalue()

# --- GIAO DIỆN ---
st.set_page_config(layout="wide", page_title="QUẢN LÝ CÔNG VIỆC")
sheet = connect_gsheet()

if sheet:
    raw_data = sheet.get_all_records()
    all_data = pd.DataFrame(raw_data) if raw_data else pd.DataFrame()

    # SIDEBAR
    st.sidebar.header("🔍 BỘ LỌC")
    sel_team = st.sidebar.selectbox("Đơn vị/Tổ:", ["Tổ 1", "Tổ 2", "Tổ 3", "Văn phòng"])
    sel_week = st.sidebar.selectbox("Tuần:", [f"Tuần {str(i).zfill(2)}" for i in range(1, 53)], index=datetime.now().isocalendar()[1]-1)
    sel_type = st.sidebar.selectbox("Loại hình:", ["Đăng ký công việc", "Báo cáo công việc"])
    staff_list = ["Văn Đức Giao", "Nguyễn Xuân Khánh", "Lê Nguyễn Hạnh Nhi", "Kiều Quang Phương", "Phan Văn Long", "Trần Hoàng Anh", "Trần Hồng Nhung", "Vũ Tuấn Anh", "Bùi Thành Tâm", "Trương Bình Minh", "Hoàng Thị Sinh", "Nguyễn Ngọc Thắng", "Đỗ Hoài Nam", "Lê Tĩnh", "Trương Thị Ngọc Linh", "Tạ Ngọc Thành", "Phùng Hữu Thọ", "Võ Xuân Quý"]
    sel_staff = st.sidebar.selectbox("Cán bộ:", staff_list)

    filtered_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff)] if not all_data.empty else pd.DataFrame()

    st.header(f"📋 {sel_type}: {sel_staff}")

    # Gợi ý STT
    suggested_stt = 1
    if not filtered_df.empty:
        try:
            suggested_stt = int(pd.to_numeric(filtered_df['stt']).max()) + 1
        except:
            suggested_stt = len(filtered_df) + 1

    options = ["-- Thêm mới --"] + sorted(filtered_df['stt'].astype(str).tolist(), key=lambda x: int(x) if x.isdigit() else 999)
    selected_stt = st.selectbox("Chọn STT để Sửa/Xóa:", options)

    # THIẾT LẬP DỮ LIỆU ĐẦU VÀO TRƯỚC FORM
    if selected_stt != "-- Thêm mới --":
        row = filtered_df[filtered_df['stt'].astype(str) == selected_stt].iloc[0]
        v_stt = str(row['stt'])
        v_leader = row['leader']
        v_status = row['status']
        v_content = row['content']
        v_progress = row['progress']
        v_product = str(row.get('product', ""))
    else:
        v_stt = str(suggested_stt)
        v_leader = ""
        v_status = "🔵 Mới"
        v_content = ""
        v_progress = ""
        v_product = ""

    # Sử dụng form với key thay đổi dựa trên selected_stt để buộc UI reset
    with st.form(key=f"form_{selected_stt}"):
        col1, col2, col3 = st.columns([1, 2, 1])
        
        stt = col1.text_input("STT (Bắt buộc)", value=v_stt)
        leader = col2.text_input("Lãnh đạo chỉ đạo", value=v_leader)
        status = col3.selectbox("Trạng thái", ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"], index=["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"].index(v_status))
        
        c_a, c_b = st.columns(2)
        content = c_a.text_area("Nội dung công việc", value=v_content)
        product = c_b.text_area("Sản phẩm", value=v_product) # Ô sản phẩm sẽ reset theo selected_stt
        
        progress = st.text_input("Tiến độ/Thời gian", value=v_progress)
        
        btn_save = st.form_submit_button("💾 LƯU DỮ LIỆU")
        btn_del = st.form_submit_button("🗑️ XÓA DÒNG")

        if btn_save:
            stt_val = stt.strip()
            if not stt_val:
                st.error("⚠️ STT không được để trống!")
            else:
                # KIỂM TRA TRÙNG STT
                is_duplicate = False
                if selected_stt == "-- Thêm mới --":
                    if stt_val in filtered_df['stt'].astype(str).values:
                        is_duplicate = True
                elif stt_val != selected_stt:
                    if stt_val in filtered_df['stt'].astype(str).values:
                        is_duplicate = True

                if is_duplicate:
                    st.error(f"❌ Lỗi: STT {stt_val} đã tồn tại!")
                else:
                    new_row = [sel_team, sel_type, sel_week, sel_staff, stt_val, content, leader, progress, status, product]
                    if selected_stt == "-- Thêm mới --":
                        sheet.append_row(new_row)
                        if sel_type == "Đăng ký công việc":
                            sync_row = [sel_team, "Báo cáo công việc", sel_week, sel_staff, stt_val, content, leader, progress, status, product]
                            sheet.append_row(sync_row)
                    else:
                        mask = (all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff) & (all_data['stt'].astype(str) == selected_stt)
                        idx = all_data[mask].index[0]
                        sheet.update(f"A{idx+2}:J{idx+2}", [new_row])
                    st.success("✅ Thành công!")
                    st.rerun()

        if btn_del and selected_stt != "-- Thêm mới --":
            mask = (all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff) & (all_data['stt'].astype(str) == selected_stt)
            idx = all_data[mask].index[0]
            sheet.delete_rows(idx + 2)
            st.rerun()

    # HIỂN THỊ BẢNG
    st.subheader("Bảng dữ liệu hiện tại")
    if not filtered_df.empty:
        try:
            filtered_df['stt_int'] = pd.to_numeric(filtered_df['stt'])
            display_df = filtered_df.sort_values('stt_int').drop(columns=['stt_int'])
        except:
            display_df = filtered_df
        def highlight_new(row):
            return ['color: red' if row['status'] == "🔵 Mới" and sel_type == "Báo cáo công việc" else 'color: black'] * len(row)
        st.dataframe(display_df.style.apply(highlight_new, axis=1), use_container_width=True)

    # XUẤT EXCEL
    st.divider()
    type_fn = "DangKy" if "Đăng ký" in sel_type else "BaoCao"
    col_ex1, col_ex2 = st.columns(2)
    with col_ex1:
        team_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not team_df.empty:
            st.download_button(f"📥 Tải Excel {sel_team}", data=export_excel_styled(team_df), file_name=f"{type_fn}_{sel_team}_{sel_week}.xlsx")
    with col_ex2:
        unit_df = all_data[(all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not unit_df.empty:
            st.download_button("📥 Tải Excel Toàn Đơn Vị", data=export_excel_styled(unit_df), file_name=f"{type_fn}_ToanDonVi_{sel_week}.xlsx")

