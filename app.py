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

# --- HÀM XUẤT EXCEL CÓ CỘT SẢN PHẨM ---
def export_excel_styled(df):
    cols_order = ['stt', 'team', 'staff', 'content', 'product', 'leader', 'progress', 'status']
    for c in cols_order:
        if c not in df.columns: df[c] = ""
    
    df_export = df[cols_order].copy()
    df_export.columns = ['STT', 'ĐƠN VỊ/TỔ', 'HỌ VÀ TÊN', 'NỘI DUNG CÔNG VIỆC', 
                         'SẢN PHẨM', 'LÃNH ĐẠO CHỈ ĐẠO', 'TIẾN ĐỘ/THỜI GIAN', 'TRẠNG THÁI']

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='BaoCao')
        workbook  = writer.book
        worksheet = writer.sheets['BaoCao']

        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#2980b9', 'font_color': 'white', 'border': 1
        })
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})

        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        worksheet.set_column('A:A', 6, cell_fmt)
        worksheet.set_column('B:C', 18, cell_fmt)
        worksheet.set_column('D:E', 40, cell_fmt) # Cột Nội dung và Sản phẩm rộng 40
        worksheet.set_column('F:H', 18, cell_fmt)

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
    sel_week = st.sidebar.selectbox("Tuần:", [f"Tuần {str(i).zfill(2)}" for i in range(1, 53)], 
                                     index=datetime.now().isocalendar()[1]-1)
    sel_type = st.sidebar.selectbox("Loại hình:", ["Đăng ký công việc", "Báo cáo công việc"])
    
    staff_list = ["Văn Đức Giao", "Nguyễn Xuân Khánh", "Lê Nguyễn Hạnh Nhi", "Kiều Quang Phương", "Phan Văn Long", "Trần Hoàng Anh", "Trần Hồng Nhung", "Vũ Tuấn Anh", "Bùi Thành Tâm", "Trương Bình Minh", "Hoàng Thị Sinh", "Nguyễn Ngọc Thắng", "Đỗ Hoài Nam", "Lê Tĩnh", "Trương Thị Ngọc Linh", "Tạ Ngọc Thành", "Phùng Hữu Thọ", "Võ Xuân Quý"]
    sel_staff = st.sidebar.selectbox("Cán bộ:", staff_list)

    # Lọc dữ liệu
    filtered_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & 
                           (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff)] if not all_data.empty else pd.DataFrame()

    st.header(f"📋 {sel_type}: {sel_staff}")

    # CHỌN DÒNG THAO TÁC
    options = ["-- Thêm mới --"] + filtered_df['stt'].astype(str).tolist()
    selected_stt = st.selectbox("Chọn STT để Sửa/Xóa:", options)

    with st.form("main_form"):
        col1, col2, col3 = st.columns([1, 2, 1])
        
        if selected_stt != "-- Thêm mới --":
            row = filtered_df[filtered_df['stt'].astype(str) == selected_stt].iloc[0]
            v_stt, v_leader, v_status = str(row['stt']), row['leader'], row['status']
            v_content, v_progress, v_product = row['content'], row['progress'], row.get('product', "")
        else:
            v_stt, v_leader, v_status, v_content, v_progress, v_product = "", "", "🔵 Mới", "", "", ""

        stt = col1.text_input("STT", value=v_stt)
        leader = col2.text_input("Lãnh đạo chỉ đạo", value=v_leader)
        status = col3.selectbox("Trạng thái", ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"], 
                                index=["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"].index(v_status))
        
        c_a, c_b = st.columns(2)
        content = c_a.text_area("Nội dung công việc", value=v_content)
        product = c_b.text_area("Sản phẩm", value=v_product)
        
        progress = st.text_input("Tiến độ/Thời gian", value=v_progress)
        
        btn_save = st.form_submit_button("💾 LƯU DỮ LIỆU")
        btn_del = st.form_submit_button("🗑️ XÓA DÒNG")

        if btn_save:
            # Gồm 10 cột theo thứ tự Sheet
            new_row = [sel_team, sel_type, sel_week, sel_staff, stt, content, leader, progress, status, product]
            
            if selected_stt == "-- Thêm mới --":
                sheet.append_row(new_row)
                # TỰ ĐỘNG ĐỒNG BỘ SANG BÁO CÁO
                if sel_type == "Đăng ký công việc":
                    sync_row = [sel_team, "Báo cáo công việc", sel_week, sel_staff, stt, content, leader, progress, status, product]
                    sheet.append_row(sync_row)
                    st.info("Đã tự động đồng bộ sang Báo cáo công việc!")
            else:
                idx = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & 
                               (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff) & 
                               (all_data['stt'].astype(str) == selected_stt)].index[0]
                sheet.update(f"A{idx+2}:J{idx+2}", [new_row])
            st.rerun()

        if btn_del and selected_stt != "-- Thêm mới --":
            idx = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & 
                           (all_data['type'] == sel_type) & (all_data['staff'] == sel_staff) & 
                           (all_data['stt'].astype(str) == selected_stt)].index[0]
            sheet.delete_rows(idx + 2)
            st.rerun()

    # HIỂN THỊ BẢNG (TÔ ĐỎ NẾU LÀ BÁO CÁO MÀ TRẠNG THÁI MỚI)
    st.subheader("Bảng dữ liệu hiện tại")
    if not filtered_df.empty:
        def highlight_new(row):
            return ['color: red' if row['status'] == "🔵 Mới" and sel_type == "Báo cáo công việc" else 'color: black'] * len(row)
        
        st.dataframe(filtered_df.style.apply(highlight_new, axis=1), use_container_width=True)

    # XUẤT EXCEL
    st.divider()
    type_fn = "DangKy" if "Đăng ký" in sel_type else "BaoCao"
    col_ex1, col_ex2 = st.columns(2)
    
    with col_ex1:
        team_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not team_df.empty:
            st.download_button(f"📥 Tải Excel {sel_team}", data=export_excel_styled(team_df), 
                               file_name=f"{type_fn}_{sel_team}_{sel_week}.xlsx")
    
    with col_ex2:
        unit_df = all_data[(all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not unit_df.empty:
            st.download_button("📥 Tải Excel Toàn Đơn Vị", data=export_excel_styled(unit_df), 
                               file_name=f"{type_fn}_ToanDonVi_{sel_week}.xlsx")

