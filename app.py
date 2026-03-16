import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import xlsxwriter
import plotly.express as px

# --- 1. KẾT NỐI HỆ THỐNG ---
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

# --- 2. HÀM XUẤT EXCEL CHUYÊN NGHIỆP ---
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
        try:
            df_export['stt_n'] = pd.to_numeric(df_export['stt'], errors='coerce')
            df_export = df_export.sort_values(by=['team', 'stt_n']).drop(columns=['stt_n'])
        except: pass

        df_export.columns = labels
        df_export.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        header_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#2980b9', 'font_color': 'white', 'border': 1})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column('A:A', 6, cell_fmt)
        worksheet.set_column('B:Z', 22, cell_fmt)
    return output.getvalue()

# --- 3. GIAO DIỆN CHÍNH ---
st.set_page_config(layout="wide", page_title="QUẢN LÝ CÔNG VIỆC")
sheet = connect_gsheet()

if sheet:
    raw_data = sheet.get_all_records()
    all_data = pd.DataFrame(raw_data) if raw_data else pd.DataFrame()

    # SIDEBAR
    st.sidebar.header("🔍 BỘ LỌC")
    sel_team = st.sidebar.selectbox("Đơn vị/Tổ:", ["Tổ 1", "Tổ 2", "Tổ 3", "OBSERVER"])
    sel_week = st.sidebar.selectbox("Tuần:", [f"Tuần {str(i).zfill(2)}" for i in range(1, 53)], index=datetime.now().isocalendar()[1]-1)
    sel_type = st.sidebar.selectbox("Loại hình:", ["Đăng ký công việc", "Báo cáo công việc", "Đăng ký lịch tuần"])

    # Danh sách cán bộ theo tổ (Cập nhật mới)
    staff_mapping = {
        "Tổ 1": ["Trần Hoàng Anh", "Trần Hồng Nhung", "Bùi Thành Tâm", "Vũ Tuấn Anh"],
        "Tổ 2": ["Nguyễn Ngọc Thắng", "Hoàng Thị Sinh", "Trương Bình Minh", "Hoàng Minh Sơn", "Lê Tĩnh", "Trương Thị Ngọc Linh", "Đỗ Hoài Nam"],
        "Tổ 3": ["Tạ Ngọc Thành", "Phùng Hữu Thọ", "Võ Xuân Quý"],
        "OBSERVER": ["Văn Đức Giao", "Lê Nguyễn Hạnh Nhi", "Nguyễn Xuân Khánh", "Phan Văn Long", "Kiều Quang Phương"]
    }
    current_staff_list = staff_mapping.get(sel_team, [])
    sel_staff = st.sidebar.selectbox("Cán bộ/Người đăng ký:", current_staff_list, key=f"staff_{sel_team}")

    # Lọc dữ liệu hiển thị
    filtered_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
    if sel_type != "Đăng ký lịch tuần":
        filtered_df = filtered_df[filtered_df['staff'] == sel_staff]

    st.header(f"📋 {sel_type}")

    # Gợi ý STT
    suggested_stt = 1
    if not filtered_df.empty:
        try: suggested_stt = int(pd.to_numeric(filtered_df['stt'], errors='coerce').max()) + 1
        except: suggested_stt = len(filtered_df) + 1

    options = ["-- Thêm mới --"] + sorted(filtered_df['stt'].astype(str).tolist(), key=lambda x: int(x) if x.isdigit() else 999)
    selected_stt = st.selectbox("Chọn STT để thao tác:", options, key=f"stt_{sel_team}_{sel_staff}")

    # --- 4. FORM NHẬP LIỆU ---
    with st.form(key=f"form_{sel_team}_{sel_staff}_{selected_stt}"):
        row_data = filtered_df[filtered_df['stt'].astype(str) == selected_stt].iloc[0] if selected_stt != "-- Thêm mới --" else {}
        
        if sel_type == "Đăng ký lịch tuần":
            c1, c2 = st.columns([1, 3])
            stt = c1.text_input("STT", value=str(row_data.get('stt', suggested_stt)))
            date_time = c2.text_input("Thứ, Ngày", value=str(row_data.get('date_time', "")))
            content = st.text_area("Nội dung đăng ký", value=str(row_data.get('content', "")))
            c3, c4 = st.columns(2)
            location = c3.text_input("Địa điểm", value=str(row_data.get('location', "")))
            host = c4.text_input("Chủ trì", value=str(row_data.get('host', "")))
            participants = st.text_area("Thành phần", value=str(row_data.get('participants', "")))
            note = st.text_input("Ghi chú", value=str(row_data.get('note', "")))
        else:
            c1, c2, c3 = st.columns([1, 2, 1])
            stt = c1.text_input("STT", value=str(row_data.get('stt', suggested_stt)))
            leader = c2.text_input("Lãnh đạo chỉ đạo", value=str(row_data.get('leader', "")))
            status = c3.selectbox("Trạng thái", ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"], index=0)
            ca, cb = st.columns(2)
            content = ca.text_area("Nội dung", value=str(row_data.get('content', "")))
            product = cb.text_area("Sản phẩm", value=str(row_data.get('product', "")))
            progress = st.text_input("Tiến độ", value=str(row_data.get('progress', "")))

        btn_save = st.form_submit_button("💾 LƯU DỮ LIỆU")
        btn_del = st.form_submit_button("🗑️ XÓA DÒNG")

        if btn_save:
            try:
                fresh_df = pd.DataFrame(sheet.get_all_records())
                data_list = [sel_team, sel_type, sel_week, sel_staff, stt, content]
                if sel_type == "Đăng ký lịch tuần":
                    data_list += ["", "", "", "", date_time, location, host, participants, note]
                else:
                    data_list += [leader, progress, status, product, "", "", "", "", ""]

                if selected_stt == "-- Thêm mới --":
                    sheet.append_row(data_list)
                    if sel_type == "Đăng ký công việc":
                        sync = data_list.copy(); sync[1] = "Báo cáo công việc"
                        sheet.append_row(sync)
                else:
                    mask = (fresh_df['team'] == sel_team) & (fresh_df['week'] == sel_week) & (fresh_df['type'] == sel_type) & (fresh_df['stt'].astype(str) == selected_stt)
                    if sel_type != "Đăng ký lịch tuần": mask &= (fresh_df['staff'] == sel_staff)
                    indices = fresh_df[mask].index.tolist()
                    if indices: sheet.update(f"A{indices[0]+2}:O{indices[0]+2}", [data_list])
                st.success("✅ Đã lưu!"); st.rerun()
            except Exception as e: st.error(f"Lỗi lưu: {e}")

        if btn_del and selected_stt != "-- Thêm mới --":
            try:
                fresh_df = pd.DataFrame(sheet.get_all_records())
                mask = (fresh_df['team'] == sel_team) & (fresh_df['week'] == sel_week) & (fresh_df['type'] == sel_type) & (fresh_df['stt'].astype(str) == selected_stt)
                if sel_type != "Đăng ký lịch tuần": mask &= (fresh_df['staff'] == sel_staff)
                indices = fresh_df[mask].index.tolist()
                if indices:
                    for idx in reversed(indices): sheet.delete_rows(int(idx) + 2)
                    st.success("✅ Đã xóa!"); st.rerun()
                else: st.warning("Không tìm thấy dòng thực tế.")
            except Exception as e: st.error(f"Lỗi xóa: {e}")

    # --- 5. HIỂN THỊ BẢNG & XUẤT EXCEL (2 NÚT) ---
    st.subheader("📊 Bảng dữ liệu hiện tại")
    st.dataframe(filtered_df, use_container_width=True)

    st.divider()
    st.subheader("📥 XUẤT FILE EXCEL")
    prefix = {"Đăng ký công việc": "DangKy", "Báo cáo công việc": "BaoCao", "Đăng ký lịch tuần": "LichTuan"}[sel_type]
    is_cal = (sel_type == "Đăng ký lịch tuần")
    
    col_ex1, col_ex2 = st.columns(2)
    with col_ex1:
        team_data = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not team_data.empty:
            st.download_button(f"📥 Xuất Excel {sel_team}", data=export_excel_flexible(team_data, is_cal), file_name=f"{prefix}_{sel_team}_{sel_week}.xlsx")
    with col_ex2:
        unit_data = all_data[(all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not unit_data.empty:
            st.download_button("📥 Xuất Excel Toàn đơn vị", data=export_excel_flexible(unit_data, is_cal), file_name=f"{prefix}_ToanDonVi_{sel_week}.xlsx")

    # --- 6. BIỂU ĐỒ HIỆU SUẤT ---
    st.divider()
    st.header("📈 HIỆU SUẤT")
    report_data = all_data[(all_data['type'] == "Báo cáo công việc") & (all_data['week'] == sel_week)]
    if not report_data.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.bar(report_data.groupby(['team', 'status']).size().reset_index(name='c'), x='team', y='c', color='status', barmode='group', title="Tiến độ theo Tổ"), use_container_width=True)
        with c2:
            staff_data = report_data[report_data['staff'] == sel_staff]
            if not staff_data.empty:
                st.plotly_chart(px.pie(staff_data['status'].value_counts().reset_index(), values='count', names='status', title=f"Tỷ lệ {sel_staff}"), use_container_width=True)
