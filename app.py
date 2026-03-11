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

# --- HÀM XUẤT EXCEL TỐI ƯU ---
def export_excel_flexible(df, is_calendar=False):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if is_calendar:
            cols = ['stt', 'team', 'date_time', 'content', 'location', 'host', 'participants', 'note']
            labels = ['STT', 'TỔ', 'THỨ, NGÀY', 'NỘI DUNG ĐĂNG KÝ', 'THỜI GIAN, ĐỊA ĐIỂM', 'CHỦ TRÌ/CHỈ ĐẠO', 'THÀNH PHẦN', 'GHI CHÚ']
        else:
            cols = ['stt', 'team', 'staff', 'content', 'leader', 'progress', 'status', 'product']
            labels = ['STT', 'ĐƠN VỊ/TỔ', 'HỌ VÀ TÊN', 'NỘI DUNG CÔNG VIỆC', 'LÃNH ĐẠO CHỈ ĐẠO', 'TIẾN ĐỘ/THỜI GIAN', 'TRẠNG THÁI', 'SẢN PHẨM']
        
        # Đảm bảo có đủ cột tránh lỗi
        for c in cols:
            if c not in df.columns: df[c] = ""
        
        df_export = df[cols].copy()
        
        # Sắp xếp theo Tổ rồi đến STT
        try:
            df_export['stt_n'] = pd.to_numeric(df_export['stt'], errors='coerce')
            df_export = df_export.sort_values(by=['team', 'stt_n']).drop(columns=['stt_n'])
        except: pass

        df_export.columns = labels
        df_export.to_excel(writer, index=False, sheet_name='Data')
        
        workbook  = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#2980b9', 'font_color': 'white', 'border': 1})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})
        
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column('A:A', 6, cell_fmt)
        worksheet.set_column('B:Z', 22, cell_fmt)
        
    return output.getvalue()

# --- GIAO DIỆN CHÍNH ---
st.set_page_config(layout="wide", page_title="QUẢN LÝ CÔNG VIỆC & HIỆU SUẤT")
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

    # Lọc dữ liệu hiển thị (Bảng và Form)
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
    selected_stt = st.selectbox("Chọn STT để thao tác:", options)

    # --- FORM NHẬP LIỆU ---
    with st.form(key=f"form_{selected_stt}_{sel_type}"):
        row_data = filtered_df[filtered_df['stt'].astype(str) == selected_stt].iloc[0] if selected_stt != "-- Thêm mới --" else {}
        
        if sel_type == "Đăng ký lịch tuần":
            c1, c2 = st.columns([1, 3])
            stt = c1.text_input("STT (Bắt buộc)", value=str(row_data.get('stt', suggested_stt)))
            date_time = c2.text_input("Thứ, Ngày", value=str(row_data.get('date_time', "")))
            content = st.text_area("Nội dung đăng ký", value=str(row_data.get('content', "")))
            c3, c4 = st.columns(2)
            location = c3.text_input("Thời gian, Địa điểm", value=str(row_data.get('location', "")))
            host = c4.text_input("Chủ trì/Chỉ đạo", value=str(row_data.get('host', "")))
            participants = st.text_area("Thành phần", value=str(row_data.get('participants', "")))
            note = st.text_input("Ghi chú", value=str(row_data.get('note', "")))
        else:
            c1, c2, c3 = st.columns([1, 2, 1])
            stt = c1.text_input("STT (Bắt buộc)", value=str(row_data.get('stt', suggested_stt)))
            leader = c2.text_input("Lãnh đạo chỉ đạo", value=str(row_data.get('leader', "")))
            status = c3.selectbox("Trạng thái", ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"], 
                                  index=["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"].index(row_data.get('status', '🔵 Mới')))
            ca, cb = st.columns(2)
            content = ca.text_area("Nội dung công việc", value=str(row_data.get('content', "")))
            product = cb.text_area("Sản phẩm", value=str(row_data.get('product', "")))
            progress = st.text_input("Tiến độ/Thời gian", value=str(row_data.get('progress', "")))

        btn_save = st.form_submit_button("💾 LƯU DỮ LIỆU")
        btn_del = st.form_submit_button("🗑️ XÓA DÒNG")

        if btn_save:
            stt_v = stt.strip()
            if not stt_v:
                st.error("⚠️ STT không được để trống!")
            elif selected_stt == "-- Thêm mới --" and stt_v in filtered_df['stt'].astype(str).values:
                st.error(f"❌ Trùng STT: {stt_v} đã tồn tại!")
            else:
                data = [sel_team, sel_type, sel_week, sel_staff, stt_v, content]
                if sel_type == "Đăng ký lịch tuần":
                    data += ["", "", "", "", date_time, location, host, participants, note]
                else:
                    data += [leader, progress, status, product, "", "", "", "", ""]
                
                if selected_stt == "-- Thêm mới --":
                    sheet.append_row(data)
                    if sel_type == "Đăng ký công việc":
                        sync_row = [sel_team, "Báo cáo công việc", sel_week, sel_staff, stt_v, content, leader, progress, status, product, "", "", "", "", ""]
                        sheet.append_row(sync_row)
                else:
                    mask = (all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type) & (all_data['stt'].astype(str) == selected_stt)
                    if sel_type != "Đăng ký lịch tuần": mask &= (all_data['staff'] == sel_staff)
                    idx = all_data[mask].index[0]
                    sheet.update(f"A{idx+2}:O{idx+2}", [data])
                st.success("✅ Đã lưu thành công!")
                st.rerun()

        if btn_del and selected_stt != "-- Thêm mới --":
            mask = (all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type) & (all_data['stt'].astype(str) == selected_stt)
            if sel_type != "Đăng ký lịch tuần": mask &= (all_data['staff'] == sel_staff)
            idx = all_data[mask].index[0]
            sheet.delete_rows(idx+2)
            st.rerun()

    # --- HIỂN THỊ BẢNG ---
    st.subheader("📊 Dữ liệu cá nhân")
    if not filtered_df.empty:
        st.dataframe(filtered_df, use_container_width=True)

    # --- PHẦN XUẤT FILE CHUYÊN SÂU ---
    st.divider()
    st.subheader("📥 XUẤT FILE EXCEL BÁO CÁO")
    
    col_ex1, col_ex2 = st.columns(2)
    is_cal = (sel_type == "Đăng ký lịch tuần")
    type_fn = "LichTuan" if is_cal else "CongViec"

    with col_ex1:
        # Xuất theo Tổ
        team_data = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not team_data.empty:
            st.info(f"Dữ liệu {sel_team} hiện có {len(team_data)} dòng.")
            st.download_button(f"📥 Tải Excel {sel_team}", data=export_excel_flexible(team_data, is_calendar=is_cal), 
                               file_name=f"{type_fn}_{sel_team}_{sel_week}.xlsx", key="btn_team")
        else:
            st.warning(f"{sel_team} chưa có dữ liệu tuần này.")

    with col_ex2:
        # Xuất Toàn đơn vị
        unit_data = all_data[(all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
        if not unit_data.empty:
            st.info(f"Dữ liệu Toàn đơn vị hiện có {len(unit_data)} dòng.")
            st.download_button("📥 Tải Excel Toàn đơn vị", data=export_excel_flexible(unit_data, is_calendar=is_cal), 
                               file_name=f"{type_fn}_ToanDonVi_{sel_week}.xlsx", key="btn_all")
        else:
            st.warning("Toàn đơn vị chưa có dữ liệu tuần này.")

    # --- PHẦN BIỂU ĐỒ (Dưới cùng) ---
    st.divider()
    st.header("📈 PHÂN TÍCH HIỆU SUẤT")
    report_data = all_data[(all_data['type'] == "Báo cáo công việc") & (all_data['week'] == sel_week)]
    if not report_data.empty:
        c_chart1, c_chart2 = st.columns(2)
        with c_chart1:
            team_stats = report_data.groupby(['team', 'status']).size().reset_index(name='counts')
            st.plotly_chart(px.bar(team_stats, x='team', y='counts', color='status', barmode='group', title="Hiệu suất theo Tổ", color_discrete_map={"🔵 Mới": "#3498db", "🟢 Hoàn thành": "#2ecc71", "🔴 Trễ hạn": "#e74c3c", "🟡 Đang làm": "#f1c40f"}), use_container_width=True)
        with c_chart2:
            ind_data = report_data[report_data['staff'] == sel_staff]
            if not ind_data.empty:
                ind_stats = ind_data['status'].value_counts().reset_index()
                ind_stats.columns = ['status', 'count']
                st.plotly_chart(px.pie(ind_stats, values='count', names='status', title=f"Tỷ lệ của {sel_staff}", color='status', color_discrete_map={"🔵 Mới": "#3498db", "🟢 Hoàn thành": "#2ecc71", "🔴 Trễ hạn": "#e74c3c", "🟡 Đang làm": "#f1c40f"}), use_container_width=True)

