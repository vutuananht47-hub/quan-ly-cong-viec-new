import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import xlsxwriter
import time
import plotly.express as px

# --- CẤU HÌNH TRANG ---
st.set_page_config(layout="wide", page_title="QUẢN LÝ CÔNG VIỆC WEB V2")

# --- KẾT NỐI GOOGLE SHEETS ---
def connect_gsheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"] 
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_id = "1B5NE0ULV9LFGw6qHNtog4jgjxtA4x2JLYgCXQ6M1P-M" 
        return client.open_by_key(spreadsheet_id).get_worksheet(0)
    except Exception as e:
        st.error(f"Lỗi kết nối hệ thống: {e}")
        return None

sheet = connect_gsheet()

# --- HÀM TÌM DÒNG CHÍNH XÁC (Tránh lỗi nhiều người dùng) ---
def find_row_index(sheet, team, week, type_job, staff, stt):
    """Tìm số dòng vật lý trên Google Sheet khớp với các tiêu chí lọc"""
    all_values = sheet.get_all_values()
    for i, row in enumerate(all_values[1:], start=2): # Bắt đầu từ dòng 2 (bỏ header)
        # Giả định thứ tự cột: team(A), type(B), week(C), staff(D), stt(E)
        if (row[0] == team and row[1] == type_job and 
            row[2] == week and row[3] == staff and str(row[4]) == str(stt)):
            return i
    return None

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

        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#2980b9', 'font_color': 'white', 'border': 1
        })
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
    # Load data
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

    # --- TỰ ĐỘNG TÍNH STT MỚI ---
    if not filtered_df.empty:
        try:
            next_stt = int(pd.to_numeric(filtered_df['stt']).max()) + 1
        except:
            next_stt = len(filtered_df) + 1
    else:
        next_stt = 1

    # --- NHẬP / SỬA / XÓA ---
    st.header(f"📋 Cán bộ: {sel_staff}")
    
    options = ["-- Thêm mới --"] + sorted(filtered_df['stt'].astype(str).tolist(), key=lambda x: int(x) if x.isdigit() else 0)
    selected_stt = st.selectbox("Chọn STT để Sửa hoặc Xóa:", options)

    with st.form("input_form"):
        c1, c2, c3 = st.columns([1, 2, 1])
        
        # Kiểm tra trước khi truy xuất hàng
        if selected_stt != "-- Thêm mới --" and not filtered_df.empty:
            match_rows = filtered_df[filtered_df['stt'].astype(str) == selected_stt]
            if not match_rows.empty:
                row = match_rows.iloc[0]
                v_stt, v_leader, v_status = str(row['stt']), row['leader'], row['status']
                v_content, v_progress = row['content'], row['progress']
            else:
                st.error("Không tìm thấy dòng dữ liệu này. Có thể nó đã bị xóa hoặc thay đổi.")
                st.stop()
        else:
            v_stt, v_leader, v_status, v_content, v_progress = str(next_stt), "", "🔵 Mới", "", ""

        stt = c1.text_input("STT (Tự động đề xuất)", value=v_stt)
        leader = c2.text_input("LÃNH ĐẠO CHỈ ĐẠO", value=v_leader)
        status = c3.selectbox("TRẠNG THÁI", ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"], 
                              index=["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"].index(v_status) if v_status in ["🔵 Mới", "🟢 Hoàn thành", "🔴 Trễ hạn", "🟡 Đang làm"] else 0)
        content = st.text_area("NỘI DUNG CÔNG VIỆC", value=v_content)
        progress = c1.text_input("TIẾN ĐỘ/THỜI GIAN", value=v_progress)
        
        col_btn1, col_btn2 = st.columns([1, 5])
        save = col_btn1.form_submit_button("💾 LƯU DỮ LIỆU")
        delete = col_btn2.form_submit_button("🗑️ XÓA DÒNG")

        if save:
            new_row = [sel_team, sel_type, sel_week, sel_staff, stt, content, leader, progress, status]
            with st.spinner('Đang đồng bộ với hệ thống...'):
                if selected_stt == "-- Thêm mới --":
                    sheet.append_row(new_row)
                    st.success("Đã thêm công việc mới thành công!")
                else:
                    # Tìm dòng vật lý dựa trên dữ liệu cũ để tránh ghi đè nhầm
                    real_row_idx = find_row_index(sheet, sel_team, sel_week, sel_type, sel_staff, selected_stt)
                    if real_row_idx:
                        sheet.update(f"A{real_row_idx}:I{real_row_idx}", [new_row])
                        st.success(f"Đã cập nhật STT {selected_stt} thành công!")
                    else:
                        st.error("Lỗi: Không tìm thấy dòng thực tế trên Sheet để sửa. Vui lòng thử lại.")
            time.sleep(1)
            st.rerun()

        if delete and selected_stt != "-- Thêm mới --":
            with st.spinner('Đang xóa dữ liệu...'):
                real_row_idx = find_row_index(sheet, sel_team, sel_week, sel_type, sel_staff, selected_stt)
                if real_row_idx:
                    sheet.delete_rows(real_row_idx)
                    st.warning(f"Đã xóa STT {selected_stt}!")
                else:
                    st.error("Lỗi: Không tìm thấy dòng thực tế trên Sheet để xóa.")
            time.sleep(1)
            st.rerun()

    # HIỂN THỊ BẢNG DỮ LIỆU
    st.subheader("📊 DỮ LIỆU ĐANG HIỂN THỊ")
    if not filtered_df.empty:
        # Sắp xếp theo STT để dễ nhìn
        display_df = filtered_df.copy()
        display_df['stt'] = pd.to_numeric(display_df['stt'], errors='coerce')
        display_df = display_df.sort_values(by='stt')
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    else:
        st.info("Chưa có dữ liệu cho các tiêu chí đã chọn. Vui lòng nhập 'Thêm mới'.")
        import plotly.express as px # Thêm thư viện này ở đầu file

# ... (Giữ nguyên các phần cũ đến đoạn hiển thị bảng dữ liệu) ...

    # --- PHẦN THỐNG KÊ & BIỂU ĐỒ ---
    st.divider()
    st.header("📊 PHÂN TÍCH TIẾN ĐỘ")
    
    # Lấy dữ liệu của toàn bộ Tổ để so sánh (không chỉ riêng 1 cán bộ)
    team_stats_df = all_data[(all_data['team'] == sel_team) & (all_data['week'] == sel_week) & (all_data['type'] == sel_type)]
    
    if not team_stats_df.empty:
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader(f"Trạng thái công việc - {sel_team}")
            status_counts = team_stats_df['status'].value_counts().reset_index()
            status_counts.columns = ['Trạng thái', 'Số lượng']
            
            # Định nghĩa màu sắc cố định cho các trạng thái
            color_map = {
                "🔵 Mới": "#3498db",
                "🟢 Hoàn thành": "#2ecc71",
                "🔴 Trễ hạn": "#e74c3c",
                "🟡 Đang làm": "#f1c40f"
            }
            
            fig_pie = px.pie(status_counts, values='Số lượng', names='Trạng thái', 
                             color='Trạng thái', color_discrete_map=color_map,
                             hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_chart2:
            st.subheader("Số lượng công việc theo Cán bộ")
            staff_stats = team_stats_df['staff'].value_counts().reset_index()
            staff_stats.columns = ['Cán bộ', 'Số lượng']
            
            fig_bar = px.bar(staff_stats, x='Cán bộ', y='Số lượng', 
                             text='Số lượng', color='Cán bộ',
                             labels={'Số lượng': 'Tổng số việc'})
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)
            
        # Thêm bảng tổng hợp nhanh
        st.markdown("**Bảng tổng hợp nhanh:**")
        summary_table = pd.crosstab(team_stats_df['staff'], team_stats_df['status'])
        st.table(summary_table)
    else:
        st.info("Không có dữ liệu thống kê cho tuần này.")


    # --- XUẤT EXCEL ---
    st.divider()
    st.subheader("📥 XUẤT BÁO CÁO EXCEL")
    
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
