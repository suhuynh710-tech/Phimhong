import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
import base64
import urllib.parse
import os
import datetime

st.set_page_config(page_title="Phím Hồng Music - PNG Generator", layout="wide")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎨 Cỗ Máy Xuất Ảnh PNG - Bản Phân Loại Học Phí")
st.write("Đã sửa hoàn toàn lỗi cú pháp chuỗi và định dạng STK / mã QR.")

# --- KIỂM TRA FILE LOGO TRÊN SERVER (GITHUB) ---
if not os.path.exists(LOGO_PATH):
    st.sidebar.error(f"❌ Không tìm thấy file: `{LOGO_PATH}`. Hãy kiểm tra lại tên file trên GitHub!")
else:
    st.sidebar.success("✅ Đã kết nối file Logo thành công!")

# --- BỘ CHỌN THỜI GIAN LINH HOẠT ---
now = datetime.datetime.now()
thang_nam_mac_dinh = f"Tháng {now.month} / {now.year}"
thang_nam_hien_thi = st.sidebar.text_input("📅 Thời gian hiển thị trên phiếu:", value=thang_nam_mac_dinh)

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi_Moi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

def clean_number(val):
    if pd.isna(val):
        return 0
    s = str(val).strip().replace(' ', '').replace(',', '')
    if s.lower() in ['', 'nan', 'none']:
        return 0
    if '.' in s:
        s = s.split('.')[0]
    try:
        return int(s)
    except:
        try:
            return int(float(s))
        except:
            return 0

def clean_stk(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# --- ICON SVG ---
svg_student = '<svg viewBox="0 0 24 24" width="24" height="24" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2.06-1.12V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>'
svg_receipt = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'
svg_calendar = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'
svg_extra = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M11 17h2v-4h4v-2h-4V7h-2v4H7v2h4v4zm1 5q-2.075 0-3.9-.788-1.825-.787-3.175-2.162-1.35-1.35-2.137-3.175Q2 14.05 2 12q0-2.075.788-3.9.787-1.825 2.137-3.175 1.35-1.35 3.175-2.137Q9.925 2 12 2q2.075 0 3.9.788 1.825.787 3.175 2.137 1.35 1.35 2.138 3.175Q22 9.925 22 12q0 2.075-.788 3.9-.787 1.825-2.138 3.175-1.35 1.35-3.175 2.162Q14.075 22 12 22Z"/></svg>'
svg_money = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-1h-1c-.55 0-1-.45-1-1v-3c0-.55.45-1 1-1h3v-1h-3V9h2V8h2v1h1c.55 0 1 .45 1 1v3c0 .55-.45 1-1 1h-3v1h3v1h-2v1z"/></svg>'
svg_thanks = '<svg viewBox="0 0 24 24" width="20" height="20" fill="#9a8a7a" style="vertical-align:middle; margin-right:8px;"><path d="M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z"/></svg>'

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names
    
    student_sheet = next((s for s in sheet_names if 'học sinh' in s.lower()), sheet_names[0])
    attendance_sheet = next((s for s in sheet_names if 'sheet1' in s.lower() or 'điểm danh' in s.lower()), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
    
    df_students = pd.read_excel(uploaded_file, sheet_name=student_sheet, dtype={'STK': str})
    df_attendance = pd.read_excel(uploaded_file, sheet_name=attendance_sheet)
    
    logo_b64 = get_base64_logo()
    
    df_students.columns = [str(c).strip() for c in df_students.columns]
    df_attendance.columns = [str(c).strip() for c in df_attendance.columns]
    
    name_col = next((col for col in df_students.columns if 'họ và tên' in col.lower()), None)
    
    if not name_col:
        st.error("❌ Không tìm thấy cột 'Họ và Tên' trong trang danh sách!")
    else:
        df_students = df_students.dropna(subset=[name_col])
        st.success(f"🎉 Kết nối thành công dữ liệu!")
        progress_bar = st.progress(0)
        
        all_receipts_html = ""
        
        lop_col = next((col for col in df_students.columns if col.lower() == 'lớp'), None)
        hp_buoi_col = next((col for col in df_students.columns if 'học phí (buổi)' in col.lower()), None)
        hp_thang_col = next((col for col in df_students.columns if col.lower() == 'học phí' or ('học phí' in col.lower() and 'buổi' not in col.lower())), None)
        phi_khac_col = next((col for col in df_students.columns if 'phí khác' in col.lower()), None)
        ghi_chu_col = next((col for col in df_students.columns if 'ghi chú' in col.lower()), None)
        nhan_xet_col = next((col for col in df_students.columns if 'nhận xét của gv' in col.lower()), None)
        ngan_hang_col = next((col for col in df_students.columns if 'ngân hàng' in col.lower()), None)
        stk_col = next((col for col in df_students.columns if 'stk' in col.lower()), None)

        date_cols = []
        for col in df_attendance.columns:
            col_str = str(col).upper()
            if isinstance(col, datetime.datetime):
                col_str = col.strftime('%d/%m')
                df_attendance.rename(columns={col: col_str}, inplace=True)
                date_cols.append(col_str)
            elif '/' in col_str or col_str.startswith('T2') or col_str.startswith('T3') or col_str.startswith('T4') or col_str.startswith('T5') or col_str.startswith('T6') or col_str.startswith('T7') or col_str.startswith('CN'):
                date_cols.append(col)

        for i, (index, row) in enumerate(df_students.iterrows()):
            ten = str(row[name_col]).strip()
            safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
            
            progress_val = min((i + 1) / len(df_students), 1.0)
            progress_bar.progress(progress_val)
