import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os
import datetime

st.set_page_config(page_title="Phím Hồng Music - PNG Generator", layout="wide")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎨 Cỗ Máy Xuất Ảnh PNG - Phím Hồng Music")
st.write("Bản sửa lỗi: Hiển thị ĐẦY ĐỦ ngày đi học, kể cả khi bạn thêm cột mới.")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

# Icon SVG
svg_student = '<svg viewBox="0 0 24 24" width="24" height="24" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2.06-1.12V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>'
svg_receipt = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'
svg_calendar = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'
svg_book = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zM6 4h5v8l-2.5-1.5L6 12V4z"/></svg>'
svg_thanks = '<svg viewBox="0 0 24 24" width="20" height="20" fill="#9a8a7a" style="vertical-align:middle; margin-right:8px;"><path d="M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z"/></svg>'

if uploaded_file:
    # Cấu hình đọc Excel: Không tự động parse header thành Date, cứ giữ nguyên là chuỗi
    df = pd.read_excel(uploaded_file, header=0).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận {len(df)} học sinh. Đang xử lý dữ liệu và thiết kế, đợi 1 xíu nhé...")
    progress_bar = st.progress(0)
    
    all_receipts_html = ""
    
    # 1. Tìm các cột Ngày Đi Học một cách thông minh (Có chữ 'T2'..'CN' hoặc chứa '/')
    date_cols = []
    for col in df.columns:
        col_str = str(col).upper()
        # Nếu cột là ngày tháng dạng datetime, chuyển nó về chuỗi ngày/tháng
        if isinstance(col, datetime.datetime):
            col_str = col.strftime('%d/%m')
            df.rename(columns={col: col_str}, inplace=True)
            date_cols.append(col_str)
        elif '/' in col_str or col_str.startswith('T2') or col_str.startswith('T3') or col_str.startswith('T4') or col_str.startswith('T5') or col_str.startswith('T6') or col_str.startswith('T7') or col_str.startswith('CN'):
            date_cols.append(col)

    for index, row in df.iterrows():
        ten = str(row['Họ và Tên']).strip()
        safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
        progress_bar.progress((index + 1) / len(df))
        
        lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Piano"
        hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
        so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
        tong_tien_goc = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else (hoc_phi * so_buoi)
        
        tien_sach = 0
        tien_sach_html = ""
        if 'Tiền sách' in df.columns:
            val = row['Tiền sách']
            if pd.notna(val) and str(val).strip() != '':
                try:
                    tien_sach = int(float(val))
                    if tien_sach > 0:
                        tien_sach_html = f'''
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_book} Tiền sách / Giáo trình:</td>
                            <td style="padding: 15px 0; font-weight: 900; color: #bc6c65; text-align: right; font-size: 24px; font-family: 'Times New Roman', serif;">+ {tien_sach:,} đ</td>
                        </tr>
                        '''
                except: pass
        
        tong_thanh_toan = tong_tien_goc + tien_sach
        nhan_xet = str(row['Nhận Xét Của GV']).strip() if pd.notna(row['Nhận Xét Của GV']) else "Bé học tập tích cực!"
        bank = str(row['Ngân Hàng']).strip()
        stk = str(row['STK']).split('.')[0] if pd.notna(row['STK']) else ""

        # --- XỬ LÝ NGÀY HỌC ---
        days_html = '<div style="display: flex; flex-wrap: wrap; gap: 8px;">'
        has_day = False
        
        for col in date_cols:
            cell_val = str(row[col]).strip().upper()
            if cell_val == 'X':
                has_day = True
                col_name = str(col).strip()
                parts = col_name.split(' ')
                
                thu = ""
                day_month = col_name
                
                if len(parts) > 1:
                    thu = parts[0]
                    try:
                        d_parts = parts[1].split('/')
                        if len(d_parts) >= 2:
                            day_month = f"{int(d_parts[0]):02d}/{int(d_parts[1]):02d}"
                        else:
                            day_month = parts[1]
                    except:
                        day_month = parts[1]
                else:
                    try:
                        d_parts = col_name.split('/')
                        if len(d_parts) >= 2:
                            day_month = f"{int(d_parts[0]):02d}/{int(d_parts[1]):02d}"
                    except:
                        pass
                
                days_html += f'''
                <div style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; display:inline-block; text-align:center;">
                    <div style="font-size:10px; color:#8e7f72; margin-bottom:2px; line-height:1;">{thu}</div>
                    <div style="font-size:13px; font-weight:bold; color:#4a2e25; line-height:1;">{day_month}</div>
                </div>
                '''
        
        days_html += '</div>'

        if not has_day:
            days_html = '<div style="color:#aaa; font-style:italic; font-size:14px; padding: 5px 0;">Chưa có dữ liệu điểm danh</div>'

        # QR Code
        qr_html = ""
        if bank and stk:
            add_info = urllib.parse.quote(ten)
            qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_thanh_toan}&addInfo={add_info}"
            try:
                resp = requests.get(qr_url, timeout=3)
                if resp.status_code == 200:
                    qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                    qr_html = f'<img src="{qr_b64}" style="width: 125px; height: 125px; border-radius: 10px;">'
            except: pass

        if not qr_html:
            qr_html = '<div style="font-size:12px; color:#999; padding:40px 0; border:1px dashed #ccc; border-radius:8px; text-align:center;">CHƯA CÓ QR</div>'

        # TEMPLATE HTML
        receipt_html = f"""
        <div class="receipt-card" data-name="{safe_name}" style="width: 850px; background: white; font-family: Arial, sans-serif; margin: 0 auto 40px auto; border-radius: 20px; overflow: hidden; box-sizing: border-box; position: relative;">
            
            <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); padding: 35px 50px; display: flex; align-items: center; justify-content: space-between; color: white;">
                <div style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid white; background-color: #fff; background-image: url('{logo_b64}'); background-size: cover; background-position: center; flex-shrink: 0;"></div>
                
                <div style="text-align: center; flex-grow: 1; padding: 0 20px;">
                    <div style="font-size: 16px; letter-spacing: 3px; font-weight: bold; opacity: 0.9; margin-bottom: 5px; text-transform: uppercase;">Lớp Nhạc Phím Hồng</div>
                    <h1 style="margin: 0; font-size: 44px; font-weight: 900; letter-spacing: 2px; font-family: 'Times New Roman', Times, serif; text-transform: uppercase; text-shadow: 1px 1px 4px rgba(0,0,0,0.2);">Phiếu Học Phí</h1>
                </div>

                <div style="text-align: center; min-width: 180px;">
                    <div style="font-size: 20px; font-weight: bold; margin-bottom: 8px;">Tháng 5 / 2026</div>
                    <div style="font-size: 28px; font-weight: 900; background: rgba(255,255,255,0.25); padding: 10px 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.5); font-family: 'Times New Roman', Times, serif;">Lớp {lop}</div>
                </div>
            </div>

            <div style="padding: 40px 60px;">
                <div style="background: #fdfaf6; border: 1px solid #f2e2b3; border-radius: 15px; padding: 30px 45px; margin: 0 auto 35px auto; width: 85%;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 20px;">
                        <tr>
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_student} Học sinh:</td>
                            <td style="padding: 15px 0; font-weight: 900; color: #2c1a16; text-align: right; font-size: 30px; font-family: 'Times New Roman', Times, serif;">{ten}</td>
                        </tr>
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_receipt} Học phí / buổi:</td>
                            <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{hoc_phi:,} đ</td>
                        </tr>
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_calendar} Số buổi học:</td>
                            <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{so_buoi} buổi</td>
                        </tr>
                        {tien_sach_html}
                    </table>
                </div>

                <div style="display: flex; gap: 45px; align-items: flex-start; justify-content: center; width: 95%; margin: 0 auto;">
                    <div style="flex: 1.3;">
                        <div style="margin-bottom: 25px;">
                            <div style="font-size: 14px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 12px;">NGÀY ĐI HỌC</div>
                            <div style="width: 100%;">{days_html}</div>
                        </div>
                        <div>
                            <div style="font-size: 14px; color: #8e7f72; font-weight: bold; margin-bottom: 12px; letter-spacing: 1px;">NHẬN XÉT CỦA GIÁO VIÊN</div>
                            <div style="background: #fffdf5; border: 1px solid #f2e2b3; border-radius: 12px; padding: 20px; color: #5a4b41; font-style: italic; line-height: 1.6; font-size: 16px;">
                                {nhan_xet}
                            </div>
                        </div>
                    </div>

                    <div style="flex: 0.7; display: flex; flex-direction: column; gap: 20px;">
                        <div style="background: #fdf6ec; border: 2px solid #ecdac8; border-radius: 15px; padding: 20px; text-align: center;">
                            <div style="font-size: 13px; color: #8e7f72; font-weight: bold;">TỔNG THANH TOÁN</div>
                            <div style="font-size: 38px; color: #4a2e25; font-weight: 900; margin-top: 10px; font-family: 'Times New Roman', Times, serif;">{tong_thanh_toan:,} đ</div>
                        </div>
                        <div style="background: white; border: 2px dashed #d49a71; border-radius: 15px; padding: 20px; text-align: center;">
                            <div style="font-size: 11px; color: #d49a71; font-weight: bold; margin-bottom: 15px;">QUÉT MÃ THANH TOÁN</div>
