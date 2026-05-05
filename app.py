import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os
from weasyprint import HTML

st.set_page_config(page_title="Phím Hồng Music - Phiếu Học Phí", layout="centered")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎯 Hệ thống xuất Phiếu Học Phí")
st.write("Đã điều chỉnh chiều ngang, thay font chữ sang Arial và nới rộng không gian để không bị rớt chữ. Giao diện xem trước được giản lược để tránh lỗi.")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

# Định nghĩa Icon SVG (Chỉ dùng cho học phí và ngày học)
icon_receipt = '''<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'''
icon_calendar = '''<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'''

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận danh sách {len(df)} học sinh. Đang xử lý, vui lòng chờ trong giây lát...")

    zip_buffer = io.BytesIO()
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_students = len(df)
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            ten = str(row['Họ và Tên']).strip()
            
            # Cập nhật thanh tiến trình
            status_text.text(f"Đang xử lý PDF cho: {ten} ({index + 1}/{total_students})...")
            progress_bar.progress((index + 1) / total_students)
            
            lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Năng Khiếu"
            hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
            so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
            tong_tien = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else 0
            nhan_xet = str(row['Nhận Xét Của GV']).strip() if pd.notna(row['Nhận Xét Của GV']) else "Bé học tập tích cực, tự tin giao tiếp."
            bank = str(row['Ngân Hàng']).strip()
            stk = str(row['STK']).split('.')[0] if pd.notna(row['STK']) else ""

            # Xử lý Ngày đi học
            date_cols = [c for c in df.columns if '/' in str(c)]
            days_html = ""
            for col in date_cols:
                if str(row[col]).strip().upper() == 'X':
                    parts = col.split(' ')
                    if len(parts) > 1:
                        thu = parts[0]
                        d_parts = parts[1].split('/')
                        day_month = f"{d_parts[0]:>02}/{d_parts[1]:>02}"
                        days_html += f'<div style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; margin:4px; display:inline-block; text-align:center;"><div style="font-size:11px; color:#8e7f72; margin-bottom:3px;">{thu}</div><div style="font-size:15px; font-weight:bold; color:#4a2e25;">{day_month}</div></div>'

            if not days_html:
                days_html = '<span style="color:#aaa; font-style:italic; font-size:14px;">Chưa có buổi học</span>'

            # QR Code lấy trực tiếp hình bằng link (Nhanh và không lỗi)
            qr_html = ""
            if bank and stk and bank != 'nan':
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_tien}&addInfo={add_info}"
                qr_html = f'<img src="{qr_url}" style="width: 120px; height: 120px; border-radius: 12px;">'
            else:
                qr_html = '<div style="font-size:11px; color:#999; padding:20px; border:1px dashed #ccc; border-radius:8px;">CHƯA CÓ QR</div>'

            logo_html = f'<div style="margin: 0 auto 15px auto; width: 85px; height: 85px; border-radius: 50%; border: 3px solid white; background-image: url({logo_b64}); background-size: cover; background-position: center; box-shadow: 0 3px 10px rgba(0,0,0,0.2);"></div>' if logo_b64 else ''

            # Bản thiết kế siêu rộng: 600px x 1350px
            html_template = f"""
            <html>
            <head>
            <meta charset="UTF-8">
            <style>
                @page {{ size: 600px 1350px; margin: 0; }}
                body {{ font-family: 'Arial', sans-serif; margin: 0; background: #fdfaf6; color: #333; line-height: 1.4; }}
                .container {{ width: 600px; height: 1350px; background: #fdfaf6; position: relative; box-sizing: border-box; }}
                
                .header {{ background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; text-align: center; padding: 60px 40px 40px; }}
                .header h1 {{ font-size: 38px; margin: 15px 0; letter-spacing: 2px; font-weight: 700; text-transform: uppercase; }}
                .badge {{ display: inline-block; background: rgba(255,255,255,0.25); padding: 8px 25px; border-radius: 25px; font-size: 16px; border: 1px solid rgba(255,255,255,0.5); font-weight: bold; letter-spacing: 1px; }}
                
                .content {{ padding: 35px 50px; }}
                .row {{ display: flex; justify-content: space-between; align-items: center; padding: 20px 0; border-bottom: 1px dashed #d8c8b8; }}
                .row .label {{ color: #7a6b5d; font-size: 18px; font-weight: 500; display:flex; align-items:center; }}
                .row .value {{ font-weight: 700; color: #2c1a16; font-size: 22px; text-align: right; width: 65%; }}
                
                .total-box {{ border: 2px solid #ecdac8; border-radius: 20px; background: #fdf6ec; text-align: center; padding: 30px; margin: 35px 0; box-shadow: 0 4px 15px rgba(236,218,200,0.3); }}
                .total-box .amount {{ font-size: 42px; color: #4a2e25; font-weight: 700; margin-top: 10px; }}
                
                .remark-box {{ border: 1px solid #f2e2b3; background: #fffdf5; border-radius: 15px; padding: 25px; text-align: center; color: #5a4b41; font-style: italic; margin-top: 20px; line-height: 1.6; font-size: 17px; }}
                
                .qr-wrap {{ display: flex; gap: 25px; align-items: center; margin-top: 45px; padding: 25px; border: 2px dashed #d49a71; background: white; border-radius: 20px; }}
                .stk-val {{ font-size: 22px; color: #bc6c65; font-weight: 800; margin: 8px 0; letter-spacing: 1px; }}
                
                .footer {{ text-align: center; padding: 40px; font-size: 16px; color: #9a8a7a; font-style: italic; position: absolute; bottom: 0; width: 100%; box-sizing:border-box; }}
            </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        {logo_html}
                        <div style="font-size: 14px; letter-spacing: 4px; font-weight:bold;">LỚP NHẠC PHÍM HỒNG</div>
                        <h1>PHIẾU HỌC PHÍ</h1>
                        <div style="margin-bottom: 20px; font-size: 18px; font-weight: 500;">Tháng 5 / 2026</div>
                        <div class="badge">{lop}</div>
                    </div>
                    
                    <div class="content">
                        <div class="row">
                            <span class="label">Học sinh</span>
                            <span class="value">{ten}</span>
                        </div>
                        <div class="row">
                            <span class="label">{icon_receipt}Học phí / buổi</span>
                            <span class="value">{hoc_phi:,} đ</span>
                        </div>
                        <div class="row">
                            <span class="label">{icon_calendar}Số buổi học</span>
                            <span class="value">{so_buoi} buổi</span>
                        </div>
                        
                        <div class="total-box">
                            <div style="font-size: 16px; color: #8e7f72; font-weight: bold; letter-spacing: 1.5px;">TỔNG HỌC PHÍ</div>
                            <div class="amount">{tong_tien:,} đ</div>
                        </div>
                        
                        <div style="text-align: center; font-size: 15px; color: #9a8a7a; font-weight: bold; margin: 30px 0 15px; letter-spacing: 1.5px;">NGÀY ĐI HỌC</div>
                        <div style="text-align: center;">{days_html}</div>
                        
                        <div style="text-align: center; font-size: 15px; color: #9a8a7a; font-weight: bold; margin: 40px 0 15px; letter-spacing: 1.5px;">— NHẬN XÉT —</div>
                        <div class="remark-box">{nhan_xet}</div>
                        
                        <div class="qr-wrap">
                            {qr_html}
                            <div style="text-align:left;">
                                <div style="color: #d49a71; font-size: 14px; font-weight: bold; margin-bottom: 5px; letter-spacing: 1px;">MÃ THANH TOÁN</div>
                                <div style="font-size: 17px; font-weight: 700; color: #4a2e25; text-transform: uppercase;">{bank if bank != 'nan' else 'CHƯA CÓ NH'}</div>
                                <div class="stk-val">{stk if stk != 'nan' else 'CHƯA CÓ STK'}</div>
                                <div style="font-size: 13px; color: #9a8a7a; font-weight:bold; letter-spacing: 1px;">LỚP NHẠC PHÍM HỒNG</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="footer">🎩 Trân trọng cảm ơn quý phụ huynh!</div>
                </div>
            </body>
            </html>
            """

            try:
                pdf_bytes = HTML(string=html_template).write_pdf()
                safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
                zip_file.writestr(f"Phieu_Hoc_Phi_{safe_name}.pdf", pdf_bytes)
            except Exception as e:
                st.error(f"Lỗi tạo PDF cho {ten}: {e}")

    status_text.text("🎉 Hoàn tất! Đã xử lý xong toàn bộ danh sách. Mời bạn nhấn nút tải về bên dưới.")
    
    st.download_button(
        label="⬇️ TẢI XUỐNG TOÀN BỘ PHIẾU PDF",
        data=zip_buffer.getvalue(),
        file_name="Phieu_Hoc_Phi_Phim_Hong.zip",
        mime="application/zip"
    )
