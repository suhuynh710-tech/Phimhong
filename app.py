import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os
from weasyprint import HTML

st.set_page_config(page_title="Phím Hồng Music - PDF VIP", layout="centered")

# --- CẤU HÌNH LOGO CỐ ĐỊNH ---
LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎯 Hệ thống xuất Phiếu Học Phí (Bản Cân Đối)")
st.write("Đã nới rộng chiều ngang, tăng chiều cao để đảm bảo 100% không bị nhảy sang trang 2.")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

# Định nghĩa các Icon SVG chuyên nghiệp
icon_student = '''<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2.06-1.12V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>'''
icon_receipt = '''<svg viewBox="0 0 24 24" width="20" height="20" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'''
icon_calendar = '''<svg viewBox="0 0 24 24" width="20" height="20" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'''

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận danh sách {len(df)} học sinh. Đang xử lý...")

    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            ten = str(row['Họ và Tên']).strip()
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
                        days_html += f'<div class="date-pill" style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:8px 15px; margin:5px; display:inline-block; text-align:center;"><div style="font-size:10px; color:#8e7f72; margin-bottom:2px;">{thu}</div><div style="font-size:14px; font-weight:bold; color:#4a2e25;">{day_month}</div></div>'

            if not days_html:
                days_html = '<span style="color:#aaa; font-style:italic; font-size:14px;">Chưa có buổi học</span>'

            # QR Code
            qr_b64 = ""
            if bank and stk and bank != 'nan':
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_tien}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url)
                    qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                except: pass

            logo_html = f'<div style="margin: 0 auto 15px auto; width: 80px; height: 80px; border-radius: 50%; border: 2px solid white; background-image: url({logo_b64}); background-size: cover; background-position: center; box-shadow: 0 2px 10px rgba(0,0,0,0.15);"></div>' if logo_b64 else ''
            qr_html = f'<img src="{qr_b64}" style="width: 120px; height: 120px; border-radius: 10px;">' if qr_b64 else '<div style="font-size:10px; color:#999;">CHƯA CÓ QR</div>'

            # Nới rộng kích thước (550px x 1250px)
            html_template = f"""
            <html>
            <head>
            <meta charset="UTF-8">
            <style>
                @page {{ size: 550px 1250px; margin: 0; }}
                body {{ font-family: 'Times New Roman', serif; margin: 0; background: #fdfaf6; color: #333; }}
                .container {{ width: 550px; height: 1250px; background: #fdfaf6; position: relative; page-break-inside: avoid; }}
                .header {{ background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; text-align: center; padding: 50px 30px 35px; position: relative; }}
                .header h1 {{ font-size: 36px; margin: 10px 0; letter-spacing: 2px; font-weight: 700; text-transform: uppercase; }}
                .badge {{ display: inline-block; background: rgba(255,255,255,0.25); padding: 6px 20px; border-radius: 20px; font-size: 15px; border: 1px solid rgba(255,255,255,0.4); font-weight: bold; font-family: Arial, sans-serif; }}
                .content {{ padding: 30px 45px; font-family: Arial, sans-serif; }}
                .row {{ display: flex; justify-content: space-between; align-items: center; padding: 18px 0; border-bottom: 1px dashed #e2d5c4; }}
                .row .label {{ color: #6d5b4b; font-size: 16px; font-weight: 400; }}
                .row .value {{ font-weight: 700; color: #2c1a16; font-size: 19px; text-align: right; width: 65%; }}
                .total-box {{ border: 2px solid #ecdac8; border-radius: 18px; background: #fcf4e8; text-align: center; padding: 25px; margin: 25px 0; }}
                .total-box .amount {{ font-size: 38px; color: #4a2e25; font-weight: 700; margin-top: 5px; }}
                .remark-box {{ border: 1px solid #f2e2b3; background: #fffdf5; border-radius: 15px; padding: 20px; text-align: center; color: #5a4b41; font-style: italic; margin-top: 15px; line-height: 1.6; font-size: 16px; }}
                .qr-wrap {{ display: flex; gap: 20px; align-items: center; margin-top: 40px; padding: 20px; border: 1px dashed #d49a71; background: white; border-radius: 18px; }}
                .stk-val {{ font-size: 20px; color: #bc6c65; font-weight: 700; margin: 5px 0; }}
                .footer {{ text-align: center; padding: 40px 30px; font-size: 15px; color: #a49688; font-style: italic; }}
            </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        {logo_html}
                        <div style="font-size: 13px; letter-spacing: 3px; font-weight:bold; font-family: Arial, sans-serif;">LỚP NHẠC PHÍM HỒNG</div>
                        <h1>PHIẾU HỌC PHÍ</h1>
                        <div style="margin-bottom: 15px; font-size: 17px; font-family: Arial, sans-serif;">Tháng 5 / 2026</div>
                        <div class="badge">{lop}</div>
                    </div>
                    <div class="content">
                        <div class="row">
                            <div style="display:flex; align-items:center;">{icon_student}<span class="label">Học sinh</span></div>
                            <span class="value">{ten}</span>
                        </div>
                        <div class="row">
                            <div style="display:flex; align-items:center;">{icon_receipt}<span class="label">Học phí / buổi</span></div>
                            <span class="value">{hoc_phi:,} đ</span>
                        </div>
                        <div class="row">
                            <div style="display:flex; align-items:center;">{icon_calendar}<span class="label">Số buổi học</span></div>
                            <span class="value">{so_buoi} buổi</span>
                        </div>
                        <div class="total-box">
                            <div style="font-size: 15px; color: #6d5b4b; font-weight: 700;">TỔNG HỌC PHÍ</div>
                            <div class="amount">{tong_tien:,} đ</div>
                        </div>
                        <div style="text-align: center; font-size: 14px; color: #8e7f72; font-weight: 700; margin: 25px 0 10px; letter-spacing: 1px;">NGÀY ĐI HỌC</div>
                        <div style="text-align: center;">{days_html}</div>
                        <div style="text-align: center; font-size: 14px; color: #8e7f72; font-weight: 700; margin: 35px 0 10px; letter-spacing: 1px;">— NHẬN XÉT —</div>
                        <div class="remark-box">{nhan_xet}</div>
                        <div class="qr-wrap">
                            {qr_html}
                            <div style="text-align:left;">
                                <div style="color: #d49a71; font-size: 13px; font-weight: 700; margin-bottom: 5px;">MÃ THANH TOÁN</div>
                                <div style="font-size: 16px; font-weight: 700; color: #4a2e25;">{bank if bank != 'nan' else 'CHƯA CÓ NH'}</div>
                                <div class="stk-val">{stk if stk != 'nan' else 'CHƯA CÓ STK'}</div>
                                <div style="font-size: 12px; color: #8e7f72; font-weight:bold;">LỚP NHẠC PHÍM HỒNG</div>
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

            # Hiển thị demo người đầu tiên lên web
            if index == 0:
                st.markdown(html_template.replace('\n', ''), unsafe_allow_html=True)
                st.write("<div style='height:50px'></div>", unsafe_allow_html=True)
