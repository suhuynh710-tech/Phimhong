import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os
from weasyprint import HTML

st.set_page_config(page_title="Phím Hồng Music - Billing", layout="centered")

# --- CẤU HÌNH CỐ ĐỊNH ---
LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎨 Hệ thống xuất Phiếu Học Phí Tự Động")
st.write("Xuất cùng lúc: Xem Web + File Ảnh (PNG) + File PDF (Chuẩn in ấn)")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận danh sách {len(df)} học sinh. Đang chuẩn bị gói tải về...")

    # Chuẩn bị bộ nhớ đệm cho file ZIP
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            # 1. Xử lý dữ liệu
            ten = str(row['Họ và Tên']).strip()
            lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Năng Khiếu"
            hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
            so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
            tong_tien = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else 0
            nhan_xet = str(row['Nhận Xét Của GV']).strip() if pd.notna(row['Nhận Xét Của GV']) else "Bé học tập tích cực!"
            bank = str(row['Ngân Hàng']).strip()
            stk = str(row['STK']).split('.')[0] if pd.notna(row['STK']) else ""

            # Xử lý ngày học
            date_cols = [c for c in df.columns if '/' in str(c)]
            days_html = ""
            for col in date_cols:
                if str(row[col]).strip().upper() == 'X':
                    parts = col.split(' ')
                    if len(parts) > 1:
                        thu = parts[0]
                        d_parts = parts[1].split('/')
                        day_month = f"{d_parts[0]:>02}/{d_parts[1]:>02}"
                        days_html += f'<div class="date-pill"><div class="thu">{thu}</div>{day_month}</div>'

            # QR Code (Chỉ để tên)
            add_info = urllib.parse.quote(ten)
            qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_tien}&addInfo={add_info}"

            # 2. Tạo mẫu thiết kế chuẩn (HTML/CSS) - Fix lỗi font bằng font-family an toàn
            html_template = f"""
            <html>
            <head>
            <meta charset="UTF-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap');
                @page {{ size: 450px 1050px; margin: 0; }}
                body {{ font-family: 'Montserrat', Arial, sans-serif; margin: 0; background: #fdfaf6; color: #333; }}
                .container {{ width: 450px; height: 1050px; position: relative; overflow: hidden; background: #fdfaf6; }}
                .header {{ background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; text-align: center; padding: 40px 20px 30px; position: relative; }}
                .logo-img {{ position: absolute; top: 20px; left: 20px; width: 60px; height: 60px; border-radius: 50%; border: 2px solid white; }}
                .header h1 {{ font-size: 32px; margin: 10px 0; letter-spacing: 1px; font-weight: 700; text-transform: uppercase; }}
                .badge {{ display: inline-block; background: rgba(255,255,255,0.2); padding: 5px 15px; border-radius: 20px; font-size: 13px; border: 1px solid rgba(255,255,255,0.4); }}
                .content {{ padding: 25px 35px; }}
                .row {{ display: flex; justify-content: space-between; align-items: center; padding: 15px 0; border-bottom: 1px dashed #e2d5c4; }}
                .row .icon {{ font-size: 20px; margin-right: 15px; width: 30px; display: inline-block; }}
                .row .label {{ color: #6d5b4b; font-weight: 400; }}
                .row .value {{ font-weight: 700; color: #2c1a16; font-size: 17px; }}
                .total-box {{ border: 2px solid #ecdac8; border-radius: 15px; background: #fcf4e8; text-align: center; padding: 20px; margin: 20px 0; }}
                .total-box .amount {{ font-size: 34px; color: #4a2e25; font-weight: 700; }}
                .date-pill {{ background: #f7f1e9; border: 1px solid #e0d1c1; border-radius: 8px; padding: 6px 12px; margin: 4px; display: inline-block; text-align: center; min-width: 50px; }}
                .date-pill .thu {{ font-size: 10px; color: #8e7f72; }}
                .remark-box {{ border: 1px solid #f2e2b3; background: #fffdf5; border-radius: 12px; padding: 15px; text-align: center; color: #5a4b41; font-style: italic; margin-top: 10px; }}
                .qr-wrap {{ display: flex; gap: 15px; align-items: center; margin-top: 30px; padding: 15px; border: 1px dashed #d49a71; background: white; border-radius: 15px; }}
                .qr-img {{ width: 110px; height: 110px; border-radius: 10px; }}
                .stk-val {{ font-size: 18px; color: #bc6c65; font-weight: 700; }}
            </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <img src="{logo_b64}" class="logo-img">
                        <div style="font-size: 12px; letter-spacing: 2px;">LỚP NHẠC PHÍM HỒNG</div>
                        <h1>PHIẾU HỌC PHÍ</h1>
                        <div style="margin-bottom: 10px;">Tháng 5 / 2026</div>
                        <div class="badge">{lop} ✨</div>
                    </div>
                    <div class="content">
                        <div class="row">
                            <span><span class="icon">🎓</span><span class="label">Học sinh</span></span>
                            <span class="value">{ten}</span>
                        </div>
                        <div class="row">
                            <span><span class="icon">🧾</span><span class="label">Học phí / buổi</span></span>
                            <span class="value">{hoc_phi_buoi:,} đ</span>
                        </div>
                        <div class="row">
                            <span><span class="icon">📅</span><span class="label">Số buổi học</span></span>
                            <span class="value">{so_buoi} buổi</span>
                        </div>
                        <div class="total-box">
                            <div style="font-size: 12px; color: #6d5b4b; font-weight: 700;">TỔNG HỌC PHÍ</div>
                            <div class="amount">{tong_tien:,} đ</div>
                        </div>
                        <div style="text-align: center; font-size: 11px; color: #8e7f72; font-weight: 700; margin-bottom: 10px;">NGÀY ĐI HỌC</div>
                        <div style="text-align: center;">{days_html if days_html else 'Chưa có dữ liệu'}</div>
                        <div style="text-align: center; font-size: 11px; color: #8e7f72; font-weight: 700; margin: 20px 0 5px;">— NHẬN XÉT —</div>
                        <div class="remark-box">{nhan_xet}</div>
                        <div class="qr-wrap">
                            <img src="{qr_url}" class="qr-img">
                            <div>
                                <div style="color: #d49a71; font-size: 11px; font-weight: 700;">MÃ THANH TOÁN</div>
                                <div style="font-size: 14px; font-weight: 700; color: #4a2e25;">{bank}</div>
                                <div class="stk-val">{stk}</div>
                                <div style="font-size: 10px; color: #8e7f72;">LỚP NHẠC PHÍM HỒNG</div>
                            </div>
                        </div>
                    </div>
                </div>
            </body>
            </html>
            """

            # 3. Hiển thị xem trước trên Web
            st.markdown(html_template.replace('\n', ''), unsafe_allow_html=True)
            st.divider()

            # 4. Xuất file PDF (Dùng WeasyPrint cực đẹp)
            pdf_bytes = HTML(string=html_template).write_pdf()
            
            # 5. Lưu vào ZIP (Cả PDF và Ảnh nếu muốn)
            safe_name = ten.replace(' ', '_')
            zip_file.writestr(f"{safe_name}.pdf", pdf_bytes)
            
            # Lưu file HTML lẻ để bạn có thể tự chụp màn hình nếu thích độ nét cao nhất
            zip_file.writestr(f"{safe_name}.html", html_template.encode('utf-8'))

    # NÚT TẢI ZIP
    st.success("🎉 Đã sẵn sàng! Bạn tải file ZIP về, mỗi học sinh sẽ có 1 file riêng biệt.")
    st.download_button(
        label="⬇️ TẢI TOÀN BỘ PHIẾU (.ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Phieu_Hoc_Phi_Phim_Hong.zip",
        mime="application/zip"
    )
