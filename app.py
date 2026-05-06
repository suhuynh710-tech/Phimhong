import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os

st.set_page_config(page_title="Phím Hồng Music - Phiếu Học Phí A4", layout="wide")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎯 Hệ thống xuất Phiếu Học Phí (Bố Cục A4 Ngang)")
st.write("Đã đưa logo qua góc trái và đổi 'Tiền phát sinh' thành 'Tiền sách'.")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận danh sách {len(df)} học sinh. Đang đóng gói file ZIP...")

    zip_buffer = io.BytesIO()
    progress_bar = st.progress(0)
    total_students = len(df)
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            ten = str(row['Họ và Tên']).strip()
            progress_bar.progress((index + 1) / total_students)
            
            lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Năng Khiếu"
            hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
            so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
            
            # Lấy Tổng học phí gốc từ Excel
            tong_tien_goc = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else (hoc_phi * so_buoi)
            
            # --- TÍNH NĂNG: TIỀN SÁCH ---
            tien_sach = 0
            phat_sinh_html = ""
            if 'Tiền sách' in df.columns:
                val = row['Tiền sách']
                if pd.notna(val) and str(val).strip() != '':
                    try:
                        tien_sach = int(float(val))
                        if tien_sach > 0:
                            phat_sinh_html = f'''
                            <tr>
                                <td style="padding: 12px 0; color: #7a6b5d; border-top: 1px dashed #e2d5c4; font-size: 17px;">Tiền sách / Giáo trình:</td>
                                <td style="padding: 12px 0; font-weight: bold; color: #bc6c65; text-align: right; border-top: 1px dashed #e2d5c4; font-size: 20px;">+ {tien_sach:,} đ</td>
                            </tr>
                            '''
                    except: pass
            
            # Tổng thanh toán cuối cùng = Tiền học + Tiền sách
            tong_thanh_toan = tong_tien_goc + tien_sach
            
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
                        days_html += f'<div style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; margin:4px; display:inline-block; text-align:center;"><div style="font-size:11px; color:#8e7f72; margin-bottom:3px;">{thu}</div><div style="font-size:14px; font-weight:bold; color:#4a2e25;">{day_month}</div></div>'

            if not days_html:
                days_html = '<span style="color:#aaa; font-style:italic; font-size:14px;">Chưa có dữ liệu điểm danh</span>'

            # QR Code VietQR
            qr_html = ""
            if bank and stk and bank != 'nan':
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_thanh_toan}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url, timeout=3)
                    if resp.status_code == 200:
                        qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                        qr_html = f'<img src="{qr_b64}" style="width: 125px; height: 125px; border-radius: 12px;">'
                except: pass
            
            if not qr_html:
                qr_html = '<div style="font-size:12px; color:#999; padding:40px 0; border:1px dashed #ccc; border-radius:8px;">CHƯA CÓ QR</div>'

            # Chỉnh logo ở absolute top-left
            logo_html = f'<div style="position: absolute; top: 25px; left: 35px; width: 70px; height: 70px; border-radius: 50%; border: 2px solid white; background-image: url({logo_b64}); background-size: cover; background-position: center; box-shadow: 0 2px 5px rgba(0,0,0,0.2);"></div>' if logo_b64 else ''

            # BỐ CỤC A4 NGANG - HEADER GỌN GÀNG HƠN
            receipt_html = f"""
            <div style="width: 800px; background: white; font-family: 'Arial', sans-serif; margin: 0 auto; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border-radius: 15px; overflow: hidden; box-sizing: border-box; position: relative;">
                
                <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); padding: 30px 45px 30px 125px; display: flex; align-items: center; justify-content: space-between; color: white; position: relative;">
                    {logo_html}
                    <div style="text-align: left;">
                        <h1 style="margin: 0; font-size: 32px; font-family: 'Times New Roman', serif; text-transform: uppercase;">Phiếu Học Phí</h1>
                        <div style="font-size: 14px; margin-top: 5px; letter-spacing: 2px; font-weight: bold; opacity: 0.9;">LỚP NHẠC PHÍM HỒNG</div>
                    </div>
                    <div style="text-align: right;">
                        <div style="font-size: 15px; font-weight: bold;">Tháng 5 / 2026</div>
                        <div style="display: inline-block; margin-top: 8px; padding: 5px 18px; background: rgba(255,255,255,0.2); border: 1px solid rgba(255,255,255,0.5); border-radius: 20px; font-weight: bold; font-size: 13px;">Lớp {lop}</div>
                    </div>
                </div>

                <div style="padding: 35px 45px;">
                    <div style="background: #fdfaf6; border: 1px solid #f2e2b3; border-radius: 15px; padding: 20px 35px; margin-bottom: 30px;">
                        <table style="width: 100%; border-collapse: collapse; font-size: 17px;">
                            <tr>
                                <td style="padding: 12px 0; color: #7a6b5d; width: 40%;">Học sinh:</td>
                                <td style="padding: 12px 0; font-weight: bold; color: #2c1a16; text-align: right; font-size: 20px;">{ten}</td>
                            </tr>
                            <tr>
                                <td style="padding: 12px 0; color: #7a6b5d; border-top: 1px dashed #e2d5c4;">Học phí / buổi:</td>
                                <td style="padding: 12px 0; font-weight: bold; color: #2c1a16; text-align: right; border-top: 1px dashed #e2d5c4;">{hoc_phi:,} đ</td>
                            </tr>
                            <tr>
                                <td style="padding: 12px 0; color: #7a6b5d; border-top: 1px dashed #e2d5c4;">Số buổi học:</td>
                                <td style="padding: 12px 0; font-weight: bold; color: #2c1a16; text-align: right; border-top: 1px dashed #e2d5c4;">{so_buoi} buổi</td>
                            </tr>
                            {phat_sinh_html}
                        </table>
                    </div>

                    <div style="display: flex; gap: 30px;">
                        
                        <div style="flex: 1;">
                            <div style="margin-bottom: 25px;">
                                <div style="font-size: 13px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 10px;">NGÀY ĐI HỌC</div>
                                <div>{days_html}</div>
                            </div>
                            <div>
                                <div style="font-size: 13px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 10px;">NHẬN XÉT CỦA GIÁO VIÊN</div>
                                <div style="background: #fffdf5; border: 1px solid #f2e2b3; border-radius: 12px; padding: 18px; color: #5a4b41; font-style: italic; line-height: 1.5; font-size: 15px;">
                                    {nhan_xet}
                                </div>
                            </div>
                        </div>

                        <div style="width: 260px; display: flex; flex-direction: column; gap: 15px;">
                            <div style="background: #fdf6ec; border: 2px solid #ecdac8; border-radius: 15px; padding: 20px 15px; text-align: center;">
                                <div style="font-size: 13px; color: #8e7f72; font-weight: bold; letter-spacing: 1px;">TỔNG THANH TOÁN</div>
                                <div style="font-size: 30px; color: #4a2e25; font-weight: 700; margin-top: 8px;">{tong_thanh_toan:,} đ</div>
                            </div>
                            <div style="background: white; border: 2px dashed #d49a71; border-radius: 15px; padding: 15px; text-align: center;">
                                <div style="font-size: 11px; color: #d49a71; font-weight: bold; margin-bottom: 12px; letter-spacing: 1px;">MÃ QUÉT THANH TOÁN</div>
                                {qr_html}
                                <div style="margin-top: 12px; font-size: 16px; font-weight: 800; color: #bc6c65; text-transform: uppercase;">{bank if bank != 'nan' else 'CHƯA CÓ NH'}</div>
                                <div style="font-size: 14px; font-weight: bold; color: #4a2e25; margin-top: 4px;">{stk if stk != 'nan' else 'CHƯA CÓ STK'}</div>
                            </div>
                        </div>
                    </div>

                    <div style="text-align: center; margin-top: 35px; font-size: 14px; color: #9a8a7a; font-style: italic;">
                        🎩 Trân trọng cảm ơn quý phụ huynh!
                    </div>
                </div>
            </div>
            """

            # Đóng gói HTML vào ZIP
            full_html_file = f"""<!DOCTYPE html>
            <html lang="vi">
            <head>
                <meta charset="UTF-8">
                <title>Phiếu Học Phí - {ten}</title>
            </head>
            <body style="background-color: #e9ecef; display: flex; justify-content: center; padding: 40px 10px; margin: 0;">
                {receipt_html}
            </body>
            </html>"""

            safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
            zip_file.writestr(f"Phieu_Hoc_Phi_{safe_name}.html", full_html_file.encode('utf-8'))

            if index == 0:
                st.markdown(receipt_html.replace('\n', ''), unsafe_allow_html=True)
                st.write("<br><br>", unsafe_allow_html=True)

    st.success("🎉 Đã hoàn tất! Mời bạn nhấn nút tải về bên dưới.")
    st.download_button(
        label="⬇️ TẢI XUỐNG FILE ZIP (BẢN A4 NGANG MỚI)",
        data=zip_buffer.getvalue(),
        file_name="Phieu_Hoc_Phi_A4.zip",
        mime="application/zip"
    )
