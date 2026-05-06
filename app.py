import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os

st.set_page_config(page_title="Phím Hồng Music - HTML", layout="centered")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("✨ Hệ thống xuất Phiếu Học Phí (Bản HTML Siêu Nét)")
st.write("Đã khôi phục icon 🎓, logo chính giữa, xuất HTML chống rớt trang 100%.")

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
            tong_tien = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else 0
            nhan_xet = str(row['Nhận Xét Của GV']).strip() if pd.notna(row['Nhận Xét Của GV']) else "Bé học tập tích cực, tự tin giao tiếp."
            bank = str(row['Ngân Hàng']).strip()
            stk = str(row['STK']).split('.')[0] if pd.notna(row['STK']) else ""

            # Xử lý Ngày
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
                days_html = '<span style="color:#aaa; font-style:italic; font-size:14px;">Chưa có buổi học</span>'

            # QR Code VietQR
            qr_html = ""
            if bank and stk and bank != 'nan':
                add_info = urllib.parse.quote(ten) # Chỉ để tên bé
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_tien}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url, timeout=3)
                    if resp.status_code == 200:
                        qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                        qr_html = f'<img src="{qr_b64}" style="width: 110px; height: 110px; border-radius: 12px;">'
                except: pass
            
            if not qr_html:
                qr_html = '<div style="font-size:11px; color:#999; padding:20px; border:1px dashed #ccc; border-radius:8px;">CHƯA CÓ QR</div>'

            # Logo chính giữa
            logo_html = f'<div style="margin: 0 auto 12px auto; width: 65px; height: 65px; border-radius: 50%; border: 2px solid white; background-image: url({logo_b64}); background-size: cover; background-position: center; box-shadow: 0 3px 8px rgba(0,0,0,0.15);"></div>' if logo_b64 else ''

            # Thẻ HTML cho Phiếu (Tự co giãn chiều cao)
            receipt_html = f"""
            <div style="width: 450px; margin: auto; background: #fdfaf6; box-shadow: 0 8px 25px rgba(0,0,0,0.15); border-radius: 15px; overflow: hidden; font-family: Arial, sans-serif; color: #333;">
                <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; text-align: center; padding: 35px 20px 25px;">
                    {logo_html}
                    <div style="font-size: 12px; letter-spacing: 2px; font-weight:bold;">LỚP NHẠC PHÍM HỒNG</div>
                    <h1 style="font-size: 32px; margin: 10px 0; font-family: 'Times New Roman', serif; text-transform: uppercase;">PHIẾU HỌC PHÍ</h1>
                    <div style="margin-bottom: 15px; font-size: 15px;">Tháng 5 / 2026</div>
                    <div style="display: inline-block; background: rgba(255,255,255,0.25); padding: 6px 20px; border-radius: 20px; font-size: 13px; border: 1px solid rgba(255,255,255,0.4); font-weight: bold;">{lop}</div>
                </div>
                
                <div style="padding: 25px 35px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 0; border-bottom: 1px dashed #e2d5c4;">
                        <div style="display:flex; align-items:center; color: #6d5b4b; font-size: 16px;"><span style="font-size: 20px; margin-right: 12px;">🎓</span> Học sinh</div>
                        <strong style="color: #2c1a16; font-size: 17px; text-align: right;">{ten}</strong>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 0; border-bottom: 1px dashed #e2d5c4;">
                        <div style="display:flex; align-items:center; color: #6d5b4b; font-size: 16px;"><span style="font-size: 20px; margin-right: 12px;">🧾</span> Học phí / buổi</div>
                        <strong style="color: #2c1a16; font-size: 17px; text-align: right;">{hoc_phi:,} đ</strong>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 0;">
                        <div style="display:flex; align-items:center; color: #6d5b4b; font-size: 16px;"><span style="font-size: 20px; margin-right: 12px;">📅</span> Số buổi học</div>
                        <strong style="color: #2c1a16; font-size: 17px; text-align: right;">{so_buoi} buổi</strong>
                    </div>
                    
                    <div style="border: 2px solid #ecdac8; border-radius: 15px; background: #fcf4e8; text-align: center; padding: 20px; margin: 25px 0;">
                        <div style="font-size: 14px; color: #8e7f72; font-weight: bold; letter-spacing: 1px;">TỔNG HỌC PHÍ</div>
                        <div style="font-size: 34px; color: #4a2e25; font-weight: 700; margin-top: 5px;">{tong_tien:,} đ</div>
                    </div>
                    
                    <div style="text-align: center; font-size: 13px; color: #8e7f72; font-weight: bold; margin: 20px 0 12px; letter-spacing: 1px;">NGÀY ĐI HỌC</div>
                    <div style="text-align: center;">{days_html}</div>
                    
                    <div style="text-align: center; font-size: 13px; color: #8e7f72; font-weight: bold; margin: 30px 0 12px; letter-spacing: 1px;">— NHẬN XÉT —</div>
                    <div style="border: 1px solid #f2e2b3; background: #fffdf5; border-radius: 12px; padding: 18px; text-align: center; color: #5a4b41; font-style: italic; line-height: 1.5; font-size: 15px;">{nhan_xet}</div>
                    
                    <div style="display: flex; gap: 20px; align-items: center; margin-top: 35px; padding: 20px; border: 1px dashed #d49a71; background: white; border-radius: 15px;">
                        {qr_html}
                        <div style="text-align:left;">
                            <div style="color: #d49a71; font-size: 12px; font-weight: bold; margin-bottom: 5px;">MÃ THANH TOÁN</div>
                            <div style="font-size: 15px; font-weight: 700; color: #4a2e25;">{bank if bank != 'nan' else 'CHƯA CÓ NH'}</div>
                            <div style="font-size: 18px; color: #bc6c65; font-weight: 800; margin: 4px 0;">{stk if stk != 'nan' else 'CHƯA CÓ STK'}</div>
                            <div style="font-size: 11px; color: #8e7f72; font-weight:bold;">LỚP NHẠC PHÍM HỒNG</div>
                        </div>
                    </div>
                </div>
                
                <div style="text-align: center; padding: 25px 20px; font-size: 14px; color: #9a8a7a; font-style: italic;">🎩 Trân trọng cảm ơn quý phụ huynh!</div>
            </div>
            """

            # Code HTML hoàn chỉnh cho file tải về (Nền xám nhạt để nổi bật tờ phiếu)
            full_html_file = f"""<!DOCTYPE html>
            <html lang="vi">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Phiếu Học Phí - {ten}</title>
            </head>
            <body style="background-color: #e9ecef; display: flex; justify-content: center; padding: 40px 10px; margin: 0;">
                {receipt_html}
            </body>
            </html>"""

            # Ép ghi file với bảng mã UTF-8 để không lỗi dấu Tiếng Việt
            safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
            zip_file.writestr(f"Phieu_Hoc_Phi_{safe_name}.html", full_html_file.encode('utf-8'))

            # Hiển thị demo phiếu đầu tiên cực sạch trên web
            if index == 0:
                clean_preview = receipt_html.replace('\n', '')
                st.markdown(clean_preview, unsafe_allow_html=True)
                st.write("<br>", unsafe_allow_html=True)

    st.success("🎉 Đã xong! Bạn tải file ZIP về, mở file HTML lên xem và chụp màn hình cực kỳ nét nhé.")
    st.download_button(
        label="⬇️ TẢI XUỐNG FILE ZIP (BẢN HTML HOÀN HẢO)",
        data=zip_buffer.getvalue(),
        file_name="Phieu_Hoc_Phi_HTML.zip",
        mime="application/zip"
    )
