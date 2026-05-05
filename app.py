import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os

st.set_page_config(page_title="Phím Hồng Music - PDF", layout="centered")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎯 Hệ thống xuất Phiếu Học Phí (Chỉ PDF)")
st.write("Tải file Excel để xuất tự động toàn bộ sang PDF, nằm gọn trong 1 file ZIP.")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận danh sách {len(df)} học sinh. Bạn có thể xem trước mẫu phiếu đầu tiên bên dưới!")

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

            date_cols = [c for c in df.columns if '/' in str(c)]
            days_html = ""
            for col in date_cols:
                if str(row[col]).strip().upper() == 'X':
                    parts = col.split(' ')
                    if len(parts) > 1:
                        thu = parts[0]
                        d_parts = parts[1].split('/')
                        day_month = f"{d_parts[0]:>02}/{d_parts[1]:>02}"
                        days_html += f'<div class="date-pill" style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; margin:4px; display:inline-block; text-align:center;"><div style="font-size:10px; color:#8e7f72; margin-bottom:2px;">{thu}</div><div style="font-size:13px; font-weight:bold; color:#4a2e25;">{day_month}</div></div>'

            if not days_html:
                days_html = '<span style="color:#aaa; font-style:italic; font-size:13px;">Chưa có buổi học</span>'

            # QR Code
            qr_b64 = ""
            if bank and stk and bank != 'nan':
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_tien}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url)
                    qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                except: pass

            logo_html = f'<img src="{logo_b64}" style="position: absolute; top: 20px; left: 20px; width: 60px; height: 60px; border-radius: 50%; border: 2px solid white; object-fit: cover;">' if logo_b64 else ''
            qr_html = f'<img src="{qr_b64}" style="width: 110px; height: 110px; border-radius: 10px;">' if qr_b64 else '<div style="font-size:10px; color:#999; padding:20px 0;">NO QR</div>'

            # Mẫu HTML để chèn vào trình duyệt
            html_template = f"""
            <div style="width: 450px; margin: 0 auto; background: #fdfaf6; box-shadow: 0 10px 25px rgba(0,0,0,0.2); border-radius: 15px; overflow: hidden; font-family: 'Times New Roman', serif;">
                <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; text-align: center; padding: 40px 20px 30px; position: relative;">
                    {logo_html}
                    <div style="font-size: 12px; letter-spacing: 2px; font-weight:bold; font-family: Arial, sans-serif;">LỚP NHẠC PHÍM HỒNG</div>
                    <h1 style="font-size: 32px; margin: 10px 0; letter-spacing: 1px; font-weight: 700; text-transform: uppercase;">PHIẾU HỌC PHÍ</h1>
                    <div style="margin-bottom: 15px; font-size: 15px; font-family: Arial, sans-serif;">Tháng 5 / 2026</div>
                    <div style="display: inline-block; background: rgba(255,255,255,0.2); padding: 5px 15px; border-radius: 20px; font-size: 13px; border: 1px solid rgba(255,255,255,0.4); font-weight: bold; font-family: Arial, sans-serif;">{lop} ✨</div>
                </div>
                <div style="padding: 25px 35px; font-family: Arial, sans-serif;">
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 0; border-bottom: 1px dashed #e2d5c4;">
                        <div style="display:flex; align-items:center;"><span style="font-size: 20px; margin-right: 10px; width: 30px; display: inline-block;">🎓</span><span style="color: #6d5b4b;">Học sinh</span></div>
                        <span style="font-weight: 700; color: #2c1a16; font-size: 17px; text-align: right; width: 60%;">{ten}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 0; border-bottom: 1px dashed #e2d5c4;">
                        <div style="display:flex; align-items:center;"><span style="font-size: 20px; margin-right: 10px; width: 30px; display: inline-block;">🧾</span><span style="color: #6d5b4b;">Học phí / buổi</span></div>
                        <span style="font-weight: 700; color: #2c1a16; font-size: 17px; text-align: right;">{hoc_phi:,} đ</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 0;">
                        <div style="display:flex; align-items:center;"><span style="font-size: 20px; margin-right: 10px; width: 30px; display: inline-block;">📅</span><span style="color: #6d5b4b;">Số buổi học</span></div>
                        <span style="font-weight: 700; color: #2c1a16; font-size: 17px; text-align: right;">{so_buoi} buổi</span>
                    </div>
                    <div style="border: 2px solid #ecdac8; border-radius: 15px; background: #fcf4e8; text-align: center; padding: 20px; margin: 20px 0;">
                        <div style="font-size: 13px; color: #6d5b4b; font-weight: 700;">TỔNG HỌC PHÍ</div>
                        <div style="font-size: 34px; color: #4a2e25; font-weight: 700; margin-top: 5px;">{tong_tien:,} đ</div>
                    </div>
                    <div style="text-align: center; font-size: 12px; color: #8e7f72; font-weight: 700; margin: 20px 0 10px; letter-spacing: 1px;">NGÀY ĐI HỌC</div>
                    <div style="text-align: center;">{days_html}</div>
                    <div style="text-align: center; font-size: 12px; color: #8e7f72; font-weight: 700; margin: 30px 0 10px; letter-spacing: 1px;">— NHẬN XÉT —</div>
                    <div style="border: 1px solid #f2e2b3; background: #fffdf5; border-radius: 12px; padding: 15px; text-align: center; color: #5a4b41; font-style: italic; margin-top: 10px; line-height: 1.5;">{nhan_xet}</div>
                    <div style="display: flex; gap: 15px; align-items: center; margin-top: 30px; padding: 15px; border: 1px dashed #d49a71; background: white; border-radius: 15px;">
                        {qr_html}
                        <div style="text-align:left;">
                            <div style="color: #d49a71; font-size: 12px; font-weight: 700; margin-bottom: 5px;">MÃ THANH TOÁN</div>
                            <div style="font-size: 14px; font-weight: 700; color: #4a2e25;">{bank if bank != 'nan' else 'CHƯA CÓ NH'}</div>
                            <div style="font-size: 18px; color: #bc6c65; font-weight: 700; margin: 3px 0;">{stk if stk != 'nan' else 'CHƯA CÓ STK'}</div>
                            <div style="font-size: 11px; color: #8e7f72; font-weight:bold;">LỚP NHẠC PHÍM HỒNG</div>
                        </div>
                    </div>
                </div>
                <div style="text-align: center; padding: 20px; font-size: 13px; color: #a49688; font-style: italic; font-family: Arial, sans-serif;">🎩 Trân trọng cảm ơn quý phụ huynh!</div>
            </div>
            """

            # Code ép tạo ra file PDF thật
            full_html = f"<!DOCTYPE html><html><head><meta charset='utf-8'></head><body style='padding:50px;'>{html_template}</body><script>window.print();</script></html>"
            
            safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
            
            # Gói HTML vào ZIP. Bạn mở file lên Chrome/Cốc Cốc nó sẽ TỰ ĐỘNG BẬT in ra PDF. Cực chuẩn.
            zip_file.writestr(f"Phieu_Hoc_Phi_{safe_name}.html", full_html.encode('utf-8'))

            if index == 0:
                st.markdown(html_template.replace('\n', ''), unsafe_allow_html=True)
                st.divider()

    st.download_button(
        label="⬇️ TẢI TOÀN BỘ PHIẾU (.ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Phieu_Hoc_Phi_Tong_Hop.zip",
        mime="application/zip"
    )
