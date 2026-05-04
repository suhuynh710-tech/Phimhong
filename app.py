import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os

st.set_page_config(page_title="Công cụ tạo Phiếu Học Phí", layout="centered")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("✨ Công cụ tạo Phiếu Học Phí & Xuất File ZIP")
st.write("Cập nhật: Bố cục hoàn hảo, icon chuẩn, QR Code tối giản.")

uploaded_file = st.file_uploader("📂 Tải file Excel (Danh_Sach_Hoc_Phi.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    
    logo_b64 = ""
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            logo_b64 = f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"

    st.success(f"Đã tải lên danh sách {len(df)} học sinh.")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            ten = str(row['Họ và Tên']).strip()
            lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Năng Khiếu"
            hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
            so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
            tong_tien = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else 0
            
            nhan_xet = str(row['Nhận Xét Của GV']).strip()
            if nhan_xet == 'nan' or not nhan_xet:
                nhan_xet = "Bé học tập tích cực, tự tin giao tiếp."
                
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
                        day = f"0{d_parts[0]}" if len(d_parts[0]) == 1 else d_parts[0]
                        month = f"0{d_parts[1]}" if len(d_parts[1]) == 1 else d_parts[1]
                        days_html += f'<div style="background: #f7f1e9; border: 1px solid #e0d1c1; border-radius: 8px; padding: 5px 10px; margin: 3px; display: inline-block; text-align: center;"><div style="font-size: 10px; color: #8e7f72; margin-bottom: 2px;">{thu}</div><div style="font-size: 13px; font-weight: bold; color: #4a2e25;">{day}/{month}</div></div>'
            
            if not days_html:
                days_html = '<span style="color:#aaa; font-style:italic; font-size:13px;">Chưa có buổi học</span>'
            
            qr_b64 = ""
            if bank and stk and bank != 'nan':
                # Đổi nội dung mã QR chỉ giữ Tên người chuyển
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_tien}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url)
                    qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                except: pass

            logo_html = f'<img src="{logo_b64}" style="position: absolute; top: 20px; left: 20px; width: 50px; height: 50px; border-radius: 50%; border: 2px solid rgba(255,255,255,0.8); object-fit: cover;">' if logo_b64 else ''
            qr_html = f'<img src="{qr_b64}" style="width: 100%; border-radius: 8px;">' if qr_b64 else '<div style="font-size:10px; color:#999; padding:20px 0;">NO QR</div>'

            html_string = f"""
            <div style="width: 480px; margin: auto; font-family: Arial, sans-serif; background: #fdfaf6; box-shadow: 0 4px 15px rgba(0,0,0,0.15); margin-bottom: 40px; overflow: hidden; position: relative; border-radius: 10px;">
                <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; text-align: center; padding: 35px 20px 25px; position: relative;">
                    {logo_html}
                    <div style="font-size: 11px; letter-spacing: 2px; font-weight: 600; margin-bottom: 8px;">LỚP NHẠC PHÍM HỒNG</div>
                    <h1 style="font-size: 30px; font-weight: bold; margin: 0 0 10px 0; font-family: 'Times New Roman', Times, serif;">PHIẾU HỌC PHÍ</h1>
                    <div style="font-size: 14px; margin-bottom: 15px;">Tháng 5 / 2026</div>
                    <span style="background: rgba(255,255,255,0.25); padding: 5px 15px; border-radius: 20px; font-size: 12px; font-weight: bold; border: 1px solid rgba(255,255,255,0.3);">{lop} ✨</span>
                </div>
                <div style="padding: 20px 35px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 14px 0; border-bottom: 1px dashed #e2d5c4; color: #6d5b4b;">
                        <div style="display: flex; align-items: center;"><span style="width: 35px; text-align: left; display: inline-block; font-size:18px;">🎓</span> <span>Học sinh</span></div>
                        <strong style="color: #2c1a16; font-size: 16px; text-align: right; width: 60%;">{ten}</strong>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 14px 0; border-bottom: 1px dashed #e2d5c4; color: #6d5b4b;">
                        <div style="display: flex; align-items: center;"><span style="width: 35px; text-align: left; display: inline-block; font-size:18px;">🧾</span> <span>Học phí / buổi</span></div>
                        <strong style="color: #2c1a16; font-size: 16px; text-align: right;">{hoc_phi:,} đ</strong>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; padding: 14px 0; color: #6d5b4b;">
                        <div style="display: flex; align-items: center;"><span style="width: 35px; text-align: left; display: inline-block; font-size:18px;">📅</span> <span>Số buổi học</span></div>
                        <strong style="color: #2c1a16; font-size: 16px; text-align: right;">{so_buoi} buổi</strong>
                    </div>
                    <div style="border: 2px solid #ecdac8; background: #fcf4e8; border-radius: 12px; text-align: center; padding: 18px; margin: 15px 0;">
                        <div style="font-size: 12px; color: #6d5b4b; font-weight: bold; margin-bottom: 6px;">TỔNG HỌC PHÍ</div>
                        <div style="font-size: 32px; color: #4a2e25; font-weight: bold;">{tong_tien:,} đ</div>
                    </div>
                    <div style="text-align: center; font-size: 11px; color: #8e7f72; font-weight: bold; margin: 15px 0 10px;">NGÀY ĐI HỌC</div>
                    <div style="text-align: center;">{days_html}</div>
                    <div style="text-align: center; font-size: 11px; color: #8e7f72; font-weight: bold; margin: 25px 0 10px;">— NHẬN XÉT —</div>
                    <div style="border: 1px solid #f2e2b3; background: #fffdf5; border-radius: 10px; padding: 15px; text-align: center; color: #5a4b41; font-style: italic;">
                        {nhan_xet}
                    </div>
                    <div style="display: flex; gap: 15px; align-items: center; margin-top: 25px; padding: 15px; border: 1px dashed #d49a71; background: #fff; border-radius: 12px;">
                        <div style="width: 100px;">{qr_html}</div>
                        <div style="flex: 1;">
                            <div style="color: #d49a71; font-size: 11px; font-weight: bold; margin-bottom: 5px;">MÃ THANH TOÁN</div>
                            <div style="font-size: 13px; font-weight: bold; color: #4a2e25;">{bank if bank != 'nan' else 'CHƯA CÓ NH'}</div>
                            <div style="font-size: 16px; font-weight: bold; color: #bc6c65; margin: 2px 0;">{stk if stk != 'nan' else 'CHƯA CÓ STK'}</div>
                            <div style="font-size: 10px; color: #8e7f72; font-weight: bold;">LỚP NHẠC PHÍM HỒNG</div>
                        </div>
                    </div>
                    <div style="text-align: center; color: #a49688; font-size: 12px; margin-top: 20px; font-style: italic;">
                        🎩 Trân trọng cảm ơn quý phụ huynh!
                    </div>
                </div>
            </div>
            """
            
            clean_html = html_string.replace('\n', '')
            st.markdown(clean_html, unsafe_allow_html=True)
            
            # Gói file HTML siêu nét vào ZIP
            safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
            full_html_file = f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>Phieu_{safe_name}</title></head><body style='background:#ccc; padding:50px;'>{clean_html}</body></html>"
            zip_file.writestr(f"Phieu_Hoc_Phi_{safe_name}.html", full_html_file.encode('utf-8'))

    st.download_button(
        label="⬇️ TẢI TOÀN BỘ PHIẾU (.ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Phieu_Hoc_Phi_Thang_5.zip",
        mime="application/zip"
    )