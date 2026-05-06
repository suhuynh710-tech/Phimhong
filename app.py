import streamlit as st
import pandas as pd
import requests
import base64
import urllib.parse
import zipfile
import io
import os

st.set_page_config(page_title="Phím Hồng Music - Phiếu Học Phí VIP", layout="wide")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎯 Hệ thống xuất Phiếu Học Phí (Bản Cân Đối Hoàn Hảo)")
st.write("Đã chỉnh: Logo & Tên lớp chung dòng, Tiêu đề nhỏ lại, Cân giữa toàn bộ bảng biểu.")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

# Icon SVG nét căng, không lỗi ô vuông
icon_receipt = '''<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:10px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'''
icon_calendar = '''<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:10px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'''

if uploaded_file:
    df = pd.read_excel(uploaded_file).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.success(f"Đã nhận {len(df)} học sinh. Đang xử lý file ZIP...")

    zip_buffer = io.BytesIO()
    progress_bar = st.progress(0)
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            ten = str(row['Họ và Tên']).strip()
            progress_bar.progress((index + 1) / len(df))
            
            lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Piano"
            hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
            so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
            tong_tien_goc = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else (hoc_phi * so_buoi)
            
            # Tiền sách
            tien_sach = 0
            tien_sach_html = ""
            if 'Tiền sách' in df.columns:
                val = row['Tiền sách']
                if pd.notna(val) and str(val).strip() != '':
                    tien_sach = int(float(val))
                    if tien_sach > 0:
                        tien_sach_html = f'''
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 12px 0; color: #7a6b5d;">Tiền sách / Giáo trình:</td>
                            <td style="padding: 12px 0; font-weight: bold; color: #bc6c65; text-align: right;">+ {tien_sach:,} đ</td>
                        </tr>
                        '''
            
            tong_thanh_toan = tong_tien_goc + tien_sach
            nhan_xet = str(row['Nhận Xét Của GV']).strip() if pd.notna(row['Nhận Xét Của GV']) else "Bé học tập tích cực!"
            bank = str(row['Ngân Hàng']).strip()
            stk = str(row['STK']).split('.')[0] if pd.notna(row['STK']) else ""

            # Ngày học
            date_cols = [c for c in df.columns if '/' in str(c)]
            days_html = ""
            for col in date_cols:
                if str(row[col]).strip().upper() == 'X':
                    parts = col.split(' ')
                    if len(parts) > 1:
                        thu = parts[0]
                        d_parts = parts[1].split('/')
                        day_month = f"{d_parts[0]:>02}/{d_parts[1]:>02}"
                        days_html += f'<div style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; margin:4px; display:inline-block; text-align:center;"><div style="font-size:10px; color:#8e7f72; margin-bottom:2px;">{thu}</div><div style="font-size:13px; font-weight:bold; color:#4a2e25;">{day_month}</div></div>'

            # QR Code
            qr_html = ""
            if bank and stk:
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_thanh_toan}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url, timeout=3)
                    if resp.status_code == 200:
                        qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode()}"
                        qr_html = f'<img src="{qr_b64}" style="width: 120px; height: 120px; border-radius: 10px;">'
                except: pass

            receipt_html = f"""
            <div style="width: 850px; background: white; font-family: 'Arial', sans-serif; margin: 0 auto; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border-radius: 20px; overflow: hidden; box-sizing: border-box;">
                
                <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); padding: 30px 50px; display: flex; align-items: center; justify-content: space-between; color: white;">
                    <div style="display: flex; align-items: center; gap: 20px;">
                        <div style="width: 85px; height: 85px; border-radius: 50%; border: 3px solid white; background-image: url({logo_b64}); background-size: cover; background-position: center;"></div>
                        <div style="font-size: 24px; letter-spacing: 2px; font-weight: 900; text-transform: uppercase; font-family: 'Times New Roman', serif;">Lớp Nhạc Phím Hồng</div>
                    </div>
                    
                    <div style="text-align: center; flex-grow: 1; padding: 0 20px;">
                        <h2 style="margin: 0; font-size: 22px; opacity: 0.95; font-weight: normal; letter-spacing: 1px;">PHIẾU HỌC PHÍ</h2>
                    </div>

                    <div style="text-align: right; min-width: 150px; padding-right: 20px;">
                        <div style="font-size: 18px; font-weight: bold; margin-bottom: 5px;">Tháng 5 / 2026</div>
                        <div style="font-size: 20px; font-weight: 900; background: rgba(255,255,255,0.25); padding: 5px 15px; border-radius: 10px;">Lớp {lop}</div>
                    </div>
                </div>

                <div style="padding: 40px 60px;">
                    <div style="background: #fdfaf6; border: 1px solid #f2e2b3; border-radius: 15px; padding: 30px 40px; margin: 0 auto 30px auto; width: 90%;">
                        <table style="width: 100%; border-collapse: collapse; font-size: 19px;">
                            <tr>
                                <td style="padding: 15px 0; color: #7a6b5d;">Học sinh:</td>
                                <td style="padding: 15px 0; font-weight: 900; color: #2c1a16; text-align: right; font-size: 24px;">{ten}</td>
                            </tr>
                            <tr style="border-top: 1px dashed #e2d5c4;">
                                <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{icon_receipt} Học phí / buổi:</td>
                                <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{hoc_phi:,} đ</td>
                            </tr>
                            <tr style="border-top: 1px dashed #e2d5c4;">
                                <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{icon_calendar} Số buổi học:</td>
                                <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{so_buoi} buổi</td>
                            </tr>
                            {tien_sach_html}
                        </table>
                    </div>

                    <div style="display: flex; gap: 40px; align-items: flex-start; width: 90%; margin: 0 auto;">
                        <div style="flex: 1.2;">
                            <div style="margin-bottom: 25px;">
                                <div style="font-size: 14px; color: #8e7f72; font-weight: bold; margin-bottom: 12px; letter-spacing: 1px;">NGÀY ĐI HỌC</div>
                                <div style="text-align: left;">{days_html}</div>
                            </div>
                            <div>
                                <div style="font-size: 14px; color: #8e7f72; font-weight: bold; margin-bottom: 12px; letter-spacing: 1px;">NHẬN XÉT CỦA GIÁO VIÊN</div>
                                <div style="background: #fffdf5; border: 1px solid #f2e2b3; border-radius: 12px; padding: 20px; color: #5a4b41; font-style: italic; line-height: 1.6; font-size: 16px;">
                                    {nhan_xet}
                                </div>
                            </div>
                        </div>

                        <div style="flex: 0.8; display: flex; flex-direction: column; gap: 20px;">
                            <div style="background: #fdf6ec; border: 2px solid #ecdac8; border-radius: 15px; padding: 20px; text-align: center;">
                                <div style="font-size: 13px; color: #8e7f72; font-weight: bold;">TỔNG THANH TOÁN</div>
                                <div style="font-size: 36px; color: #4a2e25; font-weight: 900; margin-top: 10px;">{tong_thanh_toan:,} đ</div>
                            </div>
                            <div style="background: white; border: 2px dashed #d49a71; border-radius: 15px; padding: 20px; text-align: center;">
                                <div style="font-size: 11px; color: #d49a71; font-weight: bold; margin-bottom: 15px;">QUÉT MÃ CHUYỂN KHOẢN</div>
                                {qr_html}
                                <div style="margin-top: 15px; font-size: 18px; font-weight: 900; color: #bc6c65;">{bank}</div>
                                <div style="font-size: 16px; font-weight: bold; color: #4a2e25; margin-top: 5px;">{stk}</div>
                            </div>
                        </div>
                    </div>

                    <div style="text-align: center; margin-top: 50px; font-size: 16px; color: #9a8a7a; font-style: italic;">
                        🎩 Trân trọng cảm ơn quý phụ huynh!
                    </div>
                </div>
            </div>
            """

            full_html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="background:#e9ecef; display:flex; justify-content:center; padding:50px 0;">{receipt_html}</body></html>"""
            safe_name = ten.replace(' ', '_')
            zip_file.writestr(f"Phieu_{safe_name}.html", full_html.encode('utf-8'))

            if index == 0:
                st.markdown(receipt_html.replace('\n', ''), unsafe_allow_html=True)

    st.download_button(label="⬇️ TẢI XUỐNG ZIP PHIẾU HOÀN HẢO", data=zip_buffer.getvalue(), file_name="Phieu_Hoc_Phi_Phim_Hong.zip", mime="application/zip")
