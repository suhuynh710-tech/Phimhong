import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import base64
import urllib.parse
import os
import datetime

st.set_page_config(page_title="Phím Hồng Music - PNG Generator", layout="wide")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎨 Cỗ Máy Xuất Ảnh PNG - Bản Phân Loại Học Phí")
st.write("Đã sửa hoàn toàn lỗi cú pháp chuỗi và hiển thị mã QR thanh toán chuẩn xác.")

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

def get_bank_code(bank_name):
    if not bank_name or pd.isna(bank_name):
        return "VCB"
    b = str(bank_name).strip().lower()
    if 'vietcombank' in b or 'vcb' in b:
        return "VCB"
    elif 'techcombank' in b or 'tcb' in b:
        return "TCB"
    elif 'mbbank' in b or 'mb' in b:
        return "MB"
    elif 'acb' in b:
        return "ACB"
    elif 'agribank' in b:
        return "AGRIBANK"
    elif 'bidv' in b:
        return "BIDV"
    elif 'vietinbank' in b:
        return "VIETINBANK"
    elif 'tpbank' in b or 'tpb' in b:
        return "TPB"
    elif 'sacombank' in b:
        return "SACOMBANK"
    elif 'vpbank' in b:
        return "VPB"
    return str(bank_name).strip().upper()

def get_receipt_html(safe_name, logo_b64, thang_nam_hien_thi, lop, ten, label_hoc_phi, hoc_phi_display_str, so_buoi, days_html, nhan_xet, tong_tien_goc_html, phi_khac_html, tong_thanh_toan_str, qr_html, bank_raw, stk, svg_student, svg_receipt, svg_calendar, svg_thanks):
    template = """
    <div class="receipt-card" data-name="{0}" style="width: 850px; background: white; font-family: Arial, sans-serif; margin: 0 auto 40px auto; border-radius: 20px; overflow: hidden; box-sizing: border-box; position: relative;">
        <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); padding: 35px 50px; display: flex; align-items: center; justify-content: space-between; color: white;">
            <div style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid white; background-color: #fff; overflow: hidden; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                <img src="{1}" style="width: 100%; height: 100%; object-fit: cover;" alt="Logo">
            </div>
            <div style="text-align: center; flex-grow: 1; padding: 0 20px;">
                <div style="font-size: 16px; letter-spacing: 3px; font-weight: bold; opacity: 0.9; margin-bottom: 5px; text-transform: uppercase;">Lớp Nhạc Phím Hồng</div>
                <h1 style="margin: 0; font-size: 44px; font-weight: 900; letter-spacing: 2px; font-family: 'Times New Roman', Times, serif; text-transform: uppercase; text-shadow: 1px 1px 4px rgba(0,0,0,0.2);">Phiếu Học Phí</h1>
            </div>
            <div style="text-align: center; min-width: 180px;">
                <div style="font-size: 20px; font-weight: bold; margin-bottom: 8px;">{2}</div>
                <div style="font-size: 28px; font-weight: 900; background: rgba(255,255,255,0.25); padding: 10px 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.5); font-family: 'Times New Roman', Times, serif; text-transform: capitalize;">Lớp {3}</div>
            </div>
        </div>
        <div style="padding: 40px 60px;">
            <div style="background: #fdfaf6; border: 1px solid #f2e2b3; border-radius: 15px; padding: 30px 45px; margin: 0 auto 35px auto; width: 85%;">
                <table style="width: 100%; border-collapse: collapse; font-size: 20px;">
                    <tr>
                        <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{4} Học sinh:</td>
                        <td style="padding: 15px 0; font-weight: 900; color: #2c1a16; text-align: right; font-size: 30px; font-family: 'Times New Roman', Times, serif;">{5}</td>
                    </tr>
                    <tr>
                        <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center; border-top: 1px dashed #e2d5c4;">{6} {7}:</td>
                        <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right; border-top: 1px dashed #e2d5c4;">{8} đ</td>
                    </tr>
                    <tr style="border-top: 1px dashed #e2d5c4;">
                        <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{9} Số buổi học:</td>
                        <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{10} buổi</td>
                    </tr>
                    {11}
                    {12}
                </table>
            </div>
            <div style="display: flex; gap: 45px; align-items: flex-start; justify-content: center; width: 95%; margin: 0 auto;">
                <div style="flex: 1.3;">
                    <div style="margin-bottom: 25px;">
                        <div style="font-size: 14px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 12px;">NGÀY ĐI HỌC</div>
                        <div style="width: 100%;">{13}</div>
                    </div>
                    <div>
                        <div style="font-size: 14px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 12px;">NHẬN XÉT CỦA GIÁO VIÊN</div>
                        <div style="background: #fffdf5; border: 1px solid #f2e2b3; border-radius: 12px; padding: 20px; color: #5a4b41; font-style: italic; line-height: 1.6; font-size: 16px; min-height: 20px;">
                            {14}
                        </div>
                    </div>
                </div>
                <div style="flex: 0.7; display: flex; flex-direction: column; gap: 20px;">
                    <div style="background: #fdf6ec; border: 2px solid #ecdac8; border-radius: 15px; padding: 20px; text-align: center;">
                        <div style="font-size: 13px; color: #8e7f72; font-weight: bold;">TỔNG THANH TOÁN</div>
                        <div style="font-size: 38px; color: #4a2e25; font-weight: 900; margin-top: 10px; font-family: 'Times New Roman', serif;">{15} đ</div>
                    </div>
                    <div style="background: white; border: 2px dashed #d49a71; border-radius: 15px; padding: 20px; text-align: center;">
                        <div style="font-size: 11px; color: #d49a71; font-weight: bold; margin-bottom: 15px;">QUÉT MÃ THANH TOÁN</div>
                        <div style="display: flex; justify-content: center;">{16}</div>
                        <div style="margin-top: 15px; font-size: 18px; font-weight: 900; color: #bc6c65; text-transform: uppercase;">{17}</div>
                        <div style="font-size: 16px; font-weight: bold; color: #4a2e25; margin-top: 5px;">{18}</div>
                    </div>
                </div>
            </div>
            <div
