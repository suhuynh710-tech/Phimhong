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
st.write("Đã khôi phục bản chuẩn ổn định và hiển thị QR code hoàn chỉnh.")

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

def get_receipt_html(safe_name, logo_b64, thang_nam_hien_thi, lop, ten, label_hoc_phi, hoc_phi_display, so_buoi, days_html, nhan_xet, tong_tien_goc_html, phi_khac_html, tong_thanh_toan, qr_html, bank_raw, stk, svg_student, svg_receipt, svg_calendar, svg_thanks):
    return f"""
    <div class="receipt-card" data-name="{safe_name}" style="width: 850px; background: white; font-family: Arial, sans-serif; margin: 0 auto 40px auto; border-radius: 20px; overflow: hidden; box-sizing: border-box; position: relative;">
        <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); padding: 35px 50px; display: flex; align-items: center; justify-content: space-between; color: white;">
            <div style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid white; background-color: #fff; overflow: hidden; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                <img src="{logo_b64}" style="width: 100%; height: 100%; object-fit: cover;" alt="Logo">
            </div>
            <div style="text-align: center; flex-grow: 1; padding: 0 20px;">
                <div style="font-size: 16px; letter-spacing: 3px; font-weight: bold; opacity: 0.9; margin-bottom: 5px; text-transform: uppercase;">Lớp Nhạc Phím Hồng</div>
                <h1 style="margin: 0; font-size: 44px; font-weight: 900; letter-spacing: 2px; font-family: 'Times New Roman', Times, serif; text-transform: uppercase; text-shadow: 1px 1px 4px rgba(0,0,0,0.2);">Phiếu Học Phí</h1>
            </div>
            <div style="text-align: center; min-width: 180px;">
                <div style="font-size: 20px; font-weight: bold; margin-bottom: 8px;">{thang_nam_hien_thi}</div>
                <div style="font-size: 28px; font-weight: 900; background: rgba(255,255,255,0.25); padding: 10px 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.5); font-family: 'Times New Roman', Times, serif; text-transform: capitalize;">Lớp {lop}</div>
            </div>
        </div>
        <div style="padding: 40px 60px;">
            <div style="background: #fdfaf6; border: 1px solid #f2e2b3; border-radius: 15px; padding: 30px 45px; margin: 0 auto 35px auto; width: 85%;">
                <table style="width: 100%; border-collapse: collapse; font-size: 20px;">
                    <tr>
                        <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_student} Học sinh:</td>
                        <td style="padding: 15px 0; font-weight: 900; color: #2c1a16; text-align: right; font-size: 30px; font-family: 'Times New Roman', Times, serif;">{ten}</td>
                    </tr>
                    <tr>
                        <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center; border-top: 1px dashed #e2d5c4;">{svg_receipt} {label_hoc_phi}:</td>
                        <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right; border-top: 1px dashed #e2d5c4;">{hoc_phi_display:,} đ</td>
                    </tr>
                    <tr style="border-top: 1px dashed #e2d5c4;">
                        <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_calendar} Số buổi học:</td>
                        <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{so_buoi} buổi</td>
                    </tr>
                    {tong_tien_goc_html}
                    {phi_khac_html}
                </table>
            </div>
            <div style="display: flex; gap: 45px; align-items: flex-start; justify-content: center; width: 95%; margin: 0 auto;">
                <div style="flex: 1.3;">
                    <div style="margin-bottom: 25px;">
                        <div style="font-size: 14px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 12px;">NGÀY ĐI HỌC</div>
                        <div style="width: 100%;">{days_html}</div>
                    </div>
                    <div>
                        <div style="font-size: 14px; color: #8e7f72; font-weight: bold; letter-spacing: 1px; margin-bottom: 12px;">NHẬN XÉT CỦA GIÁO VIÊN</div>
                        <div style="background: #fffdf5; border: 1px solid #f2e2b3; border-radius: 12px; padding: 20px; color: #5a4b41; font-style: italic; line-height: 1.6; font-size: 16px; min-height: 20px;">
                            {nhan_xet}
                        </div>
                    </div>
                </div>
                <div style="flex: 0.7; display: flex; flex-direction: column; gap: 20px;">
                    <div style="background: #fdf6ec; border: 2px solid #ecdac8; border-radius: 15px; padding: 20px; text-align: center;">
                        <div style="font-size: 13px; color: #8e7f72; font-weight: bold;">TỔNG THANH TOÁN</div>
                        <div style="font-size: 38px; color: #4a2e25; font-weight: 900; margin-top: 10px; font-family: 'Times New Roman', serif;">{tong_thanh_toan:,} đ</div>
                    </div>
                    <div style="background: white; border: 2px dashed #d49a71; border-radius: 15px; padding: 20px; text-align: center;">
                        <div style="font-size: 11px; color: #d49a71; font-weight: bold; margin-bottom: 15px;">QUÉT MÃ THANH TOÁN</div>
                        <div style="display: flex; justify-content: center;">{qr_html}</div>
                        <div style="margin-top: 15px; font-size: 18px; font-weight: 900; color: #bc6c65; text-transform: uppercase;">{bank_raw if bank_raw != 'nan' else ''}</div>
                        <div style="font-size: 16px; font-weight: bold; color: #4a2e25; margin-top: 5px;">{stk if stk != 'nan' else ''}</div>
                    </div>
                </div>
            </div>
            <div style="text-align: center; margin-top: 55px; font-size: 17px; color: #9a8a7a; font-style: italic;">
                {svg_thanks} Trân trọng cảm ơn quý phụ huynh!
            </div>
        </div>
    </div>
    """

def get_component_html(receipts_html):
    return f"""
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <style>
        body {{ font-family: Arial, sans-serif; text-align: center; margin: 0; padding: 20px; background: transparent; }}
        .btn-export {{ background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); color: white; border: none; padding: 25px 40px; font-size: 22px; border-radius: 15px; cursor: pointer; font-weight: bold; box-shadow: 0 8px 20px rgba(188, 108, 101, 0.4); transition: 0.3s; width: 100%; max-width: 600px; display: inline-flex; align-items: center; justify-content: center; gap: 15px; }}
        .btn-export:hover {{ transform: scale(1.03); box-shadow: 0 10px 25px rgba(188, 108, 101, 0.6); }}
        .btn-export:disabled {{ background: #ccc; cursor: not-allowed; transform: none; box-shadow: none; color: #666; }}
        #status {{ margin-top: 20px
