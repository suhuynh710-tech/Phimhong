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
st.write("Đã tối ưu hóa bộ render HTML bằng phương pháp an toàn tuyệt đối.")

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
            <div style="text-align: center; margin-top: 55px; font-size: 17px; color: #9a8a7a; font-style: italic;">
                {19} Trân trọng cảm ơn quý phụ huynh!
            </div>
        </div>
    </div>
    """
    bank_display = bank_raw if bank_raw != 'nan' else ''
    stk_display = stk if stk != 'nan' else ''
    return template.format(
        safe_name, logo_b64, thang_nam_hien_thi, lop, 
        svg_student, ten, svg_receipt, label_hoc_phi, hoc_phi_display_str, 
        svg_calendar, so_buoi, tong_tien_goc_html, phi_khac_html, 
        days_html, nhan_xet, tong_thanh_toan_str, qr_html, bank_display, stk_display, svg_thanks
    )

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
        #status {{ margin-top: 20px; font-size: 18px; font-weight: bold; color: #bc6c65; }}
        #hidden-receipts {{ position: absolute; left: -9999px; top: 0; pointer-events: none; }}
    </style>
</head>
<body>
    <button id="exportBtn" class="btn-export" onclick="startExport()">
        <svg viewBox="0 0 24 24" width="30" height="30" fill="white"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
        TẢI XUỐNG ZIP ẢNH PNG
    </button>
    <div id="status">Sẵn sàng xuất ảnh!</div>

    <div id="hidden-receipts">
        {receipts_html}
    </div>

    <script>
        async function startExport() {{
            const btn = document.getElementById('exportBtn');
            const status = document.getElementById('status');
            btn.disabled = true;
            btn.innerHTML = "⏳ Đang vẽ ảnh, vui lòng giữ tab này...";
            status.innerText = "Đang khởi động...";

            const zip = new JSZip();
            const receipts = document.querySelectorAll('.receipt-card');

            for(let i=0; i<receipts.length; i++) {{
                let name = receipts[i].getAttribute('data-name');
                status.innerText = "Đang vẽ ảnh: " + name.replace(/_/g, ' ') + " (" + (i+1) + "/" + receipts.length + ")";
                
                const canvas = await html2canvas(receipts[i], {{
                    scale: 2, 
                    useCORS: true,
                    backgroundColor: "#ffffff",
                    logging: false
                }});

                const imgData = canvas.toDataURL("image/png").split(',')[1];
                zip.file("Phieu_Hoc_Phi_" + name + ".png", imgData, {{base64: true}});
            }}

            status.innerText = "📦 Đang đóng gói file ZIP...";
            zip.generateAsync({{type:"blob"}}).then(function(content) {{
                saveAs(content, "Phieu_Hoc_Phi_PNG.zip");
                status.innerText = "✅ XUẤT ẢNH THÀNH CÔNG!";
                btn.innerHTML = "🚀 XUẤT LẠI ẢNH NẾU CẦN";
                btn.disabled = false;
            }});
        }}
    </script>
</body>
</html>
"""

# --- ICON SVG ---
svg_student = '<svg viewBox="0 0 24 24" width="24" height="24" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2.06-1.12V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>'
svg_receipt = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'
svg_calendar = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'
svg_extra = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M11 17h2v-4h4v-2h-4V7h-2v4H7v2h4v4zm1 5q-2.075 0-3.9-.788-1.825-.787-3.175-2.162-1.35-1.35-2.137-3.175Q2 14.05 2 12q0-2.075.788-3.9.787-1.825 2.137-3.175 1.35-1.35 3.175-2.137Q9.925 2 12 2q2.075 0 3.9.788 1.825.787 3.175 2.137 1.35 1.35 2.138 3.175Q22 9.925 22 12q0 2.075-.788 3.9-.787 1.825-2.138 3.175-1.35 1.35-3.175 2.162Q14.075 22 12 22Z"/></svg>'
svg_money = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10
