import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
import base64
import urllib.parse
import io
import os
import datetime

st.set_page_config(page_title="Phím Hồng Music - PNG Generator", layout="wide")

LOGO_PATH = "PHÍM HỒNG MUSIC (Nền trắng).jpg"

st.title("🎨 Cỗ Máy Xuất Ảnh PNG - Phím Hồng Music")
st.write("Bản Chốt Hạ: Upload Excel xong là xuất thẳng ra PNG ngay trên web. Có tính năng Phí Khác (+/-).")

uploaded_file = st.file_uploader("📂 Tải file Excel Danh_Sach_Hoc_Phi.xlsx", type=["xlsx"])

def get_base64_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
    return ""

# --- ICON SVG SIÊU NÉT ---
svg_student = '<svg viewBox="0 0 24 24" width="24" height="24" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2.06-1.12V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>'
svg_receipt = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'
svg_calendar = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'
svg_extra = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M11 17h2v-4h4v-2h-4V7h-2v4H7v2h4v4zm1 5q-2.075 0-3.9-.788-1.825-.787-3.175-2.162-1.35-1.35-2.137-3.175Q2 14.05 2 12q0-2.075.788-3.9.787-1.825 2.137-3.175 1.35-1.35 3.175-2.137Q9.925 2 12 2q2.075 0 3.9.788 1.825.787 3.175 2.137 1.35 1.35 2.138 3.175Q22 9.925 22 12q0 2.075-.788 3.9-.787 1.825-2.138 3.175-1.35 1.35-3.175 2.162Q14.075 22 12 22Z"/></svg>'
svg_thanks = '<svg viewBox="0 0 24 24" width="20" height="20" fill="#9a8a7a" style="vertical-align:middle; margin-right:8px;"><path d="M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z"/></svg>'

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=0).dropna(subset=['Họ và Tên'])
    logo_b64 = get_base64_logo()
    
    st.info(f"⏳ Đã nhận {len(df)} học sinh. Đang vẽ ảnh, chờ 3 giây...")
    progress_bar = st.progress(0)
    
    all_receipts_html = ""
    
    # Bắt cột Ngày đi học
    date_cols = []
    for col in df.columns:
        col_str = str(col).upper()
        if isinstance(col, datetime.datetime):
            col_str = col.strftime('%d/%m')
            df.rename(columns={col: col_str}, inplace=True)
            date_cols.append(col_str)
        elif '/' in col_str or col_str.startswith('T2') or col_str.startswith('T3') or col_str.startswith('T4') or col_str.startswith('T5') or col_str.startswith('T6') or col_str.startswith('T7') or col_str.startswith('CN'):
            date_cols.append(col)

    for index, row in df.iterrows():
        ten = str(row['Họ và Tên']).strip()
        safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
        progress_bar.progress((index + 1) / len(df))
        
        lop = str(row['Lớp']).strip() if pd.notna(row['Lớp']) else "Piano"
        hoc_phi = int(row['Học Phí (buổi)']) if pd.notna(row['Học Phí (buổi)']) else 0
        so_buoi = int(row['Tổng buổi học']) if pd.notna(row['Tổng buổi học']) else 0
        tong_tien_goc = int(row['Tổng học phí']) if pd.notna(row['Tổng học phí']) else (hoc_phi * so_buoi)
        
        # --- XỬ LÝ PHÍ KHÁC VÀ GHI CHÚ ---
        phi_khac = 0
        phi_khac_html = ""
        if 'Phí khác' in df.columns:
            val = row['Phí khác']
            if pd.notna(val) and str(val).strip() != '':
                try:
                    phi_khac = int(float(val))
                    if phi_khac != 0:
                        ghi_chu_text = ""
                        if 'Ghi chú' in df.columns:
                            raw_gc = row['Ghi chú']
                            if pd.notna(raw_gc) and str(raw_gc).strip() != "":
                                ghi_chu_text = f" <span style='font-size: 16px; font-style: italic; color: #a49688;'>({str(raw_gc).strip()})</span>"
                        
                        # Xử lý màu sắc (+) màu chuẩn, (-) màu đỏ tươi
                        d_color = "#bc6c65" if phi_khac > 0 else "#ff0000"
                        d_sign = "+" if phi_khac > 0 else "-"
                        
                        phi_khac_html = f'''
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_extra} Phí khác{ghi_chu_text}:</td>
                            <td style="padding: 15px 0; font-weight: 900; color: {d_color}; text-align: right; font-size: 24px; font-family: 'Times New Roman', serif;">{d_sign} {abs(phi_khac):,} đ</td>
                        </tr>
                        '''
                except: pass
        
        tong_thanh_toan = tong_tien_goc + phi_khac
        if tong_thanh_toan < 0: tong_thanh_toan = 0

        raw_nhan_xet = row['Nhận Xét Của GV'] if 'Nhận Xét Của GV' in df.columns else ""
        nhan_xet = str(raw_nhan_xet).strip() if (pd.notna(raw_nhan_xet) and str(raw_nhan_xet).lower() != 'nan') else ""

        bank = str(row['Ngân Hàng']).strip() if 'Ngân Hàng' in df.columns else ""
        stk = str(row['STK']).split('.')[0] if ('STK' in df.columns and pd.notna(row['STK'])) else ""

        # Ngày đi học
        days_html = '<div style="display: flex; flex-wrap: wrap; gap: 8px;">'
        has_day = False
        for col in date_cols:
            cell_val = str(row[col]).strip().upper()
            if cell_val == 'X':
                has_day = True
                col_name = str(col).strip()
                parts = col_name.split(' ')
                thu = parts[0] if len(parts) > 1 else ""
                day_month = parts[1] if len(parts) > 1 else col_name
                days_html += f'''
                <div style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; display:inline-block; text-align:center;">
                    <div style="font-size:10px; color:#8e7f72; margin-bottom:2px; line-height:1;">{thu}</div>
                    <div style="font-size:13px; font-weight:bold; color:#4a2e25; line-height:1;">{day_month}</div>
                </div>
                '''
        days_html += '</div>'
        if not has_day:
            days_html = '<div style="color:#aaa; font-style:italic; font-size:14px; padding: 5px 0;">Chưa có dữ liệu điểm danh</div>'

        # QR Code - Nhúng nội tuyến dạng Base64
        qr_html = ""
        if bank and stk and bank != 'nan':
            # QUAN TRỌNG: Chỉ lấy tên học sinh (như lệnh đã dặn)
            add_info = urllib.parse.quote(ten)
            qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_thanh_toan}&addInfo={add_info}"
            try:
                resp = requests.get(qr_url, timeout=5)
                if resp.status_code == 200:
                    qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode('utf-8')}"
                    qr_html = f'<img src="{qr_b64}" style="width: 125px; height: 125px; border-radius: 10px;">'
            except: pass
            
        if not qr_html:
            qr_html = '<div style="font-size:12px; color:#999; padding:40px 0; border:1px dashed #ccc; border-radius:8px; text-align:center;">CHƯA CÓ QR</div>'

        # --- GIAO DIỆN PHIẾU (RENDER NGẦM) ---
        receipt_html = f"""
        <div class="receipt-card" data-name="{safe_name}" style="width: 850px; background: white; font-family: Arial, sans-serif; margin: 0 auto 40px auto; border-radius: 20px; overflow: hidden; box-sizing: border-box; position: relative;">
            <div style="background: linear-gradient(135deg, #bc6c65 0%, #d49a71 100%); padding: 35px 50px; display: flex; align-items: center; justify-content: space-between; color: white;">
                <div style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid white; background-color: #fff; background-image: url('{logo_b64}'); background-size: cover; background-position: center; flex-shrink: 0;"></div>
                <div style="text-align: center; flex-grow: 1; padding: 0 20px;">
                    <div style="font-size: 16px; letter-spacing: 3px; font-weight: bold; opacity: 0.9; margin-bottom: 5px; text-transform: uppercase;">Lớp Nhạc Phím Hồng</div>
                    <h1 style="margin: 0; font-size: 44px; font-weight: 900; letter-spacing: 2px; font-family: 'Times New Roman', Times, serif; text-transform: uppercase; text-shadow: 1px 1px 4px rgba(0,0,0,0.2);">Phiếu Học Phí</h1>
                </div>
                <div style="text-align: center; min-width: 180px;">
                    <div style="font-size: 20px; font-weight: bold; margin-bottom: 8px;">Tháng 5 / 2026</div>
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
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_receipt} Học phí / buổi:</td>
                            <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{hoc_phi:,} đ</td>
                        </tr>
                        <tr style="border-top: 1px dashed #e2d5c4;">
                            <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_calendar} Số buổi học:</td>
                            <td style="padding: 15px 0; font-weight: bold; color: #2c1a16; text-align: right;">{so_buoi} buổi</td>
                        </tr>
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
                            <div style="margin-top: 15px; font-size: 18px; font-weight: 900; color: #bc6c65; text-transform: uppercase;">{bank if bank != 'nan' else ''}</div>
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
        all_receipts_html += receipt_html

    # --- NHÚNG MÃ TRỰC TIẾP VÀO STREAMLIT (CẢ CẢNH BÁO LỖI HTML) ---
    st.success(f"🎉 Hoàn tất đọc dữ liệu của {len(df)} học sinh!")
    
    component_html = f"""
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
            /* Ẩn các phiếu khỏi tầm nhìn nhưng không dùng display:none để canvas vẫn chụp được ảnh */
            #hidden-receipts {{ position: absolute; left: -9999px; top: 0; opacity: 0; pointer-events: none; }}
        </style>
    </head>
    <body>
        <button id="exportBtn" class="btn-export" onclick="startExport()">
            <svg viewBox="0 0 24 24" width="30" height="30" fill="white"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
            TẢI XUỐNG ZIP ẢNH PNG
        </button>
        <div id="status">Sẵn sàng xuất ảnh!</div>

        <div id="hidden-receipts">
            {all_receipts_html}
        </div>

        <script>
            async function startExport() {{
                const btn = document.getElementById('exportBtn');
                const status = document.getElementById('status');
                btn.disabled = true;
                btn.innerHTML = "⏳ Đang vẽ ảnh, đừng tắt trang nhé...";
                status.innerText = "Đang khởi động...";

                const zip = new JSZip();
                const receipts = document.querySelectorAll('.receipt-card');

                for(let i=0; i<receipts.length; i++) {{
                    let name = receipts[i].getAttribute('data-name');
                    status.innerText = "Đang xử lý ảnh: " + name + " (" + (i+1) + "/" + receipts.length + ")";
                    
                    const canvas = await html2canvas(receipts[i], {{
                        scale: 2, 
                        useCORS: true,
                        backgroundColor: "#ffffff",
                        logging: false
                    }});

                    const imgData = canvas.toDataURL("image/png").split(',')[1];
                    zip.file("Phieu_Hoc_Phi_" + name + ".png", imgData, {{base64: true}});
                }}

                status.innerText = "📦 Đang nén thành file ZIP...";
                zip.generateAsync({{type:"blob"}}).then(function(content) {{
                    saveAs(content, "Phieu_Hoc_Phi_PNG.zip");
                    status.innerText = "✅ XONG! BẠN KIỂM TRA THƯ MỤC DOWNLOAD NHÉ.";
                    btn.innerHTML = "🚀 XUẤT LẠI ẢNH NẾU CẦN";
                    btn.disabled = false;
                }});
            }}
        </script>
    </body>
    </html>
    """
    
    # Render giao diện này NGAY TRONG Streamlit (Cao 250px đủ chứa nút bấm)
    components.html(component_html, height=250, scrolling=False)
