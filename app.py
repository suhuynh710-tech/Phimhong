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

st.title("🎨 Cỗ Máy Xuất Ảnh PNG - Bản Phân Loại Học Phí")
st.write("Đã vá lỗi thanh tiến trình vượt quá 100% do file Excel có dòng trống.")

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

# --- ICON SVG SIÊU NÉT ---
svg_student = '<svg viewBox="0 0 24 24" width="24" height="24" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2.06-1.12V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>'
svg_receipt = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M18 17H6v-2h12v2zm0-4H6v-2h12v2zm0-4H6V7h12v2zM3 22l1.5-1.5L6 22l1.5-1.5L9 22l1.5-1.5L12 22l1.5-1.5L15 22l1.5-1.5L18 22l1.5-1.5L21 22V2l-1.5 1.5L18 2l-1.5 1.5L15 2l-1.5 1.5L12 2l-1.5 1.5L9 2l-1.5 1.5L6 2 4.5 3.5 3 2v20z"/></svg>'
svg_calendar = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M19 4h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20a2 2 0 002 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V10h14v10zm0-12H5V6h14v2z"/></svg>'
svg_extra = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M11 17h2v-4h4v-2h-4V7h-2v4H7v2h4v4zm1 5q-2.075 0-3.9-.788-1.825-.787-3.175-2.162-1.35-1.35-2.137-3.175Q2 14.05 2 12q0-2.075.788-3.9.787-1.825 2.137-3.175 1.35-1.35 3.175-2.137Q9.925 2 12 2q2.075 0 3.9.788 1.825.787 3.175 2.137 1.35 1.35 2.138 3.175Q22 9.925 22 12q0 2.075-.788 3.9-.787 1.825-2.138 3.175-1.35 1.35-3.175 2.162Q14.075 22 12 22Z"/></svg>'
svg_money = '<svg viewBox="0 0 24 24" width="22" height="22" fill="#6d5b4b" style="margin-right:12px;"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-1h-1c-.55 0-1-.45-1-1v-3c0-.55.45-1 1-1h3v-1h-3V9h2V8h2v1h1c.55 0 1 .45 1 1v3c0 .55-.45 1-1 1h-3v1h3v1h-2v1z"/></svg>'
svg_thanks = '<svg viewBox="0 0 24 24" width="20" height="20" fill="#9a8a7a" style="vertical-align:middle; margin-right:8px;"><path d="M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z"/></svg>'

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names
    
    student_sheet = next((s for s in sheet_names if 'học sinh' in s.lower()), sheet_names[0])
    attendance_sheet = next((s for s in sheet_names if 'sheet1' in s.lower() or 'điểm danh' in s.lower()), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
    
    df_students = pd.read_excel(uploaded_file, sheet_name=student_sheet)
    df_attendance = pd.read_excel(uploaded_file, sheet_name=attendance_sheet)
    
    logo_b64 = get_base64_logo()
    
    df_students.columns = [str(c).strip() for c in df_students.columns]
    df_attendance.columns = [str(c).strip() for c in df_attendance.columns]
    
    name_col = next((col for col in df_students.columns if 'họ và tên' in col.lower()), None)
    
    if not name_col:
        st.error("❌ Không tìm thấy cột 'Họ và Tên' trong trang danh sách!")
    else:
        df_students = df_students.dropna(subset=[name_col])
        st.success(f"🎉 Kết nối thành công dữ liệu!")
        progress_bar = st.progress(0)
        
        all_receipts_html = ""
        
        lop_col = next((col for col in df_students.columns if col.lower() == 'lớp'), None)
        hp_buoi_col = next((col for col in df_students.columns if 'học phí (buổi)' in col.lower()), None)
        hp_thang_col = next((col for col in df_students.columns if 'học phí (tháng)' in col.lower()), None)
        phi_khac_col = next((col for col in df_students.columns if 'phí khác' in col.lower()), None)
        ghi_chu_col = next((col for col in df_students.columns if 'ghi chú' in col.lower()), None)
        nhan_xet_col = next((col for col in df_students.columns if 'nhận xét của gv' in col.lower()), None)
        ngan_hang_col = next((col for col in df_students.columns if 'ngân hàng' in col.lower()), None)
        stk_col = next((col for col in df_students.columns if 'stk' in col.lower()), None)

        date_cols = []
        for col in df_attendance.columns:
            col_str = str(col).upper()
            if isinstance(col, datetime.datetime):
                col_str = col.strftime('%d/%m')
                df_attendance.rename(columns={col: col_str}, inplace=True)
                date_cols.append(col_str)
            elif '/' in col_str or col_str.startswith('T2') or col_str.startswith('T3') or col_str.startswith('T4') or col_str.startswith('T5') or col_str.startswith('T6') or col_str.startswith('T7') or col_str.startswith('CN'):
                date_cols.append(col)

        # Sử dụng enumerate thay vì index để thanh tiến trình chạy chuẩn xác không vượt 100%
        for i, (index, row) in enumerate(df_students.iterrows()):
            ten = str(row[name_col]).strip()
            safe_name = ten.replace(' ', '_').replace('(', '').replace(')', '')
            
            # Tính toán tiến trình theo số đếm thực tế
            progress_val = min((i + 1) / len(df_students), 1.0)
            progress_bar.progress(progress_val)
            
            lop = str(row[lop_col]).strip() if lop_col and pd.notna(row[lop_col]) else "Piano"
            
            so_buoi = 0
            days_html = '<div style="display: flex; flex-wrap: wrap; gap: 8px;">'
            has_day = False
            
            att_col = next((col for col in df_attendance.columns if ten.lower() in str(col).lower()), None)
            if att_col:
                present_rows = df_attendance[df_attendance[att_col].astype(str).str.strip().str.upper() == 'X']
                so_buoi = len(present_rows)
                
                thu_col = next((c for c in df_attendance.columns if 'THỨ' in c.upper()), None)
                ngay_col = next((c for c in df_attendance.columns if 'NGÀY' in c.upper()), None)
                
                for _, att_row in present_rows.iterrows():
                    has_day = True
                    thu = str(att_row[thu_col]).strip() if thu_col and pd.notna(att_row[thu_col]) else ""
                    ngay_val = att_row[ngay_col]
                    
                    if isinstance(ngay_val, (datetime.datetime, datetime.date)):
                        day_month = ngay_val.strftime('%d/%m')
                    else:
                        s_date = str(ngay_val).strip().split(' ')[0]
                        if '-' in s_date:
                            try:
                                p = s_date.split('-')
                                day_month = f"{int(p[2]):02d}/{int(p[1]):02d}"
                            except: day_month = s_date
                        elif '/' in s_date:
                            try:
                                p = s_date.split('/')
                                day_month = f"{int(p[0]):02d}/{int(p[1]):02d}"
                            except: day_month = s_date
                        else:
                            day_month = s_date
                    
                    days_html += f'''
                    <div style="background:#f7f1e9; border:1px solid #e0d1c1; border-radius:8px; padding:6px 12px; display:inline-block; text-align:center;">
                        <div style="font-size:10px; color:#8e7f72; margin-bottom:2px; line-height:1;">{thu}</div>
                        <div style="font-size:13px; font-weight:bold; color:#4a2e25; line-height:1;">{day_month}</div>
                    </div>
                    '''
            days_html += '</div>'
            if not has_day:
                days_html = '<div style="color:#aaa; font-style:italic; font-size:14px; padding: 5px 0;">Chưa có dữ liệu điểm danh</div>'
            
            hoc_phi_buoi = clean_number(row[hp_buoi_col]) if hp_buoi_col else 0
            hoc_phi_thang = clean_number(row[hp_thang_col]) if hp_thang_col else 0
            
            if hoc_phi_thang > 0:
                label_hoc_phi = "Học phí (tháng)"
                hoc_phi_display = hoc_phi_thang
                tong_tien_goc = hoc_phi_thang
            else:
                label_hoc_phi = "Học phí (buổi)"
                hoc_phi_display = hoc_phi_buoi
                tong_tien_goc = hoc_phi_buoi * so_buoi
            
            phi_khac = 0
            phi_khac_html = ""
            tong_tien_goc_html = ""
            
            if phi_khac_col:
                val = row[phi_khac_col]
                if pd.notna(val) and str(val).strip() != '':
                    try:
                        phi_khac = clean_number(val)
                        if phi_khac != 0:
                            ghi_chu_text = ""
                            raw_gc = str(row[ghi_chu_col]).strip() if ghi_chu_col and pd.notna(row[ghi_chu_col]) else ""
                            
                            if raw_gc and raw_gc.lower() != 'nan':
                                ghi_chu_text = f" <span style='font-size: 16px; font-style: italic; color: #a49688;'>({raw_gc})</span>"
                                if phi_khac > 0 and any(kw in raw_gc.lower() for kw in ['trừ', 'nghỉ', 'giảm']):
                                    phi_khac = -phi_khac

                            d_color = "#bc6c65" if phi_khac > 0 else "#ff0000"
                            d_sign = "+" if phi_khac > 0 else "-"
                            
                            tong_tien_goc_html = f'''
                            <tr style="border-top: 1px dashed #e2d5c4; background: rgba(247,241,233,0.5);">
                                <td style="padding: 12px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_money} Thành tiền (Tiền học):</td>
                                <td style="padding: 12px 0; font-weight: bold; color: #4a2e25; text-align: right;">{tong_tien_goc:,} đ</td>
                            </tr>
                            '''

                            phi_khac_html = f'''
                            <tr style="border-top: 1px dashed #e2d5c4;">
                                <td style="padding: 15px 0; color: #7a6b5d; display:flex; align-items:center;">{svg_extra} Phí khác{ghi_chu_text}:</td>
                                <td style="padding: 15px 0; font-weight: 900; color: {d_color}; text-align: right; font-size: 24px; font-family: 'Times New Roman', serif;">{d_sign} {abs(phi_khac):,} đ</td>
                            </tr>
                            '''
                    except: pass
            
            tong_thanh_toan = tong_tien_goc + phi_khac
            if tong_thanh_toan < 0: tong_thanh_toan = 0

            raw_nhan_xet = row[nhan_xet_col] if nhan_xet_col else ""
            nhan_xet = str(raw_nhan_xet).strip() if (pd.notna(raw_nhan_xet) and str(raw_nhan_xet).lower() != 'nan') else ""

            bank = str(row[ngan_hang_col]).strip() if ngan_hang_col else ""
            stk = str(row[stk_col]).split('.')[0] if (stk_col and pd.notna(row[stk_col])) else ""

            qr_html = ""
            if bank and stk and bank != 'nan':
                add_info = urllib.parse.quote(ten)
                qr_url = f"https://img.vietqr.io/image/{bank}-{stk}-compact2.png?amount={tong_thanh_toan}&addInfo={add_info}"
                try:
                    resp = requests.get(qr_url, timeout=5)
                    if resp.status_code == 200:
                        qr_b64 = f"data:image/png;base64,{base64.b64encode(resp.content).decode('utf-8')}"
                        qr_html = f'<img src="{qr_b64}" style="width: 125px; height: 125px; border-radius: 10px;">'
                    else:
                        qr_html = f'<img src="{qr_url}" style="width: 125px; height: 125px; border-radius: 10px;">'
                except:
                    qr_html = f'<img src="{qr_url}" style="width: 125px; height: 125px; border-radius: 10px;">'
                
            if not qr_html:
                qr_html = '<div style="font-size:12px; color:#999; padding:40px 0; border:1px dashed #ccc; border-radius:8px; text-align:center;">CHƯA CÓ QR</div>'

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
                    btn.innerHTML = "⏳ Đang vẽ ảnh, vui lòng giữ tab này...";
                    status.innerText = "Đang khởi động...";

                    const zip = new JSZip();
                    const receipts = document.querySelectorAll('.receipt-card');

                    for(let i=0; i<receipts.length; i++) {{
                        let name = receipts[i].getAttribute('data-name');
                        status.innerText = "Đang vẽ ảnh: " + name.replace('_', ' ') + " (" + (i+1) + "/" + receipts.length + ")";
                        
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
        components.html(component_html, height=250, scrolling=False)
