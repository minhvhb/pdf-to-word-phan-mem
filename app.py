import streamlit as st

# ==========================================
# CẤU HÌNH TRANG CHỦ ĐẠO (BẮT BUỘC PHẢI Ở DÒNG ĐẦU TIÊN)
# ==========================================
st.set_page_config(page_title="Hệ thống Công cụ Văn phòng tại NMTLBT Craven A", page_icon="⚙️", layout="centered", initial_sidebar_state="expanded")

import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
import os
import re
import time
import zipfile
import pandas as pd
from io import BytesIO
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from PIL import Image
from pypdf import PdfReader, PdfWriter, Transformation
import difflib
import fitz

# Ẩn nút gập thanh bên mặc định của Streamlit
st.markdown(
    """
    <style>
        [data-testid="collapsedControl"] { display: none; }
        [data-testid="stSidebarCollapseButton"] { display: none; }
    </style>
    """,
    unsafe_allow_html=True
)

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

def clear_file():
    st.session_state.uploader_key += 1

# ==========================================
# 0. HÀM GỌI AI THÔNG MINH (TỰ ĐỘNG CHUYỂN MODEL KHI QUÁ TẢI)
# ==========================================
def call_gemini_smart(client, contents_payload):
    """
    Tự động thử lần lượt các model.
    Bổ sung cơ chế chống ngắt kết nối do vượt giới hạn tốc độ (Rate Limit).
    """
    models_to_try = [
        "gemini-1.5-flash", # Hạn mức cao nhất: 1500 request/ngày
        "gemini-1.5-pro"    # Hạn mức: 50 request/ngày
    ]
    last_err = None
    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=contents_payload
            )
            return response
        except Exception as e:
            err_msg = str(e).lower()
            last_err = e
            # Chạm trần Rate Limit (429): Nghỉ 5 giây để máy chủ hồi lại rồi chạy tiếp
            if "429" in err_msg or "quota" in err_msg or "exhausted" in err_msg:
                time.sleep(5)
                continue
            # Máy chủ quá tải (503): Nghỉ 2 giây
            if "503" in err_msg or "unavailable" in err_msg or "overloaded" in err_msg:
                time.sleep(2)
                continue
            # Tên model không tồn tại: Bỏ qua để thử model dưới
            if "404" in err_msg or "not found" in err_msg:
                continue
            raise e
    raise last_err

# ==========================================
# 1. ĐỌC DỮ LIỆU TÀI KHOẢN VÀ ĐĂNG NHẬP
# ==========================================
with open('config.yaml', 'r', encoding='utf-8') as file:
    config = yaml.load(file, Loader=SafeLoader)

stauth.Hasher.hash_passwords(config['credentials'])

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days']
)

# ==========================================
# 2. GIAO DIỆN TRANG ĐĂNG NHẬP
# ==========================================
if st.session_state.get("authentication_status") != True:
    st.markdown("""
        <style>
            .stApp {
                background-image: url("https://images.unsplash.com/photo-1477414348463-c0eb7f1359b6?q=80&w=2070&auto=format&fit=crop");
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }
            [data-testid="stForm"] {
                background-color: rgba(255, 255, 255, 0.95) !important;
                padding: 40px !important;
                border-radius: 15px !important;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3) !important;
                width: 700px !important; 
                max-width: 95vw !important;
                margin: 10vh auto 0 auto !important;
                border: 1px solid rgba(255, 255, 255, 0.5) !important;
            }
            [data-testid="stForm"] h2 {
                text-align: center !important;
                color: #003366 !important;
                font-weight: 900 !important;
                font-size: 24px !important; 
                white-space: nowrap !important;
                margin-bottom: 25px !important;
                padding-bottom: 0px !important;
                width: 100% !important;
            }
            [data-testid="stForm"] [data-testid="stVerticalBlock"] > div {
                width: 100% !important;
                max-width: 100% !important;
            }
            [data-testid="stForm"] .stTextInput {
                width: 100% !important;
                min-width: 100% !important;
            }
            [data-testid="stForm"] div[data-baseweb="input"] {
                background-color: #ffffff !important;
                border: 2px solid #000000 !important;
                border-radius: 6px !important;
                width: 100% !important;
            }
            [data-testid="stForm"] div[data-baseweb="input"]:focus-within {
                border-color: #003366 !important; 
            }
            [data-testid="stForm"] input {
                background-color: transparent !important;
                color: #000 !important;
                font-size: 15px !important;
            }
            [data-testid="stForm"] .stTextInput label p {
                font-size: 14px !important;
                font-weight: bold !important;
                color: #333 !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] {
                width: 100% !important;
                max-width: 100% !important;
                display: block !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] button {
                width: 100% !important;
                min-width: 100% !important;
                background-color: #003366 !important;
                color: white !important;
                font-size: 16px !important;
                letter-spacing: 1px !important;
                border-radius: 8px !important;
                padding: 12px 0 !important;
                border: none !important;
                margin-top: 15px !important;
                display: block !important;
                transition: all 0.3s ease-in-out !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] button p {
                font-weight: 900 !important;
                margin: 0 !important;
            }
            [data-testid="stForm"] div[data-testid="stFormSubmitButton"] button:hover {
                background-color: #001f3f !important;
                box-shadow: 0px 4px 15px rgba(0,0,0,0.4) !important;
            }
            div[data-testid="stVerticalBlock"] > div:has(div[data-testid="stMarkdownContainer"]) {
                display: flex;
                justify-content: center;
                margin-top: 15px;
            }
        </style>
    """, unsafe_allow_html=True)

authenticator.login(
    location='main',
    fields={
        'Form name': 'CÔNG CỤ VĂN PHÒNG NMTLBT CRAVEN A',
        'Username': 'Tên đăng nhập',
        'Password': 'Mật khẩu',
        'Login': 'ĐĂNG NHẬP'
    }
)

if st.session_state["authentication_status"] == False:
    st.error('Tên đăng nhập hoặc mật khẩu không đúng!')
    st.stop()
elif st.session_state["authentication_status"] == None:
    st.warning('Vui lòng đăng nhập để sử dụng công cụ')
    st.stop()

# ==========================================
# 3. GIAO DIỆN THANH BÊN (SIDEBAR) TỐI ƯU
# ==========================================
st.markdown("""
    <style>
        [data-testid="stSidebarUserContent"] { 
            padding-top: 1rem !important; 
            padding-bottom: 0rem !important; 
            padding-left: 1rem !important; 
            padding-right: 1rem !important;
        }
        [data-testid="stSidebarUserContent"] div[role="radiogroup"] > label { 
            margin-bottom: -6px !important; 
        }
        [data-testid="stSidebarUserContent"] hr { 
            margin-top: 5px; 
            margin-bottom: 5px; 
        }
        .custom-alert {
            background-color: #f8d7da;
            color: #842029;
            padding: 10px;
            border-radius: 5px;
            font-size: 13.5px;
            line-height: 1.4;
            border: 1px solid #f5c2c7;
        }
    </style>
""", unsafe_allow_html=True)

col1, col2, col3 = st.sidebar.columns([1, 2, 1])
with col2:
    authenticator.logout('Đăng xuất', 'main')

st.sidebar.markdown(f"<h5 style='text-align: center; margin-top: 5px; margin-bottom: 0px;'>Chào mừng {st.session_state['name']}!</h5>", unsafe_allow_html=True)
st.sidebar.markdown("---")

st.sidebar.markdown("### 📌 Menu Công Cụ")
app_mode = st.sidebar.radio(
    "Vui lòng chọn ứng dụng:",
    [
        "📄 1. PDF sang Word", 
        "🖨️ 2. Chuyển PDF về khổ A4", 
        "📊 3. PDF/Ảnh sang Excel", 
        "🔍 4. So sánh Văn bản / Hợp đồng",
        "✂️ 5. Cắt & Ghép PDF",
        "💻 6. Chuyên gia Công thức & VBA",
        "🗜️ 7. Nén & Đóng Gói ZIP Gửi Mail",
        "🔄 8. Xoay Trang PDF",
        "🛡️ 9. Xóa Siêu Dữ Liệu Bảo Mật",
        "📑 10. PDF Scan / Ảnh sang Word"
    ],
    label_visibility="collapsed"
)
st.sidebar.markdown("---") 

st.sidebar.markdown("""
    <div class="custom-alert">
        <b>⚠️ NGUYÊN TẮC SỬ DỤNG:</b><br>
        <b>- Bảo mật:</b> KHÔNG tải lên tài liệu MẬT, TỐI MẬT, dữ liệu tài chính chưa công khai.<br>
        <b>- Tối ưu:</b> Chỉ tải file PDF <b>dưới 30 trang/lần</b>.
    </div>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown("""
    <div style="text-align: center; font-size: 12px; color: #666; margin-top: 10px; line-height: 1.5;">
        <b>BỘ CÔNG CỤ VĂN PHÒNG</b><br>
        Phiên bản: 1.0.0<br>
        Phát triển: Vũ Huỳnh Bình Minh<br>
        Nhân viên nghiệp vụ - Phòng Nghiệp vụ<br>
        © 2026 NMTLBT Craven "A"
    </div>
""", unsafe_allow_html=True)


# ==========================================
# 4. CÁC HÀM XỬ LÝ CHỨC NĂNG CỦA ỨNG DỤNG
# ==========================================
def parse_and_add_runs(paragraph, text):
    parts_bold = re.split(r'\*\*(.*?)\*\*', text)
    for i, p_bold in enumerate(parts_bold):
        is_bold = (i % 2 == 1)
        parts_italic = re.split(r'\*(.*?)\*', p_bold)
        for j, p_italic in enumerate(parts_italic):
            is_italic = (j % 2 == 1)
            parts_underline = re.split(r'<u>(.*?)</u>', p_italic)
            for k, p_underline in enumerate(parts_underline):
                is_underline = (k % 2 == 1)
                if p_underline:
                    run = paragraph.add_run(p_underline)
                    run.bold = is_bold
                    run.italic = is_italic
                    run.underline = is_underline

def clean_tags_and_align(text, paragraph, default_align):
    align = default_align
    text_clean = text.strip()
    
    if '[CENTER]' in text_clean:
        align = WD_ALIGN_PARAGRAPH.CENTER
        text_clean = text_clean.replace('[CENTER]', '').strip()
    if '[RIGHT]' in text_clean:
        align = WD_ALIGN_PARAGRAPH.RIGHT
        text_clean = text_clean.replace('[RIGHT]', '').strip()
    if '[LEFT]' in text_clean:
        align = WD_ALIGN_PARAGRAPH.LEFT
        text_clean = text_clean.replace('[LEFT]', '').strip()
        
    paragraph.alignment = align
    return text_clean

def app_pdf_to_word():
    import random
    
    try:
        # Lấy danh sách 4 key từ Streamlit Secrets
        api_keys = st.secrets["GEMINI_API_KEYS"]
        # Chọn ngẫu nhiên 1 key cho mỗi lần xử lý để né hạn mức của Google
        api_key_input = random.choice(api_keys)
    except KeyError:
        st.error("⚠️ Hệ thống chưa tìm thấy cấu hình GEMINI_API_KEYS. Vui lòng liên hệ Quản trị viên!")
        st.stop()

    st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
    
    uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"], key=f"app1_{st.session_state.uploader_key}")

    if uploaded_file is not None:
        st.success(f"Đã tải lên file: **{uploaded_file.name}**")
        
        if st.button("🚀 Bắt đầu Chuyển đổi", type="primary"):
            with st.spinner("🤖 Đang phân tích lề và vẽ lại bảng biểu phức tạp..."):
                try:
                    temp_input_path = f"temp_{uploaded_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    NHIỆM VỤ OCR - BẮT BUỘC TUÂN THỦ NGHIÊM NGẶT CÁC QUY TẮC SAU:

                    1. THỂ THỨC VĂN BẢN (QUỐC HIỆU & CHỮ KÝ):
                       - Phần Quốc hiệu trên cùng (hoặc các khối chữ ký ngang nhau): BẮT BUỘC dùng [HEADER_TABLE] (bảng tàng hình 2 cột) để chia tỷ lệ.
                       - Không dùng bảng cho Footer.

                    2. CANH LỀ ĐOẠN VĂN:
                       - Ghi [CENTER] ở đầu dòng cần canh giữa.
                       - Ghi [RIGHT] ở đầu dòng lệch phải.

                    3. ĐỊNH DẠNG CHỮ: 
                       - In đậm -> bọc trong ** (VD: **THÔNG BÁO**). 
                       - In nghiêng -> bọc trong * (VD: *Nơi nhận:*).
                       - Gạch chân -> bọc trong <u> và </u>.

                    4. BẢNG BIỂU PHỨC TẠP (QUAN TRỌNG NHẤT):
                       - Bắt buộc vẽ bằng Markdown (|...|).
                       - TẤT CẢ các dòng trong cùng 1 bảng PHẢI CÓ SỐ LƯỢNG CỘT BẰNG NHAU (cùng số lượng ký tự `|`).
                       - XỬ LÝ Ô GỘP (MERGED CELLS) TRONG BẢN GỐC:
                         + Gộp ngang (kéo dài sang phải): Ghi chữ vào ô trái cùng, các ô bị gộp phía sau điền CHÍNH XÁC chữ: `[MERGE_LEFT]`
                         + Gộp dọc (kéo dài xuống dưới): Ghi chữ vào ô trên cùng, các ô bị gộp phía dưới điền CHÍNH XÁC chữ: `[MERGE_UP]`
                       - Ví dụ bảng có tiêu đề gộp:
                         | STT | TÊN VĂN BẢN | TÌNH TRẠNG HỒ SƠ | [MERGE_LEFT] | [MERGE_LEFT] | [MERGE_LEFT] |
                         | [MERGE_UP] | [MERGE_UP] | Có | [MERGE_LEFT] | [MERGE_LEFT] | Không có |
                         | [MERGE_UP] | [MERGE_UP] | Bản chính | Bản sao | Bản photo | [MERGE_UP] |
                       - TUYỆT ĐỐI không để ô trống làm lệch số lượng cột. Dùng `<br>` để xuống dòng trong ô.

                    5. KHÔNG sinh ra mã phân trang (Start of Page). KHÔNG dùng mã LaTeX toán học.
                    """

                    response = call_gemini_smart(
                        client, 
                        [types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    is_landscape = False
                    try:
                        if uploaded_file.name.lower().endswith('.pdf'):
                            pdf_reader = PdfReader(BytesIO(file_bytes))
                            first_page = pdf_reader.pages[0]
                            w = float(first_page.mediabox.width)
                            h = float(first_page.mediabox.height)
                            is_landscape = w > h
                        else:
                            img = Image.open(BytesIO(file_bytes))
                            is_landscape = img.width > img.height
                    except Exception as e:
                        is_landscape = False
                    
                    doc = Document()
                    section = doc.sections[0]
                    
                    if is_landscape:
                        section.orientation = WD_ORIENT.LANDSCAPE
                        section.page_width = Mm(297)
                        section.page_height = Mm(210)
                    else:
                        section.orientation = WD_ORIENT.PORTRAIT
                        section.page_width = Mm(210)
                        section.page_height = Mm(297)

                    section.top_margin = Mm(20)
                    section.bottom_margin = Mm(20)
                    section.left_margin = Mm(30)
                    section.right_margin = Mm(20)

                    style = doc.styles['Normal']
                    font = style.font
                    font.name = 'Times New Roman'
                    font.size = Pt(13)
                    
                    def build_docx_table(doc_obj, buffer, is_header_table=False):
                        if not buffer: return
                        num_cols = max(len(row) for row in buffer)
                        
                        normalized_buffer = []
                        for row in buffer:
                            new_row = list(row)
                            while len(new_row) < num_cols:
                                new_row.append('')
                            normalized_buffer.append(new_row)
                            
                        current_table = doc_obj.add_table(rows=len(normalized_buffer), cols=num_cols)
                        
                        if not is_header_table:
                            current_table.style = 'Table Grid'
                            current_table.autofit = True
                        else:
                            current_table.autofit = False

                        for row_idx, row_data in enumerate(normalized_buffer):
                            row_cells = current_table.rows[row_idx].cells
                            
                            if is_header_table and num_cols == 2:
                                row_cells[0].width = Cm(6.0)
                                row_cells[1].width = Cm(10.0)

                            for col_idx, cell_data in enumerate(row_data):
                                cell_text = cell_data.strip()
                                if cell_text not in ['[MERGE_LEFT]', '[MERGE_UP]']:
                                    cell = row_cells[col_idx]
                                    cell.text = ""
                                    cell_lines = cell_text.split('<br>')
                                    for idx, c_line in enumerate(cell_lines):
                                        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
                                        p.paragraph_format.space_after = Pt(0)
                                        
                                        default_al = WD_ALIGN_PARAGRAPH.CENTER if (is_header_table or row_idx == 0) else WD_ALIGN_PARAGRAPH.LEFT
                                        clean_text = clean_tags_and_align(c_line.strip(), p, default_al)
                                        parse_and_add_runs(p, clean_text)
                                        
                        for row_idx in range(len(normalized_buffer)):
                            for col_idx in range(num_cols):
                                cell_text = normalized_buffer[row_idx][col_idx].strip()
                                try:
                                    if cell_text == '[MERGE_LEFT]' and col_idx > 0:
                                        current_table.cell(row_idx, col_idx - 1).merge(current_table.cell(row_idx, col_idx))
                                    elif cell_text == '[MERGE_UP]' and row_idx > 0:
                                        current_table.cell(row_idx - 1, col_idx).merge(current_table.cell(row_idx, col_idx))
                                except Exception:
                                    pass

                    response_text = response.text
                    table_buffer = []
                    is_next_table_header = False

                    for line in response_text.split('\n'):
                        line_stripped = re.sub(r'<td[^>]*>', '', line.strip())
                        line_stripped = re.sub(r'</td>', '', line_stripped)
                        
                        line_stripped = re.sub(r'==\s*(Start|End)\s+of\s+Page.*?==', '', line_stripped, flags=re.IGNORECASE).strip()
                        line_stripped = line_stripped.replace('$\ge$', '≥').replace('$\rightarrow$', '→').replace('$\le$', '≤').replace('$\Rightarrow$', '⇒').replace('$\leftarrow$', '←')

                        if not line_stripped or line_stripped.startswith("```"):
                            continue

                        if line_stripped == '[HEADER_TABLE]':
                            is_next_table_header = True
                            continue

                        if line_stripped.startswith('|') and line_stripped.endswith('|'):
                            if re.match(r'^[\s\|\-:]+$', line_stripped):
                                continue
                            
                            cells_data = [cell.strip() for cell in line_stripped.split('|')][1:-1]
                            table_buffer.append(cells_data)
                        else:
                            if table_buffer:
                                if all(cell == '' for cell in table_buffer[0]):
                                    table_buffer.pop(0)
                                build_docx_table(doc, table_buffer, is_header_table=is_next_table_header)
                                table_buffer = []
                                is_next_table_header = False
                            
                            if line_stripped.startswith('---'):
                                continue
                                
                            if line_stripped:
                                p = doc.add_paragraph()
                                p.paragraph_format.space_after = Pt(6)
                                clean_line = clean_tags_and_align(line_stripped, p, WD_ALIGN_PARAGRAPH.JUSTIFY)
                                parse_and_add_runs(p, clean_line)

                    if table_buffer:
                        if all(cell == '' for cell in table_buffer[0]):
                            table_buffer.pop(0)
                        build_docx_table(doc, table_buffer, is_header_table=is_next_table_header)

                    output_docx_path = "ket_qua.docx"
                    doc.save(output_docx_path)

                    st.success("🎉 Chuyển đổi thành công! Bảng biểu phức tạp đã được gộp ô chính xác.")

                    with open(output_docx_path, "rb") as file_download:
                        st.download_button(
                            label="📥 Tải xuống file Word (.docx)",
                            data=file_download,
                            file_name=f"Converted_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

def app_number_2():
    st.title("🖨️ Chuyển PDF về khổ A4")
    st.markdown("Xóa bỏ mọi khung ẩn của bản vẽ cũ, ép lại chính xác thành khổ A4 tiêu chuẩn.")

    uploaded_pdf = st.file_uploader("Tải lên bản vẽ PDF cần xử lý:", type=["pdf"], key=f"app2_{st.session_state.uploader_key}")

    if uploaded_pdf is not None:
        st.success(f"Đã tải lên file: **{uploaded_pdf.name}**")
        
        if st.button("📏 Chuyển thành A4", type="primary"):
            with st.spinner("Đang truy quét và ghi đè các khung viền ẩn..."):
                try:
                    reader = PdfReader(uploaded_pdf)
                    writer = PdfWriter()

                    A4_W = 595.276
                    A4_H = 841.890

                    for page in reader.pages:
                        orig_w = float(page.mediabox.width)
                        orig_h = float(page.mediabox.height)

                        is_landscape = orig_w > orig_h
                        target_w = A4_H if is_landscape else A4_W
                        target_h = A4_W if is_landscape else A4_H

                        scale_w = target_w / orig_w
                        scale_h = target_h / orig_h
                        scale_factor = min(scale_w, scale_h)

                        scaled_w = orig_w * scale_factor
                        scaled_h = orig_h * scale_factor
                        tx = (target_w - scaled_w) / 2.0
                        ty = (target_h - scaled_h) / 2.0

                        op = Transformation().scale(sx=scale_factor, sy=scale_factor).translate(tx=tx, ty=ty)
                        page.add_transformation(op)

                        page.mediabox.lower_left = (0, 0)
                        page.mediabox.upper_right = (target_w, target_h)
                        page.cropbox.lower_left = (0, 0)
                        page.cropbox.upper_right = (target_w, target_h)
                        
                        if "/BleedBox" in page:
                            page.bleedbox.lower_left = (0, 0)
                            page.bleedbox.upper_right = (target_w, target_h)
                        if "/TrimBox" in page:
                            page.trimbox.lower_left = (0, 0)
                            page.trimbox.upper_right = (target_w, target_h)
                        if "/ArtBox" in page:
                            page.artbox.lower_left = (0, 0)
                            page.artbox.upper_right = (target_w, target_h)

                        writer.add_page(page)

                    output_path = f"A4_Chuan_{uploaded_pdf.name}"
                    with open(output_path, "wb") as f:
                        writer.write(f)

                    st.success("🎉 Xử lý thành công! Toàn bộ khung hình đã được chuyển thành A4.")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Tải xuống bản vẽ A4 chuẩn (.pdf)",
                            data=f,
                            file_name=output_path,
                            mime="application/pdf",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

def parse_rich_text(text_val, font_name="Times New Roman", size=12):
    text_str = str(text_val) if text_val is not None else ""
    parts = re.split(r'\*\*(.*?)\*\*', text_str)

    if len(parts) == 1:
        return text_str 

    rt = CellRichText()
    font_normal = InlineFont(rFont=font_name, sz=size)
    font_bold = InlineFont(rFont=font_name, sz=size, b=True)

    for i, p in enumerate(parts):
        if not p: continue
        if i % 2 == 1: 
            rt.append(TextBlock(font=font_bold, text=p))
        else: 
            rt.append(TextBlock(font=font_normal, text=p))
    return rt

def app_number_3():
    try:
        api_key_input = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
        st.stop()

    st.title("📊 Bóc tách PDF/Ảnh sang Excel (Chuẩn A4 & Giữ Định Dạng)")
    st.markdown("Trích xuất và tự động định dạng giống PDF gốc, ép sẵn khổ in **A4 (Ngang/Dọc tự động)**.")

    uploaded_excel_file = st.file_uploader("Tải lên tài liệu (Ảnh hoặc PDF):", type=["jpg", "jpeg", "png", "pdf"], key=f"app3_{st.session_state.uploader_key}")

    if uploaded_excel_file is not None:
        st.success(f"Đã tải lên file: **{uploaded_excel_file.name}**")
        
        if st.button("🚀 Trích xuất ra Excel", type="primary"):
            with st.spinner("🤖 Vui lòng đợi..."):
                try:
                    temp_input_path = f"temp_excel_{uploaded_excel_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_excel_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_excel_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    Bạn là một chuyên gia số hóa tài liệu. Hãy đọc kỹ tài liệu và trả về kết quả DUY NHẤT dưới dạng chuỗi JSON hợp lệ. Không trả lời thêm.
                    
                    LƯU Ý QUAN TRỌNG VỀ ĐỊNH DẠNG:
                    Hãy phân tích và BỌC CÁC CHỮ ĐƯỢC IN ĐẬM trong bản gốc bằng dấu sao kép (**). 
                    Ví dụ: "- **Thời gian:** 9 giờ ngày 15 tháng 7 năm 2026"
                    
                    Cấu trúc JSON BẮT BUỘC:
                    {
                      "title": "Dòng tiêu đề trên cùng",
                      "info_lines": [
                        "Dòng thông tin 1",
                        "Dòng thông tin 2"
                      ],
                      "headers": ["Cột 1", "Cột 2", "Cột 3"],
                      "rows": [
                        ["Dữ liệu 1", "Dữ liệu 2", "Dữ liệu 3"],
                        ["Dữ liệu 1", "Dữ liệu 2", "Dữ liệu 3"]
                      ]
                    }
                    """

                    response = call_gemini_smart(
                        client, 
                        [types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    is_landscape = False
                    try:
                        if uploaded_excel_file.name.lower().endswith('.pdf'):
                            pdf_reader = PdfReader(BytesIO(file_bytes))
                            first_page = pdf_reader.pages[0]
                            w = float(first_page.mediabox.width)
                            h = float(first_page.mediabox.height)
                            is_landscape = w > h
                        else:
                            img = Image.open(BytesIO(file_bytes))
                            is_landscape = img.width > img.height
                    except Exception as e:
                        is_landscape = False

                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.startswith("```"):
                        raw_text = raw_text[3:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    data = json.loads(raw_text)

                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Danh_Sach"

                    ws.page_setup.paperSize = ws.PAPERSIZE_A4
                    if is_landscape:
                        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    else:
                        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
                    ws.print_options.horizontalCentered = True 
                    
                    font_title = Font(name="Times New Roman", size=14, bold=True)
                    font_bold = Font(name="Times New Roman", size=12, bold=True)
                    font_normal = Font(name="Times New Roman", size=12)
                    align_center = Alignment(horizontal="center", vertical="center")
                    align_left = Alignment(horizontal="left", vertical="center")
                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                    current_row = 1
                    total_cols = len(data.get("headers", [1,2,3,4,5]))

                    title = data.get("title", "")
                    if title:
                        clean_title = str(title).replace('**', '').upper()
                        cell = ws.cell(row=current_row, column=1, value=clean_title)
                        cell.font = font_title
                        cell.alignment = align_center
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_cols)
                        current_row += 1

                    for info in data.get("info_lines", []):
                        cell = ws.cell(row=current_row, column=1)
                        rich_val = parse_rich_text(info)
                        
                        cell.value = rich_val
                        if isinstance(rich_val, str): 
                            cell.font = font_normal
                            
                        cell.alignment = align_left
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_cols)
                        current_row += 1

                    current_row += 1 

                    headers = data.get("headers", [])
                    for col_idx, header in enumerate(headers, 1):
                        clean_header = str(header).replace('**', '')
                        cell = ws.cell(row=current_row, column=col_idx, value=clean_header)
                        cell.font = font_bold
                        cell.alignment = align_center
                        cell.border = thin_border
                    current_row += 1

                    for row_data in data.get("rows", []):
                        for col_idx, val in enumerate(row_data, 1):
                            cell = ws.cell(row=current_row, column=col_idx)
                            rich_val = parse_rich_text(val)
                            
                            cell.value = rich_val
                            if isinstance(rich_val, str):
                                cell.font = font_normal
                                
                            cell.border = thin_border
                            
                            if col_idx == 1 or col_idx == total_cols:
                                cell.alignment = align_center
                            else:
                                cell.alignment = align_left
                        current_row += 1

                    ws.column_dimensions['A'].width = 8   
                    ws.column_dimensions['B'].width = 25  
                    ws.column_dimensions['C'].width = 30  
                    ws.column_dimensions['D'].width = 25  
                    ws.column_dimensions['E'].width = 15  

                    output = BytesIO()
                    wb.save(output)
                    processed_data = output.getvalue()

                    st.success("🎉 Đã xuất bảng Excel thành công! Khổ giấy in đã được set tự động xoay ngang/dọc.")

                    st.download_button(
                        label="📥 Tải xuống Excel Chuẩn Format (.xlsx)",
                        data=processed_data,
                        file_name=f"Excel_Chuan_{uploaded_excel_file.name.split('.')[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        on_click=clear_file
                    )

                except Exception as e:
                    st.error(f"Đã xảy ra lỗi hệ thống: {e}")

def extract_text_from_file(uploaded_file, client):
    text = ""
    file_ext = uploaded_file.name.split('.')[-1].lower()
    
    if file_ext == "docx":
        doc = Document(uploaded_file)
        text = "\n".join([para.text for para in doc.paragraphs])
    
    elif file_ext == "pdf":
        try:
            reader = PdfReader(uploaded_file)
            extracted = "\n".join([p.extract_text() or "" for p in reader.pages]).strip()
            if len(extracted) > 50: 
                uploaded_file.seek(0) 
                return extracted
        except Exception:
            pass
            
        uploaded_file.seek(0) 
        file_bytes = uploaded_file.getbuffer()
        prompt = "Hãy trích xuất TOÀN BỘ nội dung văn bản trong tài liệu này một cách chính xác nhất. Trả về plain text."
        response = call_gemini_smart(
            client, 
            [types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"), prompt]
        )
        text = response.text
        
    elif file_ext in ["png", "jpg", "jpeg"]:
        file_bytes = uploaded_file.getbuffer()
        prompt = "Hãy trích xuất TOÀN BỘ nội dung văn bản trong tài liệu này một cách chính xác nhất. Trả về plain text."
        response = call_gemini_smart(
            client, 
            [types.Part.from_bytes(data=file_bytes, mime_type=f"image/{file_ext}"), prompt]
        )
        text = response.text
    
    return text

def app_document_compare():
    try:
        api_key_input = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
        st.stop()

    st.title("🔍 So sánh Văn bản / Hợp đồng")
    st.markdown("Đối soát độ lệch chữ giữa 2 phiên bản tài liệu. Hỗ trợ đối chiếu chéo **Word vs Word**, **PDF scan vs Word**.")

    col1, col2 = st.columns(2)
    with col1:
        st.info("📄 BẢN GỐC (V1)")
        file_v1 = st.file_uploader("Tải lên bản gốc (PDF/Word/Ảnh):", key=f"v1_{st.session_state.uploader_key}")
    with col2:
        st.warning("📝 BẢN ĐỐI TÁC GỬI LẠI (V2)")
        file_v2 = st.file_uploader("Tải lên bản chỉnh sửa (PDF/Word/Ảnh):", key=f"v2_{st.session_state.uploader_key}")

    if file_v1 and file_v2:
        if st.button("🔎 Soi Sự Khác Biệt", type="primary"):
            with st.spinner("🤖 Đang quét tài liệu và so sánh từng ký tự..."):
                try:
                    client = genai.Client(api_key=api_key_input)
                    
                    st.toast("Đang đọc và giải mã Bản Gốc (V1)...")
                    text_v1 = extract_text_from_file(file_v1, client)
                    
                    st.toast("Đang đọc và giải mã Bản Đối Tác (V2)...")
                    text_v2 = extract_text_from_file(file_v2, client)

                    words_v1 = text_v1.splitlines() 
                    words_v2 = text_v2.splitlines()
                    
                    words_v1_flat = [word for line in words_v1 for word in line.split(" ") if word.strip()]
                    words_v2_flat = [word for line in words_v2 for word in line.split(" ") if word.strip()]

                    matcher = difflib.SequenceMatcher(None, words_v1_flat, words_v2_flat)
                    
                    edits = []
                    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                        if tag != 'equal':
                            edits.append({
                                'tag': tag,
                                'old': " ".join(words_v1_flat[i1:i2]),
                                'new': " ".join(words_v2_flat[j1:j2])
                            })
                    
                    total_edits = len(edits)

                    st.toast("AI đang đánh giá rủi ro các điểm thay đổi...")
                    summary_prompt = f"""
                    Tôi có 2 phiên bản hợp đồng. Phát hiện tổng cộng {total_edits} lần thay đổi.
                    Bản gốc: {text_v1[:3000]}... 
                    Bản sửa: {text_v2[:3000]}...
                    
                    Hãy viết một đoạn văn ngắn gọn đánh giá rủi ro pháp lý/ngữ nghĩa của các điểm khác biệt này.
                    BẮT BUỘC TUÂN THỦ: 
                    1. TUYỆT ĐỐI KHÔNG dùng ký tự đặc biệt như dấu sao (*), thăng (#), gạch ngang (-), hay bảng biểu (|). 
                    2. Trả lời bằng văn bản thuần túy (plain text) để đưa trực tiếp vào Word.
                    """
                    summary_response = call_gemini_smart(
                        client, 
                        [summary_prompt]
                    )
                    ai_text = summary_response.text.replace('*', '').replace('#', '').replace('`', '').strip()

                    doc_report = Document()
                    
                    style = doc_report.styles['Normal']
                    style.font.name = 'Times New Roman'
                    style.font.size = Pt(13)

                    p_title = doc_report.add_heading(level=0)
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_title = p_title.add_run('BÁO CÁO ĐỐI CHIẾU TÀI LIỆU (V1 vs V2)')
                    r_title.font.name = 'Times New Roman'
                    r_title.font.size = Pt(14)
                    
                    p_h1 = doc_report.add_heading(level=1)
                    r_h1 = p_h1.add_run('I. Bảng Tổng kết chi tiết các lần chỉnh sửa')
                    r_h1.font.name = 'Times New Roman'
                    r_h1.font.size = Pt(12)

                    p_total = doc_report.add_paragraph()
                    run_total = p_total.add_run(f"Hệ thống ghi nhận tổng cộng {total_edits} lần chỉnh sửa từ đối tác.")
                    run_total.bold = True
                    run_total.font.color.rgb = RGBColor(255, 0, 0)
                    
                    if total_edits > 0:
                        table = doc_report.add_table(rows=1, cols=4)
                        table.style = 'Table Grid'
                        hdr_cells = table.rows[0].cells
                        hdr_cells[0].text = 'STT'
                        hdr_cells[1].text = 'Thao tác'
                        hdr_cells[2].text = 'Bản gốc (V1)'
                        hdr_cells[3].text = 'Đối tác sửa (V2)'
                        
                        for cell in hdr_cells:
                            for p in cell.paragraphs:
                                for r in p.runs:
                                    r.font.name = 'Times New Roman'
                                    r.bold = True
                        
                        for idx, edit in enumerate(edits, 1):
                            row_cells = table.add_row().cells
                            row_cells[0].text = str(idx)
                            
                            if edit['tag'] == 'delete':
                                row_cells[1].text = 'Xóa bỏ'
                                row_cells[2].text = edit['old']
                                row_cells[3].text = "(Đã xóa)"
                            elif edit['tag'] == 'insert':
                                row_cells[1].text = 'Thêm mới'
                                row_cells[2].text = "(Không có)"
                                row_cells[3].text = edit['new']
                            elif edit['tag'] == 'replace':
                                row_cells[1].text = 'Thay thế'
                                row_cells[2].text = edit['old']
                                row_cells[3].text = edit['new']

                    p_h2 = doc_report.add_heading(level=1)
                    r_h2 = p_h2.add_run('II. Nhận định rủi ro tổng quan (AI)')
                    r_h2.font.name = 'Times New Roman'
                    r_h2.font.size = Pt(12)
                    
                    doc_report.add_paragraph(ai_text)

                    p_h3 = doc_report.add_heading(level=1)
                    r_h3 = p_h3.add_run('III. Chi tiết văn bản (Kính lúp bôi màu)')
                    r_h3.font.name = 'Times New Roman'
                    r_h3.font.size = Pt(12)
                    
                    legend = doc_report.add_paragraph()
                    run_del = legend.add_run("Chữ màu đỏ có gạch ngang: Bị đối tác xóa bỏ\n")
                    run_del.font.color.rgb = RGBColor(255, 0, 0)
                    run_del.font.strike = True
                    run_add = legend.add_run("Chữ màu xanh có gạch chân: Được đối tác thêm mới")
                    run_add.font.color.rgb = RGBColor(0, 0, 255)
                    run_add.font.underline = True

                    p_diff = doc_report.add_paragraph()
                    
                    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                        if tag == 'equal':
                            p_diff.add_run(" " + " ".join(words_v1_flat[i1:i2]))
                        elif tag == 'delete':
                            run = p_diff.add_run(" " + " ".join(words_v1_flat[i1:i2]))
                            run.font.color.rgb = RGBColor(255, 0, 0)
                            run.font.strike = True
                        elif tag == 'insert':
                            run = p_diff.add_run(" " + " ".join(words_v2_flat[j1:j2]))
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.font.underline = True
                        elif tag == 'replace':
                            run_old = p_diff.add_run(" " + " ".join(words_v1_flat[i1:i2]))
                            run_old.font.color.rgb = RGBColor(255, 0, 0)
                            run_old.font.strike = True
                            
                            run_new = p_diff.add_run(" " + " ".join(words_v2_flat[j1:j2]))
                            run_new.font.color.rgb = RGBColor(0, 0, 255)
                            run_new.font.underline = True

                    output_report_path = "Bao_Cao_So_Sanh.docx"
                    doc_report.save(output_report_path)

                    st.success("🎉 So sánh hoàn tất! Báo cáo đã được định dạng chuẩn hành chính.")

                    with open(output_report_path, "rb") as file_download:
                        st.download_button(
                            label="📥 Tải xuống Báo cáo chi tiết (.docx)",
                            data=file_download,
                            file_name="Bao_Cao_So_Sanh_V1_vs_V2.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

# ==========================================
# APP 5: CẮT, TRÍCH XUẤT TỔNG HỢP & GHÉP PDF
# ==========================================
def app_pdf_split_merge():
    import fitz
    from PIL import Image
    from io import BytesIO
    import zipfile

    st.title("✂️ Cắt, Trích Xuất & Ghép Nối PDF")
    st.markdown("Công cụ thao tác PDF toàn diện: Cắt trang trực quan qua ảnh xem trước, rã file hàng loạt (Burst), và ghép nối với khả năng tùy chỉnh thứ tự file linh hoạt.")

    # CSS căn giữa checkbox: căn ở KHUNG CHA (stElementContainer), không chỉ ở bên trong checkbox
    st.markdown("""
        <style>
        /* Khung cha chứa checkbox: căn giữa */
        div[data-testid="stElementContainer"]:has(div[data-testid="stCheckbox"]),
        div.element-container:has(div[data-testid="stCheckbox"]) {
            display: flex !important;
            justify-content: center !important;
            width: 100% !important;
        }
        /* Checkbox chỉ rộng vừa nội dung để được căn giữa */
        div[data-testid="stCheckbox"] {
            width: auto !important;
        }
        div[data-testid="stCheckbox"] label {
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        /* In đậm nhãn của checkbox */
        div[data-testid="stCheckbox"] label p {
            font-weight: bold !important;
        }
        </style>
    """, unsafe_allow_html=True)

    tab_split, tab_merge = st.tabs(["✂️ Cắt Trang & Tách File (Visual Split)", "📑 Ghép Nhiều File PDF (Reorder Merge)"])

    # ----------------------------------------------------
    # TAB 1: CẮT TRANG VÀ TÁCH FILE TRỰC QUAN (GỘP ĐA FILE)
    # ----------------------------------------------------
    with tab_split:
        st.subheader("1. Chọn các file PDF cần trích xuất")
        uploaded_split_files = st.file_uploader(
            "Tải lên một hoặc nhiều file PDF (Hệ thống sẽ trải tất cả các trang ra màn hình):",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"app5_split_{st.session_state.uploader_key}"
        )

        if uploaded_split_files:
            try:
                # Gộp ngầm tất cả các file vào 1 document trong bộ nhớ để trải ra màn hình
                combined_doc = fitz.open()
                for f in uploaded_split_files:
                    temp_doc = fitz.open("pdf", f.getvalue())
                    combined_doc.insert_pdf(temp_doc)
                
                total_pages = len(combined_doc)
                st.info(f"📁 Đã tải lên **{len(uploaded_split_files)} file**. Trải dài tổng cộng: **{total_pages} trang**.")

                # Tạo key lưu trạng thái dựa trên danh sách file để không bị loạn khi đổi file
                sig = "|".join([f.name for f in uploaded_split_files])
                state_check_key = f"split_selected_{sig}_{st.session_state.uploader_key}"
                
                if state_check_key not in st.session_state:
                    # MẶC ĐỊNH LÀ FALSE (Không chọn trang nào cả)
                    st.session_state[state_check_key] = [False] * total_pages
                    for i in range(total_pages):
                        st.session_state[f"chk_page_{i}_{state_check_key}"] = False

                st.markdown("---")
                col_view1, col_view2 = st.columns([2, 1])
                with col_view1:
                    st.markdown("### 👁️ Danh sách trang tổng hợp")
                    st.caption("Tick chọn những trang bạn muốn giữ lại để trích xuất thành 1 file duy nhất.")
                with col_view2:
                    grid_cols = st.slider("🔍 Cỡ ảnh (Số trang / hàng):", min_value=1, max_value=4, value=2, help="Kéo về 1 hoặc 2 để phóng to trang giấy.")

                zoom_mat = 0.8 if grid_cols <= 2 else 0.4
                mat = fitz.Matrix(zoom_mat, zoom_mat)
                cols = st.columns(grid_cols)

                for idx in range(total_pages):
                    c_idx = idx % grid_cols
                    page = combined_doc[idx]
                    with cols[c_idx]:
                        pix = page.get_pixmap(matrix=mat)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        st.image(img, use_container_width=True)

                        st.markdown(f"<div style='text-align: center; color: red; font-size: 24px; font-weight: 900; margin-top: 5px; margin-bottom: 0px;'>Trang {idx + 1}</div>", unsafe_allow_html=True)

                        st.session_state[state_check_key][idx] = st.checkbox(
                            "Chọn trang này",
                            key=f"chk_page_{idx}_{state_check_key}"
                        )
                        st.markdown("<br>", unsafe_allow_html=True)

                st.markdown("---")
                col_act1, col_act2 = st.columns(2)

                with col_act1:
                    st.markdown("#### 🎯 Trích xuất trang đã chọn")
                    selected_indices = [i for i, val in enumerate(st.session_state[state_check_key]) if val]
                    st.caption(f"Đang chọn **{len(selected_indices)}/{total_pages}** trang.")

                    if st.button("🚀 Xuất 1 PDF Đã Chọn", type="primary", disabled=(len(selected_indices) == 0)):
                        with st.spinner("Đang trích xuất trang đã chọn..."):
                            out_doc = fitz.open()
                            out_doc.insert_pdf(combined_doc, from_page=0, to_page=total_pages - 1)
                            out_doc.select(selected_indices)
                            out_stream = BytesIO()
                            out_doc.save(out_stream, garbage=3, deflate=True)
                            
                            st.success(f"🎉 Đã trích xuất thành công {len(selected_indices)} trang thành 1 file duy nhất!")
                            st.download_button(
                                label="📥 Tải xuống PDF Đã Cắt",
                                data=out_stream.getvalue(),
                                file_name="Tai_Lieu_Trich_Xuat_Tong_Hop.pdf",
                                mime="application/pdf",
                                on_click=clear_file
                            )

                with col_act2:
                    st.markdown("#### ⚡ Tách rời toàn bộ trang (Burst)")
                    st.caption("Rã từng trang thành các file PDF riêng biệt (ZIP).")
                    if st.button("📦 Rã Từng Trang Sang ZIP"):
                        with st.spinner(f"Đang bóc tách {total_pages} trang thành các file riêng..."):
                            zip_buffer = BytesIO()
                            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                                for i in range(total_pages):
                                    single_page_doc = fitz.open()
                                    single_page_doc.insert_pdf(combined_doc, from_page=i, to_page=i)
                                    p_stream = BytesIO()
                                    single_page_doc.save(p_stream, garbage=3, deflate=True)
                                    zf.writestr(f"Trang_{i+1:03d}.pdf", p_stream.getvalue())

                            st.success("🎉 Đã rã toàn bộ trang và đóng gói ZIP thành công!")
                            st.download_button(
                                label="📥 Tải xuống File ZIP",
                                data=zip_buffer.getvalue(),
                                file_name="Tach_Roi_Tong_Hop.zip",
                                mime="application/zip",
                                on_click=clear_file
                            )

            except Exception as e:
                st.error(f"Lỗi khi xử lý file PDF: {e}")

    # ----------------------------------------------------
    # TAB 2: GHÉP NHIỀU FILE VỚI TÍNH NĂNG ĐỔI THỨ TỰ (REORDER)
    # ----------------------------------------------------
    with tab_merge:
        st.subheader("2. Tải lên Khay chứa file PDF")
        uploaded_merge_files = st.file_uploader(
            "Tải lên các file PDF (Có thể tải lên nhiều lần):",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"app5_merge_{st.session_state.uploader_key}"
        )

        if uploaded_merge_files:
            merge_state_key = f"merge_order_list_{st.session_state.uploader_key}"
            sig_key = f"sig_{merge_state_key}"

            current_sig = "|".join(sorted([f.name for f in uploaded_merge_files]))
            if sig_key not in st.session_state or st.session_state[sig_key] != current_sig:
                st.session_state[merge_state_key] = list(uploaded_merge_files)
                st.session_state[sig_key] = current_sig

            st.markdown("### 📋 DANH SÁCH FILE CHỐT ĐỂ GHÉP")
            st.warning("Các file bị xóa ở danh sách dưới đây sẽ KHÔNG được ghép, mặc dù file đó vẫn còn nằm ở Khay tải lên phía trên.")
            st.info(f"Đang có **{len(st.session_state[merge_state_key])} file** sẵn sàng. Bấm mũi tên để đổi thứ tự.")

            for i, f in enumerate(st.session_state[merge_state_key]):
                col_name, col_up, col_down, col_del = st.columns([6, 1.2, 1.2, 1.2])
                with col_name:
                    file_mb = len(f.getvalue()) / (1024 * 1024)
                    st.markdown(f"**Vị trí {i+1}:** 📄 `{f.name}` *({file_mb:.2f} MB)*")
                with col_up:
                    if st.button("⬆️ Lên", key=f"up_{i}_{merge_state_key}", disabled=(i == 0)):
                        st.session_state[merge_state_key][i - 1], st.session_state[merge_state_key][i] = (
                            st.session_state[merge_state_key][i],
                            st.session_state[merge_state_key][i - 1]
                        )
                        st.rerun()
                with col_down:
                    if st.button("⬇️ Xuống", key=f"down_{i}_{merge_state_key}", disabled=(i == len(st.session_state[merge_state_key]) - 1)):
                        st.session_state[merge_state_key][i + 1], st.session_state[merge_state_key][i] = (
                            st.session_state[merge_state_key][i],
                            st.session_state[merge_state_key][i + 1]
                        )
                        st.rerun()
                with col_del:
                    if st.button("🗑️ Xóa khỏi DS", key=f"del_{i}_{merge_state_key}"):
                        st.session_state[merge_state_key].pop(i)
                        st.rerun()

            st.markdown("---")
            merge_name = st.text_input("Tên file PDF sau khi ghép nối:", value="Tai_Lieu_Ghep_Tong_Hop.pdf")
            if not merge_name.lower().endswith(".pdf"):
                merge_name += ".pdf"

            if st.button("🚀 Bắt đầu Ghép Nối", type="primary", disabled=(len(st.session_state[merge_state_key]) == 0)):
                with st.spinner("Đang tiến hành ghép nối các file PDF theo thứ tự chỉ định..."):
                    try:
                        merged_doc = fitz.open()
                        for file_obj in st.session_state[merge_state_key]:
                            sub_doc = fitz.open("pdf", file_obj.getvalue())
                            merged_doc.insert_pdf(sub_doc)

                        out_merged = BytesIO()
                        merged_doc.save(out_merged, garbage=3, deflate=True)

                        total_pages_merged = len(merged_doc)
                        merged_size_mb = len(out_merged.getvalue()) / (1024 * 1024)

                        st.success(f"🎉 Ghép thành công **{len(st.session_state[merge_state_key])} file** thành 1 file duy nhất! (Tổng cộng **{total_pages_merged} trang**, nặng **{merged_size_mb:.2f} MB**)")
                        st.download_button(
                            label=f"📥 Tải xuống {merge_name}",
                            data=out_merged.getvalue(),
                            file_name=merge_name,
                            mime="application/pdf",
                            on_click=clear_file
                        )
                    except Exception as e:
                        st.error(f"Lỗi khi ghép các file PDF: {e}")
                        # ==========================================

def app_excel_expert():
    try:
        api_key_input = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
        st.stop()

    st.title("💻 Chuyên gia Công thức & VBA Excel")
    st.markdown("Tải file lên, mô tả bài toán và nhận ngay Công thức hoặc Mã VBA chuẩn xác.")

    uploaded_file = st.file_uploader("Tải lên bảng dữ liệu (Excel/CSV) làm mẫu:", type=["csv", "xlsx", "xls"], key=f"app6_{st.session_state.uploader_key}")
    
    if uploaded_file:
        st.success(f"Đã nạp thành công: **{uploaded_file.name}**")
        
        context_str = ""
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, nrows=3)
                context_str = f"Tên file: {uploaded_file.name}\nDữ liệu mẫu (Tiêu đề và 3 dòng đầu):\n{df.to_csv(index=False)}"
            else:
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                df = pd.read_excel(uploaded_file, sheet_name=sheet_names[0], nrows=3)
                context_str = f"Các sheets có trong file: {', '.join(sheet_names)}\nDữ liệu mẫu của sheet đầu tiên (Tiêu đề và 3 dòng đầu):\n{df.to_csv(index=False)}"
        except Exception as e:
            st.error(f"Lỗi đọc file mẫu: {e}")
            return

        user_query = st.text_area("Bạn muốn xử lý dữ liệu này như thế nào? (Ví dụ: Tính tổng cột Thành tiền nếu cột Trạng thái là Đã giao)", height=100)

        if st.button("🚀 Tạo Công thức / Mã VBA", type="primary"):
            if not user_query.strip():
                st.warning("Vui lòng nhập yêu cầu của bạn!")
                return
                
            with st.spinner("🤖 AI đang tính toán phương án tối ưu nhất..."):
                client = genai.Client(api_key=api_key_input)
                
                prompt = f"""
                Bạn là Chuyên gia Cố vấn Excel và VBA.
                
                ĐÂY LÀ CẤU TRÚC FILE CỦA NGƯỜI DÙNG:
                {context_str}
                
                YÊU CẦU CỦA NGƯỜI DÙNG:
                "{user_query}"
                
                QUY TẮC PHẢN HỒI (BẮT BUỘC TUÂN THỦ):
                1. Quy tắc Ưu tiên (Formula First): Luôn cố gắng giải quyết bằng Công thức Excel tiêu chuẩn trước. Nếu công thức ngắn và tối ưu, hãy trả về công thức.
                2. Đo lường độ phức tạp (Chuyển cấp VBA): Nếu bài toán quá khó, cần vòng lặp, cần thao tác qua nhiều sheet, hoặc công thức quá nặng -> BẮT BUỘC dùng mã VBA.
                3. Luật VBA: Nếu dùng VBA, TRƯỚC KHI in mã code, BẮT BUỘC phải in ra dòng cảnh báo này: "Yêu cầu này không thể giải quyết hiệu quả bằng công thức. Hệ thống đề xuất dùng đoạn mã VBA dưới đây để xử lý tự động và nhẹ máy hơn:".
                4. Hướng dẫn VBA: NẾU có mã VBA, phải in ra 4 bước hướng dẫn sau ở cuối cùng:
                   - Cách dùng VBA:
                   1. Nhấn Alt + F11 để mở cửa sổ VBA.
                   2. Chọn Insert > Module và dán đoạn mã này vào.
                   3. Tại file Excel, vào tab Insert > Shapes > Vẽ một hình chữ nhật.
                   4. Click chuột phải vào hình vừa vẽ > Chọn Assign Macro... > Chọn tên Macro vừa dán > OK. Bấm vào hình để chạy!
                5. Vấn đề Dấu phân cách: Trong công thức, sử dụng dấu chấm phẩy (;) để phân cách các hàm (chuẩn máy tính Việt Nam).
                6. Tính tương thích: Ưu tiên dùng các hàm phổ biến (INDEX, MATCH, IF, SUMIFS...).
                7. TUYỆT ĐỐI xuất code trong khối lệnh Markdown (ví dụ ```excel hoặc ```vba) để giao diện hiển thị nút Copy.
                8. Trả lời cực kỳ ngắn gọn, đi thẳng vào vấn đề, không giải thích dài dòng lan man.
                """
                
                try:
                    response = call_gemini_smart(
                        client, 
                        [prompt]
                    )
                    
                    st.info("💡 Kết quả từ AI (Nhấp vào biểu tượng ở góc phải khung code để sao chép):")
                    st.markdown(response.text)
                    
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi AI: {e}")

# ==========================================
# APP 7: NÉN DUNG LƯỢNG & ĐÓNG GÓI ZIP GỬI MAIL (CỐT LÕI PYMUPDF - TỐI ƯU MÀU SẮC)
# ==========================================
def compress_single_image(img_bytes, quality=50, is_color=True):
    try:
        img = Image.open(BytesIO(img_bytes))
        
        # Xử lý hệ màu theo lựa chọn của người dùng
        if not is_color:
            img = img.convert("L")
        elif img.mode != "RGB":
            img = img.convert("RGB")
            
        # Thu nhỏ khung hình nếu ảnh quá lớn
        max_dim = int(800 + (quality - 30) * (1600 - 800) / (90 - 30))
        if max(img.width, img.height) > max_dim:
            ratio = max_dim / float(max(img.width, img.height))
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS)
                
        out = BytesIO()
        img.save(out, format="JPEG", quality=quality, optimize=True)
        return out.getvalue()
    except Exception:
        return img_bytes

def compress_pdf_extreme(pdf_bytes, quality=50, is_color=True):
    """Sử dụng PyMuPDF cào bằng trang PDF, tối ưu hóa DPI và giữ nguyên hệ màu RGB"""
    try:
        doc = fitz.open("pdf", pdf_bytes)
        new_pdf = fitz.open()
        
        # Ép DPI từ 100 (siêu nhẹ) đến 200 (sắc nét) dựa trên thanh trượt
        dpi = int(100 + (quality - 30) * (200 - 100) / (90 - 30))
        
        for page in doc:
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            
            # Khởi tạo bản vẽ với RGB (Màu) hoặc GRAY (Đen trắng)
            c_space = fitz.csRGB if is_color else fitz.csGRAY
            pix = page.get_pixmap(matrix=mat, colorspace=c_space)
            
            # Ép nén bằng thuật toán phân tích hạt JPEG
            img_bytes = pix.tobytes("jpeg", jpg_quality=quality)
            
            # Gói lại vào trang PDF mới tinh (loại bỏ hoàn toàn rác mã hóa của máy scan)
            imgdoc = fitz.open("jpeg", img_bytes)
            pdfbytes = imgdoc.convert_to_pdf()
            imgpdf = fitz.open("pdf", pdfbytes)
            new_pdf.insert_pdf(imgpdf)
            
        out_stream = BytesIO()
        new_pdf.save(out_stream, garbage=4, deflate=True)
        return out_stream.getvalue()
    except Exception as e:
        return pdf_bytes

def app_compress_and_zip():
    st.title("🗜️ Nén Dung Lượng & Đóng Gói ZIP Gửi Mail")
    st.markdown("Tối ưu hóa các file PDF/Ảnh scan nặng, tự động đóng gói toàn bộ vào 1 file **.ZIP** để gửi email công ty không bị vượt trần 25MB.")

    uploaded_files = st.file_uploader(
        "Tải lên một hoặc nhiều file cần nén và đóng gói (PDF, JPG, PNG, DOCX, XLSX...):",
        accept_multiple_files=True,
        key=f"app7_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        total_orig_bytes = sum(len(f.getvalue()) for f in uploaded_files)
        total_orig_mb = total_orig_bytes / (1024 * 1024)
        st.info(f"📁 Đã chọn **{len(uploaded_files)} file** | Tổng dung lượng gốc: **{total_orig_mb:.2f} MB**")

        # Giao diện điều khiển mới chia làm 2 cột
        col1, col2 = st.columns(2)
        with col1:
            che_do_mau = st.radio(
                "Chế độ màu sắc:",
                ["🌈 Giữ nguyên màu gốc", "🔲 Chuyển Đen Trắng"],
                help="Chế độ Giữ Màu vẫn ép nén rất tốt. Chỉ dùng Đen Trắng khi bạn cần file đạt mức siêu nhẹ gọn."
            )
        with col2:
            quality = st.slider(
                "Mức chất lượng nén (Mặc định 50):",
                min_value=30,
                max_value=90,
                value=50,
                step=5,
                help="Điều chỉnh thanh trượt càng thấp, cả độ phân giải và chất lượng ảnh sẽ càng được tối ưu nhỏ lại."
            )
        
        # Bổ sung khung hướng dẫn sử dụng chi tiết
        st.markdown("""
        > 💡 **Bạn hoàn toàn làm chủ độ nét thông qua Thanh trượt chất lượng:**
        > - **Kéo lên mức 70 - 90:** Dành cho các hợp đồng quan trọng, ảnh chụp linh kiện máy móc cần độ nét cao để nghiệm thu (chấp nhận file ZIP đầu ra sẽ nặng hơn một chút).
        > - **Để mức 40 - 50:** Tối ưu hóa tối đa để đính kèm qua email nội bộ bị giới hạn 25MB.
        """)

        # Biến boolean kiểm tra chế độ màu
        is_color = "🌈" in che_do_mau

        zip_filename = st.text_input("Tên file ZIP xuất ra:", value="Tai_Lieu_Gui_Mail.zip")
        if not zip_filename.lower().endswith(".zip"):
            zip_filename += ".zip"

        if st.button("🚀 Bắt đầu Nén & Đóng Gói ZIP", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            zip_buffer = BytesIO()
            report_rows = []

            # Sử dụng thuật toán ZipDeflated cấp độ nén cao nhất
            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
                for idx, file in enumerate(uploaded_files):
                    fname = file.name
                    fext = fname.split('.')[-1].lower()
                    raw_bytes = file.getvalue()
                    orig_len = len(raw_bytes)

                    status_text.text(f"Đang xử lý ({idx+1}/{len(uploaded_files)}): {fname}...")

                    # Phân luồng nén
                    if fext == "pdf":
                        processed_bytes = compress_pdf_extreme(raw_bytes, quality=quality, is_color=is_color)
                    elif fext in ["jpg", "jpeg", "png"]:
                        processed_bytes = compress_single_image(raw_bytes, quality=quality, is_color=is_color)
                        # Đồng nhất định dạng ảnh đầu ra thành JPG để nhẹ gọn nhất
                        if not fname.lower().endswith(".jpg") and not fname.lower().endswith(".jpeg"):
                            fname = f"{fname.rsplit('.', 1)[0]}.jpg"
                    else:
                        # Giữ nguyên bản gốc đối với các file Word, Excel
                        processed_bytes = raw_bytes

                    new_len = len(processed_bytes)
                    zf.writestr(fname, processed_bytes)

                    saved_pct = ((orig_len - new_len) / orig_len * 100) if orig_len > 0 else 0
                    report_rows.append({
                        "Tên file": fname,
                        "Gốc (MB)": f"{orig_len / (1024*1024):.2f}",
                        "Sau nén (MB)": f"{new_len / (1024*1024):.2f}",
                        "Tiết kiệm": f"{saved_pct:.1f}%"
                    })

                    progress_bar.progress((idx + 1) / len(uploaded_files))

            final_zip_bytes = zip_buffer.getvalue()
            final_zip_mb = len(final_zip_bytes) / (1024 * 1024)
            total_saved_pct = ((total_orig_bytes - len(final_zip_bytes)) / total_orig_bytes * 100) if total_orig_bytes > 0 else 0

            status_text.empty()
            progress_bar.empty()

            # Hiển thị số liệu trực quan
            col1, col2, col3 = st.columns(3)
            col1.metric("Tổng dung lượng gốc", f"{total_orig_mb:.2f} MB")
            col2.metric("File ZIP hoàn thành", f"{final_zip_mb:.2f} MB")
            col3.metric("Tiết kiệm được", f"{total_saved_pct:.1f}%")

            if final_zip_mb <= 25.0:
                st.success(f"🎉 Siêu nén thành công! File ZIP nặng **{final_zip_mb:.2f} MB**, hoàn toàn đủ điều kiện gửi email công ty.")
            else:
                st.warning(f"⚠️ File ZIP hiện tại nặng **{final_zip_mb:.2f} MB** (vẫn vượt 25MB). Hãy thử giảm mức chất lượng xuống 30-40% hoặc chọn Đen trắng.")

            st.dataframe(pd.DataFrame(report_rows), use_container_width=True)

            st.download_button(
                label=f"📥 Tải xuống {zip_filename}",
                data=final_zip_bytes,
                file_name=zip_filename,
                mime="application/zip",
                on_click=clear_file
            )
# ==========================================
# ==========================================
# APP 8: XOAY TÙY Ý TỪNG TRANG & TỰ ĐỘNG GHÉP FILE
# ==========================================
def app_rotate_pdf():
    import fitz
    from io import BytesIO
    from PIL import Image

    st.title("🔄 Xoay Trang PDF & Ghép File")
    st.markdown("Tải lên nhiều file PDF, điều chỉnh góc xoay cho từng trang bị ngược. Hệ thống sẽ tự động ghép tất cả lại thành **1 file PDF duy nhất**.")

    uploaded_files = st.file_uploader(
        "Tải lên các file PDF cần kiểm tra, xoay và ghép:", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key=f"app8_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        st.info(f"📁 Đã chọn **{len(uploaded_files)} file**. Các file sẽ được ghép nối theo đúng thứ tự tải lên.")
        
        # Lặp qua từng file và tạo khung mở rộng riêng
        for file_idx, uploaded_file in enumerate(uploaded_files):
            pdf_bytes = uploaded_file.getvalue()
            
            try:
                doc = fitz.open("pdf", pdf_bytes)
                num_pages = len(doc)
                
                with st.expander(f"📄 File {file_idx + 1}: {uploaded_file.name} ({num_pages} trang)", expanded=True):
                    # Quản lý bộ nhớ trạng thái xoay riêng cho từng file
                    state_key = f"rotations_{uploaded_file.name}_{st.session_state.uploader_key}"
                    if state_key not in st.session_state:
                        st.session_state[state_key] = {i: 0 for i in range(num_pages)}

                    cols = st.columns(3)
                    
                    for i in range(num_pages):
                        col_idx = i % 3
                        page = doc[i]
                        
                        with cols[col_idx]:
                            st.markdown(f"<div style='text-align: center; color: #003366; margin-top: 15px;'><b>Trang {i+1}</b></div>", unsafe_allow_html=True)
                            
                            rot_val = st.selectbox(
                                "Hướng xoay:",
                                options=[0, 90, 180, 270],
                                index=[0, 90, 180, 270].index(st.session_state[state_key][i]),
                                format_func=lambda x: "✅ Giữ nguyên" if x == 0 else "↪️ Phải 90°" if x == 90 else "🔃 Ngược 180°" if x == 180 else "↩️ Trái 90°",
                                key=f"select_rot_{uploaded_file.name}_{i}_{state_key}",
                                label_visibility="collapsed"
                            )
                            st.session_state[state_key][i] = rot_val
                            
                            current_rot = page.rotation
                            page.set_rotation((current_rot + rot_val) % 360)
                            
                            mat = fitz.Matrix(0.3, 0.3)
                            pix = page.get_pixmap(matrix=mat)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            
                            st.image(img, use_container_width=True)
                            st.markdown("<br>", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Lỗi khi đọc file {uploaded_file.name}: {e}")

        st.markdown("---")
        
        if st.button("🚀 Lưu & Tự Động Ghép File", type="primary"):
            with st.spinner("Đang áp dụng góc xoay và ghép tất cả thành 1 file PDF duy nhất..."):
                try:
                    # Khởi tạo một file PDF tổng (rỗng)
                    merged_doc = fitz.open()
                    
                    for file in uploaded_files:
                        out_doc = fitz.open("pdf", file.getvalue())
                        state_key = f"rotations_{file.name}_{st.session_state.uploader_key}"
                        
                        if state_key in st.session_state:
                            for i in range(len(out_doc)):
                                p = out_doc[i]
                                p.set_rotation((p.rotation + st.session_state[state_key][i]) % 360)
                        
                        # Chèn file PDF đã xoay vào file tổng
                        merged_doc.insert_pdf(out_doc)
                    
                    out_stream = BytesIO()
                    merged_doc.save(out_stream, garbage=3, deflate=True)
                    
                    st.success("🎉 Đã xoay và ghép xong tất cả các file thành 1 file duy nhất!")
                    st.download_button(
                        label="📥 Tải xuống PDF Đã Ghép",
                        data=out_stream.getvalue(),
                        file_name="Tai_Lieu_Da_Xoay_Va_Ghep.pdf",
                        mime="application/pdf",
                        on_click=clear_file
                    )
                except Exception as e:
                    st.error(f"Có lỗi trong quá trình ghép file: {e}")
# ==========================================
# ==========================================
# APP 9: XÓA SIÊU DỮ LIỆU (METADATA CLEANER) BẢO MẬT TÀI LIỆU
# ==========================================
def clean_pdf_metadata(pdf_bytes):
    import fitz
    from io import BytesIO
    doc = fitz.open("pdf", pdf_bytes)
    # Gán trắng toàn bộ từ điển Metadata ẩn
    doc.set_metadata({})
    out = BytesIO()
    doc.save(out, garbage=4, deflate=True)
    return out.getvalue()

def clean_docx_metadata(docx_bytes):
    from docx import Document
    from io import BytesIO
    doc = Document(BytesIO(docx_bytes))
    props = doc.core_properties
    props.author = ""
    props.last_modified_by = ""
    props.comments = ""
    props.title = ""
    props.subject = ""
    props.category = ""
    props.keywords = ""
    out = BytesIO()
    doc.save(out)
    return out.getvalue()

def clean_xlsx_metadata(xlsx_bytes):
    import openpyxl
    from io import BytesIO
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes))
    props = wb.properties
    props.creator = ""
    props.lastModifiedBy = ""
    props.title = ""
    props.subject = ""
    props.description = ""
    props.keywords = ""
    props.category = ""
    props.company = ""
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def app_metadata_cleaner():
    import zipfile
    from io import BytesIO

    st.title("🛡️ Xóa Siêu Dữ Liệu (Metadata) Bảo Mật Tài Liệu")
    st.markdown("Tự động xóa sạch tên tác giả, người chỉnh sửa cuối cùng, thông tin máy in, phần mềm và lịch sử ẩn bên trong file **PDF, Word (.docx), Excel (.xlsx)** trước khi gửi ra ngoài.")

    uploaded_files = st.file_uploader(
        "Tải lên các file cần tẩy sạch dữ liệu ẩn (Hỗ trợ PDF, DOCX, XLSX):",
        type=["pdf", "docx", "xlsx"],
        accept_multiple_files=True,
        key=f"app9_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        st.info(f"📁 Đã chọn **{len(uploaded_files)} file** sẵn sàng làm sạch dữ liệu bảo mật.")

        if st.button("🚀 Bắt đầu Tẩy Sạch Dữ Liệu Ẩn", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            # Trường hợp 1: Tải lên 1 file -> Trả về đúng định dạng gốc
            if len(uploaded_files) == 1:
                file = uploaded_files[0]
                fname = file.name
                fext = fname.split('.')[-1].lower()
                raw_bytes = file.getvalue()

                status_text.text(f"Đang làm sạch file: {fname}...")
                try:
                    if fext == "pdf":
                        cleaned_bytes = clean_pdf_metadata(raw_bytes)
                        mime_type = "application/pdf"
                    elif fext == "docx":
                        cleaned_bytes = clean_docx_metadata(raw_bytes)
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    elif fext == "xlsx":
                        cleaned_bytes = clean_xlsx_metadata(raw_bytes)
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    else:
                        cleaned_bytes = raw_bytes
                        mime_type = "application/octet-stream"

                    progress_bar.progress(100)
                    status_text.empty()
                    st.success(f"🎉 Đã tẩy sạch toàn bộ metadata của file **{fname}**!")

                    st.download_button(
                        label=f"📥 Tải xuống {fname} (Bản sạch)",
                        data=cleaned_bytes,
                        file_name=f"Cleaned_{fname}",
                        mime=mime_type,
                        on_click=clear_file
                    )
                except Exception as e:
                    st.error(f"Lỗi khi xử lý file: {e}")

            # Trường hợp 2: Tải lên nhiều file -> Làm sạch và đóng gói ZIP
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for idx, file in enumerate(uploaded_files):
                        fname = file.name
                        fext = fname.split('.')[-1].lower()
                        raw_bytes = file.getvalue()

                        status_text.text(f"Đang xử lý ({idx+1}/{len(uploaded_files)}): {fname}...")
                        try:
                            if fext == "pdf":
                                cleaned_bytes = clean_pdf_metadata(raw_bytes)
                            elif fext == "docx":
                                cleaned_bytes = clean_docx_metadata(raw_bytes)
                            elif fext == "xlsx":
                                cleaned_bytes = clean_xlsx_metadata(raw_bytes)
                            else:
                                cleaned_bytes = raw_bytes

                            zf.writestr(f"Cleaned_{fname}", cleaned_bytes)
                        except Exception:
                            zf.writestr(f"Loi_{fname}", raw_bytes)

                        progress_bar.progress((idx + 1) / len(uploaded_files))

                progress_bar.progress(100)
                status_text.empty()
                st.success("🎉 Đã làm sạch toàn bộ metadata và đóng gói thành công!")

                st.download_button(
                    label="📥 Tải xuống File ZIP đã làm sạch",
                    data=zip_buffer.getvalue(),
                    file_name="Tai_Lieu_Da_Lam_Sach.zip",
                    mime="application/zip",
                    on_click=clear_file
                )
# ==========================================
# APP 10: CHUYỂN PDF SCAN SANG WORD (TỰ ĐỘNG SĂN MODEL API + DỰNG LAYOUT CHUẨN)
# ==========================================
# ==========================================
# APP 10: PDF SCAN / ẢNH SANG WORD (GIỮ NGUYÊN BỐ CỤC) - GEMINI ONLINE
# ==========================================
# Nguyên lý: mỗi trang -> ảnh độ phân giải cao -> Gemini (vision) trả về "bản đồ bố cục" JSON
# (đoạn văn, bảng, ảnh + toạ độ) -> cắt ảnh gốc cho logo/QR/chữ ký/hình -> dựng Word thật:
# đúng khổ giấy, lề, khoảng cách dòng, thụt lề, bảng gộp ô, ảnh nội tuyến/nổi.
import io
import json
import re
import time
import math
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

from PIL import Image
from docx import Document
from docx.shared import Pt, Mm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# Danh sách model ưu tiên dòng 1.5 để có hạn mức cao nhất (1500 request/ngày)
A10_MODELS_PRECISE = ["gemini-1.5-pro", "gemini-1.5-flash"]
A10_MODELS_FAST = ["gemini-1.5-flash"]
A10_RENDER_DPI = 250          # độ phân giải dựng ảnh trang
A10_API_MAX_SIDE = 2600       # cạnh dài tối đa của ảnh gửi lên API

_A10_SCHEMA_TXT = '''{
  "font_family": "serif",
  "blocks": [
    {"type": "paragraph", "bbox": [ymin, xmin, ymax, xmax], "align": "left|center|right|justify",
     "lines": 2, "first_indent_pt": 0,
     "runs": [{"t": "văn bản", "b": 0, "i": 0, "u": 0, "c": ""}]},
    {"type": "table", "bbox": [ymin, xmin, ymax, xmax], "borders": true, "font_size_pt": 11,
     "col_pct": [10, 40, 50], "row_pct": [20, 40, 40],
     "rows": [[{"runs": [{"t": "ô", "b": 0, "i": 0, "u": 0, "c": ""}], "align": "left|center|right",
                "valign": "top|middle|bottom", "colspan": 1, "rowspan": 1, "fill": "", "image": null}]]},
    {"type": "image", "bbox": [ymin, xmin, ymax, xmax], "floating": false}
  ]
}'''

_A10_PROMPT = """Bạn là hệ thống OCR + phân tích bố cục tài liệu có độ chính xác cao nhất. Nhiệm vụ: đọc ảnh một trang tài liệu và trả về BẢN ĐỒ BỐ CỤC dạng JSON để dựng lại file Word giống bản gốc đến từng chi tiết.

TOẠ ĐỘ: mọi "bbox" là [ymin, xmin, ymax, xmax], thang 0-1000 theo kích thước ảnh trang (0,0 là góc trên-trái). bbox phải ÔM SÁT vùng chữ/ảnh thật.

QUY TẮC:
1. Chép CHÍNH XÁC từng ký tự: dấu tiếng Việt, số, ngày tháng, số hiệu, dấu câu, chữ in hoa/thường. Không sửa, không dịch, không tóm tắt, không bỏ sót dòng nào (kể cả đầu trang, chân trang, số trang, ghi chú nhỏ).
2. KHÔNG dùng markdown (không **, *, #). Đậm/nghiêng/gạch chân biểu diễn bằng cờ b/i/u (1 hoặc 0) trong từng run. "c" là mã màu hex (vd "#C00000") chỉ khi chữ KHÔNG phải màu đen, còn lại để "".
3. Khối chữ chạy liên tục (đoạn văn tự xuống dòng theo lề) -> 1 block "paragraph", KHÔNG chèn \\n tại chỗ xuống dòng tự nhiên. Chỉ dùng \\n khi bản gốc chủ động ngắt dòng (tiêu đề nhiều dòng căn giữa, địa chỉ, danh sách từng dòng, khối chữ ký).
4. "lines" = số dòng thực tế của khối đó trong ảnh. "align" theo thực tế. "first_indent_pt" = độ thụt đầu dòng đầu tiên (pt, 0 nếu không có).
5. Nội dung đặt NGANG HÀNG nhau (quốc hiệu 2 cột, khối ký tên 2-3 cột, logo bên cạnh tên cơ quan...) -> gộp thành 1 block "table" với "borders": false. Bảng có kẻ ô thật -> "borders": true.
6. Bảng: liệt kê đủ mọi hàng/cột; ô gộp dùng colspan/rowspan (các ô bị gộp KHÔNG khai báo lại). "col_pct" = % độ rộng từng cột (tổng 100), "row_pct" = % chiều cao từng hàng (tổng 100). "fill" = màu nền hex nếu ô có tô nền (vd "#D9D9D9"), còn lại "". "font_size_pt" = cỡ chữ ước lượng trong bảng. Nếu trong ô có hình (logo, QR, chữ ký, con dấu) -> điền "image": [ymin,xmin,ymax,xmax] của hình đó (toạ độ trang), ô vẫn có thể có chữ kèm theo.
7. Hình ảnh/biểu đồ/sơ đồ/logo/mã QR/ảnh chụp/chữ ký tay/con dấu đứng riêng -> block "image" (KHÔNG chép chữ nằm bên trong logo/con dấu). Con dấu hoặc chữ ký đè lên chữ khác -> "floating": true; còn lại false.
8. Sắp xếp blocks theo thứ tự đọc từ trên xuống dưới. Các block không được chồng bbox lên nhau (trừ ảnh floating).
9. "font_family": "serif" nếu chữ có chân (kiểu Times New Roman), "sans" nếu không chân (kiểu Arial/Calibri).
10. Trang trống hoặc chỉ có nền -> "blocks": [].

Chỉ trả về MỘT đối tượng JSON thuần theo cấu trúc sau, không giải thích:
""" + _A10_SCHEMA_TXT + """
<<HINT>>"""


# ------------------------------------------------------------------ tiện ích nhỏ
def _a10_clean(t):
    if t is None:
        return ""
    t = str(t).replace("**", "").replace("`", "")
    return t.replace("\u200b", "")


def _a10_clamp(v, lo, hi):
    return max(lo, min(hi, v))


def _a10_num(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


def _a10_hex(c):
    """'#C00000' -> RGBColor, trả None nếu rỗng/đen/không hợp lệ."""
    if not c:
        return None
    c = str(c).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", c):
        return None
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    if r < 40 and g < 40 and b < 40:
        return None
    return RGBColor(r, g, b)


def _a10_bbox(b):
    try:
        y0, x0, y1, x1 = [float(v) for v in b]
    except Exception:
        return None
    y0, y1 = sorted((_a10_clamp(y0, 0, 1000), _a10_clamp(y1, 0, 1000)))
    x0, x1 = sorted((_a10_clamp(x0, 0, 1000), _a10_clamp(x1, 0, 1000)))
    if y1 - y0 < 0.5 or x1 - x0 < 0.5:
        return None
    return [y0, x0, y1, x1]


def _a10_parse_json(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    try:
        return json.loads(text)
    except Exception:
        s, e = text.find("{"), text.rfind("}")
        if s != -1 and e > s:
            return json.loads(text[s:e + 1])
        raise


def _a10_sanitize(layout):
    """Chuẩn hoá JSON từ AI: bỏ khối hỏng, ép kiểu, trả về (layout, số khối bị bỏ)."""
    dropped = 0
    out = {"font_family": "sans" if str(layout.get("font_family", "")).lower().startswith("sans") else "serif",
           "blocks": []}
    for blk in layout.get("blocks", []) or []:
        if not isinstance(blk, dict):
            dropped += 1
            continue
        bb = _a10_bbox(blk.get("bbox"))
        t = str(blk.get("type", "")).lower()
        if bb is None or t not in ("paragraph", "table", "image"):
            dropped += 1
            continue
        blk["bbox"] = bb
        blk["type"] = t
        if t == "table":
            rows = blk.get("rows") or []
            if not rows:
                dropped += 1
                continue
            for row in rows:
                for cell in row:
                    if "image" in cell:
                        cell["image"] = _a10_bbox(cell.get("image")) if cell.get("image") else None
        out["blocks"].append(blk)
    out["blocks"].sort(key=lambda b: (b["bbox"][0], b["bbox"][1]))
    return out, dropped


# ------------------------------------------------------------------ gọi Gemini online
# ------------------------------------------------------------------ gọi Gemini online
def _a10_call_gemini(api_key, img_bytes, hint_text, models):
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    hint = ""
    if hint_text:
        hint = ("\n(VĂN BẢN GỐC lấy từ lớp text của PDF (đáng tin hơn OCR về mặt ký tự; hãy dùng để chép đúng \n"
                "chính tả, số liệu, dấu tiếng Việt, nhưng bố cục/vị trí vẫn lấy theo ảnh):\n\n"
                + hint_text[:12000] + "\n)")
    prompt = _A10_PROMPT.replace("HINT", hint)
    last_err = None
    
    for model in models:
        for attempt in range(3):
            try:
                resp = client.models.generate_content(
                    model=model,
                    contents=[types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"), prompt],
                    config=types.GenerateContentConfig(
                        temperature=0.0,
                        response_mime_type="application/json",
                        max_output_tokens=8192,
                    ),
                )
                return _a10_parse_json(resp.text), model
            except json.JSONDecodeError as e:          
                last_err = e
                continue
            except Exception as e:
                last_err = e
                msg = str(e).lower()
                # Chạm hạn mức 15 yêu cầu/phút (429) -> Ngủ 10 giây để API nhả khóa
                if any(x in msg for x in ('429', 'quota', 'exhausted')):
                    time.sleep(10)
                    continue
                # Máy chủ nghẽn mạng (503) -> Ngủ 3 giây
                if any(x in msg for x in ('503', 'unavailable', 'overloaded', 'high demand')):
                    time.sleep(3)
                    continue
                # Model không tồn tại -> Lập tức thoát vòng lặp để đổi model
                if "404" in msg or "not found" in msg:
                    break 
                break
    raise RuntimeError(f"Tất cả model đều thất bại. Lỗi cuối: {last_err}")
# ------------------------------------------------------------------ đọc số liệu gốc của PDF (nếu có text layer)
_A10_FONT_MAP = [("times", "Times New Roman"), ("arial", "Arial"), ("helvet", "Arial"), ("calibri", "Calibri"),
                 ("cambria", "Cambria"), ("tahoma", "Tahoma"), ("verdana", "Verdana"), ("segoe", "Segoe UI"),
                 ("courier", "Courier New"), ("georgia", "Georgia"), ("garamond", "Garamond"),
                 ("palatino", "Palatino Linotype"), ("book antiqua", "Book Antiqua"), ("vni-times", "Times New Roman"),
                 ("timesnew", "Times New Roman")]


def _a10_map_font(raw):
    raw = (raw or "").lower()
    for key, name in _A10_FONT_MAP:
        if key in raw:
            return name
    return None


def _a10_native_stats(page, bbox_pt):
    """Cỡ chữ / font / số dòng / bước dòng thật từ text layer trong vùng bbox (pt). None nếu không có."""
    try:
        x0, y0, x1, y1 = bbox_pt
        d = page.get_text("dict")
    except Exception:
        return None
    sizes, fonts, cys = Counter(), Counter(), []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            lb = ln["bbox"]
            cx, cy = (lb[0] + lb[2]) / 2, (lb[1] + lb[3]) / 2
            if not (x0 - 3 <= cx <= x1 + 3 and y0 - 3 <= cy <= y1 + 3):
                continue
            has = False
            for s in ln.get("spans", []):
                tx = s.get("text", "").strip()
                if tx:
                    has = True
                    sizes[round(s["size"] * 2) / 2] += len(tx)
                    fonts[_a10_map_font(s.get("font", ""))] += len(tx)
            if has:
                cys.append(cy)
    if not sizes:
        return None
    cys.sort()
    lines = [cys[0]]
    for c in cys[1:]:
        if c - lines[-1] > 2.0:
            lines.append(c)
    pitch = (lines[-1] - lines[0]) / (len(lines) - 1) if len(lines) > 1 else None
    fnt = [f for f, _ in fonts.most_common() if f]
    return {"size": sizes.most_common(1)[0][0], "font": fnt[0] if fnt else None,
            "lines": len(lines), "pitch": pitch}


# ------------------------------------------------------------------ cắt ảnh
def _a10_crop_bytes(img, bbox, w_pt, transparent=False, pad=0.002):
    y0, x0, y1, x1 = bbox
    W, H = img.size
    box = (int(_a10_clamp((x0 / 1000 - pad) * W, 0, W)), int(_a10_clamp((y0 / 1000 - pad) * H, 0, H)),
           int(_a10_clamp((x1 / 1000 + pad) * W, 0, W)), int(_a10_clamp((y1 / 1000 + pad) * H, 0, H)))
    if box[2] - box[0] < 2 or box[3] - box[1] < 2:
        return None
    crop = img.crop(box).convert("RGB")
    out = io.BytesIO()
    if transparent:
        import numpy as np
        from PIL import ImageFilter
        arr = np.asarray(crop).astype("int16")
        mx, mn = arr.max(axis=2), arr.min(axis=2)
        sat = (mx - mn) / np.maximum(mx, 1)
        nonwhite = mn < 225
        colored = nonwhite & (sat > 0.30)
        if colored.sum() > 50 and colored.sum() > 0.02 * nonwhite.sum():
            # con dấu / mực màu: chỉ giữ mực màu, loại chữ đen nằm bên dưới để không bị lặp chữ
            m = Image.fromarray((colored * 255).astype("uint8")).filter(ImageFilter.MaxFilter(3))
            keep = np.asarray(m) > 0
        else:
            keep = nonwhite                                  # chữ ký mực đen
        rgba = crop.convert("RGBA")
        alpha = Image.fromarray((keep * 255).astype("uint8"))
        rgba.putalpha(alpha)
        rgba.save(out, format="PNG")
    elif crop.size[0] * crop.size[1] > 350000:
        crop.save(out, format="JPEG", quality=90)
    else:
        crop.save(out, format="PNG")
    out.seek(0)
    return out


# ------------------------------------------------------------------ thao tác docx mức thấp
def _a10_mark_size(p, half_points):
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        sect = pPr.find(qn("w:sectPr"))
        if sect is not None:
            sect.addprevious(rPr)
        else:
            pPr.append(rPr)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(half_points))


def _a10_tiny(p, height_pt=1.0):
    """Biến đoạn thành dải đệm cao đúng height_pt (dùng làm khoảng cách / đoạn cuối trang)."""
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(_a10_clamp(height_pt, 1, 1500))
    _a10_mark_size(p, 2)


def _a10_set_run_font(run, name, size, bold=False, italic=False, underline=False, color=None):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), name)
    run.font.size = Pt(size)
    run.bold = bool(bold)
    run.italic = bool(italic)
    run.underline = bool(underline)
    if color is not None:
        run.font.color.rgb = color


def _a10_add_runs(par, runs, font, size, force_bold=False):
    for r in runs or []:
        if not isinstance(r, dict):
            r = {"t": str(r)}
        text = _a10_clean(r.get("t", ""))
        for k, part in enumerate(text.split("\n")):
            if k > 0:
                par.add_run().add_break()
            if part == "":
                continue
            run = par.add_run(part)
            _a10_set_run_font(run, font, size, bool(r.get("b")) or force_bold, bool(r.get("i")), bool(r.get("u")),
                              _a10_hex(r.get("c")))


_A10_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
              "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
_A10_VALIGN = {"top": WD_ALIGN_VERTICAL.TOP, "middle": WD_ALIGN_VERTICAL.CENTER, "center": WD_ALIGN_VERTICAL.CENTER,
               "bottom": WD_ALIGN_VERTICAL.BOTTOM}


def _a10_set_page(sec, w, h, ml, mr, mt, mb):
    if abs(w - 595.28) < 8 and abs(h - 841.89) < 8:
        sec.page_width, sec.page_height = Mm(210), Mm(297)
    elif abs(w - 841.89) < 8 and abs(h - 595.28) < 8:
        sec.page_width, sec.page_height = Mm(297), Mm(210)
    else:
        sec.page_width, sec.page_height = Pt(w), Pt(h)
    sec.orientation = WD_ORIENT.LANDSCAPE if w > h else WD_ORIENT.PORTRAIT
    sec.left_margin, sec.right_margin = Pt(ml), Pt(mr)
    sec.top_margin, sec.bottom_margin = Pt(mt), Pt(mb)
    sec.header_distance = Pt(0)
    sec.footer_distance = Pt(0)


def _a10_make_floating(run, img_stream, x_pt, y_pt, w_pt, h_pt, behind=True):
    """Chèn ảnh nổi (không đẩy chữ) tại toạ độ tuyệt đối so với trang."""
    shape = run.add_picture(img_stream, width=Pt(w_pt), height=Pt(h_pt))
    inline = shape._inline
    graphic = inline.find(qn("a:graphic"))
    docPr = inline.find(qn("wp:docPr"))
    anchor = parse_xml(
        '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" '
        'behindDoc="%d" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="%s" name="%s"/><wp:cNvGraphicFramePr/></wp:anchor>'
        % (nsdecls("wp", "a"), 1 if behind else 0, int(Pt(x_pt)), int(Pt(y_pt)), int(Pt(w_pt)), int(Pt(h_pt)),
           docPr.get("id"), docPr.get("name")))
    anchor.append(graphic)
    drawing = inline.getparent()
    drawing.remove(inline)
    drawing.append(anchor)


def _a10_table_props(tbl, x_off_pt, width_pt, borders):
    tblPr = tbl._tbl.tblPr
    for tag in ("w:tblW", "w:tblInd", "w:tblLayout", "w:tblCellMar"):
        el = tblPr.find(qn(tag))
        if el is not None:
            tblPr.remove(el)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(width_pt * 20)))
    tblW.set(qn("w:type"), "dxa")
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), str(int(x_off_pt * 20)))
    ind.set(qn("w:type"), "dxa")
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", 15), ("left", 70), ("bottom", 15), ("right", 70)):
        e = OxmlElement("w:" + side)
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    # thứ tự schema: tblStyle, ..., tblW, jc, ..., tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook
    look = tblPr.find(qn("w:tblLook"))
    for el in (tblW, ind, lay, mar):
        if look is not None:
            look.addprevious(el)
        else:
            tblPr.append(el)
    if borders:
        tbl.style = "Table Grid"


def _a10_shade(cell, hexcolor):
    h = str(hexcolor).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
        return
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), h.upper())
    tcPr.append(shd)



def _a10_w_em(s, bold=False):
    """Ước lượng độ rộng chuỗi theo đơn vị em (serif, tiếng Việt)."""
    tot = 0.0
    for ch in s:
        if ch == " ":
            tot += 0.25
        elif ch.isupper():
            tot += 0.68
        elif ch.isdigit():
            tot += 0.5
        elif ch.isalpha():
            tot += 0.47
        else:
            tot += 0.33
    return tot * (1.04 if bold else 1.0)


def _a10_cell_fit_size(runs, fs, col_w, span_h):
    """Thu nhỏ nhẹ cỡ chữ trong ô (tối đa -20%) nếu chữ không vừa ô -> tránh tràn/xuống dòng thừa do khác phông."""
    plain = "".join(_a10_clean(r.get("t", "")) if isinstance(r, dict) else str(r) for r in (runs or []))
    bold = any(isinstance(r, dict) and r.get("b") for r in (runs or []))
    segs = plain.split("\n")
    avail = max(col_w - 8.0, 6.0)
    for sc in (1.0, 0.95, 0.9, 0.85, 0.8):
        f = fs * sc
        need = sum(max(1, math.ceil(_a10_w_em(sg, bold) * f / avail)) for sg in segs)
        cap = max(1, int(span_h / (f * 1.25) + 0.3))
        if need <= cap:
            return round(f * 2) / 2
    return round(fs * 0.8 * 2) / 2


# ------------------------------------------------------------------ dựng Word từ bản đồ bố cục
def a10_build_docx(pages):
    """
    pages: list dict(layout=dict|None, img=PIL.Image, w=pt, h=pt, native=fitz.Page|None, hint_ok=bool)
    Trả về (docx_bytes, danh_sách_cảnh_báo)
    """
    warnings = []
    fams = Counter((p["layout"] or {}).get("font_family", "serif") for p in pages if p.get("layout"))
    default_font = "Arial" if fams and fams.most_common(1)[0][0] == "sans" else "Times New Roman"

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = default_font
    normal.font.size = Pt(11)
    rf = normal.element.get_or_add_rPr().find(qn("w:rFonts"))
    if rf is not None:
        rf.set(qn("w:eastAsia"), default_font)
    npf = normal.paragraph_format
    npf.space_before = Pt(0)
    npf.space_after = Pt(0)
    npf.line_spacing = 1.0

    n_pages = len(pages)
    for pi, pg in enumerate(pages):
        sec = doc.sections[-1]
        w, h, img, layout, native = pg["w"], pg["h"], pg["img"], pg.get("layout"), pg.get("native")
        floats = []

        if not layout or not layout["blocks"]:
            # Dự phòng: trang không đọc được -> chèn nguyên ảnh trang để KHÔNG mất nội dung
            if layout is None:
                warnings.append(f"Trang {pi + 1}: AI không phản hồi được, đã chèn nguyên trang dạng ảnh.")
                _a10_set_page(sec, w, h, 0, 0, 0, 0)
                p = doc.add_paragraph()
                stream = io.BytesIO()
                img.convert("RGB").save(stream, format="JPEG", quality=92)
                stream.seek(0)
                p.add_run().add_picture(stream, width=Pt(w * 0.995))
            else:
                _a10_set_page(sec, w, h, 36, 36, 36, 36)
            host = None
        else:
            blocks = layout["blocks"]
            flow = [b for b in blocks if not (b["type"] == "image" and b.get("floating"))]
            P = lambda v, tot: v / 1000.0 * tot  # noqa: E731
            bbp = lambda b: (P(b["bbox"][1], w), P(b["bbox"][0], h), P(b["bbox"][3], w), P(b["bbox"][2], h))  # noqa: E731

            xs0 = [bbp(b)[0] for b in flow] or [36]
            xs1 = [bbp(b)[2] for b in flow] or [w - 36]
            ys0 = [bbp(b)[1] for b in flow] or [36]
            ys1 = [bbp(b)[3] for b in flow] or [h - 36]
            ml = _a10_clamp(min(xs0), 0, 0.4 * w)
            mr = _a10_clamp(w - max(xs1), 0, 0.4 * w)
            mt = _a10_clamp(min(ys0), 0, 0.35 * h)
            mb = _a10_clamp(h - max(ys1) - 6, 0, 54)
            _a10_set_page(sec, w, h, ml, mr, mt, mb)
            text_w = w - ml - mr
            flow_y = mt

            for blk in flow:
                x0, y0, x1, y1 = bbp(blk)
                gap_top = max(0.0, y0 - flow_y)

                # ---------------- ĐOẠN VĂN ----------------
                if blk["type"] == "paragraph":
                    runs = blk.get("runs") or []
                    plain = "".join(_a10_clean(r.get("t", "")) if isinstance(r, dict) else str(r) for r in runs)
                    if not plain.strip():
                        continue
                    n = max(1, int(_a10_num(blk.get("lines"), 1)))
                    nat = _a10_native_stats(native, (x0, y0, x1, y1)) if native is not None else None
                    fname = default_font
                    hb = y1 - y0
                    if nat:
                        n, fs, fname = nat["lines"], nat["size"], nat["font"] or default_font
                        pitch = nat["pitch"] or fs * 1.2
                        if not (fs * 1.0 <= pitch <= fs * 1.8):
                            pitch = fs * 1.2
                    else:
                        pitch = hb / (n - 0.15)
                        fs = pitch / 1.2
                        bold_any = any(isinstance(r, dict) and r.get("b") for r in runs)
                        segs = plain.split("\n")
                        bw = x1 - x0
                        if len(segs) > 1 or n == 1:      # mỗi dòng độc lập: dòng dài nhất phải lọt khung
                            wmax = max(_a10_w_em(sg, bold_any) for sg in segs) or 1
                            fs = min(fs, bw * 1.04 / wmax)
                        else:                            # đoạn tự xuống dòng: chặn trên theo tổng độ dài
                            wt = _a10_w_em(plain, bold_any) or 1
                            fs = min(fs, bw * n / wt)
                        fs = _a10_clamp(round(fs * 2) / 2, 6, 30)
                        pitch = _a10_clamp(pitch, fs * 1.1, fs * 1.6)
                    pitch = _a10_clamp(pitch, fs, fs * 2.5)
                    top_target = y0 - 0.075 * pitch
                    space_before = max(0.0, top_target - flow_y)

                    p = doc.add_paragraph()
                    pf = p.paragraph_format
                    pf.space_before = Pt(space_before)
                    pf.space_after = Pt(0)
                    al = str(blk.get("align", "left")).lower()
                    li, ri = max(0.0, x0 - ml), max(0.0, (w - mr) - x1)
                    if "\n" in plain or n == 1:         # dòng độc lập: nới khung để khỏi xuống dòng thừa
                        ext = max(8.0, 0.08 * (x1 - x0))
                        if al == "center":
                            li, ri = max(0.0, li - ext), max(0.0, ri - ext)
                        elif al == "right":
                            li = 0.0
                        else:
                            ri = 0.0
                    pf.left_indent = Pt(li)
                    pf.right_indent = Pt(ri)
                    fi = _a10_num(blk.get("first_indent_pt"), 0)
                    if fi > 1:
                        pf.first_line_indent = Pt(min(fi, 80))
                    pf.line_spacing = Pt(pitch)                      # chính xác -> giữ nguyên nhịp dòng
                    pf.alignment = _A10_ALIGN.get(str(blk.get("align", "left")).lower(), WD_ALIGN_PARAGRAPH.LEFT)
                    _a10_add_runs(p, runs, fname, fs)
                    flow_y = max(flow_y, top_target) + n * pitch
                    continue

                # ---------------- ẢNH NỘI TUYẾN ----------------
                if blk["type"] == "image":
                    stream = _a10_crop_bytes(img, blk["bbox"], x1 - x0)
                    if stream is None:
                        continue
                    wp = min(x1 - x0, text_w)
                    hp = (y1 - y0) * wp / max(x1 - x0, 1)
                    p = doc.add_paragraph()
                    pf = p.paragraph_format
                    pf.space_before = Pt(gap_top)
                    pf.space_after = Pt(0)
                    pf.left_indent = Pt(max(0.0, x0 - ml))
                    p.add_run().add_picture(stream, width=Pt(wp), height=Pt(hp))
                    flow_y = max(flow_y, y0) + hp + 2
                    continue

                # ---------------- BẢNG ----------------
                rows = blk["rows"]
                placements, occ, ncols = [], set(), 0
                for r_i, row in enumerate(rows):
                    c = 0
                    for cell in row:
                        while (r_i, c) in occ:
                            c += 1
                        cs = max(1, int(_a10_num(cell.get("colspan"), 1)))
                        rs = max(1, min(int(_a10_num(cell.get("rowspan"), 1)), len(rows) - r_i))
                        for rr in range(r_i, r_i + rs):
                            for cc in range(c, c + cs):
                                occ.add((rr, cc))
                        placements.append((r_i, c, cs, rs, cell))
                        c += cs
                        ncols = max(ncols, c)
                nrows = len(rows)

                tw = x1 - x0
                cp = blk.get("col_pct") or []
                if len(cp) != ncols or sum(_a10_num(v) for v in cp) <= 0:
                    cp = [100.0 / ncols] * ncols
                s = sum(_a10_num(v) for v in cp)
                col_w = [tw * _a10_num(v) / s for v in cp]
                rp = blk.get("row_pct") or []
                th = y1 - y0
                if len(rp) != nrows or sum(_a10_num(v) for v in rp) <= 0:
                    rp = [100.0 / nrows] * nrows
                s = sum(_a10_num(v) for v in rp)
                row_h = [th * _a10_num(v) / s for v in rp]

                nat = _a10_native_stats(native, (x0, y0, x1, y1)) if native is not None else None
                fs = nat["size"] if nat else _a10_clamp(_a10_num(blk.get("font_size_pt"), 11), 6, 18)
                fname = (nat["font"] if nat and nat["font"] else default_font)

                # dải đệm tạo khoảng cách phía trên bảng (bảng không có space_before)
                sp = doc.add_paragraph()
                _a10_tiny(sp, max(gap_top, 1.0))

                tbl = doc.add_table(rows=nrows, cols=ncols)
                tbl.autofit = False
                _a10_table_props(tbl, max(0.0, x0 - ml), tw, bool(blk.get("borders")))
                for ci, wv in enumerate(col_w):
                    tbl.columns[ci].width = Pt(wv)
                for ri, hv in enumerate(row_h):
                    tbl.rows[ri].height = Pt(max(hv, 4))
                    tbl.rows[ri].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

                for r_i, c_i, cs, rs, cell in placements:
                    if c_i >= ncols:
                        continue
                    cs = min(cs, ncols - c_i)
                    target = tbl.cell(r_i, c_i)
                    if cs > 1 or rs > 1:
                        target = target.merge(tbl.cell(r_i + rs - 1, c_i + cs - 1))
                    target.width = Pt(sum(col_w[c_i:c_i + cs]))
                    target.vertical_alignment = _A10_VALIGN.get(str(cell.get("valign", "top")).lower(),
                                                                WD_ALIGN_VERTICAL.TOP)
                    if cell.get("fill"):
                        _a10_shade(target, cell["fill"])
                    # xoá đoạn rỗng thừa sau khi gộp ô
                    while len(target.paragraphs) > 1 and not target.paragraphs[-1].text.strip():
                        target._tc.remove(target.paragraphs[-1]._p)
                    par = target.paragraphs[0]
                    par.paragraph_format.alignment = _A10_ALIGN.get(str(cell.get("align", "left")).lower(),
                                                                   WD_ALIGN_PARAGRAPH.LEFT)
                    span_h = sum(row_h[r_i:r_i + rs])
                    cfs = _a10_cell_fit_size(cell.get("runs"), fs, sum(col_w[c_i:c_i + cs]), span_h)
                    _a10_add_runs(par, cell.get("runs"), fname, cfs)
                    if cell.get("image"):
                        ib = cell["image"]
                        iw = P(ib[3] - ib[1], w)
                        ih = P(ib[2] - ib[0], h)
                        stream = _a10_crop_bytes(img, ib, iw)
                        if stream is not None:
                            iw2 = min(iw, sum(col_w[c_i:c_i + cs]) - 8)
                            ih2 = ih * iw2 / max(iw, 1)
                            ip = par if not par.text.strip() else target.add_paragraph()
                            ip.paragraph_format.alignment = par.paragraph_format.alignment
                            ip.add_run().add_picture(stream, width=Pt(iw2), height=Pt(ih2))
                flow_y = max(flow_y, y0) + th

            # ảnh nổi (con dấu / chữ ký đè lên chữ)
            for blk in blocks:
                if blk["type"] == "image" and blk.get("floating"):
                    x0, y0, x1, y1 = bbp(blk)
                    stream = _a10_crop_bytes(img, blk["bbox"], x1 - x0, transparent=True)
                    if stream is not None:
                        floats.append((stream, x0, y0, x1 - x0, y1 - y0))

        # đoạn cuối trang (chứa ngắt section + ảnh nổi)
        if pi < n_pages - 1:
            doc.add_section(WD_SECTION.NEW_PAGE)
            host = doc.paragraphs[-1]
        else:
            host = doc.add_paragraph()
        _a10_tiny(host, 1.0)
        for stream, fx, fy, fw, fh in floats:
            _a10_make_floating(host.add_run(), stream, fx, fy, fw, fh, behind=True)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), warnings


# ------------------------------------------------------------------ giao diện Streamlit
def app_pdf_image_to_word_advanced():
    import os
    import streamlit as st
    import fitz

    st.title("📑 PDF Scan / Ảnh sang Word (Giữ nguyên bố cục)")
    st.markdown(
        "Dùng **Gemini online** đọc từng trang (chữ, bảng, ô gộp, logo, mã QR, chữ ký, con dấu), "
        "rồi dựng lại file Word **đúng khổ giấy, lề, cỡ chữ, khoảng cách dòng, thụt lề và vị trí hình ảnh**. "
        "Dùng được cho **mọi loại văn bản** (công văn, quyết định, biểu mẫu, hợp đồng, báo cáo...)."
    )

    col1, col2 = st.columns(2)
    with col1:
        mode = st.radio("Chế độ xử lý:", ["🎯 Chính xác tối đa (khuyên dùng)", "⚡ Nhanh"], index=0)
    with col2:
        use_hint = st.checkbox("Dùng lớp chữ gốc của PDF (nếu có) để sửa lỗi nhận dạng", value=True)
        workers = st.slider("Số trang xử lý song song", 1, 4, 3)

    uploaded_file = st.file_uploader(
        "Tải lên file PDF Scan hoặc ảnh (PDF / PNG / JPG):",
        type=["pdf", "png", "jpg", "jpeg"],
        key=f"app10_{st.session_state.uploader_key}",
    )
    if not uploaded_file:
        return

    file_bytes = uploaded_file.getvalue()
    fname = uploaded_file.name
    st.info(f"📄 Đã nhận: **{fname}** ({len(file_bytes) / 1024:.1f} KB)")

    if not st.button("🚀 Bắt đầu chuyển đổi sang Word", type="primary"):
        return

    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        st.error("Chưa cấu hình GEMINI_API_KEY!")
        return

    models = A10_MODELS_PRECISE if mode.startswith("🎯") else A10_MODELS_FAST

    try:
        # 1) Dựng ảnh từng trang
        pages, pdf_doc = [], None
        if fname.lower().endswith(".pdf"):
            pdf_doc = fitz.open("pdf", file_bytes)
            if len(pdf_doc) == 0:
                st.error("File PDF không có trang nào.")
                return
            if len(pdf_doc) > 30:
                st.warning("File trên 30 trang - thời gian xử lý sẽ lâu và dễ chạm giới hạn API.")
            for i in range(len(pdf_doc)):
                pg = pdf_doc[i]
                pix = pg.get_pixmap(dpi=A10_RENDER_DPI)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                txt = pg.get_text("text").strip() if use_hint else ""
                pages.append({"img": img, "w": pg.rect.width, "h": pg.rect.height, "native": pg if txt else None,
                              "hint": txt if len(txt) > 20 else ""})
        else:
            img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            pages.append({"img": img, "w": img.size[0] * 72 / 200.0, "h": img.size[1] * 72 / 200.0,
                          "native": None, "hint": ""})

        # 2) Gửi Gemini (song song), có tiến trình
        def api_bytes(im):
            s = im.copy()
            s.thumbnail((A10_API_MAX_SIDE, A10_API_MAX_SIDE))
            b = io.BytesIO()
            s.save(b, format="JPEG", quality=92)
            return b.getvalue()

        def work(idx):
            data, used = _a10_call_gemini(api_key, api_bytes(pages[idx]["img"]), pages[idx]["hint"], models)
            return idx, data, used

        bar = st.progress(0.0, text="Đang gửi các trang lên Gemini...")
        notes, done = [], 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(work, i): i for i in range(len(pages))}
            for fu in as_completed(futs):
                i = futs[fu]
                try:
                    _, data, used = fu.result()
                    lay, dropped = _a10_sanitize(data)
                    pages[i]["layout"] = lay
                    if dropped:
                        notes.append(f"Trang {i + 1}: bỏ {dropped} khối dữ liệu lỗi từ AI.")
                except Exception as e:
                    pages[i]["layout"] = None
                    notes.append(f"Trang {i + 1}: {e}")
                done += 1
                bar.progress(done / len(pages), text=f"Đã xử lý {done}/{len(pages)} trang")

        # 3) Dựng Word
        with st.spinner("Đang dựng file Word..."):
            docx_bytes, warns = a10_build_docx(pages)
        bar.empty()

        st.success(f"🎉 Hoàn tất {len(pages)} trang. Hãy mở file và đối chiếu nhanh với bản gốc.")
        for m in notes + warns:
            st.warning(m)
        base = fname.rsplit(".", 1)[0]
        st.download_button(
            label="📥 Tải xuống File Word (.docx)",
            data=docx_bytes,
            file_name=f"{base}_Word.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            on_click=clear_file,
        )
    except Exception as e:
        st.error(f"Lỗi: {e}")


# ==========================================                        
# 5. KÍCH HOẠT ỨNG DỤNG THEO LỰA CHỌN MENU
# ==========================================
if "1. PDF sang Word" in app_mode:
    app_pdf_to_word()
elif "2. Chuyển PDF" in app_mode:
    app_number_2()
elif "3. PDF/Ảnh sang Excel" in app_mode:
    app_number_3()
elif "4. So sánh Văn bản" in app_mode:
    app_document_compare()
elif "5. Cắt & Ghép PDF" in app_mode:
    app_pdf_split_merge()
elif "6. Chuyên gia Công thức" in app_mode:
    app_excel_expert()
elif "7. Nén & Đóng Gói" in app_mode:
    app_compress_and_zip()
elif "8. Xoay Trang PDF" in app_mode:
    app_rotate_pdf()
elif "9. Xóa Siêu Dữ Liệu" in app_mode:
    app_metadata_cleaner()
elif "10. PDF Scan / Ảnh sang Word" in app_mode:
    app_pdf_image_to_word_advanced()
