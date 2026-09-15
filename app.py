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
        "💻 6. Chuyên gia Công thức & VBA"
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
    try:
        api_key_input = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
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
                    
                    # PROMPT ĐƯỢC TĂNG CƯỜNG LỆNH ÉP BUỘC MERGE_LEFT
                    prompt = """
                    NHIỆM VỤ OCR - BẮT BUỘC TUÂN THỦ NGHIÊM NGẶT CÁC QUY TẮC SAU:

                    1. THỂ THỨC VĂN BẢN (QUỐC HIỆU & CHỮ KÝ):
                       - Phần Quốc hiệu trên cùng (hoặc khối chữ ký ngang nhau): BẮT BUỘC dùng [HEADER_TABLE] (bảng tàng hình 2 cột) để chia tỷ lệ.
                       - TUYỆT ĐỐI KHÔNG vẽ bảng cho phần Footer (như số trang, mã tài liệu). Hãy xuất thành văn bản thường.

                    2. CANH LỀ ĐOẠN VĂN:
                       - Ghi [CENTER] ở đầu dòng cần canh giữa.
                       - Ghi [RIGHT] ở đầu dòng lệch phải.

                    3. ĐỊNH DẠNG CHỮ: 
                       - In đậm -> bọc trong ** (VD: **THÔNG BÁO**). 
                       - In nghiêng -> bọc trong * (VD: *Nơi nhận:*).
                       - Gạch chân -> bọc trong <u> và </u>.

                    4. BẢNG BIỂU PHỨC TẠP (QUAN TRỌNG NHẤT):
                       - Bắt buộc vẽ bằng Markdown (|...|).
                       - TẤT CẢ các dòng trong cùng 1 bảng PHẢI CÓ CÙNG SỐ LƯỢNG CỘT (số dấu | phải bằng nhau).
                       - XỬ LÝ Ô GỘP (MERGED CELLS):
                         + Gộp ngang (kéo sang phải): Ghi chữ vào ô trái cùng, các ô bị chiếm chỗ phía sau điền CHÍNH XÁC: `[MERGE_LEFT]`
                         + Gộp dọc (kéo xuống dưới): Ghi chữ vào ô trên cùng, các ô bị chiếm chỗ phía dưới điền CHÍNH XÁC: `[MERGE_UP]`
                       - ĐẶC BIỆT LƯU Ý VỚI DÒNG TIÊU ĐỀ NHÓM BÊN TRONG BẢNG:
                         + Các dòng phân loại (Ví dụ: I, II, III... và Tên nhóm) thường gộp hết tất cả các cột phía sau.
                         + BẠN BẮT BUỘC PHẢI DÙNG `[MERGE_LEFT]` ĐỂ LẤP ĐẦY TẤT CẢ CÁC CỘT TRỐNG ĐÓ.
                         + VÍ DỤ CHUẨN:
                           | I | Sản xuất thuốc lá theo nhãn hiệu... | [MERGE_LEFT] | [MERGE_LEFT] | [MERGE_LEFT] |
                           | 1 | Giấy chứng nhận... | x | | |
                       - TUYỆT ĐỐI KHÔNG để trống các ô gộp, nếu không bảng sẽ bị vỡ nát. Dùng `<br>` để xuống dòng.

                    5. KHÔNG sinh ra mã phân trang. KHÔNG dùng mã LaTeX toán học.
                    """

                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
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
                    
                    # THUẬT TOÁN ĐÃ ĐƯỢC TỐI ƯU CƠ CHẾ GỘP Ô THÔNG MINH
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

                        # BƯỚC 1: Đổ chữ vào ô (Né các ô chứa mã gộp ra để tránh in chữ [MERGE_LEFT] ra Word)
                        for row_idx, row_data in enumerate(normalized_buffer):
                            row_cells = current_table.rows[row_idx].cells
                            
                            if is_header_table and num_cols == 2:
                                row_cells[0].width = Cm(6.0)
                                row_cells[1].width = Cm(10.0)

                            for col_idx, cell_data in enumerate(row_data):
                                cell_text = cell_data.strip()
                                is_merge = '[MERGE_LEFT]' in cell_text.upper() or '[MERGE_UP]' in cell_text.upper()
                                
                                cell = row_cells[col_idx]
                                cell.text = "" # Dọn sạch rác mặc định của ô
                                
                                if not is_merge:
                                    cell_lines = cell_text.split('<br>')
                                    for idx, c_line in enumerate(cell_lines):
                                        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
                                        p.paragraph_format.space_after = Pt(0)
                                        
                                        default_al = WD_ALIGN_PARAGRAPH.CENTER if (is_header_table or row_idx == 0) else WD_ALIGN_PARAGRAPH.LEFT
                                        clean_text = clean_tags_and_align(c_line.strip(), p, default_al)
                                        parse_and_add_runs(p, clean_text)
                                        
                        # BƯỚC 2: Truy vết và Nối ô (Root-finding Merge)
                        for row_idx in range(len(normalized_buffer)):
                            for col_idx in range(num_cols):
                                cell_text = normalized_buffer[row_idx][col_idx].strip().upper()
                                
                                if '[MERGE_LEFT]' in cell_text:
                                    # Tìm cột gốc không phải mã gộp bên trái
                                    root_col = col_idx - 1
                                    while root_col > 0 and '[MERGE_LEFT]' in normalized_buffer[row_idx][root_col].strip().upper():
                                        root_col -= 1
                                    if root_col >= 0:
                                        try:
                                            current_table.cell(row_idx, root_col).merge(current_table.cell(row_idx, col_idx))
                                        except:
                                            pass
                                            
                                elif '[MERGE_UP]' in cell_text:
                                    # Tìm dòng gốc không phải mã gộp phía trên
                                    root_row = row_idx - 1
                                    while root_row > 0 and '[MERGE_UP]' in normalized_buffer[root_row][col_idx].strip().upper():
                                        root_row -= 1
                                    if root_row >= 0:
                                        try:
                                            current_table.cell(root_row, col_idx).merge(current_table.cell(row_idx, col_idx))
                                        except:
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
        st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
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

                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
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
                    if raw_text.endswith("
