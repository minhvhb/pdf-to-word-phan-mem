import streamlit as st
import os
import re
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

# ==========================================
# 1. CẤU HÌNH TRANG CHỦ ĐẠO
# ==========================================
st.set_page_config(page_title="Hệ thống Công cụ AI", page_icon="⚙️", layout="centered", initial_sidebar_state="expanded")
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
# 2. KHU VỰC APP 1: CHUYỂN PDF SANG WORD
# ==========================================
def app_pdf_to_word():
    try:
        api_key_input = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
        st.stop()

    st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
    st.markdown("Sử dụng **Google Gemini AI** để trích xuất văn bản và bảng biểu với độ chính xác cao.")

    uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"], key=f"app1_{st.session_state.uploader_key}")

    if uploaded_file is not None:
        st.success(f"Đã tải lên file: **{uploaded_file.name}**")
        
        if st.button("🚀 Bắt đầu Chuyển đổi", type="primary"):
            with st.spinner("🤖 AI đang phân tích hướng giấy và trích xuất dữ liệu, vui lòng đợi..."):
                try:
                    temp_input_path = f"temp_{uploaded_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    NHIỆM VỤ OCR - BẮT BUỘC TUÂN THỦ NGHIÊM NGẶT THEO THỨ TỰ SAU:

                    1. CHIỀU TRANG GIẤY (ĐIỀU KIỆN TIÊN QUYẾT):
                       - Phân tích bức ảnh. Nếu bề ngang rộng hơn bề dọc, DÒNG ĐẦU TIÊN CỦA BẠN PHẢI LÀ: [ORIENTATION: LANDSCAPE]. 
                       - Nếu bề dọc dài hơn bề ngang, DÒNG ĐẦU TIÊN CỦA BẠN PHẢI LÀ: [ORIENTATION: PORTRAIT].

                    2. CANH LỀ ĐOẠN VĂN:
                       - Tiêu đề hoặc chữ canh giữa trang -> ghi [CENTER] ở đầu dòng.
                       - Chữ nằm lệch góc phải -> ghi [RIGHT] ở đầu dòng.
                       - (KHÔNG áp dụng lệnh [CENTER] hay [RIGHT] vào bên trong bảng biểu).

                    3. ĐỊNH DẠNG CHỮ: Chữ nào in đậm trong bản gốc, phải bọc bằng dấu sao kép (Ví dụ: **DANH SÁCH**).

                    4. BẢNG BIỂU: Vẽ bảng bằng cú pháp Markdown chuẩn (|...|).
                    
                    5. TUYỆT ĐỐI KHÔNG dùng HTML, không bịa dấu ba chấm. Chữ ký tay thay bằng [Đã ký].
                    """

                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    doc = Document()
                    style = doc.styles['Normal']
                    font = style.font
                    font.name = 'Times New Roman'
                    font.size = Pt(12)
                    
                    response_text = response.text

                    section = doc.sections[0]
                    is_landscape = False
                    
                    if "[ORIENTATION: LANDSCAPE]" in response_text:
                        is_landscape = True
                        response_text = response_text.replace("[ORIENTATION: LANDSCAPE]", "").strip()
                    
                    if "[ORIENTATION: PORTRAIT]" in response_text:
                        is_landscape = False
                        response_text = response_text.replace("[ORIENTATION: PORTRAIT]", "").strip()

                    if is_landscape:
                        if section.page_height > section.page_width:
                            new_width, new_height = section.page_height, section.page_width
                            section.orientation = WD_ORIENT.LANDSCAPE
                            section.page_width = new_width
                            section.page_height = new_height
                    else:
                        if section.page_width > section.page_height:
                            new_width, new_height = section.page_height, section.page_width
                            section.orientation = WD_ORIENT.PORTRAIT
                            section.page_width = new_width
                            section.page_height = new_height
                    
                    table_buffer = []

                    for line in response_text.split('\n'):
                        clean_line = re.sub(r'<[^>]+>', '', line)
                        line_stripped = clean_line.strip()

                        if not line_stripped or line_stripped == '```markdown' or line_stripped == '
