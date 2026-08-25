import streamlit as st
import os
import time
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from bs4 import BeautifulSoup
import markdown2

# ==========================================
# 1. HÀM CHUYỂN ĐỔI MARKDOWN SANG WORD
# ==========================================
def markdown_to_docx(markdown_text, output_path="ket_qua.docx"):
    doc = Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    html_content = markdown2.markdown(markdown_text, extras=['tables'])
    soup = BeautifulSoup(html_content, 'html.parser')

    for element in soup.children:
        if element.name is None:
            continue
            
        if element.name in ['h1', 'h2', 'h3']:
            level = int(element.name[1])
            doc.add_heading(element.get_text(), level=level)

        elif element.name == 'table':
            rows = element.find_all('tr')
            if not rows:
                continue
                
            num_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
            table = doc.add_table(rows=0, cols=num_cols)
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            for row in rows:
                row_cells = table.add_row().cells
                cells_data = row.find_all(['td', 'th'])
                
                for idx, cell in enumerate(cells_data):
                    if idx < num_cols:
                        cell_p = row_cells[idx].paragraphs[0]
                        cell_p.text = cell.get_text().strip()
                        
                        if cell.name == 'th' or cell.parent.name == 'thead':
                            for r in cell_p.runs:
                                r.bold = True
            doc.add_paragraph()

        elif element.name == 'p':
            text = element.get_text().strip()
            if text:
                p = doc.add_paragraph(text)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.save(output_path)
    return output_path

# ==========================================
# 2. GIAO DIỆN CHÍNH STREAMLIT
# ==========================================
st.set_page_config(page_title="Chuyển PDF/Ảnh sang Word", page_icon="📄", layout="centered")

st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
st.markdown("Sử dụng **Google Gemini AI** để trích xuất văn bản và bảng biểu với độ chính xác cao.")

st.sidebar.header("⚙️ Cấu hình")
# Thêm key duy nhất để tránh trùng lặp
api_key_input = st.sidebar.text_input("Nhập Google Gemini API Key của bạn:", type="password", key="main_api_key")
st.sidebar.info("💡 Mẹo: Lấy API Key miễn phí tại [Google AI Studio](https://aistudio.google.com/).")

# Thêm key duy nhất cho file uploader
uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"], key="main_file_uploader")

if uploaded_file is not None:
    st.success(f"Đã tải lên file: **{uploaded_file.name}**")
    
    if st.button("🚀 Bắt đầu Chuyển đổi", type="primary", key="main_convert_btn"):
        if not api_key_input:
            st.error("⚠️ Vui lòng nhập API Key ở thanh menu bên trái trước khi chuyển đổi!")
        else:
            with st.spinner("🤖 AI đang phân tích và định dạng file Word..."):
                try:
                    client = genai.Client(api_key=api_key_input)
                    file_bytes = uploaded_file.getvalue()
                    mime_type = uploaded_file.type
                    
                    prompt = """
                    Bạn là một hệ thống OCR cao cấp. Hãy đọc file được cung cấp và chuyển đổi sang dạng Markdown Table chuẩn (| Header 1 | Header 2 |) cho toàn bộ các bảng biểu.
                    KHÔNG giải thích, chỉ trả về duy nhất mã Markdown.
                    """

                    max_retries = 3
                    response = None
                    for attempt in range(max_retries):
                        try:
                            response = client.models.generate_content(
                                model='gemini-2.5-flash',
                                contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                            )
                            break
                        except Exception as api_err:
                            if "503" in str(api_err) and attempt < max_retries - 1:
                                time.sleep(4)
                            else:
                                raise api_err

                    if response and response.text:
                        output_docx_path = "ket_qua.docx"
                        markdown_to_docx(response.text, output_docx_path)

                        st.success("🎉 Chuyển đổi và tạo bảng Word thành công!")

                        with open(output_docx_path, "rb") as file_download:
                            st.download_button(
                                label="📥 Tải xuống file Word (.docx)",
                                data=file_download,
                                file_name=f"Converted_{uploaded_file.name.split('.')[0]}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="main_download_btn"
                            )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")
