import streamlit as st
import os
import re
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

# Ép mở thanh bên và ẩn mũi tên
st.set_page_config(page_title="Chuyển PDF/Ảnh sang Word", page_icon="📄", layout="centered", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
        [data-testid="collapsedControl"] { display: none; }
        [data-testid="stSidebarCollapseButton"] { display: none; }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
st.markdown("Sử dụng **Google Gemini AI** để trích xuất văn bản và bảng biểu với độ chính xác cao.")

try:
    api_key_input = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
    st.stop()

st.sidebar.error("""
:red[**⚠️ NGUYÊN TẮC SỬ DỤNG:**]

:red[- **Bảo mật:** KHÔNG tải lên tài liệu MẬT, TỐI MẬT, dữ liệu tài chính chưa công khai và thông tin nhạy cảm của khách hàng.]

:red[- **Tối ưu:** Chỉ tải file PDF **dưới 30 trang/lần** để file Word không bị lỗi định dạng.]
""")

uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"])

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
                
                # PROMPT ĐÃ ĐƯỢC TÁI CẤU TRÚC: ÉP BUỘC LÀM THEO THỨ TỰ
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

                # --- NHẬN DIỆN VÀ XOAY TRANG GIẤY MỘT CÁCH CHẮC CHẮN ---
                section = doc.sections[0]
                is_landscape = False
                
                if "[ORIENTATION: LANDSCAPE]" in response_text:
                    is_landscape = True
                    response_text = response_text.replace("[ORIENTATION: LANDSCAPE]", "").strip()
                
                if "[ORIENTATION: PORTRAIT]" in response_text:
                    is_landscape = False
                    response_text = response_text.replace("[ORIENTATION: PORTRAIT]", "").strip()

                # Áp dụng xoay
                if is_landscape:
                    # Nếu giấy đang dọc, thì xoay ngang
                    if section.page_height > section.page_width:
                        new_width, new_height = section.page_height, section.page_width
                        section.orientation = WD_ORIENT.LANDSCAPE
                        section.page_width = new_width
                        section.page_height = new_height
                else:
                    # Nếu yêu cầu dọc nhưng lỡ đang ngang, thì xoay dọc lại
                    if section.page_width > section.page_height:
                        new_width, new_height = section.page_height, section.page_width
                        section.orientation = WD_ORIENT.PORTRAIT
                        section.page_width = new_width
                        section.page_height = new_height
                
                table_buffer = []

                for line in response_text.split('\n'):
                    clean_line = re.sub(r'<[^>]+>', '', line)
                    line_stripped = clean_line.strip()

                    # Bỏ qua dòng rỗng hoặc các lệnh dư thừa
                    if not line_stripped or line_stripped == '```markdown' or line_stripped == '```':
                        continue

                    # XỬ LÝ BẢNG BIỂU
                    if line_stripped.startswith('|') and line_stripped.endswith('|'):
                        check_line = line_stripped.replace(' ', '').replace(':', '')
                        if check_line.startswith('|---'): 
                            continue
                        
                        cells_data = [cell.strip() for cell in line_stripped.split('|')][1:-1]
                        table_buffer.append(cells_data)
                    else:
                        if table_buffer:
                            if all(cell == '' for cell in table_buffer[0]):
                                table_buffer.pop(0)
                                
                            if table_buffer:
                                num_cols = max(len(row) for row in table_buffer)
                                current_table = doc.add_table(rows=len(table_buffer), cols=num_cols)
                                current_table.style = 'Table Grid'
                                
                                for row_idx, row_data in enumerate(table_buffer):
                                    row_cells = current_table.rows[row_idx].cells
                                    for col_idx, cell_data in enumerate(row_data):
                                        if col_idx < len(row_cells):
                                            cell = row_cells[col_idx]
                                            cell.text = "" 
                                            p = cell.paragraphs[0]
                                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                                            
                                            parts = re.split(r'\*\*(.*?)\*\*', cell_data)
                                            for i, part in enumerate(parts):
                                                if part:
                                                    run = p.add_run(part)
                                                    if i % 2 == 1:
                                                        run.bold = True
                            
                            table_buffer = []
                        
                        if line_stripped.startswith('---'):
                            continue
                            
                        # XỬ LÝ CANH LỀ ĐOẠN VĂN
                        align = WD_ALIGN_PARAGRAPH.LEFT 
                        
                        if line_stripped.startswith('[CENTER]'):
                            align = WD_ALIGN_PARAGRAPH.CENTER
                            line_stripped = line_stripped.replace('[CENTER]', '').strip()
                        elif line_stripped.startswith('[RIGHT]'):
                            align = WD_ALIGN_PARAGRAPH.RIGHT
                            line_stripped = line_stripped.replace('[RIGHT]', '').strip()
                        
                        if line_stripped:
                            p = doc.add_paragraph()
                            p.alignment = align
                            
                            # XỬ LÝ IN ĐẬM
                            parts = re.split(r'\*\*(.*?)\*\*', line_stripped)
                            for i, part in enumerate(parts):
                                if part:
                                    run = p.add_run(part)
                                    if i % 2 == 1:
                                        run.bold = True

                # Vẽ nốt bảng nếu bảng nằm ở cuối
                if table_buffer:
                    if all(cell == '' for cell in table_buffer[0]):
                        table_buffer.pop(0)
                    if table_buffer:
                        num_cols = max(len(row) for row in table_buffer)
                        current_table = doc.add_table(rows=len(table_buffer), cols=num_cols)
                        current_table.style = 'Table Grid'
                        for row_idx, row_data in enumerate(table_buffer):
                            row_cells = current_table.rows[row_idx].cells
                            for col_idx, cell_data in enumerate(row_data):
                                if col_idx < len(row_cells):
                                    cell = row_cells[col_idx]
                                    cell.text = "" 
                                    p = cell.paragraphs[0]
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                                    parts = re.split(r'\*\*(.*?)\*\*', cell_data)
                                    for i, part in enumerate(parts):
                                        if part:
                                            run = p.add_run(part)
                                            if i % 2 == 1:
                                                run.bold = True

                output_docx_path = "ket_qua.docx"
                doc.save(output_docx_path)

                st.success("🎉 Chuyển đổi và định dạng chuẩn thành công!")

                with open(output_docx_path, "rb") as file_download:
                    st.download_button(
                        label="📥 Tải xuống file Word (.docx)",
                        data=file_download,
                        file_name=f"Converted_{uploaded_file.name.split('.')[0]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")
