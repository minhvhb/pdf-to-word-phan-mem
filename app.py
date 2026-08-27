import streamlit as st
import os
import re
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT  # Thư viện để xoay giấy ngang/dọc

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
        with st.spinner("🤖 AI đang phân tích bố cục, chiều trang và vẽ bảng, vui lòng đợi..."):
            try:
                temp_input_path = f"temp_{uploaded_file.name}"
                with open(temp_input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                client = genai.Client(api_key=api_key_input)
                with open(temp_input_path, "rb") as f:
                    file_bytes = f.read()
                mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                
                # PROMPT NÂNG CẤP: DẠY AI NHẬN DIỆN CHIỀU TRANG VÀ CANH LỀ
                prompt = """
                Bạn là một hệ thống OCR và số hóa tài liệu cấp cao. Nhiệm vụ của bạn là bóc tách nội dung từ ảnh/PDF sang văn bản thô để chuyển vào Word.
                
                YÊU CẦU VỀ BỐ CỤC & ĐỊNH DẠNG (BẮT BUỘC):
                1. CHIỀU TRANG (Rất quan trọng): Nhìn tổng thể tài liệu. Nếu là giấy xoay ngang (chiều rộng lớn hơn chiều cao, ví dụ: danh sách, bảng chấm công), BẮT BUỘC ghi đúng 1 dòng này ở trên cùng: [ORIENTATION: LANDSCAPE]. Nếu xoay dọc, ghi: [ORIENTATION: PORTRAIT].
                2. CANH LỀ: 
                   - Nếu văn bản canh giữa (ví dụ: Tiêu đề), ghi thêm [CENTER] ở đầu dòng đó.
                   - Nếu canh phải (ví dụ: ngày tháng năm), ghi thêm [RIGHT] ở đầu dòng đó.
                3. IN ĐẬM: Bọc các chữ được in đậm trong bản gốc bằng dấu sao kép (Ví dụ: **Nội dung in đậm**).
                
                YÊU CẦU KỶ LUẬT THÉP (BẮT BUỘC TUÂN THỦ):
                1. QUÉT SẠCH & CHÍNH XÁC: Quét không bỏ sót. Giữ nguyên lỗi sai chính tả.
                2. BẢNG BIỂU: Trình bày bằng cú pháp bảng Markdown chuẩn. TUYỆT ĐỐI KHÔNG chèn tag [CENTER] hay [RIGHT] vào bên trong bảng Markdown để tránh lỗi.
                3. XỬ LÝ CHỮ KÝ: Tự động bỏ qua hình mờ và con dấu đỏ. Chữ ký tay ghi là: [Đã ký].
                4. KHÔNG DÙNG MÃ HTML/CSS. Không tự bịa dấu chấm/gạch ngang.
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

                # --- 1. NHẬN DIỆN VÀ XOAY TRANG GIẤY TỰ ĐỘNG ---
                section = doc.sections[0]
                if "[ORIENTATION: LANDSCAPE]" in response_text:
                    new_width, new_height = section.page_height, section.page_width
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width = new_width
                    section.page_height = new_height
                
                # Dọn dẹp mã mật khẩu xoay trang ra khỏi văn bản
                response_text = response_text.replace("[ORIENTATION: LANDSCAPE]", "").replace("[ORIENTATION: PORTRAIT]", "").strip()
                
                table_buffer = []

                for line in response_text.split('\n'):
                    clean_line = re.sub(r'<[^>]+>', '', line)
                    line_stripped = clean_line.strip()

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
                                            cell.text = "" # Xóa nội dung rỗng mặc định
                                            p = cell.paragraphs[0]
                                            
                                            # Trích xuất và bôi đậm chữ bên trong bảng
                                            parts = re.split(r'\*\*(.*?)\*\*', cell_data)
                                            for i, part in enumerate(parts):
                                                if part:
                                                    run = p.add_run(part)
                                                    if i % 2 == 1:
                                                        run.bold = True
                            
                            table_buffer = []
                        
                        if line_stripped.startswith('---'):
                            continue
                            
                        # --- 2. XỬ LÝ CANH LỀ ---
                        align = WD_ALIGN_PARAGRAPH.LEFT # Mặc định canh trái cho đẹp
                        
                        if line_stripped.startswith('[CENTER]'):
                            align = WD_ALIGN_PARAGRAPH.CENTER
                            line_stripped = line_stripped.replace('[CENTER]', '').strip()
                        elif line_stripped.startswith('[RIGHT]'):
                            align = WD_ALIGN_PARAGRAPH.RIGHT
                            line_stripped = line_stripped.replace('[RIGHT]', '').strip()
                        
                        if line_stripped:
                            p = doc.add_paragraph()
                            p.alignment = align
                            
                            # --- 3. XỬ LÝ IN ĐẬM ---
                            parts = re.split(r'\*\*(.*?)\*\*', line_stripped)
                            for i, part in enumerate(parts):
                                if part:
                                    run = p.add_run(part)
                                    if i % 2 == 1:
                                        run.bold = True

                # Vẽ nốt bảng nếu bảng nằm ở cuối file
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
