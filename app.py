import streamlit as st
import os
import re
import pandas as pd
from io import BytesIO
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

                        if not line_stripped or line_stripped.startswith("```"):
                            continue

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
                                
                                parts = re.split(r'\*\*(.*?)\*\*', line_stripped)
                                for i, part in enumerate(parts):
                                    if part:
                                        run = p.add_run(part)
                                        if i % 2 == 1:
                                            run.bold = True

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
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            on_click=clear_file
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

# ==========================================
# 3. KHU VỰC APP 2: CHUYỂN PDF VỀ KHỔ A4
# ==========================================
def app_number_2():
    st.title("🖨️ Chuyển PDF về khổ A4")
    st.markdown("Xóa bỏ mọi khung ẩn của bản vẽ cũ, ép lại chính xác thành khổ A4 tiêu chuẩn.")

    uploaded_pdf = st.file_uploader("Tải lên bản vẽ PDF cần xử lý:", type=["pdf"], key=f"app2_{st.session_state.uploader_key}")

    if uploaded_pdf is not None:
        st.success(f"Đã tải lên file: **{uploaded_pdf.name}**")
        
        if st.button("📏 Chuyển thành A4", type="primary"):
            with st.spinner("Đang truy quét và ghi đè các khung viền ẩn..."):
                try:
                    from pypdf import PdfReader, PdfWriter, Transformation
                    
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

# ==========================================
# 4. KHU VỰC APP 3: BÓC TÁCH BẢNG BIỂU SANG EXCEL
# ==========================================
def app_number_3():
    try:
        api_key_input = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Quản trị viên!")
        st.stop()

    st.title("📊 Bóc tách PDF/Ảnh sang Excel")
    st.markdown("Trích xuất tự động các bảng biểu tài chính, sao kê, danh sách từ PDF hoặc Ảnh thành file **Excel (.xlsx)** sạch sẽ.")

    uploaded_excel_file = st.file_uploader("Tải lên tài liệu chứa bảng (Ảnh hoặc PDF):", type=["jpg", "jpeg", "png", "pdf"], key=f"app3_{st.session_state.uploader_key}")

    if uploaded_excel_file is not None:
        st.success(f"Đã tải lên file: **{uploaded_excel_file.name}**")
        
        if st.button("🚀 Trích xuất ra Excel", type="primary"):
            with st.spinner("🤖 AI đang quét cấu trúc bảng và bóc tách dữ liệu số liệu..."):
                try:
                    temp_input_path = f"temp_excel_{uploaded_excel_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_excel_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_excel_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    Bạn là một chuyên gia số hóa dữ liệu kế toán và hành chính. 
                    Nhiệm vụ của bạn là đọc toàn bộ bảng biểu có trong tài liệu này và xuất kết quả DUY NHẤT dưới dạng bảng Markdown chuẩn (có dấu | ở đầu và cuối mỗi dòng).
                    
                    YÊU CẦU QUAN TRỌNG:
                    1. Giữ nguyên vẹn các con số, ký hiệu tiền tệ, ngày tháng, tên riêng. Không tự ý làm tròn hoặc bịa thêm số liệu.
                    2. Nếu tài liệu có nhiều bảng, hãy bóc tách và phân tách chúng rõ ràng.
                    3. KHÔNG trả về các đoạn văn bản dài dòng ngoài lề, chỉ tập trung vào cấu trúc bảng dữ liệu Markdown.
                    """

                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    lines = response.text.split('\n')
                    table_rows = []
                    
                    for line in lines:
                        line_str = line.strip()
                        if line_str.startswith('|') and line_str.endswith('|'):
                            if re.match(r'^\|[\s\-:]+\|$', line_str.replace(' ', '')):
                                continue
                            cells = [cell.strip() for cell in line_str.split('|')][1:-1]
                            table_rows.append(cells)

                    if len(table_rows) > 1:
                        headers = table_rows[0]
                        data = table_rows[1:]
                        
                        df = pd.DataFrame(data, columns=headers if len(headers) == len(data[0]) else None)
                        
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Sheet1')
                        processed_data = output.getvalue()

                        st.success("🎉 Bóc tách dữ liệu thành công!")

                        st.download_button(
                            label="📥 Tải xuống file Excel (.xlsx)",
                            data=processed_data,
                            file_name=f"Data_{uploaded_excel_file.name.split('.')[0]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            on_click=clear_file
                        )
                    else:
                        st.warning("⚠️ Không tìm thấy bảng dữ liệu rõ ràng trong tài liệu này. Vui lòng thử lại với file có khung bảng sắc nét hơn.")

                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")

# ==========================================
# 5. THANH MENU BÊN TRÁI ĐIỀU HƯỚNG CÁC APP
# ==========================================
st.sidebar.title("📌 Menu Công Cụ")

app_mode = st.sidebar.radio(
    "Vui lòng chọn ứng dụng:",
    ["📄 1. PDF sang Word", "🖨️ 2. Chuyển PDF về khổ A4", "📊 3. PDF/Ảnh sang Excel"]
)

st.sidebar.markdown("---") 

st.sidebar.error("""
:red[**⚠️ NGUYÊN TẮC SỬ DỤNG:**]

:red[- **Bảo mật:** KHÔNG tải lên tài liệu MẬT, TỐI MẬT, dữ liệu tài chính chưa công khai và thông tin nhạy cảm của khách hàng.]

:red[- **Tối ưu:** Chỉ tải file PDF **dưới 30 trang/lần**.]
""")

# ==========================================
# 6. KÍCH HOẠT ỨNG DỤNG DỰA TRÊN LỰA CHỌN
# ==========================================
if app_mode == "📄 1. PDF sang Word":
    app_pdf_to_word()
elif app_mode == "🖨️ 2. Chuyển PDF về khổ A4":
    app_number_2()
elif app_mode == "📊 3. PDF/Ảnh sang Excel":
    app_number_3()
