import streamlit as st
import os
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

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

st.sidebar.header("⚙️ Cấu hình")
api_key_input = st.sidebar.text_input("Nhập Google Gemini API Key của bạn:", type="password")
st.sidebar.info("💡 Mẹo: Lấy API Key miễn phí tại [Google AI Studio](https://aistudio.google.com/).")

st.sidebar.error("""
:red[**⚠️ NGUYÊN TẮC SỬ DỤNG:**]

:red[- **Bảo mật:** KHÔNG tải lên tài liệu MẬT, TỐI MẬT, dữ liệu tài chính chưa công khai và thông tin nhạy cảm của khách hàng.]

:red[- **Tối ưu:** Chỉ tải file PDF **dưới 30 trang/lần** để file Word không bị lỗi định dạng.]
""")

uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"])

if uploaded_file is not None:
    st.success(f"Đã tải lên file: **{uploaded_file.name}**")
    
    if st.button("🚀 Bắt đầu Chuyển đổi", type="primary"):
        if not api_key_input:
            st.error("⚠️ Vui lòng nhập API Key ở thanh menu bên trái trước khi chuyển đổi!")
        else:
            with st.spinner("🤖 AI đang phân tích và vẽ lại bảng biểu, vui lòng đợi..."):
                try:
                    temp_input_path = f"temp_{uploaded_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)
                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()
                    mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    Bạn là một hệ thống OCR và số hóa tài liệu cấp cao. Nhiệm vụ của bạn là bóc tách toàn bộ nội dung từ ảnh/PDF sang định dạng văn bản thô (Clean Text/Markdown) để chuyển vào file Word.
                    
                    YÊU CẦU KỶ LUẬT THÉP (BẮT BUỘC TUÂN THỦ):
                    1. QUÉT SẠCH & CHÍNH XÁC 100%: Quét từ trên xuống dưới, không bỏ sót bất kỳ ký tự, con số, mã vạch nào ở góc/lề (VD: số Serial). Sao chép chính xác nguyên bản, bắt buộc giữ nguyên cả lỗi sai chính tả.
                    2. KHÔNG TỰ BỊA DỮ LIỆU: Chỉ sử dụng khoảng trắng (Space/Tab). TUYỆT ĐỐI KHÔNG tự ý chèn thêm dấu ba chấm (...).
                    3. XỬ LÝ CHỮ KÝ: Tự động bỏ qua hình mờ và con dấu đỏ. Tại vị trí có chữ ký tay, chỉ cần ghi chú chữ: [Đã ký].
                    4. BẢNG BIỂU (QUAN TRỌNG): Bất cứ khi nào tài liệu gốc là một bảng có kẻ khung (ví dụ báo cáo Total | Prints), BẮT BUỘC phải trình bày bằng cú pháp bảng Markdown chuẩn (có dấu | ở đầu và cuối mỗi dòng). Không được gộp thành văn bản thường.
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
                    
                    doc.add_heading('Kết quả trích xuất từ AI', level=1)
                    
                    in_table = False
                    current_table = None

                    for line in response.text.split('\n'):
                        line_stripped = line.strip()

                        if line_stripped.startswith('---') and not line_stripped.startswith('|'):
                            continue

                        # Phát hiện Bảng Markdown
                        if line_stripped.startswith('|') and line_stripped.endswith('|'):
                            # ĐÃ VÁ LỖI: Lọc sạch khoảng trắng để bắt chết dòng gạch ngang
                            check_line = line_stripped.replace(' ', '').replace(':', '')
                            if check_line.startswith('|---'): 
                                continue
                            
                            cells_data = [cell.strip() for cell in line_stripped.split('|')][1:-1]
                            
                            if not in_table:
                                current_table = doc.add_table(rows=1, cols=len(cells_data))
                                current_table.style = 'Table Grid'
                                hdr_cells = current_table.rows[0].cells
                                for i, data in enumerate(cells_data):
                                    if i < len(hdr_cells):
                                        hdr_cells[i].text = data
                                in_table = True
                            else:
                                row_cells = current_table.add_row().cells
                                for i, data in enumerate(cells_data):
                                    if i < len(row_cells):
                                        row_cells[i].text = data
                        else:
                            in_table = False 
                            
                            if line_stripped.startswith('#'):
                                doc.add_heading(line_stripped.replace('#', '').strip(), level=2)
                            elif line_stripped:
                                p = doc.add_paragraph(line_stripped)
                                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

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
