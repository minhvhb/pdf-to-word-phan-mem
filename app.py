import streamlit as st
import os
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ĐÃ THÊM LỆNH ÉP MỞ THANH BÊN (initial_sidebar_state="expanded")
st.set_page_config(page_title="Chuyển PDF/Ảnh sang Word", page_icon="📄", layout="centered", initial_sidebar_state="expanded")

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
            with st.spinner("🤖 AI đang phân tích và định dạng file Word chuẩn hành chính, vui lòng đợi..."):
                try:
                    temp_input_path = f"temp_{uploaded_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)

                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()

                    mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    Bạn là một hệ thống OCR cao cấp chuyên trích xuất và khôi phục định dạng tài liệu. Nhiệm vụ của bạn là đọc toàn bộ nội dung từ file PDF/Hình ảnh được cung cấp và tái tạo lại thành văn bản định dạng Markdown/HTML với mục tiêu GIỮ NGUYÊN 100% CẤU TRÚC VÀ ĐỊNH DẠNG BAN ĐẦU.

                    Hãy tuân thủ nghiêm ngặt các quy tắc sau:
                    
                    1. NGUYÊN TẮC BẢO TOÀN NỘI DUNG:
                       - Trích xuất chính xác 100% ký tự, số liệu, dấu tiếng Việt và ký tự đặc biệt.
                       - KHÔNG tự ý sửa lỗi chính tả, KHÔNG tóm tắt, KHÔNG thêm/bớt từ ngữ.
                       - Giữ nguyên cả phần tiêu đề trang (Header), chân trang (Footer), số trang và ghi chú (Footnotes).           
                    2. CẤU TRÚC & HỆ THỐNG TIÊU ĐỀ:
                       - Sử dụng thẻ Markdown (#, ##, ###, ####) tương ứng chính xác với kích thước và cấp độ phân cấp của Tiêu đề trong file gốc.
                       - Giữ nguyên sự phân đoạn văn bản, ngắt dòng và thứ tự trước/sau.     
                    3. ĐỊNH DẠNG VĂN BẢN CHI TIẾT:
                       - In đậm: Sử dụng **văn bản**
                       - In nghiêng: Sử dụng *văn bản*
                       - Gạch chân: Sử dụng <u>văn bản</u>
                       - Danh sách: Giữ nguyên các thụt lùi đầu dòng (indentation) đối với danh sách dạng đầu dòng (-) hoặc dạng số (1., a., i.).
                    4. XỬ LÝ BẢNG BIỂU (TỐI QUAN TRỌNG):
                       - Chuyển toàn bộ bảng biểu sang dạng Markdown Table chuẩn (| Cột 1 | Cột 2 |).
                       - Với các bảng phức tạp có ô gộp (Merged Cells/Span): BẮT BUỘC sử dụng thẻ HTML Table (`<table>`, `<tr>`, `<td colspan="...">`, `<td rowspan="...">`) để đảm bảo không lệch cột khi dựng lại file Word.
                       - Giữ nguyên toàn bộ số liệu và căn lề trong từng ô của bảng.
                    5. PHẦN TỬ ĐỒ HỌA & HÌNH ẢNH:
                       - Tại vị trí có hình ảnh, sơ đồ hoặc logo, chèn ký hiệu thay thế theo định dạng: [HÌNH ẢNH: Mô tả ngắn nội dung hình ảnh/sơ đồ].
                    6. ĐẦU RA (OUTPUT REQUIREMENT):
                       - Chỉ xuất ra duy nhất mã văn bản (Markdown/HTML). KHÔNG kèm lời chào, KHÔNG giải thích, KHÔNG viết bất kỳ câu dẫn dắt nào ở đầu hoặc cuối phản hồi.
                                           [VAI TRÒ & MỤC TIÊU]							
                    Bạn là một chuyên gia OCR (Chuyển đổi ký tự quang học) và Tái tạo Cấu trúc Tài liệu cấp cao.							
                    Nhiệm vụ của bạn là đọc hình ảnh/tài liệu PDF được gửi kèm và trích xuất TOÀN BỘ nội dung sang định dạng Markdown chuẩn xác 100% so với bản gốc, sẵn sàng để chuyển đổi trực tiếp sang file Microsoft Word (.docx).							
                    							
                    [QUY TẮC NỘI DUNG - KHÔNG ĐƯỢC VI PHẠM]							
                    1. NGUYÊN VĂN 100%: Trích xuất chính xác từng từ, từng câu, số liệu, ký tự đặc biệt, dấu câu và tiếng Việt (kể cả dấu thanh). KHÔNG tự ý tóm tắt, KHÔNG bỏ sót văn bản, KHÔNG thêm lời giải thích hay hội thoại.							
                    2. CHÍNH TẢ: Giữ đúng các lỗi chính tả nếu đó là văn bản gốc, trừ khi ký tự bị nhòe/mờ thì hãy khôi phục dựa trên ngữ cảnh chuẩn xác nhất.							
                    							
                    [CẤU TRÚC VÀ ĐỊNH DẠNG HÌNH THỨC]							
                    Hãy tái tạo cấu trúc trực quan của trang bằng các quy tắc Markdown sau:							
                    							
                    1. TIÊU ĐỀ (HEADINGS):							
                    - Phân cấp tiêu đề rõ ràng dựa theo kích thước chữ và độ đậm: `#` cho Tiêu đề chính (H1), `##` cho Tiêu đề phụ (H2), `###` cho Tiêu đề nhỏ (H3).							
                    2. ĐỊNH DẠNG VĂN BẢN (TEXT FORMATTING):							
                    - In đậm: Dùng `**văn bản**` cho chữ in đậm.							
                    - In nghiêng: Dùng `*văn bản*` cho chữ in nghiêng.							
                    - Gạch chân: Dùng `<u>văn bản</u>` cho chữ gạch chân.							
                    - Chữ gạch ngang: Dùng `~~văn bản~~`.							
                    3. DANH SÁCH (LISTS):							
                    - Giữ nguyên thụt lùi đầu dòng và thứ tự danh sách (dùng `-` cho danh sách không thứ tự, `1.`, `2.` cho danh sách có thứ tự).							
                    - Bảo toàn đúng các cấp độ danh sách lùi vào trong (nested lists).							
                    							
                    [XỬ LÝ BẢNG BIỂU VÀ PHẦN TỬ ĐẶC BIỆT]							
                    1. BẢNG BIỂU (TABLES):							
                    - Chuyển toàn bộ bảng biểu sang dạng Markdown Table (`| Header 1 | Header 2 |`).							
                    - Giữ đúng cấu trúc hàng và cột. Đối với các ô bị hợp nhất (merge cells), hãy lặp lại nội dung hoặc dùng định dạng HTML Table (`<table>`) nếu bảng quá phức tạp để đảm bảo khi sang Word không bị tràn/lệch cột.							
                    2. HÌNH ẢNH / SƠ ĐỒ / CHỮ KÝ:							
                    - Nếu có hình ảnh, biểu đồ hoặc con dấu, hãy đặt một thẻ thay thế theo vị trí tương ứng: `[HÌNH ẢNH: Mô tả ngắn về ảnh]` hoặc `[CON DẤU: Mô tả con dấu]`.							
                    3. CÔNG THỨC TOÁN HỌC / HÓA HỌC:							
                    - Đặt các công thức nằm trong định dạng LaTeX chuẩn: `$công_thức$` (cho inline) hoặc `$$công_thức$$` (cho dạng khối).							
                    4. NGẮT TRANG (PAGE BREAKS):							
                    - Nếu tài liệu có nhiều trang, hãy phân tách giữa các trang bằng đường gạch ngang `---`.							
                    							
                    [ĐẦU RA YÊU CẦU]							
                    Chỉ trả về DUY NHẤT mã Markdown đã được định dạng. Không viết bất kỳ câu mở đầu ("Đây là kết quả..."), không viết lời kết.							
                    """

                    # Gọi API Gemini 3.7 lấy kết quả
                    response = client.models.generate_content(
                        model='gemini-3.7-flash',
                        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    # Chuyển đổi mã Markdown/HTML thu được thành file Word thực sự
                    output_docx_path = "ket_qua.docx"
                    markdown_to_docx(response.text, output_docx_path)
                    st.success("🎉 Đã tạo bảng và định dạng Word chuẩn thành công!")
                    doc = Document()
                    
                    style = doc.styles['Normal']
                    font = style.font
                    font.name = 'Times New Roman'
                    font.size = Pt(12)
                    
                    doc.add_heading('Kết quả trích xuất từ AI', level=1)
                    
                    for line in response.text.split('\n'):
                        if line.strip().startswith('---'):
                            continue
                            
                        if line.strip().startswith('#'):
                            doc.add_heading(line.replace('#', '').strip(), level=2)
                        else:
                            p = doc.add_paragraph(line)
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
