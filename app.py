import streamlit as st
import os
from google import genai
from google.genai import types
from docx import Document

st.set_page_config(page_title="Chuyển PDF/Ảnh sang Word", page_icon="📄", layout="centered")
st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
st.markdown("Sử dụng **Google Gemini AI** để trích xuất văn bản và bảng biểu với độ chính xác cao.")

st.sidebar.header("⚙️ Cấu hình")
api_key_input = st.sidebar.text_input("Nhập Google Gemini API Key của bạn:", type="password")
st.sidebar.markdown("---")
st.sidebar.info("💡 Mẹo: Lấy API Key miễn phí tại [Google AI Studio](https://aistudio.google.com/).")

uploaded_file = st.file_uploader("Tải lên file ảnh (JPG, PNG) hoặc PDF:", type=["jpg", "jpeg", "png", "pdf"])

if uploaded_file is not None:
    st.success(f"Đã tải lên file: **{uploaded_file.name}**")
    
    if st.button("🚀 Bắt đầu Chuyển đổi", type="primary"):
        if not api_key_input:
            st.error("⚠️ Vui lòng nhập API Key ở thanh menu bên trái trước khi chuyển đổi!")
        else:
            with st.spinner("🤖 AI đang phân tích tài liệu và tạo file Word, vui lòng đợi..."):
                try:
                    temp_input_path = f"temp_{uploaded_file.name}"
                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    client = genai.Client(api_key=api_key_input)

                    with open(temp_input_path, "rb") as f:
                        file_bytes = f.read()

                    mime_type = "application/pdf" if uploaded_file.name.endswith(".pdf") else "image/jpeg"
                    
                    prompt = """
                    Bạn là một chuyên gia số hóa tài liệu và chế bản điện tử. Nhiệm vụ của bạn là đọc hiểu hình ảnh/PDF được cung cấp và trích xuất toàn bộ nội dung sang văn bản Markdown sạch (Clean Markdown) để tối ưu hóa việc xuất ra định dạng Word.Hãy tuân thủ nghiêm ngặt 5 quy tắc sau:
                    1. Trích xuất nguyên bản: Giữ nguyên 100% nội dung chữ, tuyệt đối không tóm tắt, không diễn giải, không thêm thắt. Nếu có lỗi chính tả trong bản gốc, hãy giữ nguyên.
                    2. Cấu trúc & Thứ tự đọc: Nhận diện đúng luồng văn bản (ví dụ: đọc hết cột trái rồi mới sang cột phải). Phân định rõ ràng các cấp độ tiêu đề bằng cú pháp Markdown (Sử dụng # cho H1, ## cho H2, ### cho H3).
                    3. Xử lý Bảng biểu tuyệt đối: Thể hiện bảng bằng cú pháp Markdown chuẩn xác. Phải đối chiếu để không bỏ sót bất kỳ dòng hay cột nào. Nếu ô trong bản gốc bị trống, bắt buộc phải để trống ô tương ứng trong Markdown.
                    4. Loại bỏ "Rác" định dạng: Tự động nhận diện và BỎ QUA các chi tiết không thuộc nội dung chính như: số thứ tự trang, tiêu đề đầu trang/chân trang (header/footer) lặp lại, hình mờ (watermark), hoặc dấu mộc đỏ/chữ ký tay.
                    5. Toán học & Ký tự: Giữ nguyên các ký tự đặc biệt. Với công thức khoa học phức tạp, sử dụng cú pháp LaTeX chuẩn. Đảm bảo các đoạn văn không bị ngắt dòng vô lý giữa câu.
                    """

                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt]
                    )
                    
                    doc = Document()
                    doc.add_heading('Kết quả trích xuất từ AI', level=1)
                    
                    for line in response.text.split('\n'):
                        if line.strip().startswith('#'):
                            doc.add_heading(line.replace('#', '').strip(), level=2)
                        else:
                            doc.add_paragraph(line)

                    output_docx_path = "ket_qua.docx"
                    doc.save(output_docx_path)

                    st.success("🎉 Chuyển đổi thành công!")

                    with open(output_docx_path, "rb") as file_download:
                        st.download_button(
                            label="📥 Tải xuống file Word (.docx)",
                            data=file_download,
                            file_name=f"Converted_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")
