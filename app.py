import streamlit as st
import os
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Chuyển PDF/Ảnh sang Word", page_icon="📄", layout="centered")
st.title("📄 Ứng dụng Chuyển đổi PDF & Ảnh sang Word")
st.markdown("Sử dụng **Google Gemini AI** để trích xuất văn bản và bảng biểu với độ chính xác cao.")

st.sidebar.header("⚙️ Cấu hình")
api_key_input = st.sidebar.text_input("Nhập Google Gemini API Key của bạn:", type="password")

st.sidebar.info("💡 Mẹo: Lấy API Key miễn phí tại [Google AI Studio](https://aistudio.google.com/).")

# BẮT ĐẦU PHẦN CẢNH BÁO CHỮ ĐỎ
st.sidebar.error("""
:red[**⚠️ NGUYÊN TẮC SỬ DỤNG:**]

:red[- **Bảo mật:** KHÔNG tải lên tài liệu MẬT, TỐI MẬT, dữ liệu tài chính/khách hàng nhạy cảm.]

:red[- **Tối ưu:** Chỉ tải lên **dưới 30 trang/lần** để file Word không bị lỗi định dạng.]
""")
# KẾT THÚC PHẦN CẢNH BÁO

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
                    Bạn là một hệ thống OCR và số hóa tài liệu hành chính cấp cao. Nhiệm vụ của bạn là bóc tách toàn bộ nội dung từ ảnh/PDF sang định dạng văn bản thô (Clean Text/Markdown) để chuyển vào file Word.
                    YÊU CẦU KỶ LUẬT THÉP (BẮT BUỘC TUÂN THỦ):
                    1. Trích xuất nguyên bản: Giữ nguyên 100% nội dung chữ, tuyệt đối không tóm tắt, không diễn giải, không thêm thắt. Nếu có lỗi chính tả trong bản gốc, hãy giữ nguyên.
                    2. Cấu trúc & Thứ tự đọc: Nhận diện đúng luồng văn bản (ví dụ: đọc hết cột trái rồi mới sang cột phải). Phân định rõ ràng các cấp độ tiêu đề bằng cú pháp Markdown (Sử dụng # cho H1, ## cho H2, ### cho H3).
                    3. Xử lý Bảng biểu tuyệt đối: Thể hiện bảng bằng cú pháp Markdown chuẩn xác. Phải đối chiếu để không bỏ sót bất kỳ dòng hay cột nào. Nếu ô trong bản gốc bị trống, bắt buộc phải để trống ô tương ứng trong Markdown.
                    4. Loại bỏ "Rác" định dạng: Tự động nhận diện và BỎ QUA các chi tiết không thuộc nội dung chính như: số thứ tự trang, tiêu đề đầu trang/chân trang (header/footer) lặp lại, hình mờ (watermark), hoặc dấu mộc đỏ/chữ ký tay.
                    5. Toán học & Ký tự: Giữ nguyên các ký tự đặc biệt. Với công thức khoa học phức tạp, sử dụng cú pháp LaTeX chuẩn. Đảm bảo các đoạn văn không bị ngắt dòng vô lý giữa câu.
                    6. ĐỘ CHÍNH XÁC TUYỆT ĐỐI (Cực kỳ quan trọng): Sao chép chính xác 100% từng từ, từng chữ của bản gốc. Nếu bản gốc sai chính tả, BẮT BUỘC phải giữ nguyên lỗi sai đó (Ví dụ: "lao uộng" phải giữ là "lao uộng"). Tuyệt đối không tự ý sửa lỗi từ vựng, ngữ pháp hay thay đổi từ đồng nghĩa.
                    7. KHÔNG ĐIỀN CHỖ TRỐNG: Nếu bản gốc có khoảng trắng chờ điền (Ví dụ: "ngày ... tháng ... năm"), bắt buộc phải để trống hoặc dùng dấu ba chấm (...), tuyệt đối không tự ý bịa ngày tháng, tên người hay bất kỳ dữ liệu nào để điền vào.
                    8. XỬ LÝ CHỮ KÝ & CON DẤU: Bỏ qua các hình ảnh con dấu đỏ hoặc hình mờ. Tại vị trí có chữ ký tay, hãy ghi chú là: [Đã ký].
                    9. ĐỊNH DẠNG VÀ CẤU TRÚC: Giữ nguyên cấu trúc các cấp tiêu đề (dùng #, ##, ###). Trình bày bảng biểu bằng cú pháp Markdown. 
                    10. LỌC NHIỄU: Không chèn thêm các đường kẻ ngang (---) phân cách trang. Tự động bỏ qua số trang hoặc tiêu đề đầu/chân trang bị lặp lại (header/footer).

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
