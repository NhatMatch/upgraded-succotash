import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import google.generativeai as genai

# Cấu hình trang web
st.set_page_config(page_title="Tạo Đề Kiểm Tra GDPT 2018", page_icon="📝", layout="wide")

st.title("📝 Ứng Dụng Soạn Thảo Đề Kiểm Tra Môn Toán THCS")
st.caption("Hệ thống tự động biên soạn Ma trận - Bản đặc tả - Đề thi - Hướng dẫn chấm chuẩn GDPT 2018")

# Cột bên trái: Bảng điều khiển cấu hình
with st.sidebar:
    st.header("⚙️ Thiết lập cấu trúc đề")
    api_key = st.text_input("Nhập Gemini API Key:", type="password", help="Khóa API lấy miễn phí từ Google AI Studio")
    
    khoi_lop = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"])
    loai_kt = st.selectbox("Hình thức kiểm tra:", ["Kiểm tra định kì Giữa học kì 1", "Kiểm tra định kì Cuối học kì 1", 
                                                    "Kiểm tra định kì Giữa học kì 2", "Kiểm tra định kì Cuối học kì 2", "Kiểm tra thường xuyên"])
    thoi_gian = st.select_slider("Thời gian làm bài (phút):", options=[15, 45, 60, 90], value=60)
    
    st.markdown("---")
    st.subheader("Cấu trúc điểm số")
    ti_le_tn = st.slider("Tỉ lệ Trắc nghiệm (%):", min_value=0, max_value=100, value=30, step=10)
    ti_le_tl = 100 - ti_le_tn
    st.write(f"👉 Trắc nghiệm: **{ti_le_tn}%** | Tự luận: **{ti_le_tl}%**")
    
    st.subheader("Mức độ nhận thức")
    nb = st.number_input("Nhận biết (%)", value=40, step=5)
    th = st.number_input("Thông hiểu (%)", value=30, step=5)
    vd = st.number_input("Vận dụng (%)", value=20, step=5)
    vdc = st.number_input("Vận dụng cao (%)", value=10, step=5)
    
    if nb + th + vd + vdc != 100:
        st.error(f"Tổng tỉ lệ nhận thức hiện tại là {nb + th + vd + vdc}%, cần chỉnh về đúng 100%!")

# Khu vực nội dung chính
st.subheader("📂 1. Tải lên tài liệu / Đề cương kiến thức")
uploaded_file = st.file_uploader("Chọn file tài liệu nội dung bài học (.txt, .docx, .pdf hoặc ảnh đề cương):", type=["txt", "docx", "pdf"])
noi_dung_them = st.text_area("Hoặc dán trực tiếp nội dung/chủ đề kiến thức vào đây:", placeholder="Ví dụ: Chương I: Phân số; Hình học: Góc và đường thẳng...")

# Hàm hỗ trợ xuất file Word
def export_to_docx(content_text, khoi, loai, time):
    doc = docx.Document()
    
    # Tiêu đề tài liệu
    title = doc.add_heading(f"HỒ SƠ ĐỀ KIỂM TRA MÔN TOÁN {khoi.upper()}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.add_run(f"{loai} - Thời gian làm bài: {time} phút\n").bold = True
    p_info.add_run("Khung cấu trúc: 30% Trắc nghiệm khách quan – 70% Tự luận (40% NB - 30% TH - 20% VD - 10% VDC)\n").italic = True
    
    doc.add_paragraph("--------------------------------------------------------------------------------")
    
    # Chèn nội dung được sinh ra từ AI
    lines = content_text.split('\n')
    for line in lines:
        if line.startswith("### "):
            doc.add_heading(line.replace("### ", ""), level=2)
        elif line.startswith("## "):
            doc.add_heading(line.replace("## ", ""), level=1)
        elif line.startswith("# "):
            doc.add_heading(line.replace("# ", ""), level=0)
        else:
            doc.add_paragraph(line)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

st.subheader("⚡ 2. Tiến hành khởi tạo")
if st.button("🚀 Tạo trọn bộ hồ sơ kiểm tra", type="primary"):
    if not api_key:
        st.warning("Vui lòng nhập Gemini API Key ở thanh bên trái để sử dụng!")
    else:
        with st.spinner("Đang phân tích dữ liệu và biên soạn: Ma trận, Bản đặc tả, Đề thi và Hướng dẫn chấm..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-1.5-flash")
                
                # Trích xuất sơ bộ văn bản (nếu người dùng upload file txt)
                context_data = noi_dung_them
                if uploaded_file is not None:
                    if uploaded_file.type == "text/plain":
                        context_data += "\n" + uploaded_file.getvalue().decode("utf-8")
                    else:
                        context_data += f"\n[Tài liệu đính kèm: {uploaded_file.name}]"
                
                prompt = f"""
                Bạn là một chuyên gia khảo thí và giáo viên dạy Toán THCS xuất sắc theo Chương trình GDPT 2018.
                Hãy biên soạn bộ hồ sơ kiểm tra đánh giá hoàn chỉnh với thông tin sau:
                - Khối lớp: {khoi_lop}
                - Kỳ thi: {loai_kt}
                - Thời gian: {thoi_gian} phút
                - Tỉ lệ hình thức: {ti_le_tn}% Trắc nghiệm (4 lựa chọn A, B, C, D, mỗi câu 0.25đ) và {ti_le_tl}% Tự luận
                - Tỉ lệ nhận thức: {nb}% Nhận biết, {th}% Thông hiểu, {vd}% Vận dụng, {vdc}% Vận dụng cao.
                - Nội dung/Kiến thức dựa vào: {context_data if context_data else 'Toàn bộ nội dung chuẩn theo khung phân phối chương trình của khối lớp này'}
                
                Yêu cầu trình bày gồm 4 phần rõ ràng:
                Phần 1: KHUNG MA TRẬN ĐỀ KIỂM TRA (Trình bày chi tiết các mạch kiến thức, số câu, số điểm theo 4 mức độ).
                Phần 2: BẢN ĐẶC TẢ ĐỀ KIỂM TRA (Chi tiết chuẩn kiến thức, kĩ năng cần đánh giá cho từng câu hỏi).
                Phần 3: ĐỀ KIỂM TRA ĐỀ NGHỊ (Gồm câu hỏi trắc nghiệm A, B, C, D và các bài tập tự luận có chia ý a, b, c).
                Phần 4: HƯỚNG DẪN CHẤM VÀ ĐÁP ÁN (Bảng đáp án trắc nghiệm và thang điểm chi tiết cho từng bước giải tự luận).
                """
                
                response = model.generate_content(prompt)
                ket_qua = response.text
                
                st.success("✅ Đã hoàn thành biên soạn hồ sơ kiểm tra!")
                
                # Nút tải file Word
                word_file = export_to_docx(ket_qua, khoi_lop, loai_kt, thoi_gian)
                st.download_button(
                    label="📥 TẢI XUỐNG FILE WORD (.DOCX)",
                    data=word_file,
                    file_name=f"De_Kiem_Tra_Toan_{khoi_lop.replace(' ', '_')}_{thoi_gian}phut.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Hiển thị bản xem trước
                with st.expander("Xem trước nội dung đã tạo trên web:"):
                    st.markdown(ket_qua)
                    
            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")