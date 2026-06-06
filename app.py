import streamlit as st
import openpyxl
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.text.paragraph import Paragraph
import google.generativeai as genai
import time
import io
import os
import re
from PIL import Image
from streamlit_paste_button import paste_image_button # KÍCH HOẠT LẠI NÚT PASTE

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.set_page_config(page_title="AI HMI Translator Pro", page_icon="⚙️", layout="wide")
st.title("🌐 Ứng Dụng Dịch Tài Liệu Kỹ Thuật")
st.markdown("Hệ thống: Dịch File Word/Excel và **Dịch Ảnh với Nút Paste (Ctrl+V)**.")

# ==========================================
# SIDEBAR: CẤU HÌNH API VÀ CHỌN MODEL CHUẨN ĐÚNG MÃ
# ==========================================
with st.sidebar:
    st.header("🔑 Cấu Hình Hệ Thống")
    # Xóa dòng value="AQ..." cũ đi, thay bằng:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
# Ẩn luôn ô nhập API đi vì nó tự lấy từ file secrets rồi
    
    selected_model_name = None
    if api_key:
        try:
            genai.configure(api_key=api_key)
            available_models = [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            
            if available_models:
                st.success("✅ Đã kết nối API thành công!")
                
                default_idx = 0
                for i, m_name in enumerate(available_models):
                    if m_name.strip().lower() == "gemini-3.1-flash-lite":
                        default_idx = i
                        break
                        
                selected_model_name = st.selectbox("2. Chọn Model AI:", available_models, index=default_idx)
        except Exception as e:
            st.error(f"❌ Lỗi API Key: {str(e)}")

# ==========================================
# CHIA TAB GIAO DIỆN
# ==========================================
tab1, tab2 = st.tabs(["📄 Dịch File (Word/Excel)", "🖼️ Dịch Ảnh Kỹ Thuật (Nút Paste Nhanh)"])

# ------------------------------------------
# TAB 1: DỊCH FILE
# ------------------------------------------
with tab1:
    st.header("📂 Xử Lý File Tài Liệu")
    
    with st.expander("⚙️ Cài đặt dịch File", expanded=True):
        skip_shapes = st.toggle("⏭️ Bỏ qua chữ trong Lưu đồ (Bật nếu sợ vỡ lề Word)", value=True)
        so_doan_dich = st.number_input("🔢 Số đoạn muốn dịch thử:", min_value=0, value=0, step=1)
        emergency_stop = st.toggle("🛑 CÔNG TẮC DỪNG KHẨN CẤP")

    def is_toc_paragraph(p):
        text = p.text.strip()
        if not text: return True
        style_name = p.style.name.lower() if p.style else ""
        if 'toc' in style_name or 'table of contents' in style_name: return True
        if text.upper() in ["目录", "TABLE OF CONTENTS", "CONTENTS", "MỤC LỤC", "INDEX"]: return True
        if re.search(r'[\.…_·\-]{4,}\s*\d+\s*$', text): return True
        if re.search(r'\t\s*\d+\s*$', text): return True
        return False

    def translate_text_core(text, model):
        if not text or not str(text).strip(): return text
        prompt = f"""Bạn là chuyên gia dịch thuật tài liệu kỹ thuật công nghiệp. Dịch đoạn văn bản sau từ tiếng Trung sang tiếng Việt.
YÊU CẦU: 1. KHÔNG in quá trình suy nghĩ, KHÔNG giải thích. 2. Tên công ty (VD: 麦逊) dịch sang tiếng Anh. 3. GIỮ NGUYÊN mã lỗi, mã I/O.
Văn bản gốc:\n{text}\nBản dịch tiếng Việt:"""

        for attempt in range(3):
            try:
                response = model.generate_content(prompt)
                time.sleep(4) 
                if response.text:
                    res = response.text.strip()
                    if res.startswith("Bản dịch tiếng Việt:"): res = res.replace("Bản dịch tiếng Việt:", "", 1).strip()
                    return res
                return text
            except Exception as e:
                if "429" in str(e).lower() or "quota" in str(e).lower(): time.sleep(35)
                else: raise Exception(str(e))
        return text

    def process_excel(file_bytes, model, progress_bar, status_text, preview_box):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        total_cells = sum(1 for sheet in wb.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and not cell.value.startswith('=') and cell.value.strip())
        if total_cells == 0: return file_bytes
        processed = 0
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if emergency_stop or (so_doan_dich > 0 and processed >= so_doan_dich):
                        output = io.BytesIO()
                        wb.save(output)
                        return output.getvalue()
                    val = cell.value
                    if isinstance(val, str) and not val.startswith('=') and val.strip():
                        translated = translate_text_core(val, model)
                        cell.value = translated
                        processed += 1
                        progress_bar.progress(processed / total_cells)
                        status_text.text(f"Đang xử lý Excel: {processed}/{total_cells}...")
                        preview_box.success(f"**Đang dịch Ô {processed}:**\n\n🇨🇳 `{val}`\n\n🇻🇳 `{translated}`")
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def process_word(file_bytes, model, progress_bar, status_text, preview_box):
        doc = docx.Document(io.BytesIO(file_bytes))
        def format_paragraph(p, is_body_text=True):
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in p.runs:
                run.font.name = 'Times New Roman'
                try:
                    rPr = run._r.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(docx.oxml.ns.qn('w:ascii'), 'Times New Roman')
                    rFonts.set(docx.oxml.ns.qn('w:hAnsi'), 'Times New Roman')
                except: pass
            if is_body_text:
                try: p.paragraph_format.first_line_indent = Cm(1.27)
                except: pass
        items_to_process = []
        seen_paras = set()
        for p_xml in doc.element.xpath('//w:p'):
            p = Paragraph(p_xml, doc)
            text = p.text.strip()
            if text and p not in seen_paras and not is_toc_paragraph(p):
                is_in_table = len(p_xml.xpath('ancestor::w:tbl')) > 0
                is_in_shape = len(p_xml.xpath('ancestor::w:txbxContent')) > 0
                if skip_shapes and is_in_shape:
                    seen_paras.add(p)
                    continue
                items_to_process.append({'type': 'table' if is_in_table else 'para', 'obj': p})
                seen_paras.add(p)
        total_items = len(items_to_process)
        if total_items == 0: return file_bytes
        processed = 0
        for item in items_to_process:
            if emergency_stop or (so_doan_dich > 0 and processed >= so_doan_dich): break
            p = item['obj']
            original_text = p.text.strip()
            translated_text = translate_text_core(original_text, model)
            preview_box.success(f"**Đang dịch đoạn {processed + 1}/{total_items}:**\n\n🇨🇳 **Gốc:** {original_text}\n\n🇻🇳 **Dịch:** {translated_text}")
            p.clear() 
            p.add_run(translated_text)
            format_paragraph(p, is_body_text=(item['type'] == 'para'))
            processed += 1
            progress_bar.progress(processed / total_items)
            status_text.text(f"Đang xử lý Word: {processed}/{total_items}...")
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

    uploaded_file = st.file_uploader("📂 Tải lên file Excel (.xlsx) hoặc Word (.docx)", type=['xlsx', 'docx'])
    if uploaded_file is not None:
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        if st.button("🚀 BẮT ĐẦU DỊCH FILE"):
            if not api_key or not selected_model_name:
                st.error("Vui lòng đảm bảo API Key hợp lệ và đã chọn Model!")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                preview_box = st.empty() 
                try:
                    model = genai.GenerativeModel(selected_model_name)
                    file_bytes = uploaded_file.read()
                    if file_ext == '.xlsx':
                        result_bytes = process_excel(file_bytes, model, progress_bar, status_text, preview_box)
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    elif file_ext == '.docx':
                        result_bytes = process_word(file_bytes, model, progress_bar, status_text, preview_box)
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    status_text.success("✅ Đã xử lý xong File!")
                    preview_box.empty() 
                    st.download_button("📥 TẢI XUỐNG TỆP KẾT QUẢ", data=result_bytes, file_name=f"Translated_{uploaded_file.name}", mime=mime_type)
                except Exception as e:
                    st.error(f"❌ Lỗi xử lý: {str(e)}")

# ------------------------------------------
# TAB 2: DỊCH ẢNH (NÚT BẤM DÁN TRỰC TIẾP TỪ CLIPBOARD)
# ------------------------------------------
with tab2:
    st.header("🖼️ Đọc Ảnh Lưu Đồ Siêu Tốc")
    st.info("💡 **Cách dùng:** Chụp màn hình bằng `Win + Shift + S`, sau đó bấm vào nút xanh khổng lồ bên dưới để dán ảnh vào ngay lập tức!")
    
    img_to_process = None
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🖱️ Nút Paste Nhanh")
        paste_result = paste_image_button(
            label="📋 BẤM VÀO ĐÂY ĐỂ DÁN ẢNH (CTRL+V)",
            text_color="#ffffff",
            background_color="#0068c9",
            hover_background_color="#00519c"
        )
        if paste_result.image_data is not None:
            # FIX LỖI ẢNH TRONG SUỐT VÀ SỬA LỖI USE_CONTAINER_WIDTH
            img_to_process = paste_result.image_data.convert('RGB')
            
    with col2:
        st.markdown("### 📂 Hoặc Tải File Lên")
        uploaded_image = st.file_uploader("Nếu đã lưu thành file", type=['jpg', 'jpeg', 'png'], label_visibility="collapsed")
        if uploaded_image is not None:
            img_to_process = Image.open(uploaded_image).convert('RGB')

    # Nếu đã có ảnh (từ nút Paste hoặc Tải file)
    if img_to_process is not None:
        # SỬ DỤNG USE_COLUMN_WIDTH THAY VÌ USE_CONTAINER_WIDTH ĐỂ TRÁNH LỖI PHIÊN BẢN
        st.image(img_to_process, caption="Ảnh đang chờ dịch", use_column_width=True)
        
        if st.button("🔍 TIẾN HÀNH DỊCH ẢNH NÀY", type="primary"):
            if not api_key or not selected_model_name:
                st.error("Vui lòng đảm bảo API Key hợp lệ và đã chọn Model!")
            else:
                with st.spinner("AI đang quét ảnh..."):
                    try:
                        vision_model = genai.GenerativeModel(selected_model_name)
                        vision_prompt = """Nhìn vào ảnh và xuất ra danh sách các chữ tiếng Việt sau khi dịch từ tiếng Trung. 
Trình bày dưới dạng danh sách gạch đầu dòng (Bullet points).
Ví dụ:
- Bắt đầu
- Kiểm tra nguồn
Chỉ in ra tiếng Việt, tuyệt đối KHÔNG in tiếng Trung, KHÔNG tạo bảng, KHÔNG giải thích."""

                        response = vision_model.generate_content([vision_prompt, img_to_process])
                        st.success("✅ Dịch thành công! Nhấn biểu tượng 📋 ở góc phải khung dưới đây để Copy toàn bộ nội dung.")
                        
                        st.code(response.text, language="markdown")
                        
                    except Exception as e:
                         st.error(f"❌ Có lỗi từ Google API: {str(e)}")
                         st.warning("⚠️ Lời khuyên: Model 'flash-lite' bạn đang chọn có thể KHÔNG hỗ trợ đọc hình ảnh. Bạn hãy nhìn sang MENU BÊN TRÁI, chọn model 'gemini-1.5-flash' hoặc 'gemini-1.5-pro' rồi bấm dịch lại nhé!")