import streamlit as st
import openpyxl
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Mm
from docx.text.paragraph import Paragraph
import google.generativeai as genai
import time
import io
import os
import re
from PIL import Image
from streamlit_paste_button import paste_image_button # Giữ nút Paste thần thánh

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.set_page_config(page_title="AI HMI Translator Pro", layout="wide")
st.title("🌐 Ứng Dụng Dịch Tài Liệu Kỹ Thuật")
st.markdown("Hệ thống: Dịch File (Autosave siêu tốc) | **Đã khôi phục hộp chọn Model** | Dịch Ảnh định vị.")

# ==========================================
# SIDEBAR: ĐÃ KHÔI PHỤC HỘP CHỌN MODEL CHUẨN
# ==========================================
with st.sidebar:
    st.header("🔑 Cấu Hình Hệ Thống")
    
    # Kiểm tra API an toàn từ Secrets hoặc Text Input
    api_key = None
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
    except:
        api_key = st.text_input("Nhập Google Gemini API Key:", value="AQ.Ab8RN6JNqgoDMJHPcWUUwXym-obZ03NkKeKnmBaOmnBfhS8ydw", type="password")
        if api_key: 
            genai.configure(api_key=api_key)
        
    selected_model_name = "gemini-3.1-flash-lite" # Mặc định dự phòng
    
    if api_key:
        try:
            available_models = [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            if available_models:
                st.success("✅ Đã kết nối API thành công!")
                
                # Thuật toán tự động định vị gemini-3.1-flash-lite làm mặc định
                default_idx = 0
                for i, m_name in enumerate(available_models):
                    if m_name.strip().lower() == "gemini-3.1-flash-lite":
                        default_idx = i
                        break
                
                # HIỂN THỊ LẠI HỘP CHỌN MODEL CHO ĐỨC TỰ DO THAY ĐỔI
                selected_model_name = st.selectbox("🤖 Chọn Model AI:", available_models, index=default_idx)
        except Exception as e:
            # Nếu không list được model do lỗi hệ thống, vẫn cho chọn thủ công
            selected_model_name = st.selectbox("🤖 Chọn Model AI (Thủ công):", ["gemini-3.1-flash-lite", "gemini-1.5-flash", "gemini-1.5-pro"], index=0)

    st.markdown("---")
    st.header("🛠 Tùy Chọn Chạy File")
    smart_resume = st.toggle("🧠 Chế độ Dịch Nối Tiếp (Smart Resume)", value=True)
    skip_shapes = st.toggle("⏭️ Bỏ qua chữ trong Lưu đồ", value=True)
    so_doan_dich = st.number_input("🔢 Số đoạn muốn dịch thử:", min_value=0, value=0, step=1)
    
    st.warning("🚨 **Mẹo dừng khẩn cấp:** Nếu muốn dừng dịch giữa chừng, chỉ cần **F5 (Tải lại trang web)**. File nháp lưu tự động mỗi 3 đoạn sẽ xuất hiện ở màn hình chính!")

# ==========================================
# HÀM BỎ QUA MỤC LỤC THÔNG MINH
# ==========================================
def is_toc_paragraph(p):
    text = p.text.strip()
    if not text: return True
    style_name = p.style.name.lower() if p.style else ""
    if 'toc' in style_name or 'table of contents' in style_name: return True
    if text.upper() in ["目录", "TABLE OF CONTENTS", "CONTENTS", "MỤC LỤC", "INDEX"]: return True
    if re.search(r'[\.…_·\-]{4,}\s*\d+\s*$', text): return True
    if re.search(r'\t\s*\d+\s*$', text): return True
    return False

# ==========================================
# HÀM DỊCH THUẬT CỐT LÕI
# ==========================================
def translate_text_core(text, model_name):
    if not text or not str(text).strip(): return text
    prompt = f"""Bạn là chuyên gia dịch thuật tài liệu kỹ thuật công nghiệp. Dịch đoạn văn bản sau từ tiếng Trung sang tiếng Việt.
YÊU CẦU: 1. KHÔNG in quá trình suy nghĩ, KHÔNG giải thích. 2. Tên công ty (VD: 麦逊) dịch sang tiếng Anh. 3. GIỮ NGUYÊN mã lỗi, mã I/O.
Văn bản gốc:\n{text}\nBản dịch tiếng Việt:"""

    model = genai.GenerativeModel(model_name)
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

# ==========================================
# XỬ LÝ FILE EXCEL
# ==========================================
def process_excel(file_bytes, model_name, progress_bar, status_text, preview_box, file_name):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    autosave_filename = f"autosave_{file_name}"
    total_cells = sum(1 for sheet in wb.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and not cell.value.startswith('=') and cell.value.strip())
    if total_cells == 0: return file_bytes
    
    processed = 0
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and not val.startswith('=') and val.strip():
                    if smart_resume and not re.search(r'[\u4e00-\u9fff]', val):
                        translated = val 
                    else:
                        translated = translate_text_core(val, model_name)
                        preview_box.success(f"**Đang dịch Ô {processed}:**\n\n🇨🇳 `{val}`\n\n🇻🇳 `{translated}`")
                        
                    cell.value = translated
                    processed += 1
                    progress_bar.progress(processed / total_cells)
                    status_text.text(f"Đang xử lý Excel: {processed}/{total_cells}...")
                    
                    if processed % 3 == 0:
                        try: wb.save(autosave_filename)
                        except: pass
                    if so_doan_dich > 0 and processed >= so_doan_dich:
                        output = io.BytesIO()
                        wb.save(output)
                        return output.getvalue()

    output = io.BytesIO()
    wb.save(output)
    if os.path.exists(autosave_filename):
        try: os.remove(autosave_filename) 
        except: pass
    return output.getvalue()

# ==========================================
# XỬ LÝ FILE WORD (CĂN LỀ FULL TRANG A4 + INDENT)
# ==========================================
def process_word(file_bytes, model_name, progress_bar, status_text, preview_box, file_name):
    doc = docx.Document(io.BytesIO(file_bytes))
    autosave_filename = f"autosave_{file_name}"
    
    # ÉP ĐỊNH DẠNG FULL TRANG A4 VÀ CĂN LỀ TIÊU CHUẨN KỸ THUẬT
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

    def format_paragraph(p, is_body_text=True):
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # Căn đều văn bản
        for run in p.runs:
            run.font.name = 'Times New Roman'
            try:
                rPr = run._r.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(docx.oxml.ns.qn('w:ascii'), 'Times New Roman')
                rFonts.set(docx.oxml.ns.qn('w:hAnsi'), 'Times New Roman')
            except: pass
        if is_body_text:
            try: p.paragraph_format.first_line_indent = Cm(1.27) # Thụt đầu dòng 1.27 cm
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
        p = item['obj']
        original_text = p.text.strip()
        
        if smart_resume and not re.search(r'[\u4e00-\u9fff]', original_text):
            translated_text = original_text 
        else:
            translated_text = translate_text_core(original_text, model_name)
            preview_box.success(f"**Đang dịch đoạn {processed + 1}/{total_items}:**\n\n🇨🇳 **Gốc:** {original_text}\n\n🇻🇳 **Dịch:** {translated_text}")
            
        p.clear() 
        p.add_run(translated_text)
        format_paragraph(p, is_body_text=(item['type'] == 'para'))
        
        processed += 1
        progress_bar.progress(processed / total_items)
        status_text.text(f"Đang xử lý Word: {processed}/{total_items}...")
        
        # Tự động lưu nháp cực nhanh mỗi 3 đoạn
        if processed % 3 == 0:
            try: doc.save(autosave_filename)
            except: pass
        if so_doan_dich > 0 and processed >= so_doan_dich:
            break
            
    output = io.BytesIO()
    doc.save(output)
    if os.path.exists(autosave_filename):
        try: os.remove(autosave_filename) 
        except: pass
    return output.getvalue()

# ==========================================
# CHIA TAB GIAO DIỆN
# ==========================================
tab1, tab2 = st.tabs(["📄 Dịch File (Word/Excel)", "🖼️ Dịch Ảnh (Có Định Vị)"])

# ------------------------------------------
# TAB 1: DỊCH FILE
# ------------------------------------------
with tab1:
    st.header("📂 Xử Lý File Tài Liệu")
    
    # HIỂN THỊ FILE NHÁP CỨU HỘ KHẨN CẤP
    autosave_files = [f for f in os.listdir('.') if f.startswith('autosave_')]
    if autosave_files:
        st.error("⚠️ **PHÁT HIỆN FILE DỊCH DỞ DANG DO SẬP MẠNG/TẮT MÁY!**\nHãy tải file nháp dưới đây về, sau đó ném ngược lại vào mục tải file bên dưới để DỊCH NỐI TIẾP:")
        for f in autosave_files:
            with open(f, "rb") as file_data:
                st.download_button(
                    label=f"📥 TẢI XUỐNG BẢN NHÁP ({f.replace('autosave_', '')})",
                    data=file_data,
                    file_name=f,
                    mime="application/octet-stream",
                    type="primary"
                )
        if st.button("🗑️ Đã tải xong! Xóa bản nháp này đi"):
            for f in autosave_files:
                try: os.remove(f)
                except: pass
            st.rerun()
        st.markdown("---")
        
    uploaded_file = st.file_uploader("📂 Tải file (.xlsx, .docx)", type=['xlsx', 'docx'])
    
    if uploaded_file and st.button("🚀 Bắt Đầu Dịch File", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        preview_box = st.empty()
        
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.read()
        
        try:
            if file_ext == '.xlsx':
                result_bytes = process_excel(file_bytes, selected_model_name, progress_bar, status_text, preview_box, uploaded_file.name)
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif file_ext == '.docx':
                result_bytes = process_word(file_bytes, selected_model_name, progress_bar, status_text, preview_box, uploaded_file.name)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                
            status_text.success("✅ Đã xử lý xong toàn bộ File!")
            preview_box.empty()
            st.download_button("📥 TẢI XUỐNG TỆP HOÀN CHỈNH", data=result_bytes, file_name=f"Translated_{uploaded_file.name}", mime=mime_type, type="primary", use_container_width=True)
        except Exception as e:
            st.error(f"❌ Lỗi xử lý: {str(e)}")

# ------------------------------------------
# TAB 2: DỊCH ẢNH (ĐỊNH VỊ GPS VÀ CTRL+V)
# ------------------------------------------
with tab2:
    st.header("🖼️ Đọc & Định Vị Chữ Trong Ảnh")
    st.info("💡 Mẹo: Chụp màn hình bằng `Win + Shift + S`, sau đó click vào nút xanh bên dưới để dán ảnh vào ngay lập tức!")
    
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
            img_to_process = paste_result.image_data.convert('RGB')
            
    with col2:
        st.markdown("### 📂 Hoặc Tải File Lên")
        uploaded_image = st.file_uploader("Nếu đã lưu thành file", type=['jpg', 'jpeg', 'png'], label_visibility="collapsed")
        if uploaded_image is not None:
            img_to_process = Image.open(uploaded_image).convert('RGB')

    if img_to_process is not None:
        st.image(img_to_process, caption="Ảnh đang chờ dịch", use_column_width=True)
        
        if st.button("🔍 TIẾN HÀNH DỊCH & ĐỊNH VỊ", type="primary", use_container_width=True):
            try:
                vision_model = genai.GenerativeModel(selected_model_name)
                vision_prompt = """Hãy nhìn vào bức ảnh sơ đồ/kỹ thuật này. 
Trích xuất toàn bộ chữ tiếng Trung và dịch sang tiếng Việt. 
Đặc biệt, HÃY XÁC ĐỊNH VỊ TRÍ của cụm từ đó trong ảnh (Ví dụ: Góc trên bên trái, Ở giữa, Nhánh bên phải, Hộp lệnh số 2, Dưới cùng...) để tôi dễ dàng tìm thấy trên sơ đồ khi sửa bằng phần mềm Banana.

Trình bày theo định dạng danh sách gạch đầu dòng rõ ràng:
- [Vị trí tương đối trong ảnh] Tiếng Trung -> Tiếng Việt

Chỉ in ra kết quả dịch thuật và vị trí, tuyệt đối không giải thích dài dòng."""

                with st.spinner("AI đang quét ảnh và dò tọa độ sơ đồ..."):
                    response = vision_model.generate_content([vision_prompt, img_to_process])
                    st.success("✅ Dịch thành công! Nhấn biểu tượng 📋 ở góc phải khung để Copy toàn bộ nội dung.")
                    st.code(response.text, language="markdown")
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")
                st.info("💡 Nếu gặp lỗi, hãy nhìn sang MENU BÊN TRÁI và chuyển thử sang model 'gemini-1.5-flash' để kiểm tra tính năng đọc ảnh nhé!")