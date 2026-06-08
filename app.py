import streamlit as st
import openpyxl
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Mm
from docx.text.paragraph import Paragraph
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold # THƯ VIỆN BẺ KHÓA KIỂM DUYỆT
import time
import io
import os
import re
from PIL import Image
from streamlit_paste_button import paste_image_button
import pandas as pd 

# ==========================================
# BỘ LỌC AN TOÀN - TẮT 100% KIỂM DUYỆT (VÔ CÙNG QUAN TRỌNG)
# ==========================================
SAFETY_SETTINGS = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.set_page_config(page_title="AI HMI Translator Pro", layout="wide")
st.title("🌐 Ứng Dụng Dịch Tài Liệu Kỹ Thuật")
st.markdown("Hệ thống: Dịch File (Đã Vô Hiệu Hóa Kiểm Duyệt Từ Ngữ) | **Dịch Ảnh định vị GPS**.")

# ==========================================
# SIDEBAR: CẤU HÌNH API & ĐIỀU KHIỂN
# ==========================================
with st.sidebar:
    st.header("🔑 Cấu Hình Hệ Thống")
    
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
    except:
        api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
        if api_key: 
            genai.configure(api_key=api_key)
            
    selected_model_name = "gemini-3.1-flash-lite"
    
    if api_key:
        try:
            available_models = [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            if available_models:
                st.success("✅ Đã kết nối API thành công!")
                default_idx = 0
                for i, m_name in enumerate(available_models):
                    if m_name.strip().lower() == "gemini-3.1-flash-lite":
                        default_idx = i
                        break
                selected_model_name = st.selectbox("🤖 Chọn Model AI:", available_models, index=default_idx)
        except Exception:
            selected_model_name = st.selectbox("🤖 Chọn Model AI (Thủ công):", ["gemini-3.1-flash-lite", "gemini-1.5-flash"], index=0)

    st.markdown("---")
    st.header("🛠 Tùy Chọn Chạy File")
    
    batch_size = st.slider("📦 Gộp số đoạn 1 lần dịch:", min_value=1, max_value=10, value=3)
    smart_resume = st.toggle("🧠 Dịch Nối Tiếp (Smart Resume)", value=True)
    skip_shapes = st.toggle("⏭️ Bỏ qua chữ trong Lưu đồ", value=False) 

    st.markdown("---")
    st.header("🎯 Giới Hạn Vùng Dịch")
    dich_theo_vung = st.toggle("📍 Kích hoạt chọn vùng", value=False)
    
    start_idx = 1
    end_idx = 0
    if dich_theo_vung:
        col1, col2 = st.columns(2)
        with col1: start_idx = st.number_input("Từ đoạn số:", min_value=1, value=1)
        with col2: end_idx = st.number_input("Đến đoạn số:", min_value=1, value=100)

# ==========================================
# CÁC HÀM XỬ LÝ LÕI VỚI BÁO CÁO LỖI CHI TIẾT
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

def handle_api_error(e, attempt):
    error_msg = str(e).lower()
    if "429" in error_msg or "quota" in error_msg or "exhausted" in error_msg:
        st.toast("⚠️ Google báo quá tải/Hết API. Đang nghỉ 60s...", icon="⏳")
        time.sleep(60) 
    elif "safety" in error_msg or "finish_reason" in error_msg:
        st.toast("🛑 Lỗi Kiểm Duyệt: AI từ chối dịch từ ngữ kỹ thuật nhạy cảm!", icon="🚫")
        time.sleep(2)
    else:
        st.toast(f"Lỗi API ({str(e)[:30]}). Thử lại lần {attempt+1}...", icon="🔄")
        time.sleep(5)

def translate_text_core(text, model_name):
    prompt = f"Dịch đoạn văn bản kỹ thuật sau từ tiếng Trung sang tiếng Việt. KHÔNG giải thích. GIỮ NGUYÊN mã lỗi.\nGốc:\n{text}\nDịch:"
    model = genai.GenerativeModel(model_name)
    
    for attempt in range(4):
        try:
            # GỌI API CÙNG LỆNH BỎ KIỂM DUYỆT
            res = model.generate_content(prompt, safety_settings=SAFETY_SETTINGS)
            time.sleep(4.5) 
            if res.text: return res.text.replace("Bản dịch tiếng Việt:", "").strip()
        except Exception as e:
            handle_api_error(e, attempt)
            
    return text

def translate_batch_core(texts, model_name):
    if not texts: return []
    if len(texts) == 1: return [translate_text_core(texts[0], model_name)]
    
    delimiter = "\n\n|||\n\n"
    combined_text = delimiter.join(texts)
    
    prompt = f"""Bạn là chuyên gia dịch thuật tài liệu kỹ thuật. Dịch văn bản sau từ tiếng Trung sang tiếng Việt.
YÊU CẦU BẮT BUỘC:
1. Có {len(texts)} đoạn được phân tách bằng dấu |||
2. Trả về đúng {len(texts)} đoạn, giữ nguyên dấu ||| ở giữa. KHÔNG GỘP ĐOẠN.
3. Không giải thích.

Văn bản gốc:
{combined_text}"""

    model = genai.GenerativeModel(model_name)
    for attempt in range(3):
        try:
            # GỌI API CÙNG LỆNH BỎ KIỂM DUYỆT
            response = model.generate_content(prompt, safety_settings=SAFETY_SETTINGS)
            time.sleep(4.5) 
            if response.text:
                chunks = [c.strip() for c in response.text.strip().split("|||")]
                if len(chunks) == len(texts): return chunks
                else: break 
        except Exception as e:
            handle_api_error(e, attempt)
            
    st.toast("⚠️ Rã nhóm dịch để xử lý từng câu an toàn...", icon="🔧")
    return [translate_text_core(t, model_name) for t in texts]

# ==========================================
# XỬ LÝ FILE WORD
# ==========================================
def process_word(file_bytes, model_name, progress_bar, status_text, preview_box, file_name):
    doc = docx.Document(io.BytesIO(file_bytes))
    autosave_filename = f"autosave_{file_name}"
    
    for section in doc.sections:
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.top_margin = section.bottom_margin = section.right_margin = Cm(2.0)
        section.left_margin = Cm(3.0)

    def format_paragraph(p, is_body_text=True):
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in p.runs:
            run.font.name = 'Times New Roman'
            try:
                rFonts = run._r.get_or_add_rPr().get_or_add_rFonts()
                rFonts.set(docx.oxml.ns.qn('w:ascii'), 'Times New Roman')
                rFonts.set(docx.oxml.ns.qn('w:hAnsi'), 'Times New Roman')
            except: pass
        if is_body_text:
            try: p.paragraph_format.first_line_indent = Cm(1.27)
            except: pass
            
    all_items = []
    seen_paras = set()
    for p_xml in doc.element.xpath('//w:p'):
        p = Paragraph(p_xml, doc)
        text = p.text.strip()
        if text and p not in seen_paras and not is_toc_paragraph(p):
            is_in_table = len(p_xml.xpath('ancestor::w:tbl')) > 0
            is_in_shape = len(p_xml.xpath('ancestor::w:txbxContent')) > 0
            if skip_shapes and is_in_shape: continue
            all_items.append({'type': 'table' if is_in_table else 'para', 'obj': p, 'text': text})
            seen_paras.add(p)
            
    total_doc_items = len(all_items)
    if total_doc_items == 0: return file_bytes
    
    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else total_doc_items
    
    target_items = all_items[actual_start:actual_end]
    total_targets = len(target_items)
    
    if total_targets == 0:
        st.warning(f"Không có văn bản nào trong khoảng từ đoạn {start_idx} đến {end_idx}!")
        return file_bytes
        
    status_text.info(f"📊 Tổng: {total_doc_items} đoạn. Đang dịch vùng mục tiêu: {total_targets} đoạn.")
    
    processed = 0
    for i in range(0, total_targets, batch_size):
        chunk = target_items[i:i + batch_size]
        
        texts_to_translate = []
        for item in chunk:
            if smart_resume and not re.search(r'[\u4e00-\u9fff]', item['text']):
                item['translated'] = item['text'] 
            else:
                texts_to_translate.append(item['text'])
                item['translated'] = None
                
        if texts_to_translate:
            translated_results = translate_batch_core(texts_to_translate, model_name)
            res_idx = 0
            for item in chunk:
                if item['translated'] is None:
                    item['translated'] = translated_results[res_idx] if res_idx < len(translated_results) else item['text']
                    res_idx += 1
                    
        for item in chunk:
            p = item['obj']
            original = item['text']
            translated = item['translated']
            
            preview_box.success(f"**Đã xử lý đoạn {actual_start + processed + 1}/{total_doc_items} toàn file:**\n\n🇨🇳 `{original}`\n\n🇻🇳 `{translated}`")
            p.clear() 
            p.add_run(translated)
            format_paragraph(p, is_body_text=(item['type'] == 'para'))
            
            processed += 1
            progress_bar.progress(processed / total_targets)
            
            if processed % 5 == 0:
                try: doc.save(autosave_filename)
                except: pass
                
    output = io.BytesIO()
    doc.save(output)
    if os.path.exists(autosave_filename):
        try: os.remove(autosave_filename) 
        except: pass
    return output.getvalue()

# (GIỮ NGUYÊN HÀM EXCEL)
def process_excel(file_bytes, model_name, progress_bar, status_text, preview_box, file_name):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    autosave_filename = f"autosave_{file_name}"
    cells_to_process = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and not val.startswith('=') and val.strip():
                    cells_to_process.append(cell)
    total_doc_cells = len(cells_to_process)
    if total_doc_cells == 0: return file_bytes
    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else total_doc_cells
    target_cells = cells_to_process[actual_start:actual_end]
    total_targets = len(target_cells)
    
    processed = 0
    for cell in target_cells:
        val = cell.value
        if smart_resume and not re.search(r'[\u4e00-\u9fff]', val):
            translated = val 
        else:
            translated = translate_text_core(val, model_name)
            preview_box.success(f"**Đang dịch Ô {actual_start + processed + 1}/{total_doc_cells}:**\n\n🇨🇳 `{val}`\n\n🇻🇳 `{translated}`")
        cell.value = translated
        processed += 1
        progress_bar.progress(processed / total_targets)
        if processed % 10 == 0:
            try: wb.save(autosave_filename)
            except: pass
    output = io.BytesIO()
    wb.save(output)
    if os.path.exists(autosave_filename):
        try: os.remove(autosave_filename) 
        except: pass
    return output.getvalue()

# ==========================================
# CHIA TAB GIAO DIỆN
# ==========================================
tab1, tab2 = st.tabs(["📄 Dịch File (Word/Excel)", "🖼️ Dịch Ảnh (Có Định Vị)"])

with tab1:
    st.header("📂 Xử Lý File Tài Liệu")
    
    autosave_files = [f for f in os.listdir('.') if f.startswith('autosave_')]
    if autosave_files:
        st.error("⚠️ **PHÁT HIỆN FILE DỊCH DỞ DANG!** Tải file nháp về, thả vào ô tải file bên dưới để DỊCH TIẾP:")
        for f in autosave_files:
            with open(f, "rb") as file_data:
                st.download_button(label=f"📥 TẢI XUỐNG BẢN NHÁP", data=file_data, file_name=f, mime="application/octet-stream", type="primary")
        if st.button("🗑️ Đã tải xong! Xóa bản nháp này"):
            for f in autosave_files:
                try: os.remove(f)
                except: pass
            st.rerun()
            
    uploaded_file = st.file_uploader("📂 Tải file (.xlsx, .docx)", type=['xlsx', 'docx'])
    
    if uploaded_file:
        file_bytes = uploaded_file.read()
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()
        
        if file_ext == '.docx':
            with st.expander("🔍 Bấm vào đây để SOI VỊ TRÍ các đoạn văn trong file Word"):
                if st.button("Hiển thị danh sách đoạn"):
                    doc_preview = docx.Document(io.BytesIO(file_bytes))
                    paras_preview = []
                    seen_paras_preview = set()
                    for p_xml in doc_preview.element.xpath('//w:p'):
                        p = Paragraph(p_xml, doc_preview)
                        text = p.text.strip()
                        if text and p not in seen_paras_preview and not is_toc_paragraph(p):
                            if skip_shapes and len(p_xml.xpath('ancestor::w:txbxContent')) > 0: continue
                            paras_preview.append(text)
                            seen_paras_preview.add(p)
                    
                    df = pd.DataFrame({"Đoạn số": range(1, len(paras_preview) + 1), "Nội dung": [t[:100] + "..." if len(t) > 100 else t for t in paras_preview]})
                    st.dataframe(df, use_container_width=True, hide_index=True)
        
        if st.button("🚀 Bắt Đầu Dịch File", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            preview_box = st.empty()
            
            try:
                if file_ext == '.xlsx':
                    result_bytes = process_excel(file_bytes, selected_model_name, progress_bar, status_text, preview_box, uploaded_file.name)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                elif file_ext == '.docx':
                    result_bytes = process_word(file_bytes, selected_model_name, progress_bar, status_text, preview_box, uploaded_file.name)
                    mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    
                status_text.success("✅ Đã xử lý xong vùng mục tiêu!")
                preview_box.empty()
                st.download_button("📥 TẢI XUỐNG TỆP HOÀN CHỈNH", data=result_bytes, file_name=f"Translated_{uploaded_file.name}", mime=mime_type, type="primary", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Lỗi xử lý: {str(e)}")

with tab2:
    st.header("🖼️ Đọc & Định Vị Chữ Trong Ảnh")
    img_to_process = None
    col1, col2 = st.columns(2)
    with col1:
        paste_result = paste_image_button(label="📋 BẤM VÀO ĐÂY ĐỂ DÁN ẢNH (CTRL+V)", text_color="#ffffff", background_color="#0068c9")
        if paste_result.image_data is not None: img_to_process = paste_result.image_data.convert('RGB')
    with col2:
        uploaded_image = st.file_uploader("Hoặc Tải File Lên", type=['jpg', 'jpeg', 'png'], label_visibility="collapsed")
        if uploaded_image is not None: img_to_process = Image.open(uploaded_image).convert('RGB')

    if img_to_process is not None:
        st.image(img_to_process, use_column_width=True)
        if st.button("🔍 TIẾN HÀNH DỊCH & ĐỊNH VỊ", type="primary", use_container_width=True):
            with st.spinner("AI đang quét ảnh..."):
                model = genai.GenerativeModel(selected_model_name)
                # THÊM BỎ LỌC AN TOÀN CHO CẢ PHẦN ẢNH ĐỂ TRÁNH BỊ CHẶN LƯU ĐỒ MẠCH
                res = model.generate_content(["Trích xuất và dịch chữ tiếng Trung trong ảnh. Trình bày dạng danh sách: - [Vị trí tương đối] Tiếng Trung -> Tiếng Việt", img_to_process], safety_settings=SAFETY_SETTINGS)
                st.code(res.text, language="markdown")