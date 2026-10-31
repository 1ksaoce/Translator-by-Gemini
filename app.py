import io
import os
import time
import streamlit as st
import pandas as pd
from PIL import Image
from streamlit_paste_button import paste_image_button
import google.generativeai as genai

from modules.translator import (
    LANG_MODES,
    get_translation_cache,
    configure_api_keys_pool,
    get_active_api_label,
)
from modules.word_processor import process_word, preview_word
from modules.excel_processor import process_excel, preview_excel
from modules.pptx_processor import process_pptx, preview_pptx
from modules.pdf_processor import (
    process_pdf_to_docx,
    convert_docx_bytes_to_formatted_pdf_bytes,
    process_pdf_inplace,
    preview_pdf,
)
from modules.image_processor import process_image_translation

# ==========================================
# DỌN DẸP RÁC HỆ THỐNG AN TOÀN (MỖI PHIÊN 1 LẦN)
# ==========================================
if "cleaned_temp_files" not in st.session_state:
    st.session_state.cleaned_temp_files = True
    now_ts = time.time()
    for f in os.listdir('.'):
        if f.startswith('temp_'):
            try:
                if now_ts - os.path.getmtime(f) > 600:
                    os.remove(f)
            except Exception:
                pass

get_translation_cache()

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.set_page_config(page_title="AI HMI Translator Pro", layout="wide")
st.title("🌐 Ứng Dụng Dịch Tài Liệu Kỹ Thuật Đa Năng")
st.markdown(
    "Hỗ trợ: **Word, Excel, PowerPoint, PDF, Hình Ảnh** | "
    "Đa Ngôn Ngữ | Bảo Toàn 100% Bố Cục & Tự Động Căn Chỉnh Khung."
)

# ==========================================
# SIDEBAR & MENU CÔNG CỤ
# ==========================================
with st.sidebar:
    st.header("🔑 Cấu Hình Hệ Thống (3 API)")

    def _get_secret_key(*names):
        for n in names:
            try:
                val = st.secrets.get(n, "")
                if val and str(val).strip():
                    return str(val).strip()
            except Exception:
                pass
        return ""

    default_k1 = _get_secret_key("GEMINI_API_KEY_1", "GEMINI_API_KEY")
    default_k2 = _get_secret_key("GEMINI_API_KEY_2")
    default_k3 = _get_secret_key("GEMINI_API_KEY_3")

    api_slot_labels = [
        "🔑 API Key 1 (Chính)",
        "🔑 API Key 2 (Dự phòng 1)",
        "🔑 API Key 3 (Dự phòng 2)",
    ]
    selected_api_slot = st.radio(
        "🎯 Chọn nguồn API sử dụng:",
        options=api_slot_labels,
        index=0,
        help="Chọn API Key chính để dịch. Nếu bạn nhập cả 3 API Key, hệ thống sẽ tự động xoay vòng sang API tiếp theo ngay khi API đang dùng bị lỗi hoặc hết hạn mức!",
    )
    preferred_api_idx = api_slot_labels.index(selected_api_slot)

    with st.expander("🔐 Nhập / Thay đổi 3 API Key", expanded=not bool(default_k1 or default_k2 or default_k3)):
        api_key_1 = st.text_input(
            "API Key 1:",
            value=default_k1,
            type="password",
            key="input_api_key_1",
        )
        api_key_2 = st.text_input(
            "API Key 2 (Dự phòng):",
            value=default_k2,
            type="password",
            key="input_api_key_2",
        )
        api_key_3 = st.text_input(
            "API Key 3 (Dự phòng):",
            value=default_k3,
            type="password",
            key="input_api_key_3",
        )

    key_entries = [
        {"label": "API Key 1", "key": api_key_1},
        {"label": "API Key 2", "key": api_key_2},
        {"label": "API Key 3", "key": api_key_3},
    ]
    active_api_key = configure_api_keys_pool(key_entries, preferred_idx=preferred_api_idx)
    available_key_count = sum(1 for x in key_entries if x["key"].strip())

    selected_model_name = "gemini-3.1-flash-lite"
    if active_api_key:
        try:
            available_models = [
                m.name.replace('models/', '')
                for m in genai.list_models()
                if 'generateContent' in m.supported_generation_methods
            ]
            if available_models:
                st.success(
                    f"✅ Đang dùng: **{get_active_api_label()}** (Đã nạp {available_key_count}/3 API sẵn sàng tự động dự phòng)"
                )
                default_idx = 0
                for i, m_name in enumerate(available_models):
                    if m_name.strip().lower() == "gemini-3.1-flash-lite":
                        default_idx = i
                        break
                selected_model_name = st.selectbox(
                    "🤖 Chọn Model AI:", available_models, index=default_idx
                )
        except Exception:
            st.warning(f"⚠️ Không thể lấy danh sách model từ {get_active_api_label()}, chuyển sang chọn thủ công.")
            selected_model_name = st.selectbox(
                "🤖 Chọn Model AI (Thủ công):",
                ["gemini-3.1-flash-lite", "gemini-2.5-flash", "gemini-1.5-flash"],
                index=0,
            )
    else:
        st.warning("⚠️ Vui lòng nhập ít nhất 1 Gemini API Key ở mục trên.")

    st.markdown("---")
    st.header("🌍 Chế Độ Ngôn Ngữ Dịch")
    lang_mode_keys = list(LANG_MODES.keys())
    lang_mode_labels = [LANG_MODES[k]["label"] for k in lang_mode_keys]
    selected_lang_label = st.selectbox(
        "Chọn chiều ngôn ngữ dịch:",
        options=lang_mode_labels,
        index=0,
        help="Chọn cặp ngôn ngữ bạn muốn dịch. Chế độ 'Chỉ dịch Tiếng Trung -> Tiếng Việt' sẽ giữ nguyên tiếng Anh/Việt và chỉ dịch các phần có chữ Hán.",
    )
    selected_lang_mode = lang_mode_keys[lang_mode_labels.index(selected_lang_label)]

    st.markdown("---")
    st.header("🛠 Tùy Chọn Chạy File")

    smart_resume = st.toggle(
        "🧠 Dịch Nối Tiếp (Smart Resume)",
        value=True,
        help="Quét tia X cực nhanh: Tự động bỏ qua các câu đã ở ngôn ngữ đích. Khuyên dùng khi kéo thả file nháp vào để dịch tiếp.",
    )
    skip_shapes = st.toggle(
        "⏭️ Bỏ qua chữ trong Lưu đồ & Hình ảnh",
        value=False,
        help="Khi BẬT: Giữ nguyên 100% lưu đồ và hình ảnh trong Word/PDF/Excel mà không đụng vào chữ bên trong. Khi TẮT: Tự động quét & dịch chữ trực tiếp ngay trên hình ảnh/bản vẽ đính kèm (xóa chữ cũ, chèn chữ đã dịch)!",
    )

    st.markdown("---")
    st.header("🎯 Giới Hạn Vùng Dịch")
    dich_theo_vung = st.toggle(
        "📍 Kích hoạt chọn vùng dịch",
        value=False,
        help="Chỉ dịch từ đoạn X đến đoạn Y. Phải dùng kết hợp với nút 'SOI VỊ TRÍ' ở màn hình chính để biết chính xác số đoạn.",
    )

    start_idx, end_idx = 1, 0
    if dich_theo_vung:
        col1, col2 = st.columns(2)
        with col1:
            start_idx = st.number_input("Từ mục số:", min_value=1, value=1)
        with col2:
            end_idx = st.number_input("Đến mục số:", min_value=1, value=100)

    with st.expander("📖 Cẩm Nang Hướng Dẫn Tính Năng"):
        st.markdown("""
        **1. 🌍 4 Chế Độ Ngôn Ngữ:**
        - **Chỉ dịch Tiếng Trung ➔ Tiếng Việt:** Chỉ lọc các câu có chữ Hán để dịch, giữ nguyên tiếng Anh/mã hiệu.
        - **Tất cả ngôn ngữ ➔ Tiếng Việt:** Dịch mọi ngoại ngữ sang tiếng Việt.
        - **Tiếng Trung ➔ Tiếng Anh:** Dịch tài liệu tiếng Trung sang tiếng Anh kỹ thuật.
        - **Tiếng Anh ➔ Tiếng Việt:** Dịch tài liệu tiếng Anh sang tiếng Việt.

        **2. 📐 Tự Động Căn Khung Excel & PDF:**
        - **Excel:** Tự động nới rộng cột, chiều cao dòng và giữ nguyên tỷ lệ hình ảnh đính kèm.
        - **PDF:** Chống tràn lề A4, tự động xuống dòng khi câu dài và khử 100% lỗi ký tự ô vuông `□`.

        **3. 🖼️ Dịch Ảnh Trực Tiếp:**
        - Tự động xóa chữ cũ trên ảnh, chèn chữ đã dịch vào đúng vị trí đó và liệt kê danh sách chi tiết bên dưới.
        """)

# ==========================================
# CHIA TAB GIAO DIỆN CHÍNH
# ==========================================
tab1, tab2 = st.tabs([
    "📄 Dịch Đa Định Dạng (Word, Excel, PPTX, PDF)",
    "🖼️ Dịch Ảnh Định Vị & Thay Chữ Trực Tiếp",
])

with tab1:
    st.header("📂 Hệ Thống Xử Lý File Kỹ Thuật")

    autosave_files = [f for f in os.listdir('.') if f.startswith('autosave_')]
    if autosave_files:
        st.error("⚠️ **PHÁT HIỆN FILE DỊCH DỞ DANG!** Tải file nháp về, thả vào ô bên dưới để DỊCH TIẾP:")
        for idx_f, f in enumerate(autosave_files):
            try:
                with open(f, "rb") as file_data:
                    draft_bytes = file_data.read()
                st.download_button(
                    label=f"📥 TẢI XUỐNG BẢN NHÁP ({f})",
                    data=draft_bytes,
                    file_name=f,
                    mime="application/octet-stream",
                    type="primary",
                    key=f"draft_dl_{idx_f}_{f}",
                )
            except Exception:
                pass
        if st.button("🗑️ Đã tải xong! Xóa bản nháp này"):
            for f in autosave_files:
                try:
                    os.remove(f)
                except Exception:
                    pass
            st.rerun()

    uploaded_file = st.file_uploader(
        "📂 Tải file (.docx, .xlsx, .pptx, .pdf)",
        type=['docx', 'xlsx', 'pptx', 'pdf'],
    )

    if uploaded_file:
        file_bytes = uploaded_file.read()
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()

        state_sig = f"{uploaded_file.name}_{selected_lang_mode}"
        if st.session_state.get("last_uploaded_sig") != state_sig:
            st.session_state["last_uploaded_sig"] = state_sig
            st.session_state.pop("translated_outputs", None)

        pdf_export_mode = "both_formatted"
        if file_ext == '.pdf':
            st.info(
                "💡 **Chế độ xử lý PDF Chuẩn Form Văn Bản A4:** Tự động nối các câu bị ngắt dòng, "
                "căn đều 2 bên (Justify), thụt đầu dòng chuẩn, font Times New Roman 12pt, giãn dòng 1.5, "
                "giữ nguyên Bảng biểu & Hình ảnh và khử 100% lỗi ký tự ô vuông (□)."
            )
            pdf_mode_label = st.radio(
                "🎯 Chọn định dạng đầu ra cho file PDF:",
                options=[
                    "✨ Xuất cả 2 bản: Word (.docx) + PDF (.pdf) chuẩn Form Văn Bản A4 (Căn đều 2 bên, thụt đầu dòng, Times New Roman)",
                    "📄 Chỉ xuất file Word (.docx) - Chuẩn Form Văn Bản A4 (Times New Roman, căn đều 2 bên, thụt đầu dòng)",
                    "📑 Chỉ xuất file PDF (.pdf) - Chuẩn Form Văn Bản A4 (Times New Roman, căn đều 2 bên, thụt đầu dòng)",
                    "📐 Xuất PDF giữ nguyên tọa độ gốc (Dành riêng cho bản vẽ kỹ thuật / sơ đồ máy)",
                ],
                index=0,
            )
            if pdf_mode_label.startswith("📄"):
                pdf_export_mode = "docx"
            elif pdf_mode_label.startswith("📑"):
                pdf_export_mode = "pdf_formatted"
            elif pdf_mode_label.startswith("📐"):
                pdf_export_mode = "pdf_inplace"
            else:
                pdf_export_mode = "both_formatted"

        with st.expander("🔍 Bấm vào đây để SOI VỊ TRÍ văn bản trong file"):
            if st.button("Hiển thị danh sách đánh số"):
                preview_list = []
                if file_ext == '.docx':
                    preview_list = preview_word(file_bytes, skip_shapes=skip_shapes)
                elif file_ext == '.xlsx':
                    preview_list = preview_excel(file_bytes, skip_shapes=skip_shapes)
                elif file_ext == '.pptx':
                    preview_list = preview_pptx(file_bytes)
                elif file_ext == '.pdf':
                    with st.spinner("Đang quét cấu trúc văn bản PDF..."):
                        preview_list = preview_pdf(
                            file_bytes, pdf_export_mode=pdf_export_mode, skip_shapes=skip_shapes
                        )

                if preview_list:
                    df = pd.DataFrame({
                        "Số Thứ Tự": range(1, len(preview_list) + 1),
                        "Nội dung": [
                            t[:100] + "..." if len(t) > 100 else t for t in preview_list
                        ],
                    })
                    st.dataframe(df, use_container_width=True, hide_index=True)

        if st.button("🚀 Bắt Đầu Dịch File", use_container_width=True, type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            preview_box = st.empty()

            try:
                outputs = []
                base_name = os.path.splitext(uploaded_file.name)[0]

                if file_ext == '.docx':
                    result_bytes = process_word(
                        file_bytes,
                        selected_model_name,
                        progress_bar,
                        status_text,
                        preview_box,
                        uploaded_file.name,
                        lang_mode=selected_lang_mode,
                        smart_resume=smart_resume,
                        skip_shapes=skip_shapes,
                        dich_theo_vung=dich_theo_vung,
                        start_idx=start_idx,
                        end_idx=end_idx,
                    )
                    outputs.append({
                        "label": "📥 TẢI XUỐNG FILE WORD HOÀN CHỈNH (.DOCX)",
                        "data": result_bytes,
                        "file_name": f"Translated_{uploaded_file.name}",
                        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    })
                elif file_ext == '.xlsx':
                    result_bytes = process_excel(
                        file_bytes,
                        selected_model_name,
                        progress_bar,
                        status_text,
                        preview_box,
                        uploaded_file.name,
                        lang_mode=selected_lang_mode,
                        smart_resume=smart_resume,
                        skip_shapes=skip_shapes,
                        dich_theo_vung=dich_theo_vung,
                        start_idx=start_idx,
                        end_idx=end_idx,
                    )
                    outputs.append({
                        "label": "📥 TẢI XUỐNG FILE EXCEL HOÀN CHỈNH (.XLSX)",
                        "data": result_bytes,
                        "file_name": f"Translated_{uploaded_file.name}",
                        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    })
                elif file_ext == '.pptx':
                    result_bytes = process_pptx(
                        file_bytes,
                        selected_model_name,
                        progress_bar,
                        status_text,
                        preview_box,
                        uploaded_file.name,
                        lang_mode=selected_lang_mode,
                        smart_resume=smart_resume,
                        dich_theo_vung=dich_theo_vung,
                        start_idx=start_idx,
                        end_idx=end_idx,
                    )
                    outputs.append({
                        "label": "📥 TẢI XUỐNG FILE POWERPOINT HOÀN CHỈNH (.PPTX)",
                        "data": result_bytes,
                        "file_name": f"Translated_{uploaded_file.name}",
                        "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    })
                elif file_ext == '.pdf':
                    if pdf_export_mode in ("docx", "pdf_formatted", "both_formatted"):
                        docx_res = process_pdf_to_docx(
                            file_bytes,
                            selected_model_name,
                            progress_bar,
                            status_text,
                            preview_box,
                            uploaded_file.name,
                            lang_mode=selected_lang_mode,
                            smart_resume=smart_resume,
                            skip_shapes=skip_shapes,
                            dich_theo_vung=dich_theo_vung,
                            start_idx=start_idx,
                            end_idx=end_idx,
                        )
                        if pdf_export_mode in ("docx", "both_formatted"):
                            outputs.append({
                                "label": "📥 TẢI XUỐNG FILE WORD (.DOCX - CHUẨN FORM VĂN BẢN A4)",
                                "data": docx_res,
                                "file_name": f"Translated_{base_name}.docx",
                                "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            })
                        if pdf_export_mode in ("pdf_formatted", "both_formatted"):
                            status_text.info("📑 Đang kết xuất file PDF chuẩn Form Văn Bản A4 (Căn đều 2 bên, thụt đầu dòng, Times New Roman)...")
                            pdf_formatted_res = convert_docx_bytes_to_formatted_pdf_bytes(docx_res)
                            outputs.append({
                                "label": "📥 TẢI XUỐNG FILE PDF (.PDF - CHUẨN FORM VĂN BẢN A4)",
                                "data": pdf_formatted_res,
                                "file_name": f"Translated_{base_name}.pdf",
                                "mime": "application/pdf",
                            })
                    elif pdf_export_mode == "pdf_inplace":
                        pdf_res = process_pdf_inplace(
                            file_bytes,
                            selected_model_name,
                            progress_bar,
                            status_text,
                            preview_box,
                            uploaded_file.name,
                            lang_mode=selected_lang_mode,
                            smart_resume=smart_resume,
                            skip_shapes=skip_shapes,
                            dich_theo_vung=dich_theo_vung,
                            start_idx=start_idx,
                            end_idx=end_idx,
                        )
                        outputs.append({
                            "label": "📥 TẢI XUỐNG FILE PDF (.PDF - GIỮ TỌA ĐỘ BẢN VẼ GỐC)",
                            "data": pdf_res,
                            "file_name": f"Translated_{base_name}.pdf",
                            "mime": "application/pdf",
                        })

                st.session_state["translated_outputs"] = outputs
                status_text.success("✅ Đã xử lý xong vùng mục tiêu!")
                preview_box.empty()
            except Exception as e:
                st.error(f"❌ Lỗi xử lý hệ thống: {str(e)}")

        if st.session_state.get("translated_outputs"):
            outputs = st.session_state["translated_outputs"]
            if len(outputs) == 1:
                out = outputs[0]
                st.download_button(
                    out["label"],
                    data=out["data"],
                    file_name=out["file_name"],
                    mime=out["mime"],
                    type="primary",
                    use_container_width=True,
                    key=f"dl_final_0_{out['file_name']}",
                )
            else:
                cols = st.columns(len(outputs))
                for idx_o, out in enumerate(outputs):
                    with cols[idx_o]:
                        st.download_button(
                            out["label"],
                            data=out["data"],
                            file_name=out["file_name"],
                            mime=out["mime"],
                            type="primary",
                            use_container_width=True,
                            key=f"dl_final_{idx_o}_{out['file_name']}",
                        )

with tab2:
    st.header("🖼️ Quét Ảnh, Xóa Chữ Cũ & Chèn Chữ Đã Dịch Vào Hình")
    img_to_process = None
    col1, col2 = st.columns(2)
    with col1:
        paste_result = paste_image_button(
            label="📋 BẤM VÀO ĐÂY ĐỂ DÁN ẢNH (CTRL+V)",
            text_color="#ffffff",
            background_color="#0068c9",
        )
        if paste_result.image_data is not None:
            img_to_process = paste_result.image_data.convert('RGB')
    with col2:
        uploaded_image = st.file_uploader(
            "Hoặc Tải File Ảnh Lên",
            type=['jpg', 'jpeg', 'png'],
            label_visibility="collapsed",
        )
        if uploaded_image is not None:
            img_to_process = Image.open(uploaded_image).convert('RGB')

    if img_to_process is not None:
        if not st.session_state.get("last_translated_img"):
            st.image(img_to_process, caption="Ảnh gốc cần dịch", use_column_width=True)

        if st.button("🔍 TIẾN HÀNH DỊCH & THAY CHỮ TRỰC TIẾP TRÊN ẢNH", type="primary", use_container_width=True):
            with st.spinner("AI đang quét tọa độ chữ, xóa chữ cũ và chèn bản dịch vào ảnh..."):
                translated_img, text_summary = process_image_translation(
                    img_to_process,
                    selected_model_name,
                    lang_mode=selected_lang_mode,
                )
                img_buf = io.BytesIO()
                translated_img.save(img_buf, format="PNG")
                st.session_state["last_translated_img"] = translated_img
                st.session_state["last_translated_img_bytes"] = img_buf.getvalue()
                st.session_state["last_translated_img_text"] = text_summary

        if st.session_state.get("last_translated_img") is not None:
            img_col1, img_col2 = st.columns(2)
            with img_col1:
                st.image(img_to_process, caption="📷 Ảnh Gốc", use_column_width=True)
            with img_col2:
                st.image(
                    st.session_state["last_translated_img"],
                    caption="✨ Ảnh Đã Xóa Chữ Cũ & Chèn Chữ Dịch",
                    use_column_width=True,
                )

            st.download_button(
                label="📥 TẢI XUỐNG HÌNH ẢNH ĐÃ DỊCH (.PNG)",
                data=st.session_state["last_translated_img_bytes"],
                file_name="Translated_Image.png",
                mime="image/png",
                type="primary",
                use_container_width=True,
                key="dl_translated_image_btn",
            )

            st.subheader("📋 Danh sách vị trí & nội dung văn bản trong ảnh:")
            st.code(st.session_state.get("last_translated_img_text", ""), language="markdown")
