import io
import os
from pptx import Presentation
from .translator import (
    sanitize_xml_text,
    execute_translation_batches,
    atomic_save_bytes,
)

def preview_pptx(file_bytes):
    preview_list = []
    temp_prs = Presentation(io.BytesIO(file_bytes))
    for slide in temp_prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    txt = sanitize_xml_text(p.text).strip()
                    if txt:
                        preview_list.append(txt)
    return preview_list

def process_pptx(
    file_bytes,
    model_name,
    progress_bar,
    status_text,
    preview_box,
    file_name,
    lang_mode="zh_to_vi",
    smart_resume=True,
    dich_theo_vung=False,
    start_idx=1,
    end_idx=0,
):
    prs = Presentation(io.BytesIO(file_bytes))
    autosave_filename = f"autosave_{file_name}"

    def apply_pptx(paragraph, translated_text):
        paragraph.text = sanitize_xml_text(translated_text)

    def autosave_pptx():
        try:
            buf = io.BytesIO()
            prs.save(buf)
            atomic_save_bytes(autosave_filename, buf.getvalue())
        except Exception:
            pass

    all_items = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                text = sanitize_xml_text(p.text).strip()
                if text:
                    all_items.append({
                        'obj': p,
                        'text': text,
                        'apply_func': apply_pptx,
                        'autosave_func': autosave_pptx,
                    })

    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else len(all_items)
    target_items = all_items[actual_start:actual_end]

    if target_items:
        status_text.info(f"📊 Tổng: {len(all_items)} đoạn PPT. Đang dịch: {len(target_items)} đoạn.")
        execute_translation_batches(
            target_items,
            model_name,
            progress_bar,
            status_text,
            preview_box,
            lang_mode=lang_mode,
            smart_resume=smart_resume,
        )

    output = io.BytesIO()
    prs.save(output)
    if os.path.exists(autosave_filename):
        try:
            os.remove(autosave_filename)
        except Exception:
            pass
    return output.getvalue()

