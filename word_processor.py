import io
import os
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Mm
from docx.text.paragraph import Paragraph
from .translator import (
    sanitize_xml_text,
    execute_translation_batches,
    atomic_save_bytes,
)
from .pdf_processor import (
    _normalize_all_docx_runs_symbols,
    normalize_and_protect_docx_drawings,
)
from .image_processor import translate_docx_embedded_images

XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


def _get_direct_p_text(p_xml):
    return "".join(
        t.text
        for t in p_xml.xpath('./w:r/w:t | ./w:hyperlink//w:t')
        if t.text
    )


def preview_word(file_bytes, skip_shapes=False):
    preview_list = []
    temp_doc = docx.Document(io.BytesIO(file_bytes))
    seen_p = set()
    for p_xml in temp_doc.element.xpath('//w:p'):
        if p_xml in seen_p:
            continue
        is_in_shape = len(p_xml.xpath('ancestor::w:txbxContent')) > 0
        if skip_shapes and is_in_shape:
            continue
        txt = sanitize_xml_text(_get_direct_p_text(p_xml)).strip()
        if txt:
            preview_list.append(txt)
            seen_p.add(p_xml)
    return preview_list


def process_word(
    file_bytes,
    model_name,
    progress_bar,
    status_text,
    preview_box,
    file_name,
    lang_mode="zh_to_vi",
    smart_resume=True,
    skip_shapes=False,
    dich_theo_vung=False,
    start_idx=1,
    end_idx=0,
):
    doc = docx.Document(io.BytesIO(file_bytes))
    autosave_filename = f"autosave_{file_name}"
    qn = docx.oxml.ns.qn

    _normalize_all_docx_runs_symbols(doc)
    normalize_and_protect_docx_drawings(doc)

    for section in doc.sections:
        try:
            is_land = (
                section.page_width is not None
                and section.page_height is not None
                and section.page_width > section.page_height
            )
            if is_land:
                section.page_width, section.page_height = Mm(297), Mm(210)
            else:
                section.page_width, section.page_height = Mm(210), Mm(297)
            section.top_margin = section.bottom_margin = Cm(2.0)
            section.left_margin = section.right_margin = Cm(1.75)
        except Exception:
            pass

    def apply_word(p_info, translated_text):
        p_xml = p_info['p_xml']
        is_in_table = p_info['in_table']
        is_in_shape = p_info['in_shape']
        clean_val = sanitize_xml_text(translated_text).strip()

        # Chỉ cập nhật các run văn bản, giữ nguyên 100% các run chứa hình ảnh / sơ đồ (<w:drawing>, <w:pict>, <mc:AlternateContent>)
        text_runs = [
            r for r in p_xml.xpath('./w:r | ./w:hyperlink/w:r')
            if len(r.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) == 0
        ]

        if text_runs:
            first_r = text_runs[0]
            t_list = first_r.xpath('./w:t')
            if t_list:
                t_el = t_list[0]
                for extra_t in t_list[1:]:
                    first_r.remove(extra_t)
            else:
                t_el = docx.oxml.OxmlElement('w:t')
                first_r.append(t_el)
            t_el.set(XML_SPACE_ATTR, 'preserve')
            t_el.text = clean_val

            for extra_r in text_runs[1:]:
                for extra_t in extra_r.xpath('./w:t'):
                    extra_r.remove(extra_t)
        else:
            first_r = docx.oxml.OxmlElement('w:r')
            t_el = docx.oxml.OxmlElement('w:t')
            t_el.set(XML_SPACE_ATTR, 'preserve')
            t_el.text = clean_val
            first_r.append(t_el)
            p_xml.append(first_r)

        rPr = first_r.find(qn('w:rPr'))
        if rPr is None:
            rPr = docx.oxml.OxmlElement('w:rPr')
            first_r.insert(0, rPr)
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = docx.oxml.OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')

        has_drawing_in_p = (
            len(p_xml.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0
        )
        if not is_in_table and not is_in_shape and not has_drawing_in_p:
            p = Paragraph(p_xml, doc)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            try:
                p.paragraph_format.first_line_indent = Cm(1.27)
            except Exception:
                pass

    def autosave_word():
        try:
            buf = io.BytesIO()
            doc.save(buf)
            atomic_save_bytes(autosave_filename, buf.getvalue())
        except Exception:
            pass

    all_items = []
    seen_paras = set()
    for p_xml in doc.element.xpath('//w:p'):
        if p_xml in seen_paras:
            continue
        is_in_shape = len(p_xml.xpath('ancestor::w:txbxContent')) > 0
        if skip_shapes and is_in_shape:
            continue
        text = sanitize_xml_text(_get_direct_p_text(p_xml)).strip()
        if text:
            is_in_table = len(p_xml.xpath('ancestor::w:tc')) > 0
            all_items.append({
                'obj': {
                    'p_xml': p_xml,
                    'in_table': is_in_table,
                    'in_shape': is_in_shape,
                },
                'text': text,
                'apply_func': apply_word,
                'autosave_func': autosave_word,
            })
            seen_paras.add(p_xml)

    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else len(all_items)
    target_items = all_items[actual_start:actual_end]

    if target_items:
        status_text.info(f"📊 Tổng: {len(all_items)} đoạn Word. Đang dịch: {len(target_items)} đoạn.")
        execute_translation_batches(
            target_items,
            model_name,
            progress_bar,
            status_text,
            preview_box,
            lang_mode=lang_mode,
            smart_resume=smart_resume,
        )

    # Nếu không bật 'Bỏ qua chữ trong Lưu đồ' (skip_shapes=False),
    # dịch trực tiếp chữ bên trong các hình ảnh / lưu đồ đính kèm trong file Word!
    if not skip_shapes:
        translate_docx_embedded_images(
            doc,
            model_name,
            status_text=status_text,
            preview_box=preview_box,
            lang_mode=lang_mode,
        )

    output = io.BytesIO()
    doc.save(output)
    if os.path.exists(autosave_filename):
        try:
            os.remove(autosave_filename)
        except Exception:
            pass
    return output.getvalue()
