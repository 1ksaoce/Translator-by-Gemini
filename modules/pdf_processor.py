import copy
import html
import io
import os
import re
import time
import docx
import fitz  # PyMuPDF
from PIL import Image
from lxml import etree
from pdf2docx import Converter
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from .translator import (
    sanitize_xml_text,
    normalize_special_symbols,
    execute_translation_batches,
    atomic_save_bytes,
)
from .image_processor import (
    process_image_translation,
    translate_docx_embedded_images,
)

XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

TERMINAL_PUNCT = (
    '。', '.', '！', '!', '？', '?', '：', ':', '；', ';', '”', '"', '）', ')', '】', ']'
)
HEADING_OR_LIST_RE = re.compile(
    r'^('
    r'[A-Z][\.\、\)\s]\s*'
    r'|\d{1,3}[\.\、\)]\s*'
    r'|[一二三四五六七八九十百]+[、\.\s]\s*'
    r'|[（\(]\s*\d{1,3}\s*[）\)]\s*'
    r'|[（\(]\s*[一二三四五六七八九十]+\s*[）\)]\s*'
    r'|[①-⑳•▪◆●■➢➤]\s*'
    r')'
)
PAGE_NUM_RE = re.compile(
    r'^(\-\s*\d+\s*\-|\d{1,4}|page\s*\d+(\s*of\s*\d+)?|trang\s*\d+|第\s*\d+\s*页)$',
    re.IGNORECASE,
)


def _visual_text_width(text):
    w = 0.0
    for ch in str(text):
        if '\u4e00' <= ch <= '\u9fff':
            w += 1.0
        else:
            w += 0.52
    return max(w, 0.5)


def _is_cjk_char_or_punct(ch):
    if not ch:
        return False
    cp = ord(ch)
    return (
        0x4E00 <= cp <= 0x9FFF
        or 0x3400 <= cp <= 0x4DBF
        or 0x3000 <= cp <= 0x303F
        or ch in '，。、；：！？（）【】《》“”‘’—…'
    )


def _join_wrapped_lines(prev_text, next_text):
    a = prev_text.strip()
    b = next_text.strip()
    if not a:
        return b
    if not b:
        return a
    if _is_cjk_char_or_punct(a[-1]) and _is_cjk_char_or_punct(b[0]):
        return a + b
    if a.endswith('-') and len(b) > 0 and b[0].islower():
        return a + b
    return f"{a} {b}"


def _normalize_all_docx_runs_symbols(doc):
    """
    Quét toàn bộ các thẻ <w:t> trong tài liệu Word (kể cả các ô/đoạn chỉ chứa ký hiệu không dịch)
    để chuyển đổi các ký tự đặc biệt (℃, ①, Φ, 【】, Wingdings PUA...) sang chuẩn Unicode,
    tránh 100% lỗi hiện hình ô vuông (□).
    """
    qn = docx.oxml.ns.qn
    for r_el in doc.element.xpath('//w:r'):
        t_list = r_el.xpath('./w:t')
        if not t_list:
            continue
        has_change = False
        for t_el in t_list:
            if t_el.text:
                norm_txt = normalize_special_symbols(t_el.text)
                if norm_txt != t_el.text:
                    t_el.text = norm_txt
                    has_change = True
        if has_change:
            rPr = r_el.find(qn('w:rPr'))
            if rPr is None:
                rPr = docx.oxml.OxmlElement('w:rPr')
                r_el.insert(0, rPr)
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = docx.oxml.OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:ascii'), 'Times New Roman')
            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            rFonts.set(qn('w:cs'), 'Times New Roman')


def _get_p_raw_text(p_xml):
    """Chỉ lấy văn bản trực tiếp của đoạn văn (không lấy nhầm chữ bên trong TextBox của hình vẽ)."""
    return "".join(
        t.text
        for t in p_xml.xpath('.//w:t[not(ancestor::w:txbxContent)]')
        if t.text
    )


def _format_drawing_paragraph_clean(p_xml):
    """
    Đảm bảo đoạn văn chứa hình ảnh (<w:drawing> / <w:pict>) luôn hiển thị đầy đủ 100%:
    - Xóa bỏ lineRule="exact" (nguyên nhân chính khiến Word cắt cụt mất 99% chiều cao ảnh!)
    - Căn giữa hình ảnh, xóa thụt đầu dòng.
    - Chuyển <wp:anchor> nổi thành <wp:inline> và thu phóng tỷ lệ ảnh vừa khít trang A4.
    """
    qn = docx.oxml.ns.qn
    pPr = p_xml.find(qn('w:pPr'))
    if pPr is None:
        pPr = docx.oxml.OxmlElement('w:pPr')
        p_xml.insert(0, pPr)

    for old_tag in ('w:framePr', 'w:tabs', 'w:ind', 'w:spacing', 'w:jc'):
        for old_el in pPr.findall(qn(old_tag)):
            pPr.remove(old_el)

    jc_el = docx.oxml.OxmlElement('w:jc')
    jc_el.set(qn('w:val'), 'center')
    ind_el = docx.oxml.OxmlElement('w:ind')
    ind_el.set(qn('w:left'), '0')
    ind_el.set(qn('w:right'), '0')
    ind_el.set(qn('w:firstLine'), '0')
    spc_el = docx.oxml.OxmlElement('w:spacing')
    spc_el.set(qn('w:before'), '80')
    spc_el.set(qn('w:after'), '80')
    spc_el.set(qn('w:line'), '240')
    spc_el.set(qn('w:lineRule'), 'auto')

    pPr.append(jc_el)
    pPr.append(ind_el)
    pPr.append(spc_el)

    # Chuyển các <wp:anchor> (ảnh trôi dễ bị nhảy mất khỏi trang khi dàn lại đoạn) sang <wp:inline>
    # nếu bên trong không chứa TextBox phức tạp
    for drawing_el in p_xml.xpath('.//w:drawing'):
        for anchor_el in drawing_el.xpath('./*[local-name()="anchor"]'):
            if len(anchor_el.xpath('.//*[local-name()="txbxContent"]')) > 0:
                continue
            inline_el = etree.Element(f"{{{WP_NS}}}inline")
            for attr_k in ('distT', 'distB', 'distL', 'distR'):
                inline_el.set(attr_k, anchor_el.get(attr_k, '0'))
            for sub_name in ('extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr', 'graphic'):
                for matched in anchor_el.xpath(f'./*[local-name()="{sub_name}"]'):
                    inline_el.append(copy.deepcopy(matched))
            drawing_el.replace(anchor_el, inline_el)

        # Giới hạn kích thước ảnh tối đa vừa trang A4 (Rộng <= 16.4cm = 5900000 EMU, Cao <= 21.8cm = 7850000 EMU)
        max_cx = 5900000
        max_cy = 7850000
        extents = drawing_el.xpath('.//*[local-name()="extent"]')
        if extents:
            try:
                cx = int(extents[0].get('cx', '0'))
                cy = int(extents[0].get('cy', '0'))
                if cx > max_cx or cy > max_cy:
                    scale = min(max_cx / max(1, cx), max_cy / max(1, cy))
                    new_cx = str(int(round(cx * scale)))
                    new_cy = str(int(round(cy * scale)))
                    for ext_node in drawing_el.xpath('.//*[local-name()="extent"] | .//*[local-name()="xfrm"]/*[local-name()="ext"]'):
                        if ext_node.get('cx') is not None:
                            ext_node.set('cx', new_cx)
                        if ext_node.get('cy') is not None:
                            ext_node.set('cy', new_cy)
            except Exception:
                pass


def normalize_and_protect_docx_drawings(doc):
    """
    Bảo vệ 100% hình ảnh & sơ đồ trong tài liệu Word (.docx):
    1. Gỡ các bảng 1x1 nhân tạo do pdf2docx bọc quanh khung viền bản vẽ (tránh lỗi tràn chiều cao bảng làm mất hình).
    2. Tách riêng các run chứa hình ảnh (<w:drawing>, <w:pict>) ra đoạn văn độc lập nếu bị trộn chung với chữ.
    3. Xóa bỏ lineRule="exact" và căn chỉnh kích thước hình ảnh vừa khít khổ giấy A4.
    """
    qn = docx.oxml.ns.qn
    body = doc.element.body

    # 1. Nếu bảng <w:tbl> chỉ có 1 ô (khung viền bao quanh bản vẽ như Hình 1) hoặc chỉ chứa hình ảnh mà không có chữ,
    # giải phóng các đoạn văn chứa hình ra thẳng w:body để không bị khóa chiều cao dòng bảng (trHeight exact)
    for tbl_el in list(body.xpath('./w:tbl')):
        tcs = tbl_el.xpath('./w:tr/w:tc')
        has_img = len(tbl_el.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0
        tbl_text = "".join(t.text for t in tbl_el.xpath('.//w:t[not(ancestor::w:txbxContent)]') if t.text).strip()
        if has_img and (len(tcs) == 1 or not tbl_text):
            anchor = tbl_el
            for tc_p in tbl_el.xpath('.//w:tc/w:p'):
                cloned_p = copy.deepcopy(tc_p)
                anchor.addnext(cloned_p)
                anchor = cloned_p
            body.remove(tbl_el)

    # 2. Nếu 1 đoạn văn <w:p> trong body vừa có hình ảnh vừa có chữ bên ngoài hình, tách hình ra đoạn riêng
    for p_xml in list(body.xpath('./w:p')):
        drawing_children = [
            ch for ch in p_xml
            if etree.QName(ch).localname != "pPr"
            and len(ch.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0
        ]
        text_outside = _get_p_raw_text(p_xml).strip()
        if drawing_children and text_outside:
            img_p = docx.oxml.OxmlElement('w:p')
            for d_ch in drawing_children:
                # Nếu trong cùng 1 run vừa có drawing vừa có w:t trực tiếp, tách w:t lại cho p_xml
                direct_ts = d_ch.xpath('./w:t')
                if direct_ts:
                    d_clone = copy.deepcopy(d_ch)
                    for dt in d_clone.xpath('./w:t'):
                        d_clone.remove(dt)
                    img_p.append(d_clone)
                    for d_node in d_ch.xpath('./w:drawing | ./w:pict | ./*[local-name()="AlternateContent"]'):
                        d_ch.remove(d_node)
                else:
                    p_xml.remove(d_ch)
                    img_p.append(d_ch)
            _format_drawing_paragraph_clean(img_p)
            p_xml.addprevious(img_p)

    # 3. Chuẩn hóa mọi đoạn văn chứa hình ảnh trong toàn bộ tài liệu
    for p_xml in doc.element.xpath('//w:p'):
        if len(p_xml.xpath('ancestor::w:txbxContent')) > 0:
            continue
        if len(p_xml.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0:
            _format_drawing_paragraph_clean(p_xml)


def _extract_p_lines_and_Split_if_needed(p_xml):
    """
    Tách các dòng vật lý ngăn cách bởi <w:br/> bên trong 1 đoạn <w:p> của pdf2docx.
    """
    qn = docx.oxml.ns.qn
    lines = []
    cur_parts = []
    cur_has_lead_tab = False

    for child in p_xml:
        tag = etree.QName(child).localname
        if tag == "r":
            for sub in child:
                sub_tag = etree.QName(sub).localname
                if sub_tag == "t" and sub.text:
                    cur_parts.append(sub.text)
                elif sub_tag == "tab":
                    if not "".join(cur_parts).strip():
                        cur_has_lead_tab = True
                    else:
                        cur_parts.append(" ")
                elif sub_tag in ("br", "cr"):
                    br_type = sub.get(qn('w:type'))
                    if br_type not in ("page", "column"):
                        lines.append((cur_has_lead_tab, "".join(cur_parts)))
                        cur_parts = []
                        cur_has_lead_tab = False
        elif tag == "hyperlink":
            for t_el in child.xpath('.//w:t'):
                if t_el.text:
                    cur_parts.append(t_el.text)

    if cur_parts or not lines:
        lines.append((cur_has_lead_tab, "".join(cur_parts)))

    logical_paras = []
    for has_tab, raw_line in lines:
        cleaned_line = sanitize_xml_text(raw_line)
        stripped = cleaned_line.strip(" \t\r\n\u3000\xa0")
        if not stripped:
            continue
        has_leading_indent = (
            has_tab
            or cleaned_line.startswith('\u3000')
            or cleaned_line.startswith('  ')
            or cleaned_line.startswith('\t')
            or cleaned_line.startswith('\xa0\xa0')
        )
        if not logical_paras:
            logical_paras.append((has_leading_indent, stripped))
            continue

        prev_indent, prev_text = logical_paras[-1]
        prev_ends_sentence = prev_text.endswith(TERMINAL_PUNCT)
        curr_is_marker = bool(HEADING_OR_LIST_RE.match(stripped))
        prev_is_short_heading = (
            bool(HEADING_OR_LIST_RE.match(prev_text))
            and len(prev_text) <= 42
            and not prev_text.endswith((',', '，', '、'))
        )

        if (has_leading_indent and prev_ends_sentence) or curr_is_marker or prev_is_short_heading:
            logical_paras.append((has_leading_indent, stripped))
        else:
            logical_paras[-1] = (prev_indent, _join_wrapped_lines(prev_text, stripped))

    return logical_paras


def _set_clean_paragraph_text_and_style(p_xml, doc, text_str, is_title=False, is_page_num=False, is_in_table=False):
    """
    Chuẩn hóa toàn bộ định dạng đoạn văn <w:p> theo đúng chuẩn văn bản kỹ thuật A4:
    - Font: Times New Roman, 12pt (24 half-points), màu đen (#000000).
    - Căn lề: Justify (căn đều 2 bên trái - phải).
    - Thụt đầu dòng: 780 dxa (~1.38cm), left_indent = 0, right_indent = 0.
    - Giãn dòng: 1.5 (360/240), khoảng cách đoạn: before = 0pt, after = 6pt (120 dxa).
    """
    qn = docx.oxml.ns.qn
    pPr = p_xml.find(qn('w:pPr'))
    if pPr is None:
        pPr = docx.oxml.OxmlElement('w:pPr')
        p_xml.insert(0, pPr)

    for old_tag in ('w:framePr', 'w:tabs', 'w:ind', 'w:spacing', 'w:jc', 'w:shd', 'w:rPr'):
        for old_el in pPr.findall(qn(old_tag)):
            pPr.remove(old_el)

    jc_el = docx.oxml.OxmlElement('w:jc')
    ind_el = docx.oxml.OxmlElement('w:ind')
    spc_el = docx.oxml.OxmlElement('w:spacing')

    if is_in_table:
        jc_el.set(qn('w:val'), 'left')
        ind_el.set(qn('w:left'), '40')
        ind_el.set(qn('w:right'), '40')
        ind_el.set(qn('w:firstLine'), '0')
        spc_el.set(qn('w:before'), '40')
        spc_el.set(qn('w:after'), '40')
        spc_el.set(qn('w:line'), '276')
        spc_el.set(qn('w:lineRule'), 'auto')
        font_half_pts = '22'
    elif is_page_num:
        jc_el.set(qn('w:val'), 'center')
        ind_el.set(qn('w:left'), '0')
        ind_el.set(qn('w:right'), '0')
        ind_el.set(qn('w:firstLine'), '0')
        spc_el.set(qn('w:before'), '60')
        spc_el.set(qn('w:after'), '60')
        spc_el.set(qn('w:line'), '276')
        spc_el.set(qn('w:lineRule'), 'auto')
        font_half_pts = '24'
    elif is_title:
        jc_el.set(qn('w:val'), 'center')
        ind_el.set(qn('w:left'), '0')
        ind_el.set(qn('w:right'), '0')
        ind_el.set(qn('w:firstLine'), '0')
        spc_el.set(qn('w:before'), '80')
        spc_el.set(qn('w:after'), '160')
        spc_el.set(qn('w:line'), '324')
        spc_el.set(qn('w:lineRule'), 'auto')
        font_half_pts = '28'
    else:
        jc_el.set(qn('w:val'), 'both')
        ind_el.set(qn('w:left'), '0')
        ind_el.set(qn('w:right'), '0')
        ind_el.set(qn('w:firstLine'), '780')
        spc_el.set(qn('w:before'), '0')
        spc_el.set(qn('w:after'), '120')
        spc_el.set(qn('w:line'), '360')
        spc_el.set(qn('w:lineRule'), 'auto')
        font_half_pts = '24'

    pPr.append(jc_el)
    pPr.append(ind_el)
    pPr.append(spc_el)

    for child in list(p_xml):
        tag = etree.QName(child).localname
        if tag == "pPr":
            continue
        has_drawing = len(child.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0
        if has_drawing:
            for t_el in child.xpath('./w:t'):
                child.remove(t_el)
        else:
            p_xml.remove(child)

    clean_text = text_str.strip(" \t\r\n\u3000\xa0")
    if clean_text:
        r_el = docx.oxml.OxmlElement('w:r')
        rPr = docx.oxml.OxmlElement('w:rPr')
        rFonts = docx.oxml.OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        rPr.append(rFonts)

        if is_title:
            b_el = docx.oxml.OxmlElement('w:b')
            rPr.append(b_el)

        sz_el = docx.oxml.OxmlElement('w:sz')
        sz_el.set(qn('w:val'), font_half_pts)
        szCs_el = docx.oxml.OxmlElement('w:szCs')
        szCs_el.set(qn('w:val'), font_half_pts)
        color_el = docx.oxml.OxmlElement('w:color')
        color_el.set(qn('w:val'), '000000')

        rPr.append(sz_el)
        rPr.append(szCs_el)
        rPr.append(color_el)
        r_el.append(rPr)

        t_el = docx.oxml.OxmlElement('w:t')
        t_el.set(XML_SPACE_ATTR, 'preserve')
        t_el.text = clean_text
        r_el.append(t_el)
        p_xml.append(r_el)


def _reconstruct_and_standardize_docx_layout(doc):
    """
    Chuẩn hóa toàn bộ cấu trúc Word xuất ra từ pdf2docx:
    1. Bảo vệ 100% các hình ảnh / sơ đồ đính kèm (gỡ bảng bọc 1x1, bỏ lineRule exact làm mất ảnh).
    2. Chuẩn hóa khổ giấy A4, lề chuẩn, 1 cột.
    3. Nối các câu/đoạn văn bị pdf2docx cắt vụn theo từng dòng PDF thành đoạn văn hoàn chỉnh.
    4. Định dạng chuẩn: Times New Roman 12pt, căn đều 2 bên (Justify), thụt đầu dòng, giãn dòng 1.5.
    """
    qn = docx.oxml.ns.qn
    _normalize_all_docx_runs_symbols(doc)
    normalize_and_protect_docx_drawings(doc)

    body = doc.element.body

    # 1. Gỡ bỏ các sectPr trung gian nằm trong pPr để văn bản chảy tự nhiên
    for pPr_sect in body.xpath('./w:p/w:pPr/w:sectPr'):
        orient_els = pPr_sect.xpath('./w:pgSz/@w:orient')
        is_landscape = any(str(o).lower() == 'landscape' for o in orient_els)
        if not is_landscape:
            parent_pPr = pPr_sect.getparent()
            if parent_pPr is not None:
                parent_pPr.remove(pPr_sect)

    # 2. Chuẩn hóa mọi Section về chuẩn A4 và 1 cột duy nhất
    for cols_el in doc.element.xpath('//w:sectPr/w:cols'):
        cols_el.set(qn('w:num'), '1')
        cols_el.set(qn('w:space'), '0')
        for child_col in list(cols_el):
            cols_el.remove(child_col)

    for section in doc.sections:
        try:
            is_land = (
                section.page_width is not None
                and section.page_height is not None
                and section.page_width > section.page_height
            )
            if is_land:
                section.page_width = Cm(29.7)
                section.page_height = Cm(21.0)
            else:
                section.page_width = Cm(21.0)
                section.page_height = Cm(29.7)
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(1.75)
            section.right_margin = Cm(1.75)
        except Exception:
            pass

    # 3. Đổi chiều cao dòng bảng từ 'exact' sang 'atLeast' (và gỡ chiều cao quá khổ)
    for tr_height in doc.element.xpath('//w:trPr/w:trHeight'):
        if tr_height.get(qn('w:hRule')) == 'exact':
            tr_height.set(qn('w:hRule'), 'atLeast')
        val_str = tr_height.get(qn('w:val'), '0')
        if str(val_str).isdigit() and int(val_str) > 9000:
            tr_height.getparent().remove(tr_height)

    # 4. Duyệt các đoạn văn trực tiếp trong w:body để tách/nối các dòng bị ngắt bởi <w:br/>
    body_paras = list(body.xpath('./w:p'))
    for p_xml in body_paras:
        has_drawing = len(p_xml.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0
        has_sect = len(p_xml.xpath('./w:pPr/w:sectPr')) > 0
        raw_txt = _get_p_raw_text(p_xml).strip(" \t\r\n\u3000\xa0")

        if has_drawing:
            _format_drawing_paragraph_clean(p_xml)
            continue

        if not raw_txt and not has_sect:
            body.remove(p_xml)
            continue

        if has_sect:
            continue

        logical_paras = _extract_p_lines_and_Split_if_needed(p_xml)
        if not logical_paras:
            body.remove(p_xml)
            continue

        p_obj = Paragraph(p_xml, doc)
        pf_obj = p_obj.paragraph_format
        orig_left_pt = pf_obj.left_indent.pt if pf_obj.left_indent is not None else 0.0
        orig_first_pt = pf_obj.first_line_indent.pt if pf_obj.first_line_indent is not None else 0.0
        orig_align = p_obj.alignment

        max_sz_half = 24
        for sz_val in p_xml.xpath('.//w:rPr/w:sz/@w:val'):
            if str(sz_val).isdigit():
                max_sz_half = max(max_sz_half, int(sz_val))

        first_indent_flag, first_text = logical_paras[0]
        p_xml.set('data-orig-left', f"{orig_left_pt:.1f}")
        p_xml.set('data-init-left', f"{orig_left_pt:.1f}")
        p_xml.set('data-orig-first', f"{orig_first_pt:.1f}")
        p_xml.set('data-lead-indent', '1' if (first_indent_flag or orig_first_pt >= 10.0) else '0')
        p_xml.set('data-orig-align', 'center' if orig_align == WD_ALIGN_PARAGRAPH.CENTER else 'other')
        p_xml.set('data-max-sz', str(max_sz_half))

        _set_clean_paragraph_text_and_style(p_xml, doc, first_text)

        insert_anchor = p_xml
        for extra_indent_flag, extra_text in logical_paras[1:]:
            new_p = docx.oxml.OxmlElement('w:p')
            new_p.set('data-orig-left', f"{orig_left_pt:.1f}")
            new_p.set('data-init-left', f"{orig_left_pt:.1f}")
            new_p.set('data-orig-first', '0.0')
            new_p.set('data-lead-indent', '1' if extra_indent_flag else '0')
            new_p.set('data-orig-align', 'other')
            new_p.set('data-max-sz', str(max_sz_half))
            _set_clean_paragraph_text_and_style(new_p, doc, extra_text)
            insert_anchor.addnext(new_p)
            insert_anchor = new_p

    # 5. Gộp các đoạn <w:p> liên tiếp bị pdf2docx cắt ngang câu giữa các dòng hoặc giữa 2 trang
    idx = 0
    children = list(body)
    while idx < len(children) - 1:
        curr_el = children[idx]
        next_el = children[idx + 1]
        if etree.QName(curr_el).localname != "p" or etree.QName(next_el).localname != "p":
            idx += 1
            continue

        if (
            len(curr_el.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"] | ./w:pPr/w:sectPr')) > 0
            or len(next_el.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"] | ./w:pPr/w:sectPr')) > 0
        ):
            idx += 1
            continue

        t1 = _get_p_raw_text(curr_el).strip()
        t2 = _get_p_raw_text(next_el).strip()
        if not t1 or not t2:
            idx += 1
            continue

        if PAGE_NUM_RE.match(t1) or PAGE_NUM_RE.match(t2):
            idx += 1
            continue

        c_align = curr_el.get('data-orig-align', 'other')
        n_align = next_el.get('data-orig-align', 'other')
        c_sz = int(curr_el.get('data-max-sz', '24'))
        n_sz = int(next_el.get('data-max-sz', '24'))

        if (c_align == 'center' and c_sz >= 30) or (n_align == 'center' and n_sz >= 30):
            idx += 1
            continue

        c_left = float(curr_el.get('data-orig-left', '0.0'))
        c_init_left = float(curr_el.get('data-init-left', '0.0'))
        n_left = float(next_el.get('data-orig-left', '0.0'))
        n_lead_indent = next_el.get('data-lead-indent', '0') == '1'

        t1_ends_terminal = t1.endswith(TERMINAL_PUNCT)
        t2_is_marker = bool(HEADING_OR_LIST_RE.match(t2))
        t1_is_short_heading = (
            (bool(HEADING_OR_LIST_RE.match(t1)) and len(t1) <= 42)
            or (len(t1) <= 14 and not t1.endswith((',', '，', '、')))
        )

        n_is_indented_new_para = n_lead_indent or (n_left - c_left >= 10.0)
        n_is_flush_continuation_of_indented_p1 = (
            (c_init_left - n_left >= 10.0)
            and (c_init_left - n_left <= 60.0)
            and not n_lead_indent
        )

        should_merge = (
            not t2_is_marker
            and not t1_is_short_heading
            and (
                (not t1_ends_terminal and not n_is_indented_new_para)
                or n_is_flush_continuation_of_indented_p1
            )
        )

        if should_merge:
            merged_txt = _join_wrapped_lines(t1, t2)
            curr_el.set('data-orig-left', f"{min(c_left, n_left):.1f}")
            _set_clean_paragraph_text_and_style(curr_el, doc, merged_txt)
            body.remove(next_el)
            children.pop(idx + 1)
        else:
            idx += 1

    # 6. Hoàn thiện định dạng cho tiêu đề chính / số trang / bảng biểu và xóa các attribute tạm
    for p_xml in body.xpath('./w:p'):
        for attr_name in (
            'data-orig-left',
            'data-init-left',
            'data-orig-first',
            'data-lead-indent',
            'data-orig-align',
            'data-max-sz',
        ):
            if attr_name in p_xml.attrib:
                del p_xml.attrib[attr_name]

        if len(p_xml.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0:
            continue

        txt = _get_p_raw_text(p_xml).strip()
        if not txt:
            continue
        if PAGE_NUM_RE.match(txt):
            _set_clean_paragraph_text_and_style(p_xml, doc, txt, is_page_num=True)

    for tc_p in body.xpath('.//w:tc//w:p'):
        if len(tc_p.xpath('.//w:drawing | .//w:pict | .//*[local-name()="AlternateContent"]')) > 0:
            _format_drawing_paragraph_clean(tc_p)
            continue
        for spacing in tc_p.xpath('./w:pPr/w:spacing'):
            if spacing.get(qn('w:lineRule')) == 'exact':
                spacing.set(qn('w:lineRule'), 'atLeast')
        p_tc_obj = Paragraph(tc_p, doc)
        pf_tc = p_tc_obj.paragraph_format
        try:
            pf_tc.left_indent = Pt(2)
            pf_tc.right_indent = Pt(2)
            pf_tc.first_line_indent = Pt(0)
            pf_tc.space_before = Pt(2)
            pf_tc.space_after = Pt(2)
        except Exception:
            pass


def _extract_docx_layout_segments(doc, skip_shapes_flag=False):
    """
    Chuẩn hóa bố cục chuẩn A4 và trích xuất các phân đoạn văn bản để gửi cho AI dịch.
    """
    _reconstruct_and_standardize_docx_layout(doc)

    segments = []
    seen_p = set()

    for p_xml in doc.element.xpath('//w:p'):
        if p_xml in seen_p:
            continue
        seen_p.add(p_xml)

        is_in_shape = len(p_xml.xpath('ancestor::w:txbxContent')) > 0
        if skip_shapes_flag and is_in_shape:
            continue

        is_in_table = len(p_xml.xpath('ancestor::w:tc')) > 0
        has_tabs = len(p_xml.xpath('.//w:tab')) > 0

        current_group = []
        p_segments = []

        def flush_group(grp, preceding_tab_el=None):
            if not grp:
                return
            texts = []
            for r_el in grp:
                for t_el in r_el.xpath('./w:t'):
                    if t_el.text:
                        texts.append(t_el.text)
            raw_str = "".join(texts)
            clean_str = sanitize_xml_text(raw_str).strip(" \t\r\n\u3000\xa0")
            if clean_str:
                is_first_in_body_p = (not is_in_table) and (len(p_segments) == 0)
                seg_obj = {
                    'runs': list(grp),
                    'raw_text': raw_str,
                    'clean_text': clean_str,
                    'lead_space': "" if is_first_in_body_p else (" " if raw_str[:1].isspace() else ""),
                    'trail_space': "" if not is_in_table else (" " if raw_str[-1:].isspace() else ""),
                    'in_table': is_in_table,
                    'has_tabs': has_tabs,
                    'preceding_tab': preceding_tab_el,
                    'para_segments': p_segments,
                }
                p_segments.append(seg_obj)
                segments.append(seg_obj)

        last_tab_el = None
        for child in p_xml:
            tag = etree.QName(child).localname
            if tag == "r":
                tabs_in_r = child.xpath('./w:tab')
                has_delimiter = len(child.xpath('./w:tab | ./w:br | ./w:cr | ./w:drawing | ./w:pict')) > 0
                has_text = len(child.xpath('./w:t')) > 0
                if has_delimiter and not has_text:
                    flush_group(current_group, preceding_tab_el=last_tab_el)
                    current_group = []
                    if tabs_in_r:
                        last_tab_el = tabs_in_r[-1]
                elif has_delimiter and has_text:
                    flush_group(current_group, preceding_tab_el=last_tab_el)
                    current_group = []
                    if tabs_in_r:
                        last_tab_el = tabs_in_r[-1]
                    flush_group([child], preceding_tab_el=last_tab_el)
                elif has_text:
                    current_group.append(child)
            elif tag == "hyperlink":
                flush_group(current_group, preceding_tab_el=last_tab_el)
                current_group = []
                h_runs = child.xpath('./w:r[w:t]')
                flush_group(h_runs, preceding_tab_el=last_tab_el)
        flush_group(current_group, preceding_tab_el=last_tab_el)

    return segments


def _apply_docx_segment_translation(seg_info, translated_text):
    qn = docx.oxml.ns.qn
    runs = seg_info['runs']
    if not runs:
        return

    clean_trans = sanitize_xml_text(translated_text).strip(" \t\r\n\u3000\xa0")
    seg_info['translated_text'] = clean_trans
    final_text = f"{seg_info['lead_space']}{clean_trans}{seg_info['trail_space']}"

    first_r = runs[0]
    t_elems = first_r.xpath('./w:t')
    if t_elems:
        t_el = t_elems[0]
        for extra_t in t_elems[1:]:
            first_r.remove(extra_t)
    else:
        t_el = docx.oxml.OxmlElement('w:t')
        first_r.append(t_el)

    t_el.set(XML_SPACE_ATTR, "preserve")
    t_el.text = final_text

    for extra_r in runs[1:]:
        for extra_t in extra_r.xpath('./w:t'):
            extra_r.remove(extra_t)

    rPr = first_r.find(qn('w:rPr'))
    if rPr is None:
        rPr = docx.oxml.OxmlElement('w:rPr')
        first_r.insert(0, rPr)

    for bad_tag in ('w:spacing', 'w:w', 'w:kern', 'w:position'):
        for bad_el in rPr.findall(qn(bad_tag)):
            rPr.remove(bad_el)

    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = docx.oxml.OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')

    if not seg_info['in_table']:
        sz_el = rPr.find(qn('w:sz'))
        if sz_el is None:
            sz_el = docx.oxml.OxmlElement('w:sz')
            rPr.append(sz_el)
        if sz_el.get(qn('w:val')) not in ('28', '30', '32'):
            sz_el.set(qn('w:val'), '24')

        szCs_el = rPr.find(qn('w:szCs'))
        if szCs_el is None:
            szCs_el = docx.oxml.OxmlElement('w:szCs')
            rPr.append(szCs_el)
        if szCs_el.get(qn('w:val')) not in ('28', '30', '32'):
            szCs_el.set(qn('w:val'), '24')
    else:
        orig_w = _visual_text_width(seg_info['clean_text'])
        trans_w = _visual_text_width(clean_trans)
        if trans_w > orig_w * 1.5:
            sz_el = rPr.find(qn('w:sz'))
            orig_half_pts = 22
            if sz_el is not None and sz_el.get(qn('w:val')) and str(sz_el.get(qn('w:val'))).isdigit():
                orig_half_pts = int(sz_el.get(qn('w:val')))
            else:
                sz_el = docx.oxml.OxmlElement('w:sz')
                rPr.append(sz_el)
            new_half_pts = max(19, min(22, int(round(orig_half_pts * 0.90))))
            sz_el.set(qn('w:val'), str(new_half_pts))
            szCs_el = rPr.find(qn('w:szCs'))
            if szCs_el is None:
                szCs_el = docx.oxml.OxmlElement('w:szCs')
                rPr.append(szCs_el)
            szCs_el.set(qn('w:val'), str(new_half_pts))


def _preprocess_pdf_preserve_vector_diagrams(file_bytes):
    """
    Kiểm tra các trang PDF trước khi đưa vào pdf2docx:
    Nếu một trang là bản vẽ kỹ thuật vector có chữ xoay dọc 90 độ (như Hình 1 Sơ đồ lắp đặt)
    hoặc lưu đồ vector phức tạp mà pdf2docx sẽ làm rơi mất chữ xoay / nét vẽ,
    tự động kết xuất trang bản vẽ đó thành hình ảnh độ phân giải cao (200 DPI)
    để 100% bản vẽ được giữ nguyên và sẵn sàng cho bộ dịch hình ảnh xử lý.
    """
    try:
        src_doc = fitz.open(stream=file_bytes, filetype="pdf")
        out_doc = fitz.open()
        modified = False

        for page in src_doc:
            t_dict = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
            drawings = page.get_drawings()
            images = page.get_images()

            total_lines = 0
            rotated_lines = 0
            max_block_chars = 0
            total_chars = 0

            for b in t_dict.get("blocks", []):
                b_chars = 0
                for l in b.get("lines", []):
                    txt = "".join(s.get("text", "") for s in l.get("spans", "")).strip()
                    if not txt:
                        continue
                    total_lines += 1
                    b_chars += len(txt)
                    total_chars += len(txt)
                    d_vec = l.get("dir", (1.0, 0.0))
                    if abs(d_vec[0]) < 0.5:
                        rotated_lines += 1
                if b_chars > max_block_chars:
                    max_block_chars = b_chars

            # Trang bản vẽ có chữ xoay dọc (rotate 90/270) HOẶC trang thuần bản vẽ vector chú thích ngắn
            is_rotated_diagram_page = (
                total_lines > 0
                and rotated_lines >= max(2, int(total_lines * 0.4))
                and (len(drawings) >= 5 or len(images) >= 1)
            )
            is_vector_schematic_page = (
                len(drawings) >= 25
                and max_block_chars < 65
                and total_chars < 380
            )

            if is_rotated_diagram_page or is_vector_schematic_page:
                modified = True
                pix = page.get_pixmap(dpi=200)
                img_bytes = pix.tobytes("png")
                new_p = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                margin = 18.0
                target_rect = fitz.Rect(
                    margin,
                    margin,
                    page.rect.width - margin,
                    page.rect.height - margin,
                )
                new_p.insert_image(target_rect, stream=img_bytes, keep_proportion=True)
            else:
                out_doc.insert_pdf(src_doc, from_page=page.number, to_page=page.number)

        if modified:
            res_bytes = out_doc.tobytes()
        else:
            res_bytes = file_bytes
        out_doc.close()
        src_doc.close()
        return res_bytes
    except Exception:
        return file_bytes


def convert_pdf_bytes_to_docx_bytes(file_bytes):
    prepared_pdf_bytes = _preprocess_pdf_preserve_vector_diagrams(file_bytes)
    stamp = int(time.time() * 1000)
    temp_pdf = f"temp_in_{stamp}.pdf"
    temp_docx = f"temp_out_{stamp}.docx"
    try:
        with open(temp_pdf, "wb") as f:
            f.write(prepared_pdf_bytes)
        cv = Converter(temp_pdf)
        cv.convert(
            temp_docx,
            start=0,
            end=None,
            lines_center_aligned_threshold=15.0,
            max_line_spacing_ratio=0.9,
        )
        cv.close()
        with open(temp_docx, "rb") as f:
            return f.read()
    finally:
        for p in (temp_pdf, temp_docx):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass


# ==========================================
# KẾT XUẤT PDF CHUẨN FORM VĂN BẢN A4 TỪ DOCX ĐÃ DỊCH
# ==========================================
def convert_docx_bytes_to_formatted_pdf_bytes(docx_bytes):
    """
    Chuyển đổi tài liệu Word (.docx) đã chuẩn hóa sang file PDF (.pdf) chuẩn Form Văn Bản A4:
    - Căn đều 2 bên (text-align: justify)
    - Thụt đầu dòng chuẩn cho mọi đoạn văn
    - Font Times New Roman 12pt, giãn dòng 1.5, cách đoạn 6pt
    - Bảo toàn 100% bảng biểu (<w:tbl>) và hình ảnh / bản vẽ đính kèm (<w:drawing>)
    """
    doc = docx.Document(io.BytesIO(docx_bytes))

    arch = fitz.Archive()
    font_dirs = [
        "C:/Windows/Fonts",
        "/usr/share/fonts/truetype/dejavu",
        "/usr/share/fonts/truetype/liberation",
        "/usr/share/fonts/truetype/freefont",
    ]
    for fdir in font_dirs:
        if os.path.isdir(fdir):
            try:
                arch.add(fdir)
            except Exception:
                pass

    if arch.has_entry("times.ttf"):
        css_font_faces = """
        @font-face { font-family: 'DocSerif'; src: url('times.ttf'); font-weight: normal; font-style: normal; }
        @font-face { font-family: 'DocSerif'; src: url('timesbd.ttf'); font-weight: bold; font-style: normal; }
        @font-face { font-family: 'DocSerif'; src: url('timesi.ttf'); font-weight: normal; font-style: italic; }
        @font-face { font-family: 'DocSerif'; src: url('timesbi.ttf'); font-weight: bold; font-style: italic; }
        """
    elif arch.has_entry("LiberationSerif-Regular.ttf"):
        css_font_faces = """
        @font-face { font-family: 'DocSerif'; src: url('LiberationSerif-Regular.ttf'); font-weight: normal; }
        @font-face { font-family: 'DocSerif'; src: url('LiberationSerif-Bold.ttf'); font-weight: bold; }
        """
    elif arch.has_entry("DejaVuSerif.ttf"):
        css_font_faces = """
        @font-face { font-family: 'DocSerif'; src: url('DejaVuSerif.ttf'); font-weight: normal; }
        @font-face { font-family: 'DocSerif'; src: url('DejaVuSerif-Bold.ttf'); font-weight: bold; }
        """
    else:
        css_font_faces = ""

    css = f"""
    {css_font_faces}
    body {{
        font-family: 'DocSerif', 'Times New Roman', serif;
        font-size: 12pt;
        color: #000000;
        line-height: 1.5;
    }}
    p.body-p {{
        text-align: justify;
        text-indent: 39pt;
        margin-top: 0pt;
        margin-bottom: 6pt;
        line-height: 1.5;
        font-size: 12pt;
        color: #000000;
    }}
    p.title-p {{
        text-align: center;
        text-indent: 0pt;
        font-weight: bold;
        font-size: 14pt;
        margin-top: 4pt;
        margin-bottom: 10pt;
        line-height: 1.35;
        color: #000000;
    }}
    p.page-num-p {{
        text-align: center;
        text-indent: 0pt;
        font-size: 11.5pt;
        margin-top: 4pt;
        margin-bottom: 4pt;
        color: #333333;
    }}
    div.img-box {{
        text-align: center;
        margin-top: 6pt;
        margin-bottom: 6pt;
        page-break-inside: avoid;
    }}
    div.img-box-large {{
        text-align: center;
        margin-top: 6pt;
        margin-bottom: 6pt;
        page-break-before: always;
        page-break-inside: avoid;
    }}
    table.doc-tbl {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 6pt;
        margin-bottom: 8pt;
        font-size: 11pt;
    }}
    td.doc-td {{
        border: 0.6pt solid #333333;
        padding: 4pt 5pt;
        vertical-align: middle;
        text-align: left;
        line-height: 1.25;
        font-size: 11pt;
    }}
    p.cell-p {{
        text-align: left;
        text-indent: 0pt;
        margin-top: 1pt;
        margin-bottom: 1pt;
        font-size: 11pt;
    }}
    """

    img_counter = [0]

    def extract_images_from_element(xml_el, max_w_pt=450.0, max_h_pt=640.0):
        img_tags = []
        r_ids = xml_el.xpath(
            './/*[local-name()="blip"]/@*[local-name()="embed"] | '
            './/*[local-name()="imagedata"]/@*[local-name()="id"]'
        )
        for r_id in r_ids:
            part = doc.part.related_parts.get(r_id)
            if part is None:
                continue
            try:
                pil_img = Image.open(io.BytesIO(part.blob)).convert("RGB")
                px_w, px_h = pil_img.size
                png_buf = io.BytesIO()
                pil_img.save(png_buf, format="PNG")
                png_bytes = png_buf.getvalue()

                img_counter[0] += 1
                img_name = f"docx_img_{img_counter[0]}.png"
                arch.add(png_bytes, img_name)

                orig_w_pt = px_w * 0.75
                orig_h_pt = px_h * 0.75
                extents = xml_el.xpath('.//*[local-name()="extent"]')
                if extents:
                    cx = extents[0].get('cx')
                    cy = extents[0].get('cy')
                    if cx and str(cx).isdigit() and int(cx) > 0:
                        orig_w_pt = int(cx) / 12700.0
                    if cy and str(cy).isdigit() and int(cy) > 0:
                        orig_h_pt = int(cy) / 12700.0

                scale = min(max_w_pt / max(1.0, orig_w_pt), max_h_pt / max(1.0, orig_h_pt), 1.0)
                w_pt = max(24, int(round(orig_w_pt * scale)))
                h_pt = max(24, int(round(orig_h_pt * scale)))
                box_cls = "img-box-large" if h_pt >= 320 else "img-box"
                img_tags.append(
                    f'<div class="{box_cls}"><img src="{img_name}" width="{w_pt}" height="{h_pt}"/></div>'
                )
            except Exception:
                pass
        return img_tags

    html_parts = ["<html><head><meta charset='utf-8'></head><body>"]

    for child in doc.element.body:
        tag = etree.QName(child).localname
        if tag == "p":
            for img_html in extract_images_from_element(child, max_w_pt=455.0, max_h_pt=640.0):
                html_parts.append(img_html)

            txt = _get_p_raw_text(child).strip(" \t\r\n\u3000\xa0")
            if not txt:
                continue

            esc_txt = html.escape(sanitize_xml_text(txt))
            jc_vals = child.xpath('./w:pPr/w:jc/@w:val')
            jc_val = jc_vals[0] if jc_vals else 'both'
            sz_vals = child.xpath('.//w:rPr/w:sz/@w:val')
            sz_val = int(sz_vals[0]) if (sz_vals and str(sz_vals[0]).isdigit()) else 24

            if PAGE_NUM_RE.match(txt):
                html_parts.append(f'<p class="page-num-p">{esc_txt}</p>')
            elif jc_val == 'center' and sz_val >= 28:
                html_parts.append(f'<p class="title-p">{esc_txt}</p>')
            else:
                html_parts.append(f'<p class="body-p">{esc_txt}</p>')

        elif tag == "tbl":
            rows = child.xpath('./w:tr')
            if not rows:
                continue
            html_parts.append('<table class="doc-tbl">')
            for tr in rows:
                html_parts.append('<tr>')
                tcs = tr.xpath('./w:tc')
                col_max_w = max(140.0, 440.0 / max(1, len(tcs)))
                for tc in tcs:
                    span_vals = tc.xpath('./w:tcPr/w:gridSpan/@w:val')
                    colspan_attr = (
                        f' colspan="{span_vals[0]}"'
                        if (span_vals and str(span_vals[0]).isdigit() and int(span_vals[0]) > 1)
                        else ''
                    )
                    cell_chunks = []
                    for tc_img in extract_images_from_element(tc, max_w_pt=col_max_w, max_h_pt=500.0):
                        cell_chunks.append(tc_img)
                    for tc_p in tc.xpath('./w:p'):
                        c_txt = _get_p_raw_text(tc_p).strip(" \t\r\n\u3000\xa0")
                        if c_txt:
                            cell_chunks.append(
                                f'<p class="cell-p">{html.escape(sanitize_xml_text(c_txt))}</p>'
                            )
                    cell_inner = "".join(cell_chunks) if cell_chunks else "&nbsp;"
                    html_parts.append(f'<td class="doc-td"{colspan_attr}>{cell_inner}</td>')
                html_parts.append('</tr>')
            html_parts.append('</table>')

    html_parts.append("</body></html>")
    full_html = "\n".join(html_parts)

    story = fitz.Story(html=full_html, user_css=css, archive=arch)
    pdf_buf = io.BytesIO()
    writer = fitz.DocumentWriter(pdf_buf)

    mediabox = fitz.paper_rect("a4")
    where = fitz.Rect(
        49.6,
        56.7,
        mediabox.width - 49.6,
        mediabox.height - 56.7,
    )

    more = 1
    while more:
        device = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(device)
        writer.end_page()

    writer.close()
    return pdf_buf.getvalue()


def process_pdf_to_docx(
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
    status_text.info(
        "🔄 Đang phân tích cấu trúc PDF, bảo toàn Hình ảnh / Sơ đồ & định dạng chuẩn Form Văn Bản A4..."
    )
    docx_bytes = convert_pdf_bytes_to_docx_bytes(file_bytes)

    doc = docx.Document(io.BytesIO(docx_bytes))
    autosave_filename = f"autosave_{file_name.replace('.pdf', '.docx')}"

    segments = _extract_docx_layout_segments(doc, skip_shapes_flag=skip_shapes)

    def autosave_pdf_docx():
        try:
            buf = io.BytesIO()
            doc.save(buf)
            atomic_save_bytes(autosave_filename, buf.getvalue())
        except Exception:
            pass

    all_items = []
    for seg in segments:
        all_items.append({
            'obj': seg,
            'text': seg['clean_text'],
            'apply_func': _apply_docx_segment_translation,
            'autosave_func': autosave_pdf_docx,
        })

    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else len(all_items)
    target_items = all_items[actual_start:actual_end]

    if target_items:
        status_text.info(f"📊 Tổng: {len(all_items)} đoạn văn bản chuẩn A4. Đang dịch: {len(target_items)} đoạn.")
        execute_translation_batches(
            target_items,
            model_name,
            progress_bar,
            status_text,
            preview_box,
            lang_mode=lang_mode,
            smart_resume=smart_resume,
        )

    # Nếu KHÔNG bật 'Bỏ qua chữ trong Lưu đồ' (skip_shapes=False),
    # tự động quét và dịch trực tiếp chữ bên trong các hình ảnh / bản vẽ đính kèm!
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


# ==========================================
# DỊCH TRỰC TIẾP TRÊN TỌA ĐỘ PDF GỐC (IN-PLACE PYMUPDF DÀNH CHO BẢN VẼ / SƠ ĐỒ)
# ==========================================
def _find_vietnamese_font_path():
    candidates = [
        "C:/Windows/Fonts/times.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "C:/Windows/Fonts/calibri.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _ensure_font_supported_text(text, font_obj):
    cleaned = sanitize_xml_text(text)
    if font_obj is None:
        return cleaned
    out = []
    for ch in cleaned:
        if ch in ('\n', '\r', '\t', ' '):
            out.append(ch)
        elif font_obj.has_glyph(ord(ch)):
            out.append(ch)
        else:
            out.append(' ')
    return "".join(out)


def _extract_pdf_page_vertical_walls(page):
    vwalls = []
    try:
        for d in page.get_drawings():
            for item in d.get("items", []):
                kind = item[0]
                if kind == "re":
                    r = item[1]
                    if r.height > 6:
                        if r.width < 3.5:
                            vwalls.append(((r.x0 + r.x1) / 2.0, r.y0, r.y1))
                        else:
                            vwalls.append((r.x0, r.y0, r.y1))
                            vwalls.append((r.x1, r.y0, r.y1))
                elif kind == "l":
                    p1, p2 = item[1], item[2]
                    if abs(p1.x - p2.x) < 2.5 and abs(p1.y - p2.y) > 6:
                        vwalls.append(((p1.x + p2.x) / 2.0, min(p1.y, p2.y), max(p1.y, p2.y)))
    except Exception:
        pass
    return vwalls


def _int_color_to_rgb(c_int):
    if not isinstance(c_int, int):
        return (0.0, 0.0, 0.0)
    r = ((c_int >> 16) & 255) / 255.0
    g = ((c_int >> 8) & 255) / 255.0
    b = (c_int & 255) / 255.0
    return (r, g, b)


def _wrap_text_to_width(text, font_obj, fontsize, max_width):
    words = text.split()
    if not words:
        return [text]
    lines = []
    current = words[0]
    for w in words[1:]:
        cand = f"{current} {w}"
        if font_obj.text_length(cand, fontsize=fontsize) <= max_width:
            current = cand
        else:
            lines.append(current)
            current = w
    lines.append(current)
    return lines


def process_pdf_inplace(
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
    pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
    font_path = _find_vietnamese_font_path()
    vi_font = fitz.Font(fontfile=font_path) if font_path else fitz.Font("helv")

    all_items = []
    page_records = []

    for page_idx, page in enumerate(pdf_doc):
        vwalls = _extract_pdf_page_vertical_walls(page)
        text_dict = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
        raw_segments = []

        for block in text_dict.get("blocks", []):
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                cur_spans = [spans[0]]
                for sp in spans[1:]:
                    prev_sp = cur_spans[-1]
                    gap = sp["bbox"][0] - prev_sp["bbox"][2]
                    max_gap = max(prev_sp.get("size", 10.0) * 1.2, 10.0)
                    if gap > max_gap:
                        txt = "".join(s["text"] for s in cur_spans).strip()
                        if txt:
                            raw_segments.append(cur_spans)
                        cur_spans = [sp]
                    else:
                        cur_spans.append(sp)
                txt = "".join(s["text"] for s in cur_spans).strip()
                if txt:
                    raw_segments.append(cur_spans)

        seg_boxes = []
        for sp_group in raw_segments:
            x0 = min(s["bbox"][0] for s in sp_group)
            y0 = min(s["bbox"][1] for s in sp_group)
            x1 = max(s["bbox"][2] for s in sp_group)
            y1 = max(s["bbox"][3] for s in sp_group)
            seg_boxes.append((x0, y0, x1, y1))

        safe_page_right = max(100.0, page.rect.width - 38.0)
        safe_page_left = 28.0

        page_items = []
        for idx_s, sp_group in enumerate(raw_segments):
            x0, y0, x1, y1 = seg_boxes[idx_s]
            ym = (y0 + y1) / 2.0
            origin_x, origin_y = sp_group[0].get("origin", (x0, y1 - 2.0))
            font_size = max(s.get("size", 10.0) for s in sp_group)
            color_rgb = _int_color_to_rgb(sp_group[0].get("color", 0))
            raw_extracted = "".join(s["text"] for s in sp_group)
            full_text = sanitize_xml_text(raw_extracted).strip()

            left_candidates = [min(safe_page_left, x0)]
            right_candidates = [max(safe_page_right, x1 + 5.0)]
            has_table_right_wall = False

            for vx, vy0, vy1 in vwalls:
                if (vy0 - 2.0) <= ym <= (vy1 + 2.0):
                    if vx <= x0 + 2.0:
                        left_candidates.append(vx)
                    if vx >= x1 - 2.0:
                        right_candidates.append(vx)
                        has_table_right_wall = True

            has_right_neighbor = False
            for j_s, (ox0, oy0, ox1, oy1) in enumerate(seg_boxes):
                if j_s == idx_s:
                    continue
                if not (oy1 < y0 + 2.0 or oy0 > y1 - 2.0):
                    if ox1 <= x0 + 2.0:
                        left_candidates.append(ox1 + 3.0)
                    if ox0 >= x1 - 2.0:
                        right_candidates.append(ox0 - 4.0)
                        has_right_neighbor = True

            bound_left = max(left_candidates)
            bound_right = min(right_candidates)

            orig_w = max(20.0, x1 - x0)
            if not has_table_right_wall and not has_right_neighbor:
                max_natural_right = min(safe_page_right, max(x1 + 90.0, x0 + orig_w * 2.2))
                bound_right = min(bound_right, max_natural_right)

            if bound_right <= bound_left + 15.0:
                bound_left = x0
                bound_right = max(x1 + 25.0, safe_page_right)

            d_left = x0 - bound_left
            d_right = bound_right - x1
            box_w = bound_right - bound_left
            is_centered = (
                d_left > 5.0
                and d_right > 5.0
                and abs(d_left - d_right) < max(16.0, box_w * 0.25)
            )

            rec = {
                'page_idx': page_idx,
                'orig_raw': raw_extracted.strip(),
                'orig_text': full_text,
                'bbox': (x0, y0, x1, y1),
                'origin': (origin_x, origin_y),
                'fontsize': font_size,
                'color': color_rgb,
                'bound_left': bound_left,
                'bound_right': bound_right,
                'is_centered': is_centered,
                'translated': None,
            }
            page_items.append(rec)

            def apply_pdf_seg(r_obj, trans_str):
                r_obj['translated'] = _ensure_font_supported_text(trans_str, vi_font).strip()

            all_items.append({
                'obj': rec,
                'text': full_text,
                'apply_func': apply_pdf_seg,
            })

        page_records.append((page, page_items))

    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else len(all_items)
    target_items = all_items[actual_start:actual_end]

    if target_items:
        status_text.info(f"📊 Tổng: {len(all_items)} mục tọa độ PDF. Đang dịch: {len(target_items)} mục.")
        execute_translation_batches(
            target_items,
            model_name,
            progress_bar,
            status_text,
            preview_box,
            lang_mode=lang_mode,
            smart_resume=smart_resume,
        )

    for page, p_items in page_records:
        # Nếu không bật skip_shapes, dịch luôn các hình ảnh / bản vẽ quét (bitmap) trên trang PDF
        if not skip_shapes:
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                try:
                    base_img = pdf_doc.extract_image(xref)
                    if not base_img or "image" not in base_img:
                        continue
                    pil_img = Image.open(io.BytesIO(base_img["image"]))
                    if pil_img.width < 70 or pil_img.height < 70:
                        continue
                    status_text.info(f"🖼️ Đang quét & dịch chữ trong hình ảnh trên trang PDF {page.number + 1}...")
                    trans_pil, summary_txt = process_image_translation(
                        pil_img.convert("RGB"), model_name, lang_mode=lang_mode
                    )
                    if summary_txt:
                        img_buf = io.BytesIO()
                        trans_pil.save(img_buf, format="PNG")
                        page.replace_image(xref, stream=img_buf.getvalue())
                        preview_box.success(f"**🖼️ Đã dịch ảnh trang {page.number + 1}:**\n\n{summary_txt[:500]}")
                except Exception:
                    pass

        modified_items = [
            it for it in p_items
            if it['translated'] is not None and (
                it['translated'] != it['orig_text'] or it['orig_text'] != it['orig_raw']
            )
        ]
        if not modified_items:
            continue

        for it in modified_items:
            rx0, ry0, rx1, ry1 = it['bbox']
            redact_rect = fitz.Rect(rx0 + 0.2, ry0 + 0.5, rx1 - 0.2, ry1 - 0.5)
            page.add_redact_annot(redact_rect, fill=False)

        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_NONE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
        )

        if font_path:
            page.insert_font(fontname="vi_font", fontfile=font_path)
            used_fontname = "vi_font"
        else:
            used_fontname = "helv"

        for it in modified_items:
            trans_text = it['translated']
            if not trans_text:
                continue
            x0, y0, x1, y1 = it['bbox']
            orig_x, orig_y = it['origin']
            orig_fs = it['fontsize']
            b_left = it['bound_left']
            b_right = it['bound_right']
            color = it['color']

            if it['is_centered']:
                avail_w = max(20.0, (b_right - b_left) - 4.0)
            else:
                avail_w = max(20.0, (b_right - x0) - 3.0)

            text_w = vi_font.text_length(trans_text, fontsize=orig_fs)
            if not it['is_centered'] and text_w > avail_w and (x0 - b_left) > 4.0:
                avail_w = max(avail_w, (b_right - b_left) - 4.0)

            min_readable_fs = min(orig_fs, max(8.5, orig_fs * 0.88))
            if text_w <= avail_w:
                final_fs = orig_fs
                lines = [trans_text]
            else:
                scaled_fs = orig_fs * (avail_w / text_w)
                if scaled_fs >= min_readable_fs:
                    final_fs = scaled_fs
                    lines = [trans_text]
                else:
                    final_fs = min_readable_fs
                    lines = _wrap_text_to_width(trans_text, vi_font, final_fs, avail_w)

            if len(lines) == 1:
                line_w = vi_font.text_length(lines[0], fontsize=final_fs)
                if it['is_centered']:
                    draw_x = b_left + max(1.0, (b_right - b_left - line_w) / 2.0)
                else:
                    draw_x = x0
                    if draw_x + line_w > b_right - 1.5:
                        draw_x = max(b_left + 1.5, b_right - 1.5 - line_w)
                page.insert_text(
                    (draw_x, orig_y),
                    lines[0],
                    fontname=used_fontname,
                    fontsize=final_fs,
                    color=color,
                )
            else:
                line_step = final_fs * 1.15
                total_h = line_step * (len(lines) - 1)
                start_y = orig_y - (total_h * 0.25)
                for l_idx, ln in enumerate(lines):
                    line_w = vi_font.text_length(ln, fontsize=final_fs)
                    if it['is_centered']:
                        draw_x = b_left + max(1.0, (b_right - b_left - line_w) / 2.0)
                    else:
                        draw_x = max(b_left + 1.5, min(x0, b_right - 1.5 - line_w))
                    page.insert_text(
                        (draw_x, start_y + l_idx * line_step),
                        ln,
                        fontname=used_fontname,
                        fontsize=final_fs,
                        color=color,
                    )

    out_pdf_bytes = pdf_doc.tobytes(garbage=3, deflate=True)
    pdf_doc.close()
    return out_pdf_bytes


def preview_pdf(file_bytes, pdf_export_mode="both_formatted", skip_shapes=False):
    preview_list = []
    if pdf_export_mode == "pdf_inplace":
        temp_pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in temp_pdf_doc:
            t_dict = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
            for b in t_dict.get("blocks", []):
                for l in b.get("lines", []):
                    txt = sanitize_xml_text("".join(s["text"] for s in l.get("spans", []))).strip()
                    if txt:
                        preview_list.append(txt)
        temp_pdf_doc.close()
    else:
        docx_preview_bytes = convert_pdf_bytes_to_docx_bytes(file_bytes)
        temp_doc = docx.Document(io.BytesIO(docx_preview_bytes))
        segs = _extract_docx_layout_segments(temp_doc, skip_shapes_flag=skip_shapes)
        preview_list = [s['clean_text'] for s in segs]
    return preview_list
