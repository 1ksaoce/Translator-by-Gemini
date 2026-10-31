import copy
import io
import math
import os
import re
import zipfile
import openpyxl
import openpyxl.drawing.image
from lxml import etree
from .translator import (
    sanitize_xml_text,
    execute_translation_batches,
    atomic_save_bytes,
)

# ==========================================
# VÁ LỖI LÕI OPENPYXL: CHỐNG ĐÓNG LUỒNG ẢNH KHI LƯU NHIỀU LẦN
# ==========================================
def _safe_openpyxl_image_data(self):
    img = openpyxl.drawing.image._import_image(self.ref)
    if self.format in ['gif', 'jpeg', 'png']:
        if hasattr(img, 'fp') and img.fp is not None and not getattr(img.fp, 'closed', False):
            img.fp.seek(0)
            return img.fp.read()
        elif isinstance(self.ref, io.BytesIO) and not self.ref.closed:
            self.ref.seek(0)
            return self.ref.read()
    fp = io.BytesIO()
    img.save(fp, format="png")
    fp.seek(0)
    return fp.read()

openpyxl.drawing.image.Image._data = _safe_openpyxl_image_data

EXCEL_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EXCEL_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

CELL_REF_RE = re.compile(r'^([A-Za-z]+)(\d+)$')

def _col_letters_to_idx(letters):
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx

def _parse_cell_ref(ref_str):
    if not ref_str:
        return None, None
    m = CELL_REF_RE.match(ref_str.strip())
    if not m:
        return None, None
    return _col_letters_to_idx(m.group(1)), int(m.group(2))

def _excel_visual_width(text):
    if not text:
        return 0.0
    max_line = 0.0
    for line in str(text).split('\n'):
        w = 0.0
        for ch in line:
            if '\u4e00' <= ch <= '\u9fff':
                w += 2.05
            elif ch.isupper() or ch in 'MW@#%&':
                w += 1.22
            else:
                w += 1.05
        if w > max_line:
            max_line = w
    return max_line

def _get_si_text(si_elem):
    parts = []
    for child in si_elem:
        tag = etree.QName(child).localname
        if tag == "t":
            if child.text:
                parts.append(child.text)
        elif tag == "r":
            for sub in child:
                if etree.QName(sub).localname == "t" and sub.text:
                    parts.append(sub.text)
    return "".join(parts)

def _get_inline_is_text(is_elem):
    parts = []
    for child in is_elem:
        tag = etree.QName(child).localname
        if tag == "t":
            if child.text:
                parts.append(child.text)
        elif tag == "r":
            for sub in child:
                if etree.QName(sub).localname == "t" and sub.text:
                    parts.append(sub.text)
    return "".join(parts)

def _enable_wrap_text_in_styles(styles_root):
    """Bật wrapText cho các style ô trong xl/styles.xml để chữ dài tự động xuống dòng."""
    if styles_root is None:
        return
    ns = styles_root.nsmap.get(None, EXCEL_MAIN_NS)
    for xf in styles_root.xpath('./*[local-name()="cellXfs"]/*[local-name()="xf"]'):
        xf.set("applyAlignment", "1")
        align_list = xf.xpath('./*[local-name()="alignment"]')
        if align_list:
            align_el = align_list[0]
        else:
            align_el = etree.SubElement(xf, f"{{{ns}}}alignment")
        align_el.set("wrapText", "1")
        if not align_el.get("vertical"):
            align_el.set("vertical", "center")

def _extract_drawing_size_requirements(drawing_root):
    """
    Đọc kích thước các hình ảnh / khối vẽ trong drawing*.xml để đảm bảo cột và dòng
    chứa hình có đủ độ rộng và chiều cao, đồng thời giữ nguyên tỷ lệ ảnh.
    """
    col_min_widths = {}  # col_idx (1-based) -> width in Excel units
    row_min_heights = {}  # row_idx (1-based) -> height in points

    if drawing_root is None:
        return col_min_widths, row_min_heights

    # 1. Xử lý twoCellAnchor
    for anchor in drawing_root.xpath('./*[local-name()="twoCellAnchor"]'):
        # Đặt editAs="oneCell" để khi nới rộng cột cho chữ, hình ảnh không bị kéo giãn méo tỷ lệ
        if not anchor.get("editAs") or anchor.get("editAs") == "twoCell":
            anchor.set("editAs", "oneCell")

        from_col_el = anchor.xpath('./*[local-name()="from"]/*[local-name()="col"]')
        from_row_el = anchor.xpath('./*[local-name()="from"]/*[local-name()="row"]')
        to_col_el = anchor.xpath('./*[local-name()="to"]/*[local-name()="col"]')
        to_row_el = anchor.xpath('./*[local-name()="to"]/*[local-name()="row"]')

        if not (from_col_el and from_row_el and to_col_el and to_row_el):
            continue
        try:
            c1 = int(from_col_el[0].text) + 1
            r1 = int(from_row_el[0].text) + 1
            c2 = max(c1, int(to_col_el[0].text) + 1)
            r2 = max(r1, int(to_row_el[0].text) + 1)
        except Exception:
            continue

        ext_list = anchor.xpath('.//*[local-name()="xfrm"]/*[local-name()="ext"]')
        cx, cy = 0, 0
        if ext_list:
            try:
                cx = int(ext_list[0].get("cx", "0"))
                cy = int(ext_list[0].get("cy", "0"))
            except Exception:
                pass

        if cx > 0:
            total_w_units = (cx / 66675.0) + 3.0
            if total_w_units <= 55.0:
                col_min_widths[c1] = max(col_min_widths.get(c1, 0.0), total_w_units)
            else:
                n_cols = max(1, c2 - c1 + 1)
                per_col_w = min(55.0, total_w_units / n_cols)
                for c in range(c1, c2 + 1):
                    col_min_widths[c] = max(col_min_widths.get(c, 0.0), per_col_w)

        if cy > 0:
            total_h_pts = (cy / 12700.0) + 6.0
            if total_h_pts <= 260.0:
                row_min_heights[r1] = max(row_min_heights.get(r1, 0.0), total_h_pts)
            else:
                n_rows = max(1, r2 - r1 + 1)
                per_row_h = min(260.0, total_h_pts / n_rows)
                for r in range(r1, r2 + 1):
                    row_min_heights[r] = max(row_min_heights.get(r, 0.0), per_row_h)

    # 2. Xử lý oneCellAnchor
    for anchor in drawing_root.xpath('./*[local-name()="oneCellAnchor"]'):
        from_col_el = anchor.xpath('./*[local-name()="from"]/*[local-name()="col"]')
        from_row_el = anchor.xpath('./*[local-name()="from"]/*[local-name()="row"]')
        ext_el = anchor.xpath('./*[local-name()="ext"]')
        if not (from_col_el and from_row_el and ext_el):
            continue
        try:
            c1 = int(from_col_el[0].text) + 1
            r1 = int(from_row_el[0].text) + 1
            cx = int(ext_el[0].get("cx", "0"))
            cy = int(ext_el[0].get("cy", "0"))
            if cx > 0:
                w_units = min(60.0, (cx / 66675.0) + 3.0)
                col_min_widths[c1] = max(col_min_widths.get(c1, 0.0), w_units)
            if cy > 0:
                h_pts = min(280.0, (cy / 12700.0) + 6.0)
                row_min_heights[r1] = max(row_min_heights.get(r1, 0.0), h_pts)
        except Exception:
            pass

    return col_min_widths, row_min_heights

def _autofit_worksheet_columns_and_rows(s_root, si_list, drawing_root=None):
    """
    Tự động căn chỉnh độ rộng cột (<cols>) và chiều cao dòng (<row>) trong sheet Excel
    sao cho vừa đủ cả văn bản tiếng Việt/Anh sau khi dịch lẫn hình ảnh đính kèm.
    """
    ns = s_root.nsmap.get(None, EXCEL_MAIN_NS)

    # 1. Đọc danh sách các vùng gộp ô (<mergeCells>)
    merged_spans = {}  # (r1, c1) -> (col_span, row_span)
    for mc in s_root.xpath('./*[local-name()="mergeCells"]/*[local-name()="mergeCell"]'):
        ref = mc.get("ref", "")
        if ":" in ref:
            p1, p2 = ref.split(":", 1)
            c1, r1 = _parse_cell_ref(p1)
            c2, r2 = _parse_cell_ref(p2)
            if c1 and r1 and c2 and r2:
                merged_spans[(r1, c1)] = (max(1, c2 - c1 + 1), max(1, r2 - r1 + 1))

    # 2. Lấy kích thước tối thiểu từ các hình ảnh/shapes trên sheet
    img_col_widths, img_row_heights = _extract_drawing_size_requirements(drawing_root)
    col_widths = dict(img_col_widths)

    # 3. Đọc độ rộng cột hiện tại của sheet (không bao giờ thu hẹp cột sẵn có)
    cols_list = s_root.xpath('./*[local-name()="cols"]')
    if cols_list:
        for col_el in cols_list[0].xpath('./*[local-name()="col"]'):
            try:
                c_min = int(col_el.get("min", "1"))
                c_max = int(col_el.get("max", str(c_min)))
                w = float(col_el.get("width", "11.0"))
                if c_max - c_min <= 50:
                    for c in range(c_min, c_max + 1):
                        col_widths[c] = max(col_widths.get(c, 0.0), w)
            except Exception:
                pass

    # 4. Quét toàn bộ văn bản trong các ô để tính độ rộng cột cần thiết
    cell_texts_by_row = {}  # row_idx -> list of (col_idx, col_span, row_span, text)
    for row_el in s_root.xpath('./*[local-name()="sheetData"]/*[local-name()="row"]'):
        try:
            r_idx = int(row_el.get("r", "0"))
        except Exception:
            r_idx = 0

        for c_el in row_el.xpath('./*[local-name()="c"]'):
            c_idx, parsed_r = _parse_cell_ref(c_el.get("r", ""))
            if not c_idx:
                continue
            if not r_idx and parsed_r:
                r_idx = parsed_r

            c_type = c_el.get("t", "")
            txt = ""
            if c_type == "s":
                v_el = c_el.xpath('./*[local-name()="v"]')
                if v_el and v_el[0].text and v_el[0].text.strip().isdigit():
                    si_i = int(v_el[0].text.strip())
                    if 0 <= si_i < len(si_list):
                        txt = _get_si_text(si_list[si_i])
            elif c_type == "inlineStr":
                is_el = c_el.xpath('./*[local-name()="is"]')
                if is_el:
                    txt = _get_inline_is_text(is_el[0])
            else:
                v_el = c_el.xpath('./*[local-name()="v"]')
                if v_el and v_el[0].text:
                    txt = v_el[0].text

            txt = txt.strip()
            if not txt:
                continue

            col_span, row_span = merged_spans.get((r_idx, c_idx), (1, 1))
            vis_w = _excel_visual_width(txt) + 4.5
            per_col_needed = vis_w / col_span
            # Giới hạn độ rộng tự động tối đa là 55 đơn vị (để câu quá dài sẽ tự động xuống dòng)
            capped_w = min(55.0, max(10.0, per_col_needed))
            for c in range(c_idx, c_idx + col_span):
                col_widths[c] = max(col_widths.get(c, 10.0), capped_w)

            if r_idx > 0:
                cell_texts_by_row.setdefault(r_idx, []).append((c_idx, col_span, row_span, txt))

    # 5. Ghi lại thẻ <cols> chuẩn xác ngay trước <sheetData>
    if col_widths:
        sheet_data_list = s_root.xpath('./*[local-name()="sheetData"]')
        if cols_list:
            cols_el = cols_list[0]
            for child in list(cols_el):
                cols_el.remove(child)
        else:
            cols_el = etree.Element(f"{{{ns}}}cols")
            if sheet_data_list:
                idx_sd = list(s_root).index(sheet_data_list[0])
                s_root.insert(idx_sd, cols_el)
            else:
                s_root.insert(0, cols_el)

        for c in sorted(col_widths.keys()):
            w_val = round(col_widths[c], 2)
            col_item = etree.SubElement(cols_el, f"{{{ns}}}col")
            col_item.set("min", str(c))
            col_item.set("max", str(c))
            col_item.set("width", str(w_val))
            col_item.set("customWidth", "1")
            col_item.set("bestFit", "1")

    # 6. Tự động nới chiều cao dòng (<row ht="...">) khi văn bản xuống dòng hoặc có hình ảnh
    sheet_data_list = s_root.xpath('./*[local-name()="sheetData"]')
    existing_rows = {}
    if sheet_data_list:
        sd_el = sheet_data_list[0]
        for row_el in sd_el.xpath('./*[local-name()="row"]'):
            try:
                r_idx = int(row_el.get("r", "0"))
                if r_idx > 0:
                    existing_rows[r_idx] = row_el
            except Exception:
                pass

        # Nếu có dòng chứa hình ảnh nhưng chưa có thẻ <row> trong sheetData thì tạo thêm thẻ <row>
        for img_r in sorted(img_row_heights.keys()):
            if img_r not in existing_rows and img_r > 0:
                new_row = etree.Element(f"{{{ns}}}row")
                new_row.set("r", str(img_r))
                inserted = False
                for child_r in sd_el.xpath('./*[local-name()="row"]'):
                    try:
                        cr_idx = int(child_r.get("r", "0"))
                        if cr_idx > img_r:
                            idx_pos = list(sd_el).index(child_r)
                            sd_el.insert(idx_pos, new_row)
                            inserted = True
                            break
                    except Exception:
                        pass
                if not inserted:
                    sd_el.append(new_row)
                existing_rows[img_r] = new_row

    for r_idx, row_el in existing_rows.items():
        cur_ht = 0.0
        if row_el.get("ht"):
            try:
                cur_ht = float(row_el.get("ht"))
            except Exception:
                cur_ht = 0.0

        max_lines_in_row = 1
        for c_idx, col_span, row_span, txt in cell_texts_by_row.get(r_idx, []):
            if row_span > 1:
                continue
            avail_col_w = sum(col_widths.get(c, 12.0) for c in range(c_idx, c_idx + col_span))
            avail_text_w = max(8.0, avail_col_w - 3.0)
            total_lines = 0
            for raw_line in txt.split('\n'):
                lw = _excel_visual_width(raw_line)
                total_lines += max(1, int(math.ceil(lw / avail_text_w)))
            if total_lines > max_lines_in_row:
                max_lines_in_row = total_lines

        req_text_ht = (max_lines_in_row * 16.5 + 5.0) if max_lines_in_row > 1 else 20.0
        req_img_ht = img_row_heights.get(r_idx, 0.0)
        final_ht = max(cur_ht, req_text_ht, req_img_ht)

        if final_ht > cur_ht or max_lines_in_row > 1 or req_img_ht > 0:
            row_el.set("ht", f"{final_ht:.1f}")
            row_el.set("customHeight", "1")


def _build_excel_xml_context(file_bytes, include_shapes=True):
    in_zip = zipfile.ZipFile(io.BytesIO(file_bytes), 'r')
    archive_files = {name: in_zip.read(name) for name in in_zip.namelist()}
    in_zip.close()

    xml_trees = {}
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)

    # 1. Đọc sharedStrings.xml
    sst_path = "xl/sharedStrings.xml"
    si_list = []
    if sst_path in archive_files:
        sst_root = etree.fromstring(archive_files[sst_path], parser=parser)
        xml_trees[sst_path] = sst_root
        si_list = sst_root.xpath('./*[local-name()="si"]')
    else:
        sst_root = None

    # 2. Đọc styles.xml để bật tự động xuống dòng (wrapText)
    styles_path = "xl/styles.xml"
    if styles_path in archive_files:
        styles_root = etree.fromstring(archive_files[styles_path], parser=parser)
        xml_trees[styles_path] = styles_root
    else:
        styles_root = None

    # 3. Xác định danh sách sheet theo thứ tự
    sheet_paths = []
    wb_path = "xl/workbook.xml"
    wb_rels_path = "xl/_rels/workbook.xml.rels"
    if wb_path in archive_files and wb_rels_path in archive_files:
        wb_root = etree.fromstring(archive_files[wb_path], parser=parser)
        rels_root = etree.fromstring(archive_files[wb_rels_path], parser=parser)
        rid_to_target = {}
        for rel in rels_root.xpath('./*[local-name()="Relationship"]'):
            rid = rel.get("Id")
            target = rel.get("Target", "").lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            rid_to_target[rid] = os.path.normpath(target).replace("\\", "/")
        for sheet_el in wb_root.xpath('.//*[local-name()="sheet"]'):
            rid = sheet_el.get(f"{{{EXCEL_REL_NS}}}id")
            if rid in rid_to_target and rid_to_target[rid] in archive_files:
                sheet_paths.append(rid_to_target[rid])

    if not sheet_paths:
        sheet_paths = sorted([
            p for p in archive_files
            if p.startswith("xl/worksheets/sheet") and p.endswith(".xml")
        ])

    # 4. Đọc toàn bộ drawing*.xml và liên kết sheet -> drawing
    drawing_paths = sorted([
        p for p in archive_files
        if p.startswith("xl/drawings/drawing") and p.endswith(".xml")
    ])
    for dp in drawing_paths:
        xml_trees[dp] = etree.fromstring(archive_files[dp], parser=parser)

    sheet_to_drawing = {}
    for sp in sheet_paths:
        sheet_dir = os.path.dirname(sp)
        sheet_base = os.path.basename(sp)
        s_rels_path = f"{sheet_dir}/_rels/{sheet_base}.rels"
        if s_rels_path in archive_files:
            sr_root = etree.fromstring(archive_files[s_rels_path], parser=parser)
            for rel in sr_root.xpath('./*[local-name()="Relationship"]'):
                t_val = rel.get("Target", "")
                if "drawings/drawing" in t_val:
                    norm_dp = os.path.normpath(os.path.join(sheet_dir, t_val)).replace("\\", "/")
                    if norm_dp in xml_trees:
                        sheet_to_drawing[sp] = xml_trees[norm_dp]

    items = []
    translated_si_map = {}

    def ensure_sst_root():
        nonlocal sst_root
        if sst_root is None:
            sst_root = etree.Element(f"{{{EXCEL_MAIN_NS}}}sst")
            xml_trees[sst_path] = sst_root
            ct_path = "[Content_Types].xml"
            if ct_path in archive_files:
                ct_root = xml_trees.get(ct_path) or etree.fromstring(archive_files[ct_path], parser=parser)
                xml_trees[ct_path] = ct_root
                existing = ct_root.xpath('.//*[local-name()="Override" and @PartName="/xl/sharedStrings.xml"]')
                if not existing:
                    ov = etree.SubElement(ct_root, f"{{{ct_root.nsmap.get(None, '')}}}Override")
                    ov.set("PartName", "/xl/sharedStrings.xml")
                    ov.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        return sst_root

    def apply_cell_translation(target_info, translated_text):
        clean_val = sanitize_xml_text(translated_text)
        kind = target_info['kind']

        if kind == 's':
            orig_idx = target_info['si_idx']
            key = (orig_idx, clean_val)
            if key in translated_si_map:
                new_idx = translated_si_map[key]
            else:
                root = ensure_sst_root()
                orig_si = si_list[orig_idx] if orig_idx < len(si_list) else None
                new_si = etree.Element(f"{{{EXCEL_MAIN_NS}}}si")
                first_rpr = None
                if orig_si is not None:
                    rpr_matches = orig_si.xpath('./*[local-name()="r"]/*[local-name()="rPr"]')
                    if rpr_matches:
                        first_rpr = copy.deepcopy(rpr_matches[0])
                if first_rpr is not None:
                    r_el = etree.SubElement(new_si, f"{{{EXCEL_MAIN_NS}}}r")
                    r_el.append(first_rpr)
                    t_el = etree.SubElement(r_el, f"{{{EXCEL_MAIN_NS}}}t")
                    t_el.set(XML_SPACE_ATTR, "preserve")
                    t_el.text = clean_val
                else:
                    t_el = etree.SubElement(new_si, f"{{{EXCEL_MAIN_NS}}}t")
                    t_el.set(XML_SPACE_ATTR, "preserve")
                    t_el.text = clean_val

                root.append(new_si)
                new_idx = len(si_list)
                si_list.append(new_si)
                translated_si_map[key] = new_idx

            target_info['v_elem'].text = str(new_idx)

        elif kind == 'inlineStr':
            is_elem = target_info['is_elem']
            for child in list(is_elem):
                is_elem.remove(child)
            t_el = etree.SubElement(is_elem, f"{{{EXCEL_MAIN_NS}}}t")
            t_el.set(XML_SPACE_ATTR, "preserve")
            t_el.text = clean_val

        elif kind == 'str':
            target_info['v_elem'].text = clean_val

        elif kind == 'drawing_p':
            a_p = target_info['a_p']
            runs = a_p.xpath('./*[local-name()="r" or local-name()="fld"]')
            if runs:
                first_run = runs[0]
                t_list = first_run.xpath('./*[local-name()="t"]')
                if t_list:
                    t_list[0].text = clean_val
                else:
                    t_el = etree.SubElement(first_run, f"{{{DRAWING_A_NS}}}t")
                    t_el.text = clean_val
                for extra_r in runs[1:]:
                    a_p.remove(extra_r)

    # 5. Quét toàn bộ ô trong từng Worksheet
    for sp in sheet_paths:
        s_root = etree.fromstring(archive_files[sp], parser=parser)
        xml_trees[sp] = s_root
        for c_elem in s_root.xpath('.//*[local-name()="sheetData"]//*[local-name()="c"]'):
            if len(c_elem.xpath('./*[local-name()="f"]')) > 0:
                continue
            c_type = c_elem.get("t", "")
            if c_type == "s":
                v_list = c_elem.xpath('./*[local-name()="v"]')
                if v_list and v_list[0].text is not None:
                    try:
                        si_idx = int(v_list[0].text.strip())
                        if 0 <= si_idx < len(si_list):
                            txt = sanitize_xml_text(_get_si_text(si_list[si_idx])).strip()
                            if txt and not txt.startswith('='):
                                items.append({
                                    'obj': {
                                        'kind': 's',
                                        'c_elem': c_elem,
                                        'v_elem': v_list[0],
                                        'si_idx': si_idx,
                                    },
                                    'text': txt,
                                    'apply_func': apply_cell_translation,
                                })
                    except ValueError:
                        pass
            elif c_type == "inlineStr":
                is_list = c_elem.xpath('./*[local-name()="is"]')
                if is_list:
                    txt = sanitize_xml_text(_get_inline_is_text(is_list[0])).strip()
                    if txt and not txt.startswith('='):
                        items.append({
                            'obj': {
                                'kind': 'inlineStr',
                                'c_elem': c_elem,
                                'is_elem': is_list[0],
                            },
                            'text': txt,
                            'apply_func': apply_cell_translation,
                        })
            elif c_type == "str":
                v_list = c_elem.xpath('./*[local-name()="v"]')
                if v_list and v_list[0].text:
                    txt = sanitize_xml_text(v_list[0].text).strip()
                    if txt and not txt.startswith('='):
                        items.append({
                            'obj': {
                                'kind': 'str',
                                'c_elem': c_elem,
                                'v_elem': v_list[0],
                            },
                            'text': txt,
                            'apply_func': apply_cell_translation,
                        })

    # 6. Quét thêm chữ trong các khối Hình vẽ / TextBox (nếu không bật skip_shapes)
    if include_shapes:
        for dp in drawing_paths:
            d_root = xml_trees[dp]
            for a_p in d_root.xpath('.//*[local-name()="txBody"]/*[local-name()="p"]'):
                t_texts = [
                    t_el.text
                    for t_el in a_p.xpath('./*[local-name()="r" or local-name()="fld"]/*[local-name()="t"]')
                    if t_el.text
                ]
                full_p_text = sanitize_xml_text("".join(t_texts)).strip()
                if full_p_text:
                    items.append({
                        'obj': {
                            'kind': 'drawing_p',
                            'c_elem': None,
                            'a_p': a_p,
                        },
                        'text': full_p_text,
                        'apply_func': apply_cell_translation,
                    })

    def serialize_to_bytes(autofit=True):
        if autofit:
            _enable_wrap_text_in_styles(styles_root)
            for sp in sheet_paths:
                _autofit_worksheet_columns_and_rows(
                    xml_trees[sp],
                    si_list,
                    drawing_root=sheet_to_drawing.get(sp),
                )

        if sst_root is not None:
            if "count" in sst_root.attrib:
                del sst_root.attrib["count"]
            if "uniqueCount" in sst_root.attrib:
                sst_root.set("uniqueCount", str(len(si_list)))

        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, orig_data in archive_files.items():
                if name in xml_trees:
                    xml_bytes = etree.tostring(
                        xml_trees[name],
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    zf.writestr(name, xml_bytes)
                else:
                    zf.writestr(name, orig_data)
            if sst_path in xml_trees and sst_path not in archive_files:
                xml_bytes = etree.tostring(
                    xml_trees[sst_path],
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
                zf.writestr(sst_path, xml_bytes)
        return out_buf.getvalue()

    return items, serialize_to_bytes

def preview_excel(file_bytes, skip_shapes=False):
    excel_items, _ = _build_excel_xml_context(file_bytes, include_shapes=not skip_shapes)
    return [it['text'] for it in excel_items]

def process_excel(
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
    autosave_filename = f"autosave_{file_name}"
    all_items, serialize_to_bytes = _build_excel_xml_context(
        file_bytes, include_shapes=not skip_shapes
    )

    def autosave_excel():
        try:
            atomic_save_bytes(autosave_filename, serialize_to_bytes(autofit=False))
        except Exception:
            pass

    for item in all_items:
        item['autosave_func'] = autosave_excel

    actual_start = max(0, start_idx - 1) if dich_theo_vung else 0
    actual_end = end_idx if (dich_theo_vung and end_idx > 0) else len(all_items)
    target_items = all_items[actual_start:actual_end]

    if target_items:
        status_text.info(f"📊 Tổng: {len(all_items)} mục Excel. Đang dịch: {len(target_items)} mục.")
        execute_translation_batches(
            target_items,
            model_name,
            progress_bar,
            status_text,
            preview_box,
            lang_mode=lang_mode,
            smart_resume=smart_resume,
        )

    status_text.info("📐 Đang tự động căn chỉnh độ rộng cột, chiều cao dòng & khung hình ảnh Excel...")
    result_bytes = serialize_to_bytes(autofit=True)
    if os.path.exists(autosave_filename):
        try:
            os.remove(autosave_filename)
        except Exception:
            pass
    return result_bytes
