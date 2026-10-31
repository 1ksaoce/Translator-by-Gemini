import hashlib
import io
import json
import os
import re
import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import google.generativeai as genai
from .translator import (
    SAFETY_SETTINGS,
    LANG_MODES,
    CHINESE_CHARS_RE,
    sanitize_xml_text,
    handle_api_error,
)

_IMAGE_BLOB_CACHE = {}


def _get_image_font(size):
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "C:/Windows/Fonts/times.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    px_size = max(9, int(round(size)))
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, px_size)
            except Exception:
                pass
    return ImageFont.load_default()


def _measure_pil_text(draw, text, font):
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        return max(1, bbox[2] - bbox[0]), max(1, bbox[3] - bbox[1])
    except Exception:
        return len(text) * 8, 14


def _wrap_pil_text(draw, text, font, max_w):
    words = text.split()
    if not words:
        return [text]
    lines = []
    cur = words[0]
    for w in words[1:]:
        cand = f"{cur} {w}"
        cw, _ = _measure_pil_text(draw, cand, font)
        if cw <= max_w:
            cur = cand
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    return lines


def render_translated_image(pil_img, detections):
    """
    Xóa chữ cũ trong từng vùng bounding box bằng OpenCV + màu nền cục bộ,
    bảo vệ các đường gióng chỉ dẫn / đường viền bảng, sau đó chèn chữ đã dịch
    vào đúng vị trí & đúng chiều (hỗ trợ cả chữ nằm ngang và chữ xoay dọc 90 độ trên bản vẽ).
    """
    rgb_img = pil_img.convert("RGB")
    img_np = np.array(rgb_img, dtype=np.uint8)
    H, W, _ = img_np.shape

    valid_boxes = []
    for item in detections:
        box = item.get("box_2d")
        trans_txt = sanitize_xml_text(item.get("translated", "")).strip()
        orig_txt = sanitize_xml_text(item.get("original", "")).strip()
        if not box or not isinstance(box, (list, tuple)) or len(box) != 4 or not trans_txt:
            continue
        try:
            ymin, xmin, ymax, xmax = [float(v) for v in box]
            x1 = int(max(0, min(W - 2, round(xmin * W / 1000.0))))
            y1 = int(max(0, min(H - 2, round(ymin * H / 1000.0))))
            x2 = int(max(x1 + 4, min(W, round(xmax * W / 1000.0))))
            y2 = int(max(y1 + 4, min(H, round(ymax * H / 1000.0))))
            valid_boxes.append({
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,
                "translated": trans_txt,
                "original": orig_txt,
                "orientation": str(item.get("orientation", "")).lower(),
            })
        except Exception:
            continue

    # BƯỚC 1: Xóa chữ cũ trên ảnh, bảo vệ đường kẻ gióng / viền bảng nằm sát mép chữ
    for b in valid_boxes:
        x1, y1, x2, y2 = b["x1"], b["y1"], b["x2"], b["y2"]
        pad = 2
        ox1 = max(0, x1 - pad)
        oy1 = max(0, y1 - pad)
        ox2 = min(W, x2 + pad)
        oy2 = min(H, y2 + pad)

        roi_outer = img_np[oy1:oy2, ox1:ox2]
        if roi_outer.shape[0] > 2 and roi_outer.shape[1] > 2:
            border_concat = np.concatenate([
                roi_outer[0, :, :],
                roi_outer[-1, :, :],
                roi_outer[:, 0, :],
                roi_outer[:, -1, :],
            ], axis=0)
            bg_rgb = np.median(border_concat, axis=0).astype(np.uint8)
            bg_std = float(np.std(border_concat))
        else:
            bg_rgb = np.array([255, 255, 255], dtype=np.uint8)
            bg_std = 0.0

        b["bg_color"] = (int(bg_rgb[0]), int(bg_rgb[1]), int(bg_rgb[2]))
        lum = 0.299 * bg_rgb[0] + 0.587 * bg_rgb[1] + 0.114 * bg_rgb[2]
        b["text_color"] = (12, 12, 12) if lum > 135 else (250, 250, 250)

        ix1, iy1, ix2, iy2 = x1, y1, x2, y2
        roi = img_np[iy1:iy2, ix1:ix2]
        if roi.size == 0:
            b["ix1"], b["iy1"], b["ix2"], b["iy2"] = ix1, iy1, ix2, iy2
            continue

        rh, rw, _ = roi.shape
        dark_mask = np.linalg.norm(roi.astype(np.float32) - bg_rgb.astype(np.float32), axis=2) > 40.0
        while ix1 < ix2 - 4 and rw > 4 and np.mean(dark_mask[:, ix1 - x1]) > 0.75:
            ix1 += 1
        while ix2 > ix1 + 4 and rw > 4 and np.mean(dark_mask[:, ix2 - 1 - x1]) > 0.75:
            ix2 -= 1
        while iy1 < iy2 - 4 and rh > 4 and np.mean(dark_mask[iy1 - y1, :]) > 0.75:
            iy1 += 1
        while iy2 > iy1 + 4 and rh > 4 and np.mean(dark_mask[iy2 - 1 - y1, :]) > 0.75:
            iy2 -= 1

        b["ix1"], b["iy1"], b["ix2"], b["iy2"] = ix1, iy1, ix2, iy2
        roi = img_np[iy1:iy2, ix1:ix2]
        if roi.size == 0:
            continue

        if bg_std < 35.0:
            img_np[iy1:iy2, ix1:ix2] = bg_rgb
        else:
            diff = np.linalg.norm(roi.astype(np.float32) - bg_rgb.astype(np.float32), axis=2)
            mask = (diff > 28.0).astype(np.uint8) * 255
            kernel = np.ones((3, 3), np.uint8)
            mask = cv2.dilate(mask, kernel, iterations=1)
            inpainted = cv2.inpaint(roi, mask, 3, cv2.INPAINT_TELEA)
            blended = (
                inpainted.astype(np.float32) * 0.42 + bg_rgb.astype(np.float32) * 0.58
            ).astype(np.uint8)
            img_np[iy1:iy2, ix1:ix2] = blended

    # BƯỚC 2: Chèn chữ đã dịch vào đúng vị trí & hướng chữ (ngang hoặc dọc 90 độ)
    out_pil = Image.fromarray(img_np)
    draw = ImageDraw.Draw(out_pil)

    tall_count = sum(1 for b in valid_boxes if (b["y2"] - b["y1"]) > (b["x2"] - b["x1"]) * 1.5)
    mostly_vertical_diagram = len(valid_boxes) >= 3 and tall_count >= len(valid_boxes) * 0.4

    for b in valid_boxes:
        x1, y1, x2, y2 = b["x1"], b["y1"], b["x2"], b["y2"]
        ix1 = b.get("ix1", x1)
        trans_txt = b["translated"]
        box_w = max(10, x2 - x1)
        box_h = max(10, y2 - y1)

        ratio_thresh = 1.15 if mostly_vertical_diagram else 1.65
        is_vertical = (
            b["orientation"] in ("vertical", "vertical_cw", "90", "270")
            or (box_h > box_w * ratio_thresh and len(trans_txt) >= 2)
        )

        if is_vertical:
            max_y_below = H - 6
            for other in valid_boxes:
                if other is b:
                    continue
                if other["y1"] > y1 and abs(other["x1"] - x1) < box_w * 1.5:
                    max_y_below = min(max_y_below, other["y1"] - 4)

            avail_len = max(box_h, min(max_y_below - y1, max(box_h + 24, int(box_h * 1.55))))
            avail_thick = max(18, int(box_w * 1.5))
            start_fs = min(28, max(11, int(box_w * 0.72)))
            best_font = None
            best_lines = None

            # Ưu tiên giữ trên 1 dòng dọc duy nhất để không tràn ngang vào đường gióng
            for fs in range(start_fs, 9, -1):
                font_cand = _get_image_font(fs)
                w_single, _ = _measure_pil_text(draw, trans_txt, font_cand)
                if w_single <= avail_len:
                    best_font = font_cand
                    best_lines = [trans_txt]
                    break

            if best_font is None:
                for fs in range(start_fs, 8, -1):
                    font_cand = _get_image_font(fs)
                    lines_cand = _wrap_pil_text(draw, trans_txt, font_cand, avail_len - 2)
                    max_lw = max(_measure_pil_text(draw, ln, font_cand)[0] for ln in lines_cand)
                    _, lh = _measure_pil_text(draw, "Ag", font_cand)
                    total_lh = len(lines_cand) * (lh + 2)
                    if max_lw <= avail_len and (total_lh <= avail_thick or len(lines_cand) <= 2):
                        best_font = font_cand
                        best_lines = lines_cand
                        break
                else:
                    best_font = _get_image_font(9)
                    best_lines = _wrap_pil_text(draw, trans_txt, best_font, avail_len)

            _, line_h = _measure_pil_text(draw, "Ag", best_font)
            step_h = line_h + 2
            max_lw = max(_measure_pil_text(draw, ln, best_font)[0] for ln in best_lines)
            total_th = step_h * len(best_lines)

            canvas_w = max(8, max_lw + 6)
            canvas_h = max(8, total_th + 4)
            txt_canvas = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
            t_draw = ImageDraw.Draw(txt_canvas)

            cy = 2
            for ln in best_lines:
                t_draw.text(
                    (2, cy),
                    ln,
                    fill=b["text_color"] + (255,),
                    font=best_font,
                    stroke_width=1,
                    stroke_fill=b["bg_color"] + (255,),
                )
                cy += step_h

            rot_canvas = txt_canvas.transpose(Image.Transpose.ROTATE_270)
            px = max(2, min(W - rot_canvas.width - 2, ix1 + 1))
            py = max(2, min(H - rot_canvas.height - 2, y1))
            out_pil.paste(rot_canvas, (px, py), rot_canvas)
        else:
            avail_w = max(box_w, min(W - x1 - 4, int(box_w * 1.25)))
            start_fs = min(44, max(11, int(box_h * 0.78)))
            best_font = _get_image_font(11)
            best_lines = [trans_txt]

            for fs in range(start_fs, 9, -1):
                font_cand = _get_image_font(fs)
                lines_cand = _wrap_pil_text(draw, trans_txt, font_cand, avail_w - 2)
                max_lw = max(_measure_pil_text(draw, ln, font_cand)[0] for ln in lines_cand)
                _, lh = _measure_pil_text(draw, "Ag", font_cand)
                total_lh = len(lines_cand) * (lh + 2)
                if max_lw <= avail_w and (total_lh <= box_h * 1.35 or len(lines_cand) == 1):
                    best_font = font_cand
                    best_lines = lines_cand
                    break
            else:
                best_font = _get_image_font(10)
                best_lines = _wrap_pil_text(draw, trans_txt, best_font, avail_w)

            _, line_h = _measure_pil_text(draw, "Ag", best_font)
            step_h = line_h + 2
            total_text_h = step_h * len(best_lines) - 2
            cur_y = y1 + max(0, (box_h - total_text_h) // 2)

            for ln in best_lines:
                lw, _ = _measure_pil_text(draw, ln, best_font)
                cur_x = x1 + max(0, (box_w - lw) // 2)
                if cur_x + lw > W - 2:
                    cur_x = max(2, W - 2 - lw)
                draw.text(
                    (cur_x, cur_y),
                    ln,
                    fill=b["text_color"],
                    font=best_font,
                    stroke_width=1,
                    stroke_fill=b["bg_color"],
                )
                cur_y += step_h

    return out_pil


def process_image_translation(pil_img, model_name, lang_mode="zh_to_vi"):
    """
    Quét ảnh bằng Gemini Vision:
    1. Trích xuất tọa độ bounding box [ymin, xmin, ymax, xmax] + chữ gốc + chữ dịch
       (hỗ trợ cả ảnh nằm ngang và bản vẽ kỹ thuật xoay dọc 90 độ)
    2. Xóa chữ cũ trên ảnh và chèn chữ đã dịch vào đúng vị trí
    3. Trả về (Ảnh đã dịch, chuỗi văn bản liệt kê vị trí)
    """
    mode_cfg = LANG_MODES.get(lang_mode, LANG_MODES["zh_to_vi"])
    if lang_mode == "zh_to_vi":
        task_desc = "Trích xuất tất cả các cụm chữ tiếng Trung trong ảnh/bản vẽ và dịch sang tiếng Việt."
    elif lang_mode == "zh_to_en":
        task_desc = "Extract all Chinese text blocks in the image/diagram and translate them into concise technical English."
    elif lang_mode == "en_to_vi":
        task_desc = "Trích xuất tất cả các cụm chữ tiếng Anh trong ảnh/bản vẽ và dịch sang tiếng Việt."
    else:
        task_desc = f"Trích xuất tất cả các cụm chữ ({mode_cfg['source_desc']}) trong ảnh/bản vẽ và dịch sang {mode_cfg['target_desc']}."

    prompt = f"""{task_desc}
Lưu ý quan trọng: Hình ảnh có thể là bản vẽ kỹ thuật nằm ngang hoặc xoay dọc 90 độ (chữ chạy dọc từ trên xuống dưới). Hãy phát hiện đầy đủ mọi chú thích/nhãn chữ trên bản vẽ.
Nếu không có bất kỳ chữ nào cần dịch trong ảnh, hãy trả về mảng rỗng: []

YÊU CẦU ĐẦU RA BẮT BUỘC:
Trả về ĐÚNG MỘT mảng JSON hợp lệ (KHÔNG giải thích thêm), mỗi phần tử có cấu trúc:
[
  {{
    "box_2d": [ymin, xmin, ymax, xmax],
    "orientation": "horizontal hoặc vertical",
    "position": "Mô tả vị trí tương đối bằng tiếng Việt",
    "original": "Chữ gốc trong ảnh",
    "translated": "Bản dịch tương ứng"
  }}
]
Trong đó box_2d là tọa độ khung chứa chữ chuẩn hóa từ 0 đến 1000 theo thứ tự [ymin, xmin, ymax, xmax]."""

    detections = []
    raw_fallback_text = ""

    for attempt in range(4):
        try:
            model = genai.GenerativeModel(model_name)
            res = model.generate_content(
                [prompt, pil_img],
                safety_settings=SAFETY_SETTINGS,
            )
            if res.text:
                raw_fallback_text = res.text.strip()
                clean_json = raw_fallback_text.strip("` \n")
                if clean_json.lower().startswith("json"):
                    clean_json = clean_json[4:].strip()
                m = re.search(r'\[.*\]', clean_json, re.DOTALL)
                if m:
                    clean_json = m.group(0)
                parsed = json.loads(clean_json)
                if isinstance(parsed, list):
                    detections = parsed
                    break
        except Exception as e:
            if attempt < 3:
                handle_api_error(e, attempt)

    # Nếu chế độ chỉ dịch tiếng Trung, lọc bỏ các mục không có chữ Hán
    if lang_mode in ("zh_to_vi", "zh_to_en") and detections:
        detections = [
            d for d in detections
            if CHINESE_CHARS_RE.search(str(d.get("original", "")))
        ]

    if detections:
        translated_img = render_translated_image(pil_img, detections)
        lines_out = []
        for d in detections:
            pos = sanitize_xml_text(d.get("position", "Vị trí")).strip()
            orig = sanitize_xml_text(d.get("original", "")).strip()
            trans = sanitize_xml_text(d.get("translated", "")).strip()
            if orig or trans:
                lines_out.append(f"- [{pos}] {orig} -> {trans}")
        formatted_text = "\n".join(lines_out) if lines_out else raw_fallback_text
        return translated_img, formatted_text

    return pil_img, ""


def translate_docx_embedded_images(
    doc,
    model_name,
    status_text=None,
    preview_box=None,
    lang_mode="zh_to_vi",
):
    """
    Quét toàn bộ các hình ảnh / sơ đồ đính kèm trong tài liệu Word (.docx)
    và tự động dịch trực tiếp chữ bên trong hình ảnh (xóa chữ cũ, chèn chữ mới)
    khi người dùng KHÔNG bật 'Bỏ qua chữ trong Lưu đồ'.
    """
    r_ids = []
    seen_rids = set()
    for r_id in doc.element.xpath(
        './/*[local-name()="blip"]/@*[local-name()="embed"] | '
        './/*[local-name()="imagedata"]/@*[local-name()="id"]'
    ):
        if r_id not in seen_rids:
            seen_rids.add(r_id)
            r_ids.append(r_id)

    candidate_parts = []
    for r_id in r_ids:
        part = doc.part.related_parts.get(r_id)
        if part is None or not hasattr(part, "blob"):
            continue
        try:
            pil_img = Image.open(io.BytesIO(part.blob))
            w, h = pil_img.size
            # Bỏ qua các icon nhỏ / logo trang trí quá bé
            if w < 70 or h < 70:
                continue
            gray_arr = np.array(pil_img.convert("L"))
            if float(np.std(gray_arr)) < 4.0:
                continue
            candidate_parts.append((r_id, part, pil_img.convert("RGB")))
        except Exception:
            continue

    if not candidate_parts:
        return

    total_imgs = len(candidate_parts)
    for idx, (r_id, part, rgb_img) in enumerate(candidate_parts, start=1):
        try:
            blob_hash = hashlib.sha1(part.blob).hexdigest()
            cache_key = f"{lang_mode}||{blob_hash}"
            if cache_key in _IMAGE_BLOB_CACHE:
                part._blob = _IMAGE_BLOB_CACHE[cache_key]
                continue

            if status_text is not None:
                status_text.info(
                    f"🖼️ Đang quét & dịch chữ trực tiếp trong Hình ảnh / Lưu đồ ({idx}/{total_imgs})..."
                )

            trans_img, summary_txt = process_image_translation(
                rgb_img, model_name, lang_mode=lang_mode
            )
            if summary_txt:
                out_buf = io.BytesIO()
                trans_img.save(out_buf, format="PNG")
                new_blob = out_buf.getvalue()
                part._blob = new_blob
                _IMAGE_BLOB_CACHE[cache_key] = new_blob
                if preview_box is not None:
                    preview_box.success(
                        f"**🖼️ Đã dịch chữ trong Hình ảnh / Lưu đồ ({idx}/{total_imgs}):**\n\n{summary_txt[:600]}"
                    )
            else:
                _IMAGE_BLOB_CACHE[cache_key] = part.blob
        except Exception:
            continue
