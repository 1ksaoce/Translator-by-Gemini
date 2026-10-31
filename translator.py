import json
import os
import re
import time
import streamlit as st
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold

# ==========================================
# BỘ LỌC AN TOÀN - TẮT 100% KIỂM DUYỆT
# ==========================================
SAFETY_SETTINGS = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}

# ==========================================
# CÁC CHẾ ĐỘ NGÔN NGỮ DỊCH
# ==========================================
LANG_MODES = {
    "zh_to_vi": {
        "label": "🇨🇳 Chỉ dịch Tiếng Trung ➔ 🇻🇳 Tiếng Việt",
        "source_desc": "tiếng Trung",
        "target_desc": "tiếng Việt",
        "only_chinese": True,
    },
    "all_to_vi": {
        "label": "🌐 Tất cả ngôn ngữ ➔ 🇻🇳 Tiếng Việt",
        "source_desc": "mọi ngôn ngữ (Trung, Anh, Nhật, Hàn...)",
        "target_desc": "tiếng Việt",
        "only_chinese": False,
    },
    "zh_to_en": {
        "label": "🇨🇳 Tiếng Trung ➔ 🇬🇧 Tiếng Anh",
        "source_desc": "tiếng Trung",
        "target_desc": "tiếng Anh (Technical English)",
        "only_chinese": True,
    },
    "en_to_vi": {
        "label": "🇬🇧 Tiếng Anh ➔ 🇻🇳 Tiếng Việt",
        "source_desc": "tiếng Anh",
        "target_desc": "tiếng Việt",
        "only_chinese": False,
    },
}

_GLOBAL_TRANSLATION_CACHE = {}
_API_KEYS_POOL = []  # List of dicts: [{"label": "API 1", "key": "..."}, ...]
_ACTIVE_API_IDX = 0


def configure_api_keys_pool(key_entries, preferred_idx=0):
    """
    Cấu hình danh sách tối đa 3 API Key và đặt API được chọn làm API chính.
    Các API còn lại đóng vai trò dự phòng tự động (Auto-Failover) khi API chính bị lỗi hoặc hết hạn mức.
    """
    global _API_KEYS_POOL, _ACTIVE_API_IDX
    valid_entries = []
    # Đưa API được người dùng chọn lên đầu danh sách ưu tiên
    if 0 <= preferred_idx < len(key_entries):
        pref = key_entries[preferred_idx]
        if pref.get("key", "").strip():
            valid_entries.append({
                "label": pref.get("label", f"API {preferred_idx + 1}"),
                "key": pref["key"].strip(),
            })
    for idx, entry in enumerate(key_entries):
        if idx == preferred_idx:
            continue
        k = entry.get("key", "").strip()
        if k and all(x["key"] != k for x in valid_entries):
            valid_entries.append({
                "label": entry.get("label", f"API {idx + 1}"),
                "key": k,
            })

    _API_KEYS_POOL = valid_entries
    _ACTIVE_API_IDX = 0
    if _API_KEYS_POOL:
        genai.configure(api_key=_API_KEYS_POOL[0]["key"])
        return _API_KEYS_POOL[0]["key"]
    return ""


def get_active_api_label():
    if not _API_KEYS_POOL:
        return "Chưa có API"
    idx = _ACTIVE_API_IDX % len(_API_KEYS_POOL)
    return _API_KEYS_POOL[idx]["label"]


def rotate_to_backup_api():
    """Tự động chuyển sang API dự phòng tiếp theo nếu có từ 2 API trở lên."""
    global _ACTIVE_API_IDX
    if len(_API_KEYS_POOL) <= 1:
        return False
    _ACTIVE_API_IDX = (_ACTIVE_API_IDX + 1) % len(_API_KEYS_POOL)
    next_entry = _API_KEYS_POOL[_ACTIVE_API_IDX]
    genai.configure(api_key=next_entry["key"])
    try:
        st.toast(
            f"🔑 API hiện tại gặp sự cố/hết hạn mức ➔ Đã tự động chuyển sang {next_entry['label']}!",
            icon="🔄",
        )
    except Exception:
        pass
    return True


def get_translation_cache():
    global _GLOBAL_TRANSLATION_CACHE
    try:
        if "translation_cache" not in st.session_state:
            st.session_state.translation_cache = _GLOBAL_TRANSLATION_CACHE
        else:
            _GLOBAL_TRANSLATION_CACHE = st.session_state.translation_cache
    except Exception:
        pass
    return _GLOBAL_TRANSLATION_CACHE

# ==========================================
# CHUẨN HÓA KÝ HIỆU ĐẶC BIỆT (CHỐNG LỖI Ô VUÔNG □ TRONG PDF & WORD)
# ==========================================
ILLEGAL_XML_CHARS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

# Bảng ánh xạ các ký tự đặc biệt hay bị lỗi ô vuông (□) trong font Times New Roman / Arial
SPECIAL_SYMBOL_MAP = {
    # Đơn vị đo & Kỹ thuật
    '℃': '°C',
    '℉': '°F',
    '⌀': 'Φ',
    '∅': 'Φ',
    'φ': 'Φ',
    'ϕ': 'Φ',
    'Ω': 'Ω',
    'µ': 'μ',
    'ℓ': 'l',
    '‰': '‰',
    '±': '±',
    '×': '×',
    '÷': '÷',
    '≤': '≤',
    '≥': '≥',
    '≦': '≤',
    '≧': '≥',
    '≠': '≠',
    '≈': '≈',
    '≒': '≈',
    '≡': '≡',
    '√': '√',
    '∞': '∞',
    # Dấu ngoặc & chấm câu CJK
    '【': '[',
    '】': ']',
    '〖': '[',
    '〗': ']',
    '〔': '[',
    '〕': ']',
    '《': '«',
    '》': '»',
    '〈': '<',
    '〉': '>',
    '「': '"',
    '」': '"',
    '『': '"',
    '』': '"',
    '、': ', ',
    '。': '.',
    '：': ':',
    '；': ';',
    '（': '(',
    '）': ')',
    '！': '!',
    '？': '?',
    '～': '~',
    '—': '-',
    '–': '-',
    '…': '...',
    # Biểu tượng hình học / Checkbox / Mũi tên
    '☑': '[√]',
    '✓': '√',
    '✔': '√',
    '✗': '[X]',
    '✘': '[X]',
    '☒': '[X]',
    '☐': '[ ]',
    '□': '[ ]',
    '■': '•',
    '●': '•',
    '○': 'o',
    '◆': '•',
    '◇': 'o',
    '▲': '•',
    '△': 'Δ',
    '▼': '•',
    '▽': '∇',
    '★': '*',
    '☆': '*',
    '❖': '•',
    '➢': '->',
    '➣': '->',
    '➤': '->',
    '➔': '->',
    '➜': '->',
    '➡': '->',
    '⇒': '=>',
    '⇐': '<=',
    '⇔': '<=>',
    # Số La Mã Unicode (Ⅰ..Ⅻ)
    'Ⅰ': 'I', 'Ⅱ': 'II', 'Ⅲ': 'III', 'Ⅳ': 'IV', 'Ⅴ': 'V', 'Ⅵ': 'VI',
    'Ⅶ': 'VII', 'Ⅷ': 'VIII', 'Ⅸ': 'IX', 'Ⅹ': 'X', 'Ⅺ': 'XI', 'Ⅻ': 'XII',
    'ⅰ': 'i', 'ⅱ': 'ii', 'ⅲ': 'iii', 'ⅳ': 'iv', 'ⅴ': 'v', 'ⅵ': 'vi',
    'ⅶ': 'vii', 'ⅷ': 'viii', 'ⅸ': 'ix', 'ⅹ': 'x', 'ⅺ': 'xi', 'ⅻ': 'xii',
    # Ký tự vùng Private Use Area (Wingdings / Symbol trong PDF kỹ thuật)
    '\uf0b7': '•',
    '\uf06c': '•',
    '\uf06e': '•',
    '\uf075': '•',
    '\uf076': '•',
    '\uf0a7': '•',
    '\uf0a8': '[ ]',
    '\uf0fe': '[√]',
    '\uf0fc': '√',
    '\uf0fb': 'X',
    '\uf0d8': '->',
    '\uf0e0': '->',
    '\uf0e8': '->',
    '\uf0b0': '°',
    '\uf0b1': '±',
    '\uf0b2': '²',
    '\uf0b3': '≥',
    '\uf0a3': '≤',
    '\uf0b9': '≠',
    '\uf0bb': '≈',
    '\uf0c4': '×',
    '\uf0b8': '÷',
    '\uf066': 'Φ',
    '\uf046': 'Φ',
    '\uf063': 'Φ',
    '\uf057': 'Ω',
    '\uf077': 'Ω',
    '\uf06d': 'μ',
    '\uf0ab': '<->',
    '\uf0ac': '<-',
    '\uf0ad': '↑',
    '\uf0ae': '->',
    '\uf0af': '↓',
}

# Thêm các số tròn ①..⑳ (U+2460..U+2473), ⑴..⒇ (U+2474..U+2487), ⒈..⒛ (U+2488..U+249B), ❶..❿ (U+2776..U+277F)
for _i in range(1, 21):
    SPECIAL_SYMBOL_MAP[chr(0x245F + _i)] = f"({_i})"
    SPECIAL_SYMBOL_MAP[chr(0x2473 + _i)] = f"({_i})"
    SPECIAL_SYMBOL_MAP[chr(0x2487 + _i)] = f"{_i}."
for _i in range(1, 11):
    SPECIAL_SYMBOL_MAP[chr(0x2775 + _i)] = f"({_i})"
    SPECIAL_SYMBOL_MAP[chr(0x2789 + _i)] = f"({_i})"
    SPECIAL_SYMBOL_MAP[chr(0x277F + _i)] = f"({_i})"

def normalize_special_symbols(text):
    """Chuyển đổi các ký hiệu đặc biệt/Wingdings/CJK sang ký tự chuẩn không bao giờ bị lỗi ô vuông."""
    if not text:
        return ""
    out = []
    for ch in str(text):
        code = ord(ch)
        if ch in SPECIAL_SYMBOL_MAP:
            out.append(SPECIAL_SYMBOL_MAP[ch])
        elif 0xFF01 <= code <= 0xFF5E:
            # Chuyển ký tự Fullwidth (ＡＢＣ１２３％...) về ASCII chuẩn
            out.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            out.append(' ')
        elif 0xE000 <= code <= 0xF8FF:
            # Bỏ các ký tự Private Use Area lạ còn sót để không hiện ô vuông □
            out.append('•')
        else:
            out.append(ch)
    return "".join(out)

def sanitize_xml_text(text):
    if text is None:
        return ""
    cleaned = ILLEGAL_XML_CHARS_RE.sub('', str(text))
    cleaned = cleaned.replace('\xa0', ' ')
    return normalize_special_symbols(cleaned)

# ==========================================
# BỘ LỌC NGÔN NGỮ & SMART RESUME
# ==========================================
VIETNAMESE_CHARS_RE = re.compile(
    r'[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ'
    r'ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]'
)
CHINESE_CHARS_RE = re.compile(r'[\u4e00-\u9fff]')
CJK_EXTRA_RE = re.compile(r'[\u3040-\u30ff\uac00-\ud7af]')
LATIN_WORDS_RE = re.compile(r'[A-Za-z]{2,}')

def should_skip_translation(text, lang_mode="zh_to_vi", smart_resume=True):
    if not text or not text.strip():
        return True
    # Không dịch nếu chỉ có số, ký hiệu toán học, mã hiệu ngắn không có chữ
    if not re.search(r'[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7afA-Za-z\u0400-\u04ff]', text):
        return True

    has_zh = bool(CHINESE_CHARS_RE.search(text))
    has_cjk = has_zh or bool(CJK_EXTRA_RE.search(text))
    has_vi = bool(VIETNAMESE_CHARS_RE.search(text))
    has_en = bool(LATIN_WORDS_RE.search(text))

    if lang_mode in ("zh_to_vi", "zh_to_en"):
        # Chỉ dịch nếu có chứa chữ Hán
        return not has_zh

    if lang_mode == "en_to_vi":
        # Chỉ dịch nếu có từ tiếng Anh và không phải là câu thuần tiếng Trung
        if not has_en:
            return True
        if smart_resume and has_vi and not has_zh:
            return True
        return False

    if lang_mode == "all_to_vi":
        if smart_resume and has_vi and not has_cjk:
            return True
        return False

    return False

# ==========================================
# XỬ LÝ GỌI API GEMINI & DYNAMIC CHUNKING
# ==========================================
def handle_api_error(e, attempt):
    error_msg = str(e).lower()
    # Thử xoay vòng sang API dự phòng ngay lập tức nếu người dùng đã nhập nhiều API
    switched = rotate_to_backup_api()
    if switched:
        time.sleep(1.5)
        return

    if "429" in error_msg or "quota" in error_msg or "exhausted" in error_msg:
        st.toast("⚠️ Quá tải hạn mức API! Đang tạm nghỉ 30s để phục hồi (Hãy thêm API 2, API 3 ở thanh bên để tự động đổi API không cần chờ)...", icon="⏳")
        time.sleep(30)
    elif "safety" in error_msg or "block" in error_msg:
        st.toast("🛑 Lỗi Kiểm Duyệt: AI từ chối cụm từ nhạy cảm!", icon="🚫")
        time.sleep(2)
    else:
        st.toast(f"Lỗi API ({get_active_api_label()}). Thử lại lần {attempt+1}...", icon="🔄")
        time.sleep(3)

def _build_single_prompt(text, lang_mode):
    mode_cfg = LANG_MODES.get(lang_mode, LANG_MODES["zh_to_vi"])
    if lang_mode == "zh_to_vi":
        rule = "Dịch văn bản kỹ thuật sau từ tiếng Trung sang tiếng Việt. Giữ nguyên các từ tiếng Anh/mã kỹ thuật nếu không phải tiếng Trung."
    elif lang_mode == "zh_to_en":
        rule = "Translate the following technical text from Chinese into accurate, concise technical English."
    elif lang_mode == "en_to_vi":
        rule = "Dịch văn bản kỹ thuật sau từ tiếng Anh sang tiếng Việt chuẩn chuyên ngành."
    else:
        rule = f"Dịch văn bản kỹ thuật sau từ {mode_cfg['source_desc']} sang {mode_cfg['target_desc']} chuẩn chuyên ngành."
    return f"{rule} KHÔNG giải thích, chỉ trả về kết quả dịch.\nGốc:\n{text}\nDịch:"

def _build_batch_prompt(texts, lang_mode):
    json_input = json.dumps(texts, ensure_ascii=False)
    if lang_mode == "zh_to_vi":
        instruction = "Dịch mảng JSON chứa văn bản kỹ thuật sau từ tiếng Trung sang tiếng Việt (chỉ dịch phần tiếng Trung sang tiếng Việt, giữ nguyên mã hiệu/ký hiệu kỹ thuật)."
    elif lang_mode == "zh_to_en":
        instruction = "Translate the following JSON array of technical texts from Chinese into accurate, concise technical English."
    elif lang_mode == "en_to_vi":
        instruction = "Dịch mảng JSON chứa văn bản kỹ thuật sau từ tiếng Anh sang tiếng Việt chuẩn chuyên ngành."
    else:
        instruction = "Dịch toàn bộ mảng JSON chứa văn bản kỹ thuật sau (từ mọi ngôn ngữ) sang tiếng Việt chuẩn chuyên ngành."

    return f"""{instruction}
YÊU CẦU BẮT BUỘC:
1. Trả về ĐÚNG CẤU TRÚC mảng JSON gồm các chuỗi (string).
2. Số lượng phần tử đầu ra PHẢI BẰNG ĐÚNG đầu vào ({len(texts)} phần tử).
3. KHÔNG dùng ký tự đặc biệt lỗi font (dùng °C thay cho ℃, dùng (1)(2) thay cho ①②, dùng [ ] thay cho 【】).
4. KHÔNG có markdown, KHÔNG giải thích. Chỉ in ra mảng JSON hợp lệ.

Đầu vào:
{json_input}

Đầu ra:"""

def translate_text_core(text, model_name, lang_mode="zh_to_vi"):
    cache = get_translation_cache()
    cache_key = f"{lang_mode}||{text}"
    if cache_key in cache:
        return cache[cache_key]

    prompt = _build_single_prompt(text, lang_mode)
    for attempt in range(4):
        try:
            time.sleep(4)
            model = genai.GenerativeModel(model_name)
            res = model.generate_content(prompt, safety_settings=SAFETY_SETTINGS)
            if res.text:
                cleaned = sanitize_xml_text(
                    res.text.replace("Bản dịch tiếng Việt:", "").replace("Dịch:", "").strip()
                )
                cache[cache_key] = cleaned
                return cleaned
        except Exception as e:
            handle_api_error(e, attempt)
    return sanitize_xml_text(text)

def translate_batch_core(texts, model_name, lang_mode="zh_to_vi"):
    if not texts:
        return []
    cache = get_translation_cache()

    results = [None] * len(texts)
    missing_indices = []
    missing_texts = []
    seen_missing = {}

    for i, t in enumerate(texts):
        ck = f"{lang_mode}||{t}"
        if ck in cache:
            results[i] = cache[ck]
        elif t in seen_missing:
            missing_indices.append((i, seen_missing[t]))
        else:
            idx_in_missing = len(missing_texts)
            seen_missing[t] = idx_in_missing
            missing_texts.append(t)
            missing_indices.append((i, idx_in_missing))

    if not missing_texts:
        return results

    if len(missing_texts) == 1:
        single_res = translate_text_core(missing_texts[0], model_name, lang_mode=lang_mode)
        for orig_i, _ in missing_indices:
            results[orig_i] = single_res
        return results

    prompt = _build_batch_prompt(missing_texts, lang_mode)
    translated_missing = None

    for attempt in range(4):
        try:
            time.sleep(4)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt, safety_settings=SAFETY_SETTINGS)
            if response.text:
                res_text = response.text.strip().strip("` \n")
                if res_text.lower().startswith("json"):
                    res_text = res_text[4:].strip()
                try:
                    translated_array = json.loads(res_text)
                    if isinstance(translated_array, list) and len(translated_array) == len(missing_texts):
                        translated_missing = [
                            sanitize_xml_text(str(x) if x is not None else "")
                            for x in translated_array
                        ]
                        break
                except Exception:
                    pass
        except Exception as e:
            handle_api_error(e, attempt)

    if translated_missing is None:
        st.toast("⚠️ Phân tách gói dịch để tránh lỗi API...", icon="🔧")
        mid = len(missing_texts) // 2
        translated_missing = (
            translate_batch_core(missing_texts[:mid], model_name, lang_mode=lang_mode)
            + translate_batch_core(missing_texts[mid:], model_name, lang_mode=lang_mode)
        )

    for src_t, trans_t in zip(missing_texts, translated_missing):
        cache[f"{lang_mode}||{src_t}"] = trans_t

    for orig_i, miss_i in missing_indices:
        results[orig_i] = translated_missing[miss_i]

    return results

def dynamic_chunking(items, max_chars=1500, max_items=25):
    batches = []
    current_batch = []
    current_len = 0
    for item in items:
        clean_text = sanitize_xml_text(item['text']).replace('|||', '---')
        item['text'] = clean_text
        l = len(clean_text)

        if (current_len + l > max_chars or len(current_batch) >= max_items) and current_batch:
            batches.append(current_batch)
            current_batch = [item]
            current_len = l
        else:
            current_batch.append(item)
            current_len += l
    if current_batch:
        batches.append(current_batch)
    return batches

def execute_translation_batches(
    target_items,
    model_name,
    progress_bar,
    status_text,
    preview_box,
    lang_mode="zh_to_vi",
    smart_resume=True
):
    batches = dynamic_chunking(target_items)
    processed = 0
    total_targets = len(target_items)
    if total_targets == 0:
        return

    for chunk in batches:
        texts_to_translate = []
        for item in chunk:
            if should_skip_translation(item['text'], lang_mode=lang_mode, smart_resume=smart_resume):
                item['translated'] = sanitize_xml_text(item['text'])
            else:
                texts_to_translate.append(item['text'])
                item['translated'] = None

        if texts_to_translate:
            translated_results = translate_batch_core(
                texts_to_translate, model_name, lang_mode=lang_mode
            )
            res_idx = 0
            for item in chunk:
                if item['translated'] is None:
                    item['translated'] = (
                        translated_results[res_idx]
                        if res_idx < len(translated_results)
                        else sanitize_xml_text(item['text'])
                    )
                    res_idx += 1

        for item in chunk:
            original = item['text']
            translated = item['translated']
            if 'apply_func' in item:
                item['apply_func'](item['obj'], translated)

            processed += 1
            progress_bar.progress(min(1.0, processed / total_targets))
            preview_box.success(
                f"**Đã xử lý {processed}/{total_targets} mục tiêu:**\n\n📄 `{original}`\n\n✅ `{translated}`"
            )

            if 'autosave_func' in item and processed % 5 == 0:
                item['autosave_func']()

def atomic_save_bytes(filename, data_bytes):
    temp_path = f"{filename}.tmp_{int(time.time() * 1000)}"
    try:
        with open(temp_path, "wb") as f:
            f.write(data_bytes)
        os.replace(temp_path, filename)
    except Exception:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass

