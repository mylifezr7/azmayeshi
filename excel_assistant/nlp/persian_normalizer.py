"""
Persian Text Normalizer, Tokenizer, and Semantic Slot Extractor for Excel Commands.
Handles Persian character standardization, digit normalization, typo correction,
colloquial transformation, and robust extraction of target and source columns.
"""
import re
from typing import Dict, List, Optional, Tuple

PERSIAN_TO_ENGLISH_DIGITS = {
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9',
    '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
    '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9',
}

PERSIAN_CHAR_MAP = {
    'ي': 'ی',
    'ك': 'ک',
    'ة': 'ه',
    'ؤ': 'و',
    'إ': 'ا',
    'أ': 'ا',
    'ء': '',
    '\u200c': ' ',  # zero-width non-joiner to space for easier token matching
    '\xa0': ' ',
}

PERSIAN_TYPOS_MAP = {
    'شیلد': 'شیت',
    'حمع': 'جمع',
    'صرب': 'ضرب',
    'تقریق': 'تفریق',
    'تقصیم': 'تقسیم',
    'فیلطر': 'فیلتر',
    'تکراریها': 'تکراری ها',
    'تکراری‌ها': 'تکراری ها',
    'ردیفها': 'ردیف ها',
    'ردیف‌های': 'ردیف های',
    'ردیفهای': 'ردیف های',
    'ردیف‌ها': 'ردیف ها',
    'سطرها': 'سطر ها',
    'سطرهای': 'سطر های',
    'سطر‌های': 'سطر های',
    'ستونن': 'ستون',
    'ستوم': 'ستون',
    'ماکزمم': 'ماکزیمم',
    'ماکزیموم': 'ماکزیمم',
    'مینمم': 'مینیمم',
    'فورمول': 'فرمول',
    'فرموال': 'فرمول',
    'اتوسام': 'اتو سام',
    'حاصلضرب': 'حاصل ضرب',
    'بعلاوه': 'به علاوه',
    'مرتبسازی': 'مرتب سازی',
    'مغایرتگیری': 'مغایرت گیری',
}

PERSIAN_COLORS = {
    'زرد': '#FFFF00',
    'سبز': '#90EE90',
    'قرمز': '#FF6B6B',
    'آبی': '#87CEFA',
    'ابی': '#87CEFA',
    'نارنجی': '#FFA500',
    'خاکستری': '#D3D3D3',
    'طوسی': '#D3D3D3',
    'سفید': '#FFFFFF',
    'مشکی': '#000000',
    'سیاه': '#000000',
    'بنفش': '#DDA0DD',
    'صورتی': '#FFC0CB',
}

# Comprehensive Persian phonetic and transliteration map for columns
PERSIAN_COL_MAP = {
    'الف': 'A', 'آ': 'A', 'ای': 'A', 'اِی': 'A', 'اول': 'A', 'یک': 'A', '۱': 'A', '1': 'A',
    'بی': 'B', 'ب': 'B', 'دوم': 'B', 'دو': 'B', '۲': 'B', '2': 'B',
    'سی': 'C', 'ج': 'C', 'پ': 'C', 'سوم': 'C', 'سه': 'C', '۳': 'C', '3': 'C',
    'دی': 'D', 'د': 'D', 'چهارم': 'D', 'چهار': 'D', '۴': 'D', '4': 'D',
    'ئی': 'E', 'ه': 'E', 'پنجم': 'E', 'پنج': 'E', '۵': 'E', '5': 'E',
    'اف': 'F', 'ف': 'F', 'ششم': 'F', 'شش': 'F', '۶': 'F', '6': 'F',
    'جی': 'G', 'گ': 'G', 'هفتم': 'G', 'هفت': 'G', '۷': 'G', '7': 'G',
    'اچ': 'H', 'ح': 'H', 'هشتم': 'H', 'هشت': 'H', '۸': 'H', '8': 'H',
    'آی': 'I', 'ط': 'I', 'نهم': 'I', 'نه': 'I', '۹': 'I', '9': 'I',
    'جی_جی': 'J', 'دهم': 'J', 'ده': 'J', '۱۰': 'J', '10': 'J',
    'کی': 'K', 'کا': 'K', 'ک': 'K',
    'ال': 'L', 'ل': 'L',
    'ام': 'M', 'م': 'M',
    'ان': 'N', 'ن': 'N',
    'او': 'O',
    'پی': 'P',
    'کیو': 'Q',
    'آر': 'R', 'ار': 'R', 'ر': 'R',
    'اس': 'S', 'س': 'S',
    'تی': 'T', 'ت': 'T',
    'یو': 'U',
    'وی': 'V',
    'دبلیو': 'W', 'دابلیو': 'W',
    'ایکس': 'X',
    'وای': 'Y',
    'زد': 'Z', 'ز': 'Z',
}


def normalize_colloquial(text: str) -> str:
    """Normalizes colloquial Persian expressions to standard forms."""
    if not text:
        return ""

    # Replace standalone 'رو' with 'را' (handles 'ستون D رو', 'جواب رو', etc.)
    text = re.sub(r'\bرو\b', 'را', text)

    # Replace standalone 'توی' or 'تو' with 'در'
    text = re.sub(r'\b(?:توی|تو)\b', 'در', text)

    # Normalize colloquial verb forms
    text = re.sub(r'\bبکن\b', 'کن', text)
    text = re.sub(r'\bبکنید\b', 'کنید', text)
    text = re.sub(r'\bبزن\b', 'کن', text)

    # Normalize colloquial pronominal forms
    text = re.sub(r'\bپاسخشو\b', 'پاسخ را', text)
    text = re.sub(r'\bجوابشو\b', 'جواب را', text)
    text = re.sub(r'\bنتیجشو\b', 'نتیجه را', text)
    text = re.sub(r'\bپاسخش\b', 'پاسخ', text)
    text = re.sub(r'\bجوابش\b', 'جواب', text)
    text = re.sub(r'\bنتیجه‌اش\b', 'نتیجه', text)
    text = re.sub(r'\bنتیجه اش\b', 'نتیجه', text)

    # Normalize percentages: e.g. '15درصد' or '15 ٪' -> '15 درصد'
    text = re.sub(r'(\d+)\s*(?:درصد|٪|%)', r'\1 درصد', text)

    return text


def normalize_persian_text(text: str) -> str:
    """Normalizes Persian characters, converts digits, corrects typos, and cleans whitespace."""
    if not text:
        return ""

    # Replace digits
    for p_d, e_d in PERSIAN_TO_ENGLISH_DIGITS.items():
        text = text.replace(p_d, e_d)

    # Replace characters
    for p_c, target in PERSIAN_CHAR_MAP.items():
        text = text.replace(p_c, target)

    # Remove Arabic diacritics
    text = re.sub(r'[\u064B-\u065F\u0670]', '', text)

    # Fix common typos
    for typo, correct in PERSIAN_TYPOS_MAP.items():
        text = re.sub(r'(?:(?<=^)|(?<=\s))' + re.escape(typo) + r'(?=(?:\s|$))', correct, text)

    # Normalize colloquial phrasing
    text = normalize_colloquial(text)

    # Normalize multiple whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_column_identifier(token: str) -> Optional[str]:
    """
    Extracts column identifier (e.g. 'A', 'B', 'AA', or Persian phonetic 'دی', 'سی', 'بی', 1, 2) from a token.
    """
    if not token:
        return None
    clean = token.strip()

    # Check numeric index e.g. "1" -> "A", "2" -> "B"
    if clean.isdigit():
        num = int(clean)
        if 1 <= num <= 26:
            return chr(64 + num)

    clean_up = clean.upper()
    # If standard English column letter (A, B, C, ..., Z, AA, AB...)
    if re.match(r'^[A-Z]{1,3}$', clean_up):
        return clean_up

    clean_norm = clean.lower()
    if clean_norm in PERSIAN_COL_MAP:
        return PERSIAN_COL_MAP[clean_norm]
    return None


def resolve_column(token: str, header_map: Optional[Dict[str, str]] = None) -> Optional[str]:
    """Resolves a token into a column letter, checking headers, numbers, and phonetic names."""
    if not token:
        return None
    col = extract_column_identifier(token)
    if col:
        return col
    clean = token.strip()
    if header_map:
        if clean in header_map:
            return header_map[clean]
        clean_norm = normalize_persian_text(clean).lower()
        for h_name, col_letter in header_map.items():
            h_norm = normalize_persian_text(h_name).lower()
            if clean_norm == h_norm or clean_norm in h_norm or h_norm in clean_norm:
                return col_letter
    return None


def extract_target_column(text: str, header_map: Optional[Dict[str, str]] = None) -> Tuple[Optional[str], str]:
    """
    Extracts target column from anywhere in the Persian sentence and returns (target_column, remaining_text).
    The extracted target clause is removed from remaining_text to prevent picking it up as a source column.
    """
    text_clean = text.strip()

    # 0. Check header_map matches first (covers multi-word headers like 'جمع کل')
    if header_map:
        for h_name, col_letter in sorted(header_map.items(), key=lambda x: len(x[0]), reverse=True):
            # Header at end with verb
            m_h = re.search(
                rf'(?:و\s+)?(?:در|توی|تو|داخل|به|روی)\s+(?:ستون\s+)?{re.escape(h_name)}\s*(?:بنویس|بذار|بزار|قرار\s*بده|قرار\s*ده|قرار\s*بگذار|بریز|درج\s*کن|ثبت\s*کن)?\s*[.]*$',
                text_clean
            )
            if m_h:
                rem = text_clean[:m_h.start()].strip()
                rem = re.sub(r'\s+(?:و|را|رو)$', '', rem).strip()
                return col_letter, rem

            # Header at end with result noun
            m_h_res = re.search(
                rf'(?:و\s+)?(?:جواب|نتیجه|پاسخ|پاسخش|جوابش|نتایج)\s*(?:را|رو)?\s*(?:رو|در|توی|داخل|به|روی)?\s*(?:ستون\s+)?{re.escape(h_name)}\s*(?:بذار|بزار|بنویس|قرار\s*بده|قرار\s*ده|بریز|درج\s*کن|ثبت\s*کن)?\s*[.]*$',
                text_clean
            )
            if m_h_res:
                rem = text_clean[:m_h_res.start()].strip()
                rem = re.sub(r'\s+(?:و|را|رو)$', '', rem).strip()
                return col_letter, rem

            # Header at start
            m_h_start = re.match(
                rf'^(?:در|توی|تو|داخل|به|روی)\s+(?:ستون\s+)?{re.escape(h_name)}\s*(?:بنویس|بذار|بزار|قرار\s*بده|قرار\s*ده|بریز|درج\s*کن|ثبت\s*کن)\s*(?:،|,)?\s*(.*)$',
                text_clean
            )
            if m_h_start:
                return col_letter, m_h_start.group(1).strip()

    # 1. Target at beginning: e.g. 'در E بنویس ...', 'روی ستون E بذار ...'
    m_start = re.match(
        r'^(?:در|توی|داخل|تو|به|روی)\s+(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*(?:بنویس|بذار|بزار|قرار\s*بده|بریز|درج\s*کن|ثبت\s*کن)\s*(?:،|,)?\s*(.*)$',
        text_clean
    )
    if m_start:
        c = resolve_column(m_start.group(1), header_map)
        if c:
            return c, m_start.group(2).strip()

    # 2. Target at end:
    # Pattern A: '... و جواب / نتیجه / پاسخ [را] [در / توی / به / داخل / روی / بره به / بره تو]? [ستون]? X [بذار / بنویس / قرار بده / بریز]?'
    m_end_res = re.search(
        r'(?:و\s+)?(?:جواب|نتیجه|پاسخ|پاسخش|جوابش|نتایج)\s*(?:را|رو)?\s*(?:رو|در|توی|داخل|به|روی|بره\s+به|بره\s+تو|بره\s+توی|بریز\s+تو|بریز\s+در)?\s*(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*(?:بذار|بزار|بنویس|قرار\s*بده|قرار\s*ده|بریز|درج\s*کن|ثبت\s*کن|بره|بشینه|باشه)?\s*[.]*$',
        text_clean
    )
    if m_end_res:
        c = resolve_column(m_end_res.group(1), header_map)
        if c:
            rem = text_clean[:m_end_res.start()].strip()
            rem = re.sub(r'\s+(?:و|را|رو)$', '', rem).strip()
            return c, rem

    # Pattern B: '... بریز / بذار / بنویس [در / تو / توی / داخل / به / روی] [ستون] X'
    m_end_verb = re.search(
        r'(?:و\s+)?(?:بریز|بذار|بزار|بنویس|قرار\s*بده|درج\s*کن|ثبت\s*کن|بره)\s+(?:در|توی|تو|داخل|به|روی)?\s*(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*[.]*$',
        text_clean
    )
    if m_end_verb:
        c = resolve_column(m_end_verb.group(1), header_map)
        if c:
            rem = text_clean[:m_end_verb.start()].strip()
            rem = re.sub(r'\s+(?:و|را|رو)$', '', rem).strip()
            return c, rem

    # Pattern C: Standard '[و] [در / توی / داخل / تو / به / روی] [ستون] X [بنویس / بذار / قرار بده / بریز]'
    m_end_std = re.search(
        r'(?:و\s+)?(?:در|توی|تو|داخل|به|روی)\s+(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*(?:بنویس|بذار|بزار|قرار\s*بده|قرار\s*ده|قرار\s*بگذار|درج\s*کن|بریز|بگذار|ثبت\s*کن|قرار\s*بگیرد)?\s*[.]*$',
        text_clean
    )
    if m_end_std:
        c = resolve_column(m_end_std.group(1), header_map)
        if c:
            rem = text_clean[:m_end_std.start()].strip()
            rem = re.sub(r'\s+(?:و|را|رو)$', '', rem).strip()
            return c, rem

    return None, text_clean


def extract_source_columns(rem_text: str, header_map: Optional[Dict[str, str]] = None) -> List[str]:
    """
    Extracts source columns from remaining text in order of appearance.
    Supports multi-column list, explicit 'ستون X', header names, Latin letters, and phonetic letters.
    Special handling for 'X را از Y کم کن' to return [Y, X] for correct mathematical direction (Y - X).
    """
    # Check subtraction pattern 'X را از Y کم کن' (means Y - X)
    sub_m = re.search(
        r'([A-Za-z0-9\u0600-\u06FF]+)\s+(?:را|رو)?\s*از\s+([A-Za-z0-9\u0600-\u06FF]+)\s+(?:کم|منها|کسر|تفریق)',
        rem_text
    )
    if sub_m:
        c_sub = resolve_column(sub_m.group(1), header_map)
        c_base = resolve_column(sub_m.group(2), header_map)
        if c_sub and c_base and c_sub != c_base:
            return [c_base, c_sub]

    found: List[Tuple[int, str]] = []

    # 1. Multi-column list e.g. 'ستون های 2 و 4 و 5' or 'ستون B و D و E'
    for m in re.finditer(
        r'ستون(?:\s*های)?\s+([A-Za-z0-9\s،,و\u0600-\u06FF]+?)(?:\s+(?:را|رو|با|در|جمع|ضرب|تفریق|منهای|تقسیم|\*|\+|-|\/)|$)',
        rem_text
    ):
        tokens = re.split(r'[\s،,و]+', m.group(1).strip())
        base_pos = m.start()
        for idx, t in enumerate(tokens):
            c = resolve_column(t, header_map)
            if c:
                found.append((base_pos + idx, c))

    # 2. Explicit 'ستون X'
    for m in re.finditer(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', rem_text):
        c = resolve_column(m.group(1), header_map)
        if c:
            found.append((m.start(), c))

    # 3. Header names
    if header_map:
        for h_name, col_let in sorted(header_map.items(), key=lambda x: len(x[0]), reverse=True):
            for m in re.finditer(r'(?:(?<=^)|(?<=\s))' + re.escape(h_name) + r'(?=(?:\s|$))', rem_text):
                found.append((m.start(), col_let))

    # 4. Isolated Latin column letters
    for m in re.finditer(r'\b([A-Za-z]{1,2})\b', rem_text):
        c = m.group(1).upper()
        found.append((m.start(), c))

    # 5. Standalone phonetic column names
    for m in re.finditer(r'\b(دی|بی|سی|ای|الف|اف|جی|اچ|آی|کی|ال|ام|ان|پی|تی|وی|زد)\b', rem_text):
        c = PERSIAN_COL_MAP[m.group(1)]
        found.append((m.start(), c))

    # Sort in order of appearance
    found.sort(key=lambda x: x[0])
    ordered_cols: List[str] = []
    for pos, col in found:
        if col not in ordered_cols:
            ordered_cols.append(col)
    return ordered_cols


def extract_percentage(text: str) -> Optional[float]:
    """Extracts percentage value from Persian text, e.g. '15 درصد', '20%', '15 ٪'."""
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:درصد|٪|%)', text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


def extract_row_range(text: str) -> Tuple[Optional[int], Optional[int]]:
    """Extracts start and end row numbers from text, e.g. 'از سطر 1 تا 30'."""
    m = re.search(r'از\s+(?:سطر|ردیف)\s*(\d+)\s+تا\s*(?:سطر|ردیف)?\s*(\d+)', text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def extract_color_hex(text: str, default: str = "#FFFF00") -> str:
    """Finds color in Persian text or returns default."""
    for word, hex_val in PERSIAN_COLORS.items():
        if word in text:
            return hex_val
    # Hex regex
    hex_match = re.search(r'#(?:[0-9a-fA-F]{3}){1,2}\b', text)
    if hex_match:
        return hex_match.group(0)
    return default
