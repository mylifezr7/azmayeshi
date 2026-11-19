"""
Persian Text Normalizer and Tokenizer for Excel Commands.
Handles Persian character standardization, digit normalization, and entity extraction.
"""
import re
from typing import Dict, Optional, Tuple

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


def normalize_persian_text(text: str) -> str:
    """Normalizes Persian characters, converts digits, and cleans whitespace."""
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

    # Comprehensive Persian phonetic and transliteration map for columns
    persian_col_map = {
        'الف': 'A', 'آ': 'A', 'ای': 'A', 'اِی': 'A', 'اول': 'A', 'یک': 'A', '۱': 'A',
        'بی': 'B', 'ب': 'B', 'دوم': 'B', 'دو': 'B', '۲': 'B',
        'سی': 'C', 'ج': 'C', 'پ': 'C', 'سوم': 'C', 'سه': 'C', '۳': 'C',
        'دی': 'D', 'د': 'D', 'چهارم': 'D', 'چهار': 'D', '۴': 'D',
        'ئی': 'E', 'ه': 'E', 'پنجم': 'E', 'پنج': 'E', '۵': 'E',
        'اف': 'F', 'ف': 'F', 'ششم': 'F', 'شش': 'F', '۶': 'F',
        'جی': 'G', 'گ': 'G', 'هفتم': 'G', 'هفت': 'G', '۷': 'G',
        'اچ': 'H', 'ح': 'H', 'هشتم': 'H', 'هشت': 'H', '۸': 'H',
        'آی': 'I', 'ط': 'I', 'نهم': 'I', 'نه': 'I', '۹': 'I',
        'جی_جی': 'J', 'دهم': 'J', 'ده': 'J', '۱۰': 'J',
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
    clean_norm = clean.lower()
    if clean_norm in persian_col_map:
        return persian_col_map[clean_norm]
    return None


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
