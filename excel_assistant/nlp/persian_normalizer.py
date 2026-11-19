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
    Extracts column identifier (e.g. 'A', 'B', 'AA', or 1, 2) from a token.
    """
    clean = token.strip().upper()
    # If standard English column letter
    if re.match(r'^[A-Z]{1,3}$', clean):
        return clean
    # If Persian column letter transliteration (e.g. 'الف' -> 'A', 'ب' -> 'B', 'ج' -> 'C', 'د' -> 'D')
    persian_col_map = {
        'الف': 'A',
        'ب': 'B',
        'پ': 'C',
        'ج': 'C',
        'ت': 'D',
        'د': 'D',
        'ه': 'E',
    }
    if clean in persian_col_map:
        return persian_col_map[clean]
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
