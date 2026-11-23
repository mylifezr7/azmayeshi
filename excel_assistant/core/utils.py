"""
Utilities for Excel coordinate manipulation, column/row conversions, and formatting.
"""
import re
from typing import Optional, Tuple, Union


def col_letter_to_index(col_letter: str) -> int:
    """Convert Excel column letter ('A', 'B', 'AA') to 1-based index."""
    col_clean = col_letter.strip().upper()
    idx = 0
    for char in col_clean:
        if 'A' <= char <= 'Z':
            idx = idx * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid column character '{char}' in '{col_letter}'")
    return idx


def col_index_to_letter(col_index: int) -> str:
    """Convert 1-based index (e.g. 1 -> 'A', 28 -> 'AB') to column letter."""
    if col_index < 1:
        raise ValueError(f"Column index must be >= 1, got {col_index}")
    result = []
    idx = col_index
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def parse_cell_address(cell: str) -> Tuple[int, int]:
    """Parse cell address like 'B5' or '$C$10' to (row, col) (1-indexed)."""
    clean = cell.replace('$', '').strip().upper()
    match = re.match(r"^([A-Z]+)([0-9]+)$", clean)
    if not match:
        raise ValueError(f"Invalid cell address: '{cell}'")
    col_str, row_str = match.groups()
    return int(row_str), col_letter_to_index(col_str)


def parse_range_address(range_str: str) -> Tuple[int, int, int, int]:
    """
    Parse range like 'A1:C10' or single cell 'B2' or whole column 'A:A'.
    Returns (min_row, min_col, max_row, max_col) 1-indexed.
    """
    clean = range_str.replace('$', '').strip().upper()
    if ':' in clean:
        parts = clean.split(':')
        start, end = parts[0].strip(), parts[1].strip()
        # Check if full column e.g. "A:B"
        if re.match(r"^[A-Z]+$", start) and re.match(r"^[A-Z]+$", end):
            min_col = col_letter_to_index(start)
            max_col = col_letter_to_index(end)
            return 1, min(min_col, max_col), 1048576, max(min_col, max_col)
        # Check if full row e.g. "1:10"
        if re.match(r"^[0-9]+$", start) and re.match(r"^[0-9]+$", end):
            return int(start), 1, int(end), 16384
        r1, c1 = parse_cell_address(start)
        r2, c2 = parse_cell_address(end)
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
    else:
        # Check if single column letter like 'A'
        if re.match(r"^[A-Z]+$", clean):
            c = col_letter_to_index(clean)
            return 1, c, 1048576, c
        r, c = parse_cell_address(clean)
        return r, c, r, c


def bounds_to_range_str(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    """Format (min_row, min_col, max_row, max_col) into 'A1:B10'."""
    c1 = col_index_to_letter(min_col)
    c2 = col_index_to_letter(max_col)
    if min_row == max_row and min_col == max_col:
        return f"{c1}{min_row}"
    return f"{c1}{min_row}:{c2}{max_row}"


def normalize_hex_color(hex_str: str) -> str:
    """Normalize '#FFFF00', 'FFFF00', 'yellow' to clean 6-char hex 'FFFF00'."""
    h = hex_str.strip().lower()
    color_map = {
        'yellow': 'FFFF00',
        'green': '90EE90',
        'light_green': 'D4EDDA',
        'red': 'FF6B6B',
        'light_red': 'F8D7DA',
        'blue': '87CEFA',
        'orange': 'FFA500',
        'gray': 'D3D3D3',
        'white': 'FFFFFF',
        'زرد': 'FFFF00',
        'سبز': '90EE90',
        'قرمز': 'FF6B6B',
        'آبی': '87CEFA',
        'نارنجی': 'FFA500',
        'طوسی': 'D3D3D3',
        'خاکستری': 'D3D3D3',
    }
    if h in color_map:
        return color_map[h]
    h = h.lstrip('#')
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    if len(h) == 6:
        return h.upper()
    return "FFFF00"  # default fallback yellow


def hex_to_rgb(hex_code: str) -> Tuple[int, int, int]:
    """Convert hex string e.g. 'FFFF00' to RGB tuple (255, 255, 0)."""
    clean = normalize_hex_color(hex_code)
    return int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16)
