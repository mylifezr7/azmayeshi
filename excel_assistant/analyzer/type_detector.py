"""
Column Data Type and Statistics Detector.
Analyzes column values and determines dominant types, stats, and anomalies.
"""
from dataclasses import dataclass, field
import datetime
import re
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class ColumnSummary:
    column_letter: str
    column_index: int
    header_name: str
    detected_type: str  # numeric, integer, float, currency, percentage, date, boolean, text, empty, mixed
    total_count: int = 0
    non_null_count: int = 0
    null_count: int = 0
    unique_count: int = 0
    duplicate_count: int = 0
    sample_values: List[Any] = field(default_factory=list)
    min_val: Optional[Any] = None
    max_val: Optional[Any] = None
    numeric_values: List[float] = field(default_factory=list)
    error_count: int = 0
    is_numeric: bool = False


# Common Excel error values
EXCEL_ERRORS = {
    "#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#CALC!"
}

CURRENCY_SYMBOLS = {'$', '€', '£', '¥', '﷼', 'تومان', 'ریال'}


def is_excel_error(val: Any) -> bool:
    if val is None:
        return False
    s = str(val).strip().upper()
    return s in EXCEL_ERRORS or s.startswith("#")


def detect_cell_type(val: Any) -> str:
    if val is None or val == "":
        return "empty"
    if is_excel_error(val):
        return "error"
    if isinstance(val, bool):
        return "boolean"
    if isinstance(val, (int, float)):
        return "integer" if isinstance(val, int) or float(val).is_integer() else "float"
    if isinstance(val, (datetime.date, datetime.datetime)):
        return "date"

    s = str(val).strip()
    if not s:
        return "empty"

    # Check for percentage
    if s.endswith("%"):
        try:
            float(s[:-1].replace(',', ''))
            return "percentage"
        except ValueError:
            pass

    # Check for currency
    for sym in CURRENCY_SYMBOLS:
        if sym in s:
            clean = s.replace(sym, '').replace(',', '').strip()
            try:
                float(clean)
                return "currency"
            except ValueError:
                pass

    # Check for standard numeric string
    clean_num = s.replace(',', '')
    try:
        f = float(clean_num)
        return "integer" if f.is_integer() else "float"
    except ValueError:
        pass

    # Check for date formats (YYYY-MM-DD, DD/MM/YYYY, etc.)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$", s) or re.match(r"^\d{1,2}[-/]\d{1,2}[-/]\d{4}$", s):
        return "date"

    # Check for boolean string
    if s.lower() in ("true", "false", "صحیح", "غلط", "بله", "خیر"):
        return "boolean"

    return "text"


def analyze_column(column_letter: str, column_index: int, header_name: str, values: List[Any]) -> ColumnSummary:
    total = len(values)
    non_null_values = [v for v in values if v is not None and str(v).strip() != ""]
    null_count = total - len(non_null_values)

    type_counts: Dict[str, int] = {}
    numeric_vals: List[float] = []
    seen = set()
    dup_count = 0
    error_cnt = 0

    for v in non_null_values:
        if is_excel_error(v):
            error_cnt += 1
            continue

        c_type = detect_cell_type(v)
        type_counts[c_type] = type_counts.get(c_type, 0) + 1

        # Check uniqueness
        v_str = str(v).strip().lower()
        if v_str in seen:
            dup_count += 1
        else:
            seen.add(v_str)

        # Collect numbers for stats
        if c_type in ("integer", "float", "percentage", "currency"):
            try:
                if isinstance(v, (int, float)):
                    numeric_vals.append(float(v))
                else:
                    s_clean = str(v).replace(',', '').replace('%', '').strip()
                    for sym in CURRENCY_SYMBOLS:
                        s_clean = s_clean.replace(sym, '').strip()
                    numeric_vals.append(float(s_clean))
            except ValueError:
                pass

    # Determine dominant type
    if not non_null_values:
        dom_type = "empty"
    else:
        # Group numeric sub-types
        num_types_sum = sum(type_counts.get(t, 0) for t in ("integer", "float", "percentage", "currency"))
        text_sum = type_counts.get("text", 0)
        date_sum = type_counts.get("date", 0)
        bool_sum = type_counts.get("boolean", 0)

        effective_count = len(non_null_values) - error_cnt
        if effective_count <= 0:
            dom_type = "error"
        elif num_types_sum / effective_count >= 0.7:
            dom_type = "numeric"
        elif date_sum / effective_count >= 0.7:
            dom_type = "date"
        elif bool_sum / effective_count >= 0.7:
            dom_type = "boolean"
        elif text_sum / effective_count >= 0.7:
            dom_type = "text"
        else:
            dom_type = "mixed"

    min_v = min(numeric_vals) if numeric_vals else (min(non_null_values) if non_null_values and dom_type != "mixed" else None)
    max_v = max(numeric_vals) if numeric_vals else (max(non_null_values) if non_null_values and dom_type != "mixed" else None)

    return ColumnSummary(
        column_letter=column_letter,
        column_index=column_index,
        header_name=header_name,
        detected_type=dom_type,
        total_count=total,
        non_null_count=len(non_null_values),
        null_count=null_count,
        unique_count=len(seen),
        duplicate_count=dup_count,
        sample_values=values[:5],
        min_val=min_v,
        max_val=max_v,
        numeric_values=numeric_vals,
        error_count=error_cnt,
        is_numeric=(dom_type in ("numeric", "integer", "float", "currency", "percentage"))
    )
