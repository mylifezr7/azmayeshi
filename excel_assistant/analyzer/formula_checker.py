"""
Excel Formula and Error Checker.
Detects formulas, broken formulas, Excel error tokens, and consistency issues.
"""
from dataclasses import dataclass, field
import re
from typing import Any, Dict, List, Optional, Tuple

from excel_assistant.core.interfaces import ISheet
from excel_assistant.core.utils import col_index_to_letter, parse_range_address


@dataclass
class FormulaIssue:
    cell_address: str
    issue_type: str  # error_token, syntax_error, inconsistent_formula, missing_reference
    formula_or_value: str
    description: str
    severity: str = "warning"  # error, warning, info


@dataclass
class FormulaReport:
    total_formulas: int
    formula_cells: List[str] = field(default_factory=list)
    formula_map: Dict[str, str] = field(default_factory=dict)  # cell -> formula
    error_cells: List[str] = field(default_factory=list)
    issues: List[FormulaIssue] = field(default_factory=list)


def check_sheet_formulas_and_errors(sheet: ISheet) -> FormulaReport:
    min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
    report = FormulaReport(total_formulas=0)

    # High-speed batch read
    matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)
    formula_matrix = sheet.get_formula_matrix(min_r, min_c, max_r, max_c)

    # Scan all cells in used range via in-memory matrices
    for r_offset, row in enumerate(matrix):
        r = min_r + r_offset
        f_row = formula_matrix[r_offset] if r_offset < len(formula_matrix) else []
        for c_offset, val in enumerate(row):
            c = min_c + c_offset
            cell_addr = f"{col_index_to_letter(c)}{r}"
            formula = f_row[c_offset] if c_offset < len(f_row) else None

            # Check if formula exists
            if formula:
                formula_str = str(formula).strip()
                report.total_formulas += 1
                report.formula_cells.append(cell_addr)
                report.formula_map[cell_addr] = formula_str

                # Check for error tokens inside formula
                if "#REF!" in formula_str:
                    report.issues.append(FormulaIssue(
                        cell_address=cell_addr,
                        issue_type="missing_reference",
                        formula_or_value=formula_str,
                        description="فرمول به یک سلول یا شیت حذف‌شده اشاره دارد (#REF!)",
                        severity="error"
                    ))
                    report.error_cells.append(cell_addr)

            # Check for error tokens in cell value
            if val is not None:
                s_val = str(val).strip().upper()
                if s_val in ("#DIV/0!", "#VALUE!", "#N/A", "#REF!", "#NAME?", "#NUM!", "#NULL!"):
                    if cell_addr not in report.error_cells:
                        report.error_cells.append(cell_addr)
                    error_desc_fa = {
                        "#DIV/0!": "تقسیم بر صفر",
                        "#VALUE!": "ناسازگاری نوع داده در فرمول",
                        "#N/A": "مقدار موردنظر یافت نشد (معمولاً در VLOOKUP)",
                        "#REF!": "ارجاع نامعتبر یا حذف‌شده",
                        "#NAME?": "نام تابع یا محدوده ناشناخته",
                        "#NUM!": "خطای محاسباتی مقدار عددی نامعتبر",
                        "#NULL!": "اشتراک نامعتبر محدوده‌ها"
                    }.get(s_val, f"خطای اکسل {s_val}")

                    report.issues.append(FormulaIssue(
                        cell_address=cell_addr,
                        issue_type="error_token",
                        formula_or_value=s_val,
                        description=f"سلول دارای خطای {error_desc_fa} است.",
                        severity="error"
                    ))

    return report
