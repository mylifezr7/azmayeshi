"""
Deep File Auditor for MUYUM Excel Assistant.
Performs smart multi-dimensional health check:
- Empty cells (سلول‌های خالی)
- Duplicate information (اطلاعات تکراری)
- Formula errors (فرمول‌های اشتباه #REF!, #DIV/0!, #VALUE!)
- Incomplete columns (ستون‌های ناقص)
- Invalid data types (داده‌های نامعتبر)
- Calculation discrepancies (خطاهای محاسباتی)
"""
from dataclasses import dataclass, field
import re
from typing import Any, Dict, List, Optional, Tuple

from excel_assistant.core.interfaces import ISheet
from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index


@dataclass
class DeepAuditResult:
    health_score: int  # 0 to 100
    empty_cells: List[str] = field(default_factory=list)
    duplicate_rows: List[int] = field(default_factory=list)
    formula_errors: List[Dict[str, str]] = field(default_factory=list)
    incomplete_columns: List[Dict[str, Any]] = field(default_factory=list)
    invalid_data_cells: List[Dict[str, str]] = field(default_factory=list)
    calculation_errors: List[Dict[str, Any]] = field(default_factory=list)
    persian_summary: str = ""


class DeepFileAuditor:
    """Performs comprehensive diagnostics on an Excel sheet."""

    @classmethod
    def audit_sheet(cls, sheet: ISheet) -> DeepAuditResult:
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)
        formula_matrix = sheet.get_formula_matrix(min_r, min_c, max_r, max_c)

        empty_cells: List[str] = []
        formula_errors: List[Dict[str, str]] = []
        invalid_data_cells: List[Dict[str, str]] = []
        calculation_errors: List[Dict[str, Any]] = []

        header_row_offset = 0
        total_data_rows = max(0, len(matrix) - 1)

        # 1. Formula & cell-level scan
        error_tokens = {"#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}

        for r_idx, row in enumerate(matrix):
            r_num = min_r + r_idx
            f_row = formula_matrix[r_idx] if r_idx < len(formula_matrix) else []

            # Check if row is not completely empty
            row_has_data = any(v is not None and str(v).strip() != "" for v in row)
            if not row_has_data:
                continue

            for c_idx, val in enumerate(row):
                c_num = min_c + c_idx
                cell_addr = f"{col_index_to_letter(c_num)}{r_num}"
                form = f_row[c_idx] if c_idx < len(f_row) else None

                # Empty cell in a row that has data (skip header row)
                if r_idx > 0 and (val is None or str(val).strip() == ""):
                    empty_cells.append(cell_addr)

                # Check formula errors
                if form:
                    form_s = str(form).strip()
                    if any(tok in form_s for tok in ["#REF!", "#VALUE!"]):
                        formula_errors.append({
                            "cell": cell_addr,
                            "error": "خطای مرجع یا مقدار در فرمول",
                            "formula": form_s
                        })

                # Check cell value error tokens
                if val is not None:
                    val_s = str(val).strip().upper()
                    if val_s in error_tokens:
                        formula_errors.append({
                            "cell": cell_addr,
                            "error": f"خطای اکسل {val_s}",
                            "formula": str(form or val)
                        })

        # 2. Duplicate rows check
        row_hashes: Dict[str, int] = {}
        duplicate_rows: List[int] = []
        for r_idx in range(1, len(matrix)):
            row = matrix[r_idx]
            if all(v is None or str(v).strip() == "" for v in row):
                continue
            r_str = "|".join(str(v).strip().lower() for v in row if v is not None)
            r_num = min_r + r_idx
            if r_str in row_hashes:
                duplicate_rows.append(r_num)
            else:
                row_hashes[r_str] = r_num

        # 3. Incomplete columns & invalid data detection
        incomplete_columns: List[Dict[str, Any]] = []
        for c_idx in range(len(matrix[0]) if matrix else 0):
            c_num = min_c + c_idx
            col_letter = col_index_to_letter(c_num)
            header_val = str(matrix[0][c_idx] or f"ستون {col_letter}").strip()

            filled_count = 0
            numeric_count = 0
            text_in_numeric = []

            for r_idx in range(1, len(matrix)):
                r_num = min_r + r_idx
                val = matrix[r_idx][c_idx]
                if val is not None and str(val).strip() != "":
                    filled_count += 1
                    s_v = str(val).replace(",", "").strip()
                    try:
                        float(s_v)
                        numeric_count += 1
                    except ValueError:
                        # Non-numeric text
                        text_in_numeric.append((f"{col_letter}{r_num}", str(val)))

            fill_pct = (filled_count / total_data_rows * 100) if total_data_rows > 0 else 100
            if fill_pct < 90 and total_data_rows > 3:
                incomplete_columns.append({
                    "column": col_letter,
                    "header": header_val,
                    "fill_rate": round(fill_pct, 1),
                    "missing_count": total_data_rows - filled_count
                })

            # If column is mostly numeric (>75%) but has stray text strings
            if total_data_rows > 0 and numeric_count >= (filled_count * 0.75) and text_in_numeric:
                for c_addr, val_str in text_in_numeric[:5]:
                    invalid_data_cells.append({
                        "cell": c_addr,
                        "value": val_str,
                        "issue": f"ورود متن در ستون عددی «{header_val}»"
                    })

        # 4. Calculation discrepancy check (e.g. Qty * Unit Price != Total)
        # Scan if there are pairs of columns that look like Qty & Price and a Total column
        if len(matrix) > 1 and len(matrix[0]) >= 3:
            for r_idx in range(1, min(len(matrix), 50)):
                row = matrix[r_idx]
                r_num = min_r + r_idx
                num_vals = []
                for c_idx, val in enumerate(row):
                    try:
                        num_vals.append((c_idx, float(str(val).replace(",", "").strip())))
                    except (ValueError, TypeError):
                        pass

                # Check if any valA * valB roughly equals valC
                if len(num_vals) >= 3:
                    for i in range(len(num_vals)):
                        for j in range(i + 1, len(num_vals)):
                            for k in range(len(num_vals)):
                                if k != i and k != j:
                                    cA, vA = num_vals[i]
                                    cB, vB = num_vals[j]
                                    cC, vC = num_vals[k]
                                    if vA > 0 and vB > 0 and vC > 0:
                                        expected = vA * vB
                                        if abs(expected - vC) > 1 and abs(expected - vC) / max(expected, 1) < 0.2:
                                            # Close but deviates
                                            calculation_errors.append({
                                                "row": r_num,
                                                "cell": f"{col_index_to_letter(min_c + cC)}{r_num}",
                                                "expected": expected,
                                                "actual": vC,
                                                "diff": round(vC - expected, 2)
                                            })

        # Calculate overall health score
        penalty = 0
        penalty += min(25, len(empty_cells) * 2)
        penalty += min(25, len(duplicate_rows) * 5)
        penalty += min(30, len(formula_errors) * 10)
        penalty += min(15, len(invalid_data_cells) * 3)
        penalty += min(15, len(calculation_errors) * 5)
        health_score = max(0, 100 - penalty)

        # Build Persian summary report
        score_emoji = "🟢" if health_score >= 85 else ("🟡" if health_score >= 60 else "🔴")
        status_label = "عالی و بدون مشکل اساسی" if health_score >= 85 else ("نیازمند اصلاحات جزئی" if health_score >= 60 else "دارای خطاهای متعدد و نیازمند بررسی فوری")

        lines = [
            f"🩺 **گزارش تحلیل هوشمند و سلامت فایل اکسل:**\n",
            f"• نمره سلامت ساختار فایل: {score_emoji} **{health_score} از ۱۰۰** ({status_label})\n",
            f"• کل ردیف‌های داده: **{total_data_rows:,}** ردیف | ستون‌ها: **{len(matrix[0]) if matrix else 0}** ستون\n"
        ]

        # 1. Empty cells
        if empty_cells:
            sample_empty = ", ".join(empty_cells[:8])
            lines.append(f"⚠️ **سلول‌های خالی:** {len(empty_cells)} سلول فاقد داده (نمونه: {sample_empty})")
        else:
            lines.append("✅ **سلول‌های خالی:** هیچ سلول خالی در ردیف‌های فعال یافت نشد.")

        # 2. Duplicates
        if duplicate_rows:
            sample_dup = ", ".join(str(r) for r in duplicate_rows[:6])
            lines.append(f"⚠️ **اطلاعات تکراری:** {len(duplicate_rows)} ردیف تکراری شناسایی شد (ردیف‌های: {sample_dup})")
        else:
            lines.append("✅ **داده‌های تکراری:** ردیف تکراری یافت نشد.")

        # 3. Formula errors
        if formula_errors:
            err_details = ", ".join(f"{item['cell']} ({item['error']})" for item in formula_errors[:5])
            lines.append(f"❌ **فرمول‌های اشتباه:** {len(formula_errors)} مورد خطای محاسباتی در فرمول‌ها ({err_details})")
        else:
            lines.append("✅ **فرمول‌های اکسل:** کلیه فرمول‌ها سالم و بدون خطای توکن (#REF! / #DIV/0!) می‌باشند.")

        # 4. Incomplete columns
        if incomplete_columns:
            inc_details = ", ".join(f"{c['header']} (ستون {c['column']}: {c['fill_rate']}٪)" for c in incomplete_columns[:4])
            lines.append(f"⚠️ **ستون‌های ناقص:** ستون‌های {inc_details} دارای درصد تکمیل پایینی هستند.")
        else:
            lines.append("✅ **کامل بودن ستون‌ها:** تمامی ستون‌ها پوشش کامل دارند.")

        # 5. Invalid data
        if invalid_data_cells:
            inv_details = ", ".join(f"{c['cell']} مقدار «{c['value']}»" for c in invalid_data_cells[:4])
            lines.append(f"⚠️ **داده‌های نامعتبر:** ورود متن نامناسب در ستون عددی ({inv_details})")

        # 6. Calculation discrepancies
        if calculation_errors:
            lines.append(f"⚠️ **خطاهای محاسباتی:** {len(calculation_errors)} مورد اختلاف محاسباتی مقادیر شناسایی شد.")

        lines.append("\n💡 **پیشنهاد هوشمند MUYUM:** می‌توانید با یک کلیک یا دستور، سطرهای تکراری را حذف، سلول‌های خالی را پر، یا فرمت ستون‌ها را استانداردسازی کنید.")

        summary_persian = "\n".join(lines)

        return DeepAuditResult(
            health_score=health_score,
            empty_cells=empty_cells,
            duplicate_rows=duplicate_rows,
            formula_errors=formula_errors,
            incomplete_columns=incomplete_columns,
            invalid_data_cells=invalid_data_cells,
            calculation_errors=calculation_errors,
            persian_summary=summary_persian
        )
