"""
Mathematical and Statistical Actions for Excel Assistant.
Handles column arithmetic (multiply, add, subtract, divide) and statistical summaries.
"""
import statistics
from typing import Any, List, Optional

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.nlp.intent_models import ColumnArithmeticAction, ColumnStatsAction
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, bounds_to_range_str
)


class MathActionHandler:
    """Executes mathematical formulas and statistical calculations on Excel sheets."""

    @classmethod
    def execute_arithmetic(cls, action: ColumnArithmeticAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        c1 = action.col1.upper()
        c2 = action.col2.upper()

        # Target column determination
        if action.target_col:
            target_col = action.target_col.upper()
        else:
            # Safely pick the next empty column after all operands and used range
            next_idx = max(col_letter_to_index(c1), col_letter_to_index(c2), max_c) + 1
            target_col = col_index_to_letter(next_idx)

        op_symbols = {"multiply": "*", "add": "+", "subtract": "-", "divide": "/"}
        op_symbol = op_symbols.get(action.operation, "*")
        op_fa_names = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم"}
        op_fa = op_fa_names.get(action.operation, "محاسبه")

        # Determine whether row 1 contains header text or numeric data
        val1_r1 = sheet.get_value(f"{c1}1")
        val2_r1 = sheet.get_value(f"{c2}1")

        def is_numeric_value(v):
            if v is None:
                return False
            if isinstance(v, (int, float)):
                return True
            try:
                float(str(v).replace(',', '').strip())
                return True
            except ValueError:
                return False

        has_header = not (is_numeric_value(val1_r1) or is_numeric_value(val2_r1))

        if has_header:
            start_row = 2
            # Add header to target column
            header_title = f"{op_fa} {c1} و {c2}"
            sheet.set_value(f"{target_col}1", header_title)
            sheet.set_font(f"{target_col}1", bold=True)
        else:
            start_row = 1

        # Determine last row with actual data in either c1 or c2
        col1_vals = sheet.get_column_values(c1, include_header=True)
        col2_vals = sheet.get_column_values(c2, include_header=True)
        data_len = max(len(col1_vals), len(col2_vals))
        last_row = start_row
        for r_i in range(data_len, 0, -1):
            v1 = col1_vals[r_i - 1] if r_i <= len(col1_vals) else None
            v2 = col2_vals[r_i - 1] if r_i <= len(col2_vals) else None
            if (v1 is not None and str(v1).strip() != "") or (v2 is not None and str(v2).strip() != ""):
                last_row = max(start_row, r_i)
                break
        end_row = last_row

        # Circular Reference Guard: If target column is one of the source columns
        is_circular = (target_col == c1 or target_col == c2)

        if is_circular:
            # Compute values in Python and write numeric results so values are never zeroed
            calc_matrix = []
            for r in range(start_row, end_row + 1):
                v1 = sheet.get_value(f"{c1}{r}")
                v2 = sheet.get_value(f"{c2}{r}")
                try:
                    n1 = float(str(v1).replace(',', '').strip()) if v1 is not None else 0.0
                    n2 = float(str(v2).replace(',', '').strip()) if v2 is not None else 0.0
                    if action.operation == "multiply":
                        res_val = n1 * n2
                    elif action.operation == "add":
                        res_val = n1 + n2
                    elif action.operation == "subtract":
                        res_val = n1 - n2
                    elif action.operation == "divide":
                        res_val = (n1 / n2) if n2 != 0 else "#DIV/0!"
                    else:
                        res_val = 0.0
                except Exception:
                    res_val = 0.0
                calc_matrix.append([res_val])
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            sheet.set_value(target_range, calc_matrix)
        else:
            # Set exact formula matrix for each row
            formula_matrix = [[f"={c1}{r}{op_symbol}{c2}{r}"] for r in range(start_row, end_row + 1)]
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            sheet.set_formula(target_range, formula_matrix)

        sheet.auto_fit_columns([target_col])
        affected = f"{target_col}{start_row}:{target_col}{end_row}"

        # Trigger Excel calculation refresh if running via xlwings COM
        try:
            if hasattr(sheet, "_sheet") and hasattr(sheet._sheet, "book"):
                sheet._sheet.book.app.calculate()
        except Exception:
            pass

        msg = (
            f"✅ عملیات **{op_fa}** ستون `{c1}` در ستون `{c2}` با موفقیت انجام شد.\n"
            f"📌 نتیجه در ستون `{target_col}` (محدوده `{affected}`) درج شد."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=affected)

    @classmethod
    def execute_stats(cls, action: ColumnStatsAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        source_col = action.source_col.upper()
        values = sheet.get_column_values(source_col, include_header=False)

        # Extract numeric values
        num_vals: List[float] = []
        for v in values:
            if v is not None and not str(v).startswith('='):
                try:
                    num_vals.append(float(str(v).replace(',', '').strip()))
                except ValueError:
                    pass

        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        start_r = min_r + 1
        end_r = max_r
        range_ref = f"{source_col}{start_r}:{source_col}{end_r}"

        formula_map = {
            "sum": f"=SUM({range_ref})",
            "average": f"=AVERAGE({range_ref})",
            "max": f"=MAX({range_ref})",
            "min": f"=MIN({range_ref})",
            "median": f"=MEDIAN({range_ref})",
            "count": f"=COUNT({range_ref})"
        }

        stat_fa_map = {
            "sum": "مجموع",
            "average": "میانگین",
            "median": "میانه",
            "max": "بیشترین مقدار (حداکثر)",
            "min": "کمترین مقدار (حداقل)",
            "count": "تعداد داده‌ها",
            "percentage": "محاسبه درصد"
        }
        stat_fa = stat_fa_map.get(action.stat_type, "آمار")

        # Compute calculated value in Python
        calc_val = None
        if num_vals:
            if action.stat_type == "sum":
                calc_val = sum(num_vals)
            elif action.stat_type == "average":
                calc_val = statistics.mean(num_vals)
            elif action.stat_type == "median":
                calc_val = statistics.median(num_vals)
            elif action.stat_type == "max":
                calc_val = max(num_vals)
            elif action.stat_type == "min":
                calc_val = min(num_vals)
            elif action.stat_type == "count":
                calc_val = len(num_vals)

        # If user specified a cell or if we place summary at the bottom
        target_cell = action.target_cell
        if not target_cell:
            target_cell = f"{source_col}{max_r + 1}"
            label_cell = f"{col_index_to_letter(max(1, col_letter_to_index(source_col) - 1))}{max_r + 1}"
            sheet.set_value(label_cell, f"{stat_fa}:")
            sheet.set_font(label_cell, bold=True)

        if action.stat_type in formula_map:
            sheet.set_formula(target_cell, formula_map[action.stat_type])
            sheet.set_font(target_cell, bold=True)

        val_display = f"{calc_val:,.2f}".rstrip('0').rstrip('.') if calc_val is not None else "محاسبه با فرمول"
        msg = (
            f"📊 **{stat_fa}** ستون `{source_col}` محاسبه گردید:\n"
            f"🔹 مقدار محاسبه‌شده: **{val_display}**\n"
            f"🔹 فرمول اکسل `{formula_map.get(action.stat_type, '')}` در سلول `{target_cell}` ثبت شد."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=target_cell, data=calc_val)
