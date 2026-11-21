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

        # Extract and normalize all source columns (support letters or numbers e.g. '2', '4', '5' -> 'B', 'D', 'E')
        if action.columns and len(action.columns) >= 2:
            raw_cols = action.columns
        else:
            raw_cols = [action.col1 or "A", action.col2 or "B"]

        cols: List[str] = []
        for c in raw_cols:
            c_str = str(c).strip().upper()
            if c_str.isdigit():
                c_str = col_index_to_letter(int(c_str))
            if c_str and c_str not in cols:
                cols.append(c_str)

        if len(cols) < 2:
            cols = ["A", "B"]

        # Target column determination (support letter or number e.g. '8' -> 'H')
        if action.target_col:
            t_str = str(action.target_col).strip().upper()
            if t_str.isdigit():
                target_col = col_index_to_letter(int(t_str))
            else:
                target_col = t_str
        else:
            # Safely pick the next empty column after all operands and used range
            max_col_idx = max(col_letter_to_index(c) for c in cols)
            next_idx = max(max_col_idx, max_c) + 1
            target_col = col_index_to_letter(next_idx)

        op_symbols = {"multiply": "*", "add": "+", "subtract": "-", "divide": "/"}
        op_symbol = op_symbols.get(action.operation, "*")
        op_fa_names = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم"}
        op_fa = op_fa_names.get(action.operation, "محاسبه")

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

        # Determine whether row 1 contains header text or numeric data
        has_header = False
        for c in cols:
            val_r1 = sheet.get_value(f"{c}1")
            if val_r1 is not None and not is_numeric_value(val_r1):
                has_header = True
                break

        if has_header:
            start_row = 2
            # Add header to target column
            header_title = f"{op_fa} {' و '.join(cols)}"
            sheet.set_value(f"{target_col}1", header_title)
            sheet.set_font(f"{target_col}1", bold=True)
        else:
            start_row = 1

        # Determine last row with actual data in any operand column
        max_data_len = start_row
        for c in cols:
            c_vals = sheet.get_column_values(c, include_header=True)
            for r_i in range(len(c_vals), 0, -1):
                v = c_vals[r_i - 1]
                if v is not None and str(v).strip() != "":
                    max_data_len = max(max_data_len, r_i)
                    break
        end_row = max(start_row, max_data_len)

        # Circular Reference Guard: If target column is one of the source columns
        is_circular = (target_col in cols)

        if is_circular:
            # Compute values in Python and write numeric results so values are never zeroed
            calc_matrix = []
            for r in range(start_row, end_row + 1):
                row_nums = []
                for c in cols:
                    v = sheet.get_value(f"{c}{r}")
                    try:
                        n = float(str(v).replace(',', '').strip()) if v is not None else 0.0
                    except Exception:
                        n = 0.0
                    row_nums.append(n)

                if action.operation == "add":
                    res_val = sum(row_nums)
                elif action.operation == "multiply":
                    res_val = 1.0
                    for n in row_nums:
                        res_val *= n
                elif action.operation == "subtract":
                    res_val = row_nums[0]
                    for n in row_nums[1:]:
                        res_val -= n
                elif action.operation == "divide":
                    res_val = row_nums[0]
                    for n in row_nums[1:]:
                        res_val = (res_val / n) if n != 0 else "#DIV/0!"
                else:
                    res_val = sum(row_nums)

                calc_matrix.append([res_val])
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            sheet.set_value(target_range, calc_matrix)
        else:
            # Set exact formula matrix for each row e.g. =B2+D2+E2
            formula_matrix = [
                [f"={op_symbol.join(f'{c}{r}' for c in cols)}"]
                for r in range(start_row, end_row + 1)
            ]
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

        cols_display = "، ".join(cols)
        msg = (
            f"✅ عملیات **{op_fa}** ستون‌های `{cols_display}` با موفقیت انجام شد.\n"
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
