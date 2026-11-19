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
        header_r = 1

        c1 = action.col1.upper()
        c2 = action.col2.upper()

        # If target column not specified, append as new column after max_c
        if not action.target_col:
            target_col = col_index_to_letter(max_c + 1)
        else:
            target_col = action.target_col.upper()

        op_symbols = {"multiply": "*", "add": "+", "subtract": "-", "divide": "/"}
        op_symbol = op_symbols.get(action.operation, "*")
        op_fa_names = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم"}
        op_fa = op_fa_names.get(action.operation, "محاسبه")

        # Set target column header
        header_title = f"{op_fa} {c1} و {c2}"
        sheet.set_value(f"{target_col}{header_r}", header_title)
        sheet.set_font(f"{target_col}{header_r}", bold=True)

        start_row = header_r + 1
        end_row = max(max_r, start_row)

        # High-speed batch formula assignment
        if end_row >= start_row:
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            base_formula = f"={c1}{start_row}{op_symbol}{c2}{start_row}"
            sheet.set_formula(target_range, base_formula)

        sheet.auto_fit_columns([target_col])
        affected = f"{target_col}{start_row}:{target_col}{end_row}"

        msg = (
            f"✅ عملیات **{op_fa}** ستون `{c1}` در ستون `{c2}` با موفقیت انجام شد.\n"
            f"📌 نتیجه به صورت فرمول خودکار در ستون `{target_col}` (محدوده `{affected}`) درج شد."
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
