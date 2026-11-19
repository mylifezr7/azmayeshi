"""
Formatting Actions for Excel Assistant.
Handles conditional formatting, duplicate highlighting, font styling, column widths, and table creation.
"""
from collections import Counter
from typing import Any, List, Optional

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.nlp.intent_models import (
    HighlightConditionAction, HighlightDuplicatesAction,
    FormatCellsAction, AutoFitColumnsAction, CreateTableAction
)
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, parse_range_address, bounds_to_range_str
)


class FormatActionHandler:
    """Executes formatting and visual styling actions on Excel sheets."""

    @classmethod
    def execute_highlight_condition(cls, action: HighlightConditionAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        header_r = min_r
        start_r = min_r + 1

        if action.target_col:
            col_idx = col_letter_to_index(action.target_col)
            target_cols = [col_idx]
        else:
            target_cols = list(range(min_c, max_c + 1))

        matched_cells: List[str] = []
        thresh = float(action.value) if action.value is not None else 0.0

        # Fast in-memory evaluation using batch matrix
        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)

        for c in target_cols:
            col_letter = col_index_to_letter(c)
            c_offset = c - min_c
            for r in range(start_r, max_r + 1):
                r_offset = r - min_r
                if r_offset < len(matrix) and c_offset < len(matrix[r_offset]):
                    val = matrix[r_offset][c_offset]
                    if val is not None and not str(val).startswith('='):
                        try:
                            num = float(str(val).replace(',', '').strip())
                            is_match = False
                            if action.condition in ("gt", "gte"):
                                is_match = (num >= thresh) if action.condition == "gte" else (num > thresh)
                            elif action.condition in ("lt", "lte"):
                                is_match = (num <= thresh) if action.condition == "lte" else (num < thresh)
                            elif action.condition == "eq":
                                is_match = (num == thresh)

                            if is_match:
                                addr = f"{col_letter}{r}"
                                sheet.set_cell_color(addr, action.color_hex)
                                matched_cells.append(addr)
                        except ValueError:
                            pass

        cond_fa = {"gt": "بیشتر از", "lt": "کمتر از", "eq": "مساوی با"}.get(action.condition, "مطابق شرط")
        msg = (
            f"🎨 تعداد **{len(matched_cells)}** سلول {cond_fa} `{action.value}` شناسایی شد و با رنگ مشخص‌شده هایلایت گردید.\n"
            f"📍 سلول‌ها: {', '.join(matched_cells[:10])}{' ...' if len(matched_cells) > 10 else ''}"
        )
        return ActionResult(success=True, message_persian=msg, data=matched_cells)

    @classmethod
    def execute_highlight_duplicates(cls, action: HighlightDuplicatesAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        start_r = min_r + 1

        if action.target_col:
            col_idx = col_letter_to_index(action.target_col)
            target_cols = [col_idx]
        else:
            target_cols = list(range(min_c, max_c + 1))

        highlighted_cells: List[str] = []
        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)

        for c in target_cols:
            col_letter = col_index_to_letter(c)
            c_offset = c - min_c
            cell_map = {}
            counts = Counter()

            for r in range(start_r, max_r + 1):
                r_offset = r - min_r
                if r_offset < len(matrix) and c_offset < len(matrix[r_offset]):
                    val = matrix[r_offset][c_offset]
                    if val is not None and str(val).strip() != "":
                        addr = f"{col_letter}{r}"
                        norm_val = str(val).strip().lower()
                        counts[norm_val] += 1
                        cell_map.setdefault(norm_val, []).append(addr)

            # Highlight cells whose value count > 1
            for norm_val, count in counts.items():
                if count > 1:
                    for addr in cell_map[norm_val]:
                        sheet.set_cell_color(addr, action.color_hex)
                        highlighted_cells.append(addr)

        msg = (
            f"🎨 تعداد **{len(highlighted_cells)}** سلول تکراری شناسایی و با رنگ `{action.color_hex}` هایلایت شدند.\n"
            f"📍 نمونه سلول‌ها: {', '.join(highlighted_cells[:10])}{' ...' if len(highlighted_cells) > 10 else ''}"
        )
        return ActionResult(success=True, message_persian=msg, data=highlighted_cells)

    @classmethod
    def execute_autofit(cls, action: AutoFitColumnsAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        sheet.auto_fit_columns(action.columns)
        return ActionResult(success=True, message_persian="📐 عرض تمام ستون‌ها به طور خودکار بر اساس طول محتوا تنظیم گردید.")

    @classmethod
    def execute_create_table(cls, action: CreateTableAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        range_str = action.range_str or bounds_to_range_str(min_r, min_c, max_r, max_c)
        table_name = sheet.create_table(range_str, name=action.table_name)
        msg = f"📊 محدوده `{range_str}` با موفقیت به جدول رسمی اکسل با نام `{table_name}` تبدیل شد و قابلیت فیلترگذاری فعال گردید."
        return ActionResult(success=True, message_persian=msg, affected_range=range_str)

    @classmethod
    def execute_format_cells(cls, action: FormatCellsAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        target_range = action.target_range
        if action.bold is not None or action.italic is not None or action.font_size or action.font_name:
            sheet.set_font(target_range, name=action.font_name, size=action.font_size, bold=action.bold, italic=action.italic)
        if action.bg_color:
            sheet.set_cell_color(target_range, action.bg_color)
        return ActionResult(success=True, message_persian=f"قالب‌بندی سلول‌های محدوده `{target_range}` با موفقیت اعمال شد.")
