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
            raw_c = str(action.target_col).strip().upper()
            col_idx = int(raw_c) if raw_c.isdigit() else col_letter_to_index(raw_c)
            target_cols = [col_idx]
        else:
            target_cols = list(range(min_c, max_c + 1))

        highlighted_cells: List[str] = []
        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)

        # Distinct soft pastel color palette for grouping different duplicate items
        PASTEL_PALETTE = [
            "#FEF08A",  # زرد روشن (Light Yellow)
            "#BAE6FD",  # آبی روشن (Light Sky Blue)
            "#BBF7D0",  # سبز نعنایی روشن (Light Mint Green)
            "#FED7AA",  # هلویی روشن (Soft Peach / Orange)
            "#E9D5FF",  # یاسی روشن (Soft Lavender)
            "#FBCFE8",  # صورتی روشن (Light Rose Pink)
            "#A7F3D0",  # زمردی روشن (Light Emerald)
            "#DDD6FE",  # نیلی روشن (Light Violet)
            "#FECDD3",  # گلبهی روشن (Soft Coral)
            "#CFFAFE",  # فیروزه‌ای روشن (Light Cyan)
        ]

        def is_numeric_data(val) -> bool:
            if val is None:
                return False
            if isinstance(val, (int, float)):
                return True
            s = str(val).replace(',', '').strip()
            try:
                float(s)
                return True
            except ValueError:
                return False

        color_index = 0
        val_to_color = {}
        dup_summary = []

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
                        # Ignore pure numbers (e.g. 8, 10, quantities, prices)
                        if action.text_only and is_numeric_data(val):
                            continue
                        addr = f"{col_letter}{r}"
                        norm_val = str(val).strip()
                        key = norm_val.lower()
                        counts[key] += 1
                        cell_map.setdefault(key, []).append((addr, norm_val))

            # Highlight cells whose value count > 1
            for key, count in counts.items():
                if count > 1:
                    orig_display_val = cell_map[key][0][1]
                    if action.distinct_colors:
                        if key not in val_to_color:
                            val_to_color[key] = PASTEL_PALETTE[color_index % len(PASTEL_PALETTE)]
                            color_index += 1
                        cell_color = val_to_color[key]
                    else:
                        cell_color = action.color_hex

                    for addr, _ in cell_map[key]:
                        sheet.set_cell_color(addr, cell_color)
                        highlighted_cells.append(addr)

                    dup_summary.append(f"«{orig_display_val}» ({count} بار)")

        if not highlighted_cells:
            msg = "✅ هیچ دادهٔ متنی/حروفی تکراری در شیت یافت نشد (داده‌های عددی طبق درخواست نادیده گرفته شدند)."
            return ActionResult(success=True, message_persian=msg, data=[])

        dup_items_str = "، ".join(dup_summary[:8])
        if len(dup_summary) > 8:
            dup_items_str += " و موارد دیگر..."

        msg = (
            f"🎨 تعداد **{len(highlighted_cells)}** سلول مربوط به **{len(dup_summary)}** عبارت متنی تکراری با رنگ‌های مجزا هایلایت شدند.\n"
            f"🔹 موارد تکراری یافت‌شده: {dup_items_str}\n"
            f"💡 (نکته: طبق درخواست، داده‌های عددی مانند ۸ و ۱۰ رنگ نشدند و هر کالای متنی رنگ اختصاصی دریافت کرد)."
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
