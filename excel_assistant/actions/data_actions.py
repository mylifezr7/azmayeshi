"""
Data Management Actions for Excel Assistant.
Handles deleting empty rows, removing duplicate rows, and sorting data.
"""
from typing import Any, Dict, List, Set, Union

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.nlp.intent_models import DeleteEmptyRowsAction, DeleteDuplicatesAction, SortDataAction
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, bounds_to_range_str
)


class DataActionHandler:
    """Executes data cleaning, deduplication, and sorting actions on Excel sheets."""

    @classmethod
    def execute_delete_empty_rows(cls, action: DeleteEmptyRowsAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        header_r = min_r

        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)
        empty_rows: List[int] = []

        for r in range(header_r + 1, max_r + 1):
            r_offset = r - min_r
            if r_offset < len(matrix):
                row_vals = matrix[r_offset]
                if all(v is None or str(v).strip() == "" for v in row_vals):
                    empty_rows.append(r)

        if not empty_rows:
            return ActionResult(success=True, message_persian="ℹ️ هیچ ردیف کاملاً خالی در بین داده‌ها یافت نشد.")

        deleted_count = sheet.delete_rows(empty_rows)
        msg = f"🧹 تعداد **{deleted_count}** ردیف خالی با موفقیت از شیت حذف شد و داده‌ها یکپارچه‌سازی گردیدند."
        return ActionResult(success=True, message_persian=msg, data=deleted_count)

    @classmethod
    def execute_delete_duplicates(cls, action: DeleteDuplicatesAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        header_r = min_r

        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)
        seen: Set[str] = set()
        duplicate_rows: List[int] = []

        for r in range(header_r + 1, max_r + 1):
            r_offset = r - min_r
            if r_offset < len(matrix):
                row_vals = matrix[r_offset]
                if action.target_col:
                    target_idx = (col_letter_to_index(action.target_col) - min_c)
                    val = row_vals[target_idx] if 0 <= target_idx < len(row_vals) else None
                    key = str(val).strip().lower() if val is not None else ""
                else:
                    key = "|".join(str(v).strip().lower() for v in row_vals if v is not None)

                if key in seen:
                    duplicate_rows.append(r)
                else:
                    seen.add(key)

        if not duplicate_rows:
            return ActionResult(success=True, message_persian="ℹ️ هیچ ردیف تکراری برای حذف یافت نشد.")

        deleted_count = sheet.delete_rows(duplicate_rows)
        msg = f"🗑️ تعداد **{deleted_count}** ردیف تکراری حذف گردید و نسخه اول هر ردیف حفظ شد."
        return ActionResult(success=True, message_persian=msg, data=deleted_count)

    @classmethod
    def execute_sort(cls, action: SortDataAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        range_str = bounds_to_range_str(min_r, min_c, max_r, max_c)
        key_col = action.key_col.upper()

        sheet.sort_range(
            range_str=range_str,
            key_col=key_col,
            ascending=action.ascending,
            has_headers=True
        )

        dir_str = "صعودی (کمترین به بیشترین)" if action.ascending else "نزولی (بیشترین به کمترین)"
        msg = f"🔄 داده‌های محدوده `{range_str}` با موفقیت بر اساس ستون `{key_col}` به صورت **{dir_str}** مرتب شدند."
        return ActionResult(success=True, message_persian=msg, affected_range=range_str)
