"""
Data Management Actions for Excel Assistant.
Handles deleting empty rows, removing duplicate rows, and sorting data.
"""
from typing import Any, Dict, List, Set, Union

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.nlp.intent_models import (
    DeleteEmptyRowsAction, DeleteDuplicatesAction, SortDataAction, InsertDateAction, PreventDuplicatesAction
)
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

    @classmethod
    def execute_insert_date(cls, action: InsertDateAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet

        # Normalize target column (e.g. '1' -> 'A', 'A' -> 'A')
        raw_c = str(action.target_col).strip().upper()
        col_let = col_index_to_letter(int(raw_c)) if raw_c.isdigit() else raw_c

        start_r = max(1, action.start_row)
        end_r = max(start_r, action.end_row)

        curr_y = action.year
        curr_m = action.month
        curr_d = action.day

        # Days in Shamsi months
        def days_in_shamsi_month(y, m):
            if 1 <= m <= 6:
                return 31
            elif 7 <= m <= 11:
                return 30
            else:
                return 30 if (y % 33) in [1, 5, 9, 13, 17, 22, 26, 30] else 29

        date_matrix = []
        for r in range(start_r, end_r + 1):
            date_str = f"{curr_y:04d}/{curr_m:02d}/{curr_d:02d}"
            date_matrix.append([date_str])

            if action.mode == "sequential":
                max_d = days_in_shamsi_month(curr_y, curr_m)
                curr_d += 1
                if curr_d > max_d:
                    curr_d = 1
                    curr_m += 1
                    if curr_m > 12:
                        curr_m = 1
                        curr_y += 1

        range_str = f"{col_let}{start_r}:{col_let}{end_r}"
        sheet.set_value(range_str, date_matrix)
        sheet.auto_fit_columns([col_let])

        # Optional header if start_r > 1
        if action.add_header and start_r > 1:
            sheet.set_value(f"{col_let}1", "تاریخ")
            sheet.set_font(f"{col_let}1", bold=True)

        mode_text = "به صورت متوالی و روزشمار" if action.mode == "sequential" else "به صورت تاریخ ثابت"
        first_date = date_matrix[0][0]
        last_date = date_matrix[-1][0]
        date_span = f"از `{first_date}` تا `{last_date}`" if action.mode == "sequential" else f"`{first_date}`"

        msg = (
            f"📅 **درج تاریخ در اکسل با موفقیت انجام شد.**\n"
            f"📌 ستون: `{col_let}` | ردیف‌های: `{start_r}` تا `{end_r}` ({end_r - start_r + 1} سلول)\n"
            f"📌 تاریخ: {date_span} ({mode_text})\n"
            f"📌 محدوده سلول‌ها: `{range_str}`"
        )
        return ActionResult(success=True, message_persian=msg, affected_range=range_str)

    @classmethod
    def execute_prevent_duplicates(cls, action: PreventDuplicatesAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        from collections import Counter

        # Normalize target column (e.g. '12' -> 'L', 'L' -> 'L')
        raw_c = str(action.target_col).strip().upper()
        col_let = col_index_to_letter(int(raw_c)) if raw_c.isdigit() else raw_c

        # 1. Install Excel Data Validation rule
        sheet.add_unique_validation(
            col=col_let,
            error_title=action.error_title,
            error_message=action.error_message
        )

        # 2. Check for any existing duplicates in this column
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        existing_dups = []
        if max_r > 1:
            vals = sheet.get_column_values(col_let, include_header=False)
            val_counts = Counter(str(v).strip().lower() for v in vals if v is not None and str(v).strip() != "")
            dup_keys = {k for k, count in val_counts.items() if count > 1}

            if dup_keys:
                for idx, v in enumerate(vals, start=2):
                    if v is not None and str(v).strip().lower() in dup_keys:
                        cell_addr = f"{col_let}{idx}"
                        existing_dups.append((cell_addr, str(v).strip()))
                        if action.highlight_existing:
                            sheet.set_cell_color(cell_addr, "#FECDD3")  # Soft red alert tint

        header_val = sheet.get_value(f"{col_let}1")
        col_desc = f"`{col_let}`" + (f" («{header_val}»)" if header_val else "")

        if existing_dups:
            dup_items_str = "، ".join(f"{addr}: «{val}»" for addr, val in existing_dups[:5])
            if len(existing_dups) > 5:
                dup_items_str += " و موارد دیگر..."
            msg = (
                f"🔒 **قاعدهٔ جلوگیری از ورود دادهٔ تکراری روی ستون {col_desc} فعال شد.**\n"
                f"🛑 **کنترل اکسل:** از این پس چنانچه مقداری تکراری در ستون `{col_let}` وارد شود، اکسل بلافاصله مانع شده و خطای «{action.error_title}» را به کاربر نمایش می‌دهد.\n\n"
                f"⚠️ **توجه:** تعداد **{len(existing_dups)}** مقدار تکراری از قبل در این ستون وجود دارد که با رنگ مشخص هایلایت شدند تا بررسی نمایید:\n"
                f"📍 نمونه‌ها: {dup_items_str}"
            )
        else:
            msg = (
                f"🔒 **قاعدهٔ جلوگیری از ورود دادهٔ تکراری روی ستون {col_desc} با موفقیت فعال شد.**\n"
                f"✅ تمام مقادیر فعلی ستون `{col_let}` یکتا هستند.\n"
                f"🛑 **کنترل اکسل:** از این پس اگر کاربری بخواهد یک کد محصول یا کد ملی تکراری وارد کند، اکسل بلافاصله خطای «{action.error_title}» نمایش داده و مانع از ثبت تکراری می‌شود."
            )

        return ActionResult(success=True, message_persian=msg, data={"column": col_let, "existing_duplicates": existing_dups})
