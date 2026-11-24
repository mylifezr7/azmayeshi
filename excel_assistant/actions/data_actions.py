"""
Data Management Actions for Excel Assistant.
Handles deleting empty rows, removing duplicate rows, sorting data, filtering, text conversions, splitting, merging, and formatting.
"""
from collections import Counter
import re
from typing import Any, Dict, List, Set, Union

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.nlp.intent_models import (
    DeleteEmptyRowsAction, DeleteDuplicatesAction, SortDataAction, InsertDateAction, PreventDuplicatesAction,
    RemoveUniqueValidationAction, FilterDataAction, ConvertTextToNumberAction, SplitTextAction,
    MergeCellsAction, SetCellFormatAction
)
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, bounds_to_range_str, parse_range_address
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

    @classmethod
    def execute_remove_unique_validation(cls, action: RemoveUniqueValidationAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet

        # Normalize target column (e.g. '12' -> 'L', 'L' -> 'L')
        raw_c = str(action.target_col).strip().upper()
        col_let = col_index_to_letter(int(raw_c)) if raw_c.isdigit() else raw_c

        # 1. Remove Excel Data Validation rule
        sheet.remove_unique_validation(col=col_let)

        # 2. Clear duplicate warning highlights if requested
        cleared_count = 0
        if action.clear_highlight:
            min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
            if max_r > 1:
                for idx in range(2, max_r + 1):
                    cell_addr = f"{col_let}{idx}"
                    curr_color = sheet.get_cell_color(cell_addr)
                    if curr_color and curr_color.upper() in ["#FECDD3", "FECDD3"]:
                        sheet.set_cell_color(cell_addr, None)
                        cleared_count += 1

        header_val = sheet.get_value(f"{col_let}1")
        col_desc = f"`{col_let}`" + (f" («{header_val}»)" if header_val else "")

        msg = (
            f"🔓 **قفل یکتاسازی ستون {col_desc} با موفقیت برداشته شد.**\n"
            f"✅ محدودیت اعتبارسنجی داده (Data Validation) از ستون `{col_let}` حذف گردید.\n"
            f"📝 اکنون کاربران می‌توانند مقادیر آزاد (حتی تکراری) را در این ستون بدون دریافت خطای مانع‌شونده ثبت نمایند."
        )
        return ActionResult(
            success=True,
            message_persian=msg,
            data={"column": col_let, "cleared_highlight_cells": cleared_count}
        )

    @classmethod
    def execute_filter(cls, action: FilterDataAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        start_r = min_r + 1
        range_str = action.range_str or bounds_to_range_str(min_r, min_c, max_r, max_c)

        if action.condition == "clear":
            # Unhide all rows
            for r in range(start_r, max_r + 1):
                sheet.set_row_hidden(r, False)
            return ActionResult(
                success=True,
                message_persian="🔄 تمام فیلترها پاک شدند و کلیه سطرهای جدول مجدداً نمایش داده شدند.",
                affected_range=range_str
            )

        if action.condition == "autofilter" or not action.target_col:
            sheet.apply_filter(range_str)
            return ActionResult(
                success=True,
                message_persian="🔘 **فیلتر هوشمند (AutoFilter)** روی سرستون‌های جدول با موفقیت فعال شد. اکنون می‌توانید روی علامت فلش هر ستون در اکسل کلیک کرده و داده‌ها را فیلتر کنید.",
                affected_range=range_str
            )

        # Apply specific filter condition
        sheet.apply_filter(range_str)
        raw_c = str(action.target_col).strip().upper()
        col_let = col_index_to_letter(int(raw_c)) if raw_c.isdigit() else raw_c

        vals = sheet.get_column_values(col_let, include_header=True)
        thresh = None
        if action.value is not None:
            try:
                thresh = float(str(action.value).replace(',', '').strip())
            except ValueError:
                pass
        str_val = str(action.value).strip().lower() if action.value is not None else ""

        matched_count = 0
        hidden_count = 0

        for r in range(start_r, max_r + 1):
            idx_in_vals = r - 1
            v = vals[idx_in_vals] if idx_in_vals < len(vals) else None
            is_match = False

            if v is not None:
                v_str = str(v).strip().lower()
                num_v = None
                try:
                    num_v = float(str(v).replace(',', '').strip())
                except ValueError:
                    pass

                if action.condition in ("gt", "gte") and thresh is not None and num_v is not None:
                    is_match = (num_v >= thresh) if action.condition == "gte" else (num_v > thresh)
                elif action.condition in ("lt", "lte") and thresh is not None and num_v is not None:
                    is_match = (num_v <= thresh) if action.condition == "lte" else (num_v < thresh)
                elif action.condition == "eq":
                    if thresh is not None and num_v is not None:
                        is_match = (num_v == thresh)
                    else:
                        is_match = (v_str == str_val)
                elif action.condition == "neq":
                    if thresh is not None and num_v is not None:
                        is_match = (num_v != thresh)
                    else:
                        is_match = (v_str != str_val)
                elif action.condition == "contains":
                    is_match = (str_val in v_str)
            else:
                is_match = False

            if is_match:
                sheet.set_row_hidden(r, False)
                matched_count += 1
            else:
                sheet.set_row_hidden(r, True)
                hidden_count += 1

        cond_fa_map = {
            "gt": "بزرگتر از",
            "gte": "بزرگتر یا مساوی با",
            "lt": "کوچکتر از",
            "lte": "کوچکتر یا مساوی با",
            "eq": "برابر با",
            "neq": "نامساوی با",
            "contains": "شامل عبارت"
        }
        cond_fa = cond_fa_map.get(action.condition, "مطابق شرط")

        msg = (
            f"🔍 **فیلتر هوشمند روی ستون `{col_let}` اعمال شد:**\n"
            f"📌 شرط: مقادیر **{cond_fa} `{action.value}`**\n"
            f"👁️ تعداد **{matched_count}** سطر مطابق شرط نمایش داده شد (و {hidden_count} سطر دیگر مخفی گردید)."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=range_str, data={"matched": matched_count, "hidden": hidden_count})

    @classmethod
    def execute_convert_text_to_number(cls, action: ConvertTextToNumberAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        start_r = min_r + 1

        raw_c = str(action.target_col or "A").strip().upper()
        col_let = col_index_to_letter(int(raw_c)) if raw_c.isdigit() else raw_c

        digit_trans = str.maketrans(
            "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
            "01234567890123456789"
        )

        converted_count = 0
        for r in range(start_r, max_r + 1):
            cell_addr = f"{col_let}{r}"
            val = sheet.get_value(cell_addr)
            if val is None or str(val).startswith('='):
                continue

            val_str = str(val).translate(digit_trans)
            clean = val_str.replace(',', '').replace('،', '').replace(' ', '').replace('تومان', '').replace('ریال', '').replace('$', '').replace('€', '').strip()
            if not clean:
                continue

            try:
                if '.' in clean:
                    num_val = float(clean)
                else:
                    num_val = int(clean)

                sheet.set_value(cell_addr, num_val)
                sheet.set_number_format(cell_addr, "#,##0" if isinstance(num_val, int) else "#,##0.00")
                converted_count += 1
            except ValueError:
                pass

        col_range = f"{col_let}{start_r}:{col_let}{max_r}"
        sheet.auto_fit_columns([col_let])
        msg = (
            f"🔢 **اصلاح و تبدیل متن به عدد با موفقیت انجام شد:**\n"
            f"✅ تعداد **{converted_count}** سلول متنی در ستون `{col_let}` به **عدد واقعی اکسل** تبدیل شد.\n"
            f"✨ فرمت جداسازی ارقام (سه‌رقم‌سه‌رقم) نیز روی آن‌ها اعمال گردید."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=col_range, data=converted_count)

    @classmethod
    def execute_split_text(cls, action: SplitTextAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        header_r = min_r
        start_r = min_r + 1

        raw_c = str(action.source_col).strip().upper()
        s_col = col_index_to_letter(int(raw_c)) if raw_c.isdigit() else raw_c

        if action.target_col1 and action.target_col2:
            t1_col = action.target_col1.upper()
            t2_col = action.target_col2.upper()
        else:
            t1_idx = max_c + 1
            t2_idx = max_c + 2
            t1_col = col_index_to_letter(t1_idx)
            t2_col = col_index_to_letter(t2_idx)

        sheet.set_value(f"{t1_col}{header_r}", action.header1 or "بخش ۱")
        sheet.set_value(f"{t2_col}{header_r}", action.header2 or "بخش ۲")
        sheet.set_font(f"{t1_col}{header_r}", bold=True)
        sheet.set_font(f"{t2_col}{header_r}", bold=True)

        split_count = 0
        for r in range(start_r, max_r + 1):
            val = sheet.get_value(f"{s_col}{r}")
            if val is not None and str(val).strip():
                txt = str(val).strip()
                if action.delimiter == " ":
                    parts = txt.split(None, 1)
                else:
                    parts = txt.split(action.delimiter, 1)

                part1 = parts[0].strip() if len(parts) > 0 else ""
                part2 = parts[1].strip() if len(parts) > 1 else ""

                sheet.set_value(f"{t1_col}{r}", part1)
                sheet.set_value(f"{t2_col}{r}", part2)
                split_count += 1

        sheet.auto_fit_columns([t1_col, t2_col])
        affected = f"{t1_col}{header_r}:{t2_col}{max_r}"
        msg = (
            f"✂️ **تفکیک و جداسازی متن ستون `{s_col}` با موفقیت انجام شد:**\n"
            f"🔹 بخش اول («{action.header1}») در ستون `{t1_col}` درج شد.\n"
            f"🔹 بخش دوم («{action.header2}») در ستون `{t2_col}` درج شد.\n"
            f"📊 تعداد ردیف‌های تفکیک‌شده: **{split_count}** مورد."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=affected, data=split_count)

    @classmethod
    def execute_merge_cells(cls, action: MergeCellsAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        clean_rng = action.range_str.strip().upper()
        sheet.merge_cells(clean_rng)
        sheet.set_font(clean_rng, bold=True)

        msg = f"🔗 سلول‌های محدوده **`{clean_rng}`** با موفقیت ادغام (Merge) و محتوای آن در مرکز تراز گردید."
        return ActionResult(success=True, message_persian=msg, affected_range=clean_rng)

    @classmethod
    def execute_set_cell_format(cls, action: SetCellFormatAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        target = action.target.strip().upper()
        if ":" not in target and re.match(r"^[A-Z]+$", target):
            target_range = f"{target}1:{target}{max(max_r, 1)}"
            col_letter = target
        else:
            target_range = target
            col_letter = target.split(":")[0] if ":" in target else target

        fmt_type = action.format_type
        if action.custom_format:
            fmt_code = action.custom_format
            fmt_name_fa = f"سفارشی ({fmt_code})"
        elif fmt_type == "currency":
            fmt_code = '#,##0 "تومان"'
            fmt_name_fa = "واحد پول (تومان)"
        elif fmt_type in ("number", "thousands"):
            fmt_code = "#,##0"
            fmt_name_fa = "عدد با جداکننده هزارگان"
        elif fmt_type == "percentage":
            fmt_code = "0.0%"
            fmt_name_fa = "درصد (%)"
        elif fmt_type == "date":
            fmt_code = "yyyy/mm/dd"
            fmt_name_fa = "تاریخ (yyyy/mm/dd)"
        elif fmt_type == "text":
            fmt_code = "@"
            fmt_name_fa = "متنی (Text)"
        else:
            fmt_code = "#,##0"
            fmt_name_fa = fmt_type

        sheet.set_number_format(target_range, fmt_code)
        if re.match(r"^[A-Z]+$", col_letter):
            sheet.auto_fit_columns([col_letter])

        msg = f"🎨 فرمت سلول‌های محدوده `{target_range}` با موفقیت به **{fmt_name_fa}** تغییر یافت."
        return ActionResult(success=True, message_persian=msg, affected_range=target_range)

