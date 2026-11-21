"""
Deterministic Semantic Rule Parser for Persian Excel Commands.
Parses natural colloquial Persian instructions into structured Action models without external APIs.
"""
import re
from typing import Dict, List, Optional, Tuple, Union

from excel_assistant.nlp.intent_models import (
    ExcelAction,
    ColumnArithmeticAction,
    ColumnStatsAction,
    ApplyDiscountAction,
    HighlightConditionAction,
    HighlightDuplicatesAction,
    DeleteEmptyRowsAction,
    DeleteDuplicatesAction,
    DeleteRowsOrColsAction,
    SortDataAction,
    AutoFitColumnsAction,
    CreateTableAction,
    FormatCellsAction,
    AnalyzeSheetAction,
    UndoLastAction,
)
from excel_assistant.nlp.persian_normalizer import (
    normalize_persian_text, extract_column_identifier, extract_color_hex
)
from excel_assistant.core.utils import col_index_to_letter


def resolve_column(token: str, header_map: Optional[Dict[str, str]] = None) -> Optional[str]:
    """
    Resolves a column token either as a letter (e.g. 'A') or by matching a known header name.
    """
    clean = token.strip()
    # Check if letter
    col = extract_column_identifier(clean)
    if col:
        return col

    # Check headers
    if header_map:
        clean_norm = normalize_persian_text(clean).lower()
        for h_name, col_letter in header_map.items():
            h_norm = normalize_persian_text(h_name).lower()
            if clean_norm == h_norm or clean_norm in h_norm or h_norm in clean_norm:
                return col_letter

    return None


class PersianRuleParser:
    """Parses Persian natural text commands into typed ExcelAction instances."""

    @classmethod
    def parse(cls, raw_command: str, header_map: Optional[Dict[str, str]] = None) -> Optional[ExcelAction]:
        text = normalize_persian_text(raw_command)
        text_lower = text.lower()

        # 1. Undo command
        if any(w in text_lower for w in ["لغو", "اندو", "آندو", "برگرد عقب", "بازگرداندن"]):
            return UndoLastAction(description_persian="لغو آخرین عملیات انجام‌شده")

        # 2. Analyze sheet command
        if any(w in text_lower for w in ["تحلیل", "بررسی ساختار", "بررسی فایل", "انالیز"]):
            return AnalyzeSheetAction(description_persian="تحلیل هوشمند ساختار شیت و داده‌ها")

        # 3. Delete empty rows
        if ("ردیف" in text_lower or "سطر" in text_lower) and ("خالی" in text_lower) and any(w in text_lower for w in ["حذف", "پاک", "دلیت"]):
            return DeleteEmptyRowsAction(
                description_persian="حذف ردیف‌های خالی از شیت",
                is_destructive=True
            )

        # 4. Highlight duplicates / Delete duplicates
        if "تکراری" in text_lower:
            # Check target column if specified
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None

            if any(w in text_lower for w in ["حذف", "پاک", "دلیت"]):
                return DeleteDuplicatesAction(
                    target_col=target_col,
                    description_persian=f"حذف ردیف‌های تکراری{' در ستون ' + target_col if target_col else ''}",
                    is_destructive=True
                )
            else:
                color = extract_color_hex(text, default="#FFFF00")
                return HighlightDuplicatesAction(
                    target_col=target_col,
                    color_hex=color,
                    text_only=True,
                    distinct_colors=True,
                    description_persian="هایلایت داده‌های متنی تکراری با رنگ‌های متمایز (بدون رنگ کردن اعداد)"
                )

        # 5. Autofit columns
        if any(w in text_lower for w in ["عرض ستون", "تنظیم ستون", "اتوفیت", "fit"]):
            return AutoFitColumnsAction(description_persian="تنظیم خودکار عرض ستون‌ها بر اساس محتوا")

        # 6. Create table
        if "جدول" in text_lower and any(w in text_lower for w in ["بساز", "ایجاد", "تبدیل", "درست کن"]):
            return CreateTableAction(description_persian="تبدیل محدوده داده‌ها به جدول اکسل (Excel Table)")

        # 7. Sorting
        if any(w in text_lower for w in ["مرتب", "سورت", "sort"]):
            # Check direction
            ascending = True
            if any(w in text_lower for w in ["بیشترین به کمترین", "نزولی", "زیاد به کم", "بزرگ به کوچک"]):
                ascending = False
            elif any(w in text_lower for w in ["کمترین به بیشترین", "صعودی", "کم به زیاد", "کوچک به بزرگ"]):
                ascending = True

            # Find key column
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            key_col = resolve_column(col_match.group(1), header_map) if col_match else "A"
            dir_str = "صعودی (کم به زیاد)" if ascending else "نزولی (بیش از کم)"
            return SortDataAction(
                key_col=key_col,
                ascending=ascending,
                description_persian=f"مرتب‌سازی داده‌ها بر اساس ستون {key_col} به صورت {dir_str}"
            )

        # 7.5 Apply Discount: تخفیف کالاها
        # e.g.: تخفیف ستون B با ستون C در ستون D
        # e.g.: 20 درصد تخفیف روی ستون B در ستون D
        # e.g.: ستون B با تخفیف ستون C در ستون D
        if "تخفیف" in text:
            pct_m = re.search(r'(\d+(?:\.\d+)?)\s*(?:درصد|%)', text)
            fixed_pct = float(pct_m.group(1)) if pct_m else None

            target_c = None
            target_match = re.search(
                r'(?:در|توی|داخل|به)\s+(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*(?:قرار\s*بده|قرار\s*ده|بنویس|درج\s*کن|بریز|بگذار|ثبت\s*کن|قرار\s*بگیرد|بذار)?\s*[.]*$',
                text
            )
            if target_match:
                target_c = resolve_column(target_match.group(1), header_map)

            found_cols = []
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text):
                res = resolve_column(c_str, header_map)
                if res and res not in found_cols and res != target_c:
                    found_cols.append(res)
            for c_let in re.findall(r'\b([A-Za-z]{1,2})\b', text):
                c_up = c_let.upper()
                if c_up not in found_cols and c_up != target_c:
                    found_cols.append(c_up)

            price_c = found_cols[0] if len(found_cols) >= 1 else "B"
            disc_c = found_cols[1] if len(found_cols) >= 2 else None

            d_str = f"ستون {disc_c}" if disc_c else (f"{fixed_pct}%" if fixed_pct else "تخفیف")
            return ApplyDiscountAction(
                price_col=price_c,
                discount_col=disc_c,
                discount_percent=fixed_pct,
                target_col=target_c,
                description_persian=f"محاسبه تخفیف کالاها: ستون قیمت {price_c} با درصد {d_str}{' و درج در ستون ' + target_c if target_c else ''}"
            )

        # 8. Column Arithmetic: ضرب، جمع، تفریق، تقسیم
        # e.g.: ستون B را در C ضرب کن و در D قرار بده
        # e.g.: ستون A را در ستون B ضرب کن و جواب را در C بنویس
        # e.g.: ستون قیمت را در ستون تعداد ضرب کن
        arith_ops = {
            "ضرب": "multiply",
            "ضربدر": "multiply",
            "*": "multiply",
            "جمع": "add",
            "به علاوه": "add",
            "+": "add",
            "منهای": "subtract",
            "منها": "subtract",
            "تفریق": "subtract",
            "-": "subtract",
            "تقسیم": "divide",
            "/": "divide",
        }

        detected_op = None
        for op_word, op_type in arith_ops.items():
            if op_word in text:
                detected_op = op_type
                break

        if detected_op:
            text_clean = text.strip()

            # Step 1: Detect target column if present
            target_c = None
            target_match = re.search(
                r'(?:در|توی|داخل|به)\s+(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*(?:قرار\s*بده|قرار\s*ده|قرار\s*بگذار|بنویس|درج\s*کن|بریز|بگذار|ثبت\s*کن|قرار\s*بگیرد|بذار)?\s*[.]*$',
                text_clean
            )

            target_end_pos = len(text_clean)
            if target_match:
                cand_target = resolve_column(target_match.group(1), header_map)
                if cand_target:
                    target_c = cand_target
                    target_end_pos = target_match.start()

            prefix_text = text_clean[:target_end_pos]

            # Find columns mentioned in prefix (with 'ستون' or after prepositions or as isolated letters/headers)
            prefix_cols = []

            # 1. Multi-column pattern e.g. 'ستون های 2 و 4 و 5' or 'ستون B و D و E'
            multi_m = re.search(
                r'ستون(?:\s*های)?\s+([A-Za-z0-9\s،,و\u0600-\u06FF]+?)(?:\s+(?:را|رو|با|در|جمع|ضرب|تفریق|منهای|تقسیم|\*|\+|-|\/|$))',
                prefix_text
            )
            if multi_m:
                tokens = re.split(r'[\s،,و]+', multi_m.group(1).strip())
                for t in tokens:
                    res = resolve_column(t, header_map)
                    if res and res not in prefix_cols:
                        prefix_cols.append(res)

            # 2. Check explicit 'ستون X'
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', prefix_text):
                res = resolve_column(c_str, header_map)
                if res and res not in prefix_cols:
                    prefix_cols.append(res)

            # 3. Check after prepositions (e.g. 'در C', 'در سی', 'با B', 'با بی')
            for c_str in re.findall(r'(?:در|با|ضربدر|منهای|تقسیم بر|به علاوه|و)\s+(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)', prefix_text):
                res = resolve_column(c_str, header_map)
                if res and res not in prefix_cols:
                    prefix_cols.append(res)

            # 4. Check isolated Latin letters in prefix (e.g. 'در C ضرب کن')
            for c_letter in re.findall(r'\b([A-Za-z]{1,2})\b', prefix_text):
                c_up = c_letter.upper()
                if c_up not in prefix_cols:
                    prefix_cols.append(c_up)

            # 5. Check headers in prefix
            if header_map:
                for h_name, col_let in header_map.items():
                    if h_name in prefix_text and col_let not in prefix_cols:
                        prefix_cols.append(col_let)

            if len(prefix_cols) >= 2:
                c1, c2 = prefix_cols[0], prefix_cols[1]
                op_fa = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم"}[detected_op]
                return ColumnArithmeticAction(
                    operation=detected_op,
                    col1=c1,
                    col2=c2,
                    columns=prefix_cols,
                    target_col=target_c,
                    as_formula=True,
                    description_persian=f"{op_fa} ستون‌های {'، '.join(prefix_cols)}{' و درج نتیجه در ستون ' + target_c if target_c else ''}"
                )

            # Fallback check across whole text
            all_cols = []
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text_clean):
                res = resolve_column(c_str, header_map)
                if res and res not in all_cols:
                    all_cols.append(res)
            for c_let in re.findall(r'\b([A-Za-z]{1,2})\b', text_clean):
                c_up = c_let.upper()
                if c_up not in all_cols:
                    all_cols.append(c_up)

            if len(all_cols) >= 2:
                c1, c2 = all_cols[0], all_cols[1]
                if not target_c and len(all_cols) >= 3:
                    target_c = all_cols[-1]
                    source_cols = all_cols[:-1]
                else:
                    source_cols = all_cols
                op_fa = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم"}[detected_op]
                return ColumnArithmeticAction(
                    operation=detected_op,
                    col1=c1,
                    col2=c2,
                    columns=source_cols,
                    target_col=target_c,
                    as_formula=True,
                    description_persian=f"{op_fa} ستون‌های {'، '.join(source_cols)}{' و درج نتیجه در ستون ' + target_c if target_c else ''}"
                )

        # 9. Statistical calculations: میانگین، مجموع، بیشترین، کمترین، میانه، تعداد
        # e.g.: میانگین ستون D را حساب کن
        # e.g.: مجموع ستون B
        stats_map = {
            "میانگین": "average",
            "معدل": "average",
            "مجموع": "sum",
            "جمع": "sum",
            "کل": "sum",
            "بیشترین": "max",
            "حداکثر": "max",
            "ماکزیمم": "max",
            "کمترین": "min",
            "حداقل": "min",
            "مینیمم": "min",
            "میانه": "median",
            "تعداد": "count",
            "شمارش": "count",
            "درصد": "percentage",
        }

        for stat_word, stat_code in stats_map.items():
            if stat_word in text:
                # Find source column
                col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
                source_col = resolve_column(col_match.group(1), header_map) if col_match else None
                if not source_col:
                    # Try single letter token after word e.g. "میانگین D"
                    letter_match = re.search(r'\b([A-Za-z]{1,2})\b', text)
                    if letter_match:
                        source_col = letter_match.group(1).upper()

                if source_col:
                    stat_fa = {"average": "میانگین", "sum": "مجموع", "max": "بیشترین", "min": "کمترین",
                               "median": "میانه", "count": "تعداد", "percentage": "درصد"}[stat_code]
                    return ColumnStatsAction(
                        stat_type=stat_code,
                        source_col=source_col,
                        description_persian=f"محاسبه {stat_fa} برای ستون {source_col}"
                    )

        # 10. Conditional Highlighting:
        # e.g.: سلول‌های بیشتر از 100 را سبز کن
        # e.g.: ستون B کمتر از 50 را قرمز کن
        cond_matches = [
            ("بیشتر از", "gt"),
            ("بالای", "gt"),
            ("بزرگتر از", "gt"),
            ("کمتر از", "lt"),
            ("پایین تر از", "lt"),
            ("کوچکتر از", "lt"),
            ("مساوی", "eq"),
            ("برابر", "eq"),
        ]

        for cond_phrase, cond_type in cond_matches:
            if cond_phrase in text:
                # Extract value
                val_match = re.search(rf'{cond_phrase}\s*([-+]?\d*\.?\d+)', text)
                if val_match:
                    threshold_val = float(val_match.group(1))
                    color = extract_color_hex(text, default="#90EE90")
                    col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
                    target_col = resolve_column(col_match.group(1), header_map) if col_match else None

                    cond_fa = {"gt": "بیشتر از", "lt": "کمتر از", "eq": "مساوی"}[cond_type]
                    return HighlightConditionAction(
                        target_col=target_col,
                        condition=cond_type,
                        value=threshold_val,
                        color_hex=color,
                        description_persian=f"هایلایت سلول‌های {cond_fa} {threshold_val} با رنگ {color}"
                    )

        # 11. Format cells (Font bold, size, color)
        # e.g.: سلول‌های A1:D1 را بولد کن
        # e.g.: ردیف اول را زرد کن
        if any(w in text_lower for w in ["بولد", "پررنگ", "bold", "ایتالیک", "اندازه فونت", "فونت"]):
            range_match = re.search(r'\b([A-Za-z]+\d*(?::[A-Za-z]+\d*)?)\b', text)
            target_rng = range_match.group(1).upper() if range_match else "A1"
            bold = True if any(w in text_lower for w in ["بولد", "پررنگ", "bold"]) else None
            return FormatCellsAction(
                target_range=target_rng,
                bold=bold,
                description_persian=f"قالب‌بندی فونت محدوده {target_rng}"
            )

        return None
