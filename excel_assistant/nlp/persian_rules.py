"""
Deterministic Semantic Rule Parser for Persian Excel Commands.
Parses natural colloquial Persian instructions into structured Action models without external APIs.
Employs robust slot extraction, colloquial normalizations, and extensive synonym matching.
"""
import re
from typing import Dict, List, Optional, Tuple, Union

from excel_assistant.nlp.intent_models import (
    ExcelAction,
    ColumnArithmeticAction,
    ColumnStatsAction,
    ApplyDiscountAction,
    InsertDateAction,
    PreventDuplicatesAction,
    RemoveUniqueValidationAction,
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
    SalesSummaryAction,
    ProfitCalculationAction,
    FilterDataAction,
    ConvertTextToNumberAction,
    SplitTextAction,
    MergeCellsAction,
    SetCellFormatAction,
    PriceAdjustmentAction,
    ReconciliationAction,
    InvoiceAction,
    InventoryAlertAction,
    InventoryValuationAction,
    TopSellingAction,
    CustomerLedgerAction,
    TrialBalanceAction,
    AccountBalanceAction,
    PayrollAction,
    AttendanceAction,
    DepreciationAction,
    FinancialStatementAction,
    DeepFileAuditAction,
    AutoChartAction,
    BackupAction,
    SaveWorkbookAction,
    PdfExportAction,
    UndoLastAction,
)
from excel_assistant.nlp.persian_normalizer import (
    normalize_persian_text,
    extract_column_identifier,
    extract_color_hex,
    extract_target_column,
    extract_source_columns,
    extract_percentage,
    extract_row_range,
    resolve_column,
)
from excel_assistant.nlp.synonyms import (
    ARITHMETIC_SYNONYMS,
    STATS_SYNONYMS,
    EMPTY_ROW_ROW_WORDS,
    EMPTY_ROW_EMPTY_WORDS,
    DELETE_VERBS,
    DUPLICATE_WORDS,
    UNLOCK_UNIQUE_WORDS,
    PREVENT_WORDS,
    SORT_WORDS,
    SORT_DESCENDING_WORDS,
    SORT_ASCENDING_WORDS,
    FILTER_WORDS,
    CLEAR_FILTER_WORDS,
    AUTOFIT_WORDS,
    TABLE_WORDS,
    CURRENCY_TOMAN_WORDS,
    PERCENTAGE_WORDS,
    TEXT_TO_NUMBER_WORDS,
    SPLIT_WORDS,
    MERGE_WORDS,
    RECONCILIATION_WORDS,
    UNDO_WORDS,
    SAVE_WORDS,
    BACKUP_WORDS,
    PDF_WORDS,
)
from excel_assistant.core.utils import col_index_to_letter


class PersianRuleParser:
    """
    Parses Persian commands into executable ExcelAction objects deterministically.
    Handles colloquial language, variable word orders, and entity/slot extraction.
    """

    @classmethod
    def parse(cls, raw_command: str, header_map: Optional[Dict[str, str]] = None) -> Optional[ExcelAction]:
        text = normalize_persian_text(raw_command)
        text_lower = text.lower()

        # 1. Undo command
        if any(w in text_lower for w in UNDO_WORDS):
            return UndoLastAction(description_persian="لغو آخرین عملیات انجام‌شده")

        # 1.5 Deep File Audit / سلامت و خطاهای فایل اکسل
        if any(w in text_lower for w in [
            "سلامت فایل", "خطاهای فایل", "بررسی خطاهای فایل", "چکاپ فایل",
            "سلول های خالی", "فرمول های اشتباه", "تحلیل هوشمند فایل",
            "خطاهای محاسباتی", "داده های نامعتبر"
        ]):
            return DeepFileAuditAction(description_persian="تحلیل هوشمند و عمیق سلامت فایل اکسل")

        # 1.6 Backup / نسخه پشتیبان
        if any(w in text_lower for w in BACKUP_WORDS):
            return BackupAction(description_persian="ذخیره نسخه پشتیبان از فایل اکسل")

        # 1.65 Save Workbook / ذخیره فایل اکسل
        if any(w in text_lower for w in SAVE_WORDS) or text_lower.strip() in ["ذخیره", "سیو", "save"]:
            return SaveWorkbookAction(description_persian="ذخیره تغییرات فایل اکسل")

        # 1.7 PDF Export / خروجی PDF
        if any(w in text_lower for w in PDF_WORDS) or (
            any(w in text_lower for w in ["pdf", "پی دی اف", "پرینت"]) and
            any(w in text_lower for w in ["خروجی", "دانلود", "تبدیل", "صادر", "گرفتن", "بگیر"])
        ):
            return PdfExportAction(description_persian="تولید خروجی PDF استاندارد از شیت فعال")

        # 1.8 Reconciliation / مغایرت‌گیری دو ستون یا دو فهرست
        if any(w in text_lower for w in RECONCILIATION_WORDS):
            cols_found = extract_source_columns(text, header_map)
            c1 = cols_found[0] if len(cols_found) >= 1 else "A"
            c2 = cols_found[1] if len(cols_found) >= 2 else "B"
            return ReconciliationAction(
                col1=c1,
                col2=c2,
                description_persian=f"مغایرت‌گیری بین ستون {c1} و ستون {c2}"
            )

        # 1.9 Invoice generation / فاکتور فروش و خرید
        if "فاکتور" in text_lower:
            inv_type = "purchase" if "خرید" in text_lower else "sales"
            inv_title = "فاکتور خرید" if inv_type == "purchase" else "فاکتور فروش"
            return InvoiceAction(
                invoice_type=inv_type,
                description_persian=f"ثبت و صدور {inv_title} رسمی با محاسبات ارزش افزوده و تخفیف"
            )

        # 1.10 Attendance & Timesheet / ورود و خروج، کارکرد، اضافه‌کار و کسری کار پرسنل
        if any(w in text_lower for w in [
            "ورود و خروج", "حضور و غیاب", "حضور وغیاب", "ورود خروج",
            "ساعت کارکرد", "کارکرد ماهانه", "تایم شیت", "تایمشیت", "timesheet",
            "کسری کار", "تاخیر و اضافه کار", "تأخیر و اضافه کار", "ثبت ماهانه",
            "خلاصه کارمندان"
        ]) or (
            ("کارکرد" in text_lower or "ساعت کار" in text_lower) and any(w in text_lower for w in ["کارمند", "پرسنل", "ورود", "خروج", "ماهانه", "روزانه"])
        ):
            emp_name = None
            for cand_name in ["علی رضایی", "مریم احمدی", "حسین کریمی", "سارا محمدی", "رضا نوری", "نرگس حیدری", "امیر صادقی", "زهرا موسوی"]:
                if cand_name in text:
                    emp_name = cand_name
                    break
            return AttendanceAction(
                employee_name=emp_name,
                description_persian="محاسبه دقیق ورود و خروج، کارکرد، اضافه‌کار، کسری کار و حضور و غیاب پرسنل"
            )

        # 1.11 Payroll, Overtime, Insurance & Salary Tax / حقوق و دستمزد
        if any(w in text_lower for w in [
            "حقوق و دستمزد", "حقوق", "دستمزد", "اضافه کاری", "اضافه‌کاری",
            "بیمه کارگر", "بیمه تامین اجتماعی", "مالیات حقوق", "فیش حقوق"
        ]):
            return PayrollAction(description_persian="محاسبه لیست حقوق، اضافه‌کاری، بیمه و مالیات")

        # 1.109 Inventory Valuation (FIFO, LIFO, Weighted Average) / ارزش‌گذاری انبار به روش‌های فایفو، لایفو و میانگین
        if any(w in text_lower for w in [
            "ارزش‌گذاری انبار", "ارزش گذاری انبار", "ارزش‌گذاری موجودی", "ارزش گذاری موجودی",
            "کاردکس", "کاردکس کالا"
        ]) or (
            any(w in text_lower for w in ["فایفو", "fifo", "لایفو", "lifo", "میانگین موزون"])
            and not any(w in text_lower for w in ["سود", "زیان", "سودآوری", "حاشیه سود"])
        ):
            v_method = "fifo"
            if any(w in text_lower for w in ["مقایسه", "هر سه", "هر ۳", "کاردکس"]) or ("فایفو" in text_lower and "لایفو" in text_lower):
                v_method = "compare"
            elif any(w in text_lower for w in ["لایفو", "lifo", "آخرین وارده"]):
                v_method = "lifo"
            elif any(w in text_lower for w in ["میانگین موزون", "میانگین متحرک", "میانگین", "average", "موزون"]):
                v_method = "average"
            elif any(w in text_lower for w in ["فایفو", "fifo", "اولین وارده"]):
                v_method = "fifo"
            elif "ارزش‌گذاری" in text_lower or "ارزش گذاری" in text_lower:
                v_method = "compare"

            thresh = 5.0
            th_m = re.search(r'(?:زیر|کمتر از|آستانه|حداقل)\s*([0-9]+)', text)
            if th_m:
                try:
                    thresh = float(th_m.group(1))
                except ValueError:
                    thresh = 5.0

            sales_s = None
            purchase_s = None
            sheets_found = re.findall(r'(?:شیت|شیلد|sheet|صفحه)\s*([0-9]+|[A-Za-z0-9_\u0600-\u06FF]+)', text, flags=re.IGNORECASE)
            if len(sheets_found) >= 2:
                if "فروش" in text and "خرید" in text:
                    f_pos = text.find("فروش")
                    k_pos = text.find("خرید")
                    if f_pos < k_pos:
                        sales_s = sheets_found[0]
                        purchase_s = sheets_found[1]
                    else:
                        purchase_s = sheets_found[0]
                        sales_s = sheets_found[1]
                else:
                    sales_s = sheets_found[0]
                    purchase_s = sheets_found[1]
            elif len(sheets_found) == 1:
                if "فروش" in text:
                    sales_s = sheets_found[0]
                elif "خرید" in text or "انبار" in text:
                    purchase_s = sheets_found[0]

            method_fa = {"fifo": "فایفو (FIFO)", "lifo": "لایفو (LIFO)", "average": "میانگین موزون", "compare": "فایفو، لایفو و میانگین موزون"}.get(v_method, "فایفو")
            return InventoryValuationAction(
                method=v_method,
                min_threshold=thresh,
                sales_sheet=sales_s,
                purchase_sheet=purchase_s,
                description_persian=f"ارزش‌گذاری موجودی انبار به روش {method_fa}"
            )

        # 1.11 Inventory Alert & Valuation / موجودی انبار و هشدار کمبود موجودی
        if any(w in text_lower for w in [
            "کمبود موجودی", "هشدار موجودی", "کسری موجودی", "موجودی انبار",
            "کنترل انبار", "کنترل موجودی", "ارزش موجودی", "ارزش کل موجودی", "انبارگردانی",
            "مانده موجودی", "موجودی باقی مانده", "موجودی باقیمانده", "مانده انبار", "موجودی کالا", "موجودی"
        ]):
            thresh = 5.0
            th_m = re.search(r'(?:زیر|کمتر از|آستانه|حداقل)\s*([0-9]+)', text)
            if th_m:
                try:
                    thresh = float(th_m.group(1))
                except ValueError:
                    thresh = 5.0

            v_method = "fifo"
            if any(w in text_lower for w in ["لایفو", "lifo"]):
                v_method = "lifo"
            elif any(w in text_lower for w in ["میانگین موزون", "میانگین", "average"]):
                v_method = "average"
            elif any(w in text_lower for w in ["فایفو", "fifo"]):
                v_method = "fifo"

            sales_s = None
            purchase_s = None
            sheets_found = re.findall(r'(?:شیت|شیلد|sheet|صفحه)\s*([0-9]+|[A-Za-z0-9_\u0600-\u06FF]+)', text, flags=re.IGNORECASE)
            if len(sheets_found) >= 2:
                if "فروش" in text and "خرید" in text:
                    f_pos = text.find("فروش")
                    k_pos = text.find("خرید")
                    if f_pos < k_pos:
                        sales_s = sheets_found[0]
                        purchase_s = sheets_found[1]
                    else:
                        purchase_s = sheets_found[0]
                        sales_s = sheets_found[1]
                else:
                    sales_s = sheets_found[0]
                    purchase_s = sheets_found[1]
            elif len(sheets_found) == 1:
                if "فروش" in text:
                    sales_s = sheets_found[0]
                elif "خرید" in text or "انبار" in text:
                    purchase_s = sheets_found[0]

            return InventoryAlertAction(
                min_threshold=thresh,
                sales_sheet=sales_s,
                purchase_sheet=purchase_s,
                valuation_method=v_method,
                description_persian="مدیریت موجودی انبار و هشدار کمبود موجودی"
            )

        # 1.12 Top Selling Products / پرفروش‌ترین کالاها
        if any(w in text_lower for w in [
            "پرفروش ترین", "پرفروش‌ترین", "کالاهای پرفروش", "بیشترین فروش کالا", "پرفروش"
        ]):
            return TopSellingAction(description_persian="گزارش پرفروش‌ترین کالاهای ثبت‌شده")

        # 1.13 Customer Ledger / بدهی و طلب مشتریان
        if any(w in text_lower for w in [
            "بدهی مشتری", "طلب مشتری", "بدهی و طلب", "مانده مشتری", "طرف حساب"
        ]):
            return CustomerLedgerAction(description_persian="محاسبه مانده بدهی و طلب مشتریان")

        # 1.14 Trial Balance / تراز آزمایشی
        if any(w in text_lower for w in [
            "تراز آزمایشی", "تراز ازمایشی", "تراز ۴ ستونه", "تراز چهار ستونه", "تراز ۶ ستونه", "تراز شش ستونه"
        ]):
            return TrialBalanceAction(description_persian="تهیه تراز آزمایشی چهار ستونه حسابداری")

        # 1.15 Account Balance / محاسبه مانده حساب (بدهکار و بستانکار)
        if any(w in text_lower for w in [
            "مانده حساب", "مانده بدهکار و بستانکار", "محاسبه مانده بدهکار", "مانده گیری", "ستون مانده"
        ]):
            return AccountBalanceAction(description_persian="محاسبه مانده حساب جاری بدهکار و بستانکار")

        # 1.16 Depreciation / محاسبه استهلاک دارایی‌ها
        if "استهلاک" in text_lower:
            method = "declining_balance" if any(w in text_lower for w in ["نزولی", "مانده نزولی"]) else "straight_line"
            return DepreciationAction(
                method=method,
                description_persian=f"محاسبه جدول استهلاک دارایی ({'نزولی' if method == 'declining_balance' else 'خط مستقیم'})"
            )

        # 1.17 Financial Statements / صورت سود و زیان و ترازنامه
        if "صورت سود و زیان" in text_lower or "سود و زیان جامع" in text_lower:
            return FinancialStatementAction(
                statement_type="income_statement",
                description_persian="تهیه صورت سود و زیان جامع استاندارد"
            )
        if "ترازنامه" in text_lower or "بیلان مالی" in text_lower:
            return FinancialStatementAction(
                statement_type="balance_sheet",
                description_persian="تهیه ترازنامه مالی و وضعیت دارایی‌ها و بدهی‌ها"
            )

        # 1.18 Auto Chart / رسم نمودار گرافیکی
        if any(w in text_lower for w in ["نمودار", "چارت"]) and any(w in text_lower for w in [
            "رسم", "بکش", "ایجاد", "نمایش", "تولید", "ساز"
        ]):
            return AutoChartAction(description_persian="تولید نمودار گرافیکی خودکار اکسل")

        # 1.19 Percentage Price Adjustment / افزایش یا کاهش درصدی قیمت‌ها
        # e.g.: قیمت‌های ستون C را ۱۵ درصد افزایش بده
        # e.g.: ستون B را ۱۰ درصد کاهش بده
        # e.g.: ستون B رو 10% گرون کن
        # e.g.: ستون C رو ۱۵ درصد ارزون کن
        if ("درصد" in text or "%" in text or "٪" in text) and any(w in text_lower for w in [
            "افزایش", "کاهش", "گران", "گرون", "ارزان", "ارزون", "بالا ببر", "کم کن", "اضافه کن"
        ]):
            pct_val = extract_percentage(text) or 10.0
            is_decrease = any(w in text_lower for w in ["کاهش", "ارزان", "ارزون", "کم کن", "پایین بیار"])
            op = "decrease" if is_decrease else "increase"

            col_m = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_m.group(1), header_map) if col_m else None
            if not target_col:
                cols = extract_source_columns(text, header_map)
                target_col = cols[0] if cols else "C"

            return PriceAdjustmentAction(
                target_col=target_col,
                percentage=pct_val,
                operation=op,
                description_persian=f"{'افزایش' if op == 'increase' else 'کاهش'} {pct_val:.0f} درصدی ستون {target_col}"
            )

        # 1.20 Row-by-Row Multi-Column Average / میانگین سطری بین ستون‌ها
        # e.g.: از ستون B و C میانگین بگیر و نتیجه را در D بنویس
        # e.g.: میانگین B و C رو بریز توی D
        if any(w in text_lower for w in ["میانگین", "معدل"]) and (" و " in text or "در " in text or "توی " in text or "تو " in text):
            target_c, rem_text = extract_target_column(text, header_map)
            source_cols = extract_source_columns(rem_text, header_map)
            if len(source_cols) >= 2 and target_c:
                return ColumnArithmeticAction(
                    operation="average",
                    col1=source_cols[0],
                    col2=source_cols[1],
                    columns=source_cols,
                    target_col=target_c,
                    as_formula=True,
                    description_persian=f"میانگین سطری ستون‌های {' و '.join(source_cols)} و درج در ستون {target_c}"
                )

        # 2. Analyze sheet command
        if any(w in text_lower for w in ["تحلیل", "بررسی ساختار", "بررسی فایل", "انالیز"]):
            return AnalyzeSheetAction(description_persian="تحلیل هوشمند ساختار شیت و داده‌ها")

        # 3. Delete empty rows
        has_empty_word = any(w in text_lower for w in EMPTY_ROW_EMPTY_WORDS)
        has_row_word = any(w in text_lower for w in EMPTY_ROW_ROW_WORDS)
        has_delete_verb = any(w in text_lower for w in DELETE_VERBS)
        if (has_empty_word and (has_row_word or "خالی" in text_lower) and has_delete_verb) or (has_row_word and "خالی" in text_lower and has_delete_verb):
            return DeleteEmptyRowsAction(
                description_persian="حذف ردیف‌های خالی از شیت",
                is_destructive=True
            )

        # 4. Prevent duplicates / Remove unique lock / Delete duplicates / Highlight duplicates
        is_unlock_unique = (
            any(w in text_lower for w in UNLOCK_UNIQUE_WORDS) and
            any(w in text_lower for w in ["یکتا", "یکتایی", "یکتاسازی", "قفل", "محدودیت", "ستون"]) and
            not any(w in text_lower for w in ["ردیف", "سطر", "سطرهای"])
        )
        if is_unlock_unique:
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not target_col:
                cols = extract_source_columns(text, header_map)
                target_col = cols[0] if cols else "A"
            return RemoveUniqueValidationAction(
                target_col=target_col,
                description_persian=f"برداشتن قفل یکتایی ستون {target_col}"
            )

        if any(w in text_lower for w in DUPLICATE_WORDS) or any(w in text_lower for w in ["یکتا کن", "یکتاسازی"]):
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not target_col:
                cols = extract_source_columns(text, header_map)
                target_col = cols[0] if cols else None

            if any(w in text_lower for w in PREVENT_WORDS):
                return PreventDuplicatesAction(
                    target_col=target_col or "A",
                    description_persian=f"جلوگیری از ورود دادهٔ تکراری و قفل یکتاسازی ستون {target_col or 'A'}"
                )

            if any(w in text_lower for w in DELETE_VERBS):
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
        if any(w in text_lower for w in AUTOFIT_WORDS) or ("عرض" in text_lower and any(w in text_lower for w in ["تنظیم", "اندازه", "خودکار", "اتوماتیک"])):
            return AutoFitColumnsAction(description_persian="تنظیم خودکار عرض ستون‌ها بر اساس محتوا")

        # 6. Create table (checked only if NOT a filter command)
        if "فیلتر" not in text_lower and (
            any(w in text_lower for w in TABLE_WORDS) or
            ("جدول" in text_lower and any(w in text_lower for w in ["بساز", "ایجاد", "تبدیل", "درست کن", "درست"]))
        ):
            return CreateTableAction(description_persian="تبدیل محدوده داده‌ها به جدول اکسل (Excel Table)")

        # 7. Sorting
        if any(w in text_lower for w in SORT_WORDS):
            ascending = True
            if any(w in text_lower for w in SORT_DESCENDING_WORDS):
                ascending = False
            elif any(w in text_lower for w in SORT_ASCENDING_WORDS):
                ascending = True

            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            key_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not key_col:
                cols = extract_source_columns(text, header_map)
                key_col = cols[0] if cols else "A"

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
        if "تخفیف" in text_lower:
            fixed_pct = extract_percentage(text)
            target_c, rem_text = extract_target_column(text, header_map)
            source_cols = extract_source_columns(rem_text, header_map)

            price_c = source_cols[0] if len(source_cols) >= 1 else "B"
            disc_c = source_cols[1] if len(source_cols) >= 2 else None
            if fixed_pct is not None:
                disc_c = None

            d_str = f"ستون {disc_c}" if disc_c else (f"{fixed_pct}%" if fixed_pct else "تخفیف")
            return ApplyDiscountAction(
                price_col=price_c,
                discount_col=disc_c,
                discount_percent=fixed_pct,
                target_col=target_c,
                description_persian=f"محاسبه تخفیف کالاها: ستون قیمت {price_c} با درصد {d_str}{' و درج در ستون ' + target_c if target_c else ''}"
            )

        # 7.55 Profit Calculation: محاسبه سود (روزانه، ماهانه، سالانه، کل / فروش منهای خرید)
        if "سود" in text or (("فروش" in text and "خرید" in text and any(w in text for w in ["منهای", "منها", "تفریق", "-", "اختلاف", "به دست"])) and not any(inv_kw in text for inv_kw in ["موجودی", "انبار"])):
            period = "all"
            if "روزانه" in text and not any(w in text for w in ["ماهانه", "سالانه", "کل"]):
                period = "daily"
            elif "ماهانه" in text and not any(w in text for w in ["روزانه", "سالانه", "کل"]):
                period = "monthly"
            elif "سالانه" in text and not any(w in text for w in ["روزانه", "ماهانه", "کل"]):
                period = "yearly"
            elif "کل" in text and not any(w in text for w in ["روزانه", "ماهانه", "سالانه"]):
                period = "total"

            sales_s = None
            purchase_s = None

            sheets_found = re.findall(r'(?:شیت|شیلد|sheet|صفحه)\s*([0-9]+|[A-Za-z0-9_]+)', text, flags=re.IGNORECASE)
            if len(sheets_found) >= 2:
                sales_s = sheets_found[0]
                purchase_s = sheets_found[1]
                if not any(w in text for w in ["روزانه", "ماهانه", "سالانه", "کل"]):
                    period = "by_item"
            elif len(sheets_found) == 1:
                sales_s = sheets_found[0]

            date_col = None
            profit_col = None
            sales_col = None
            purchase_col = None

            date_m = re.search(r'(?:ستون|بر اساس)?\s*تاریخ\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if date_m and date_m.group(1):
                date_col = resolve_column(date_m.group(1), header_map)

            profit_m = re.search(r'(?:ستون|مبلغ)?\s*سود\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if profit_m and profit_m.group(1):
                profit_col = resolve_column(profit_m.group(1), header_map)

            sales_m = re.search(r'(?:ستون|مبلغ)?\s*فروش\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if sales_m and sales_m.group(1):
                sales_col = resolve_column(sales_m.group(1), header_map)

            purchase_m = re.search(r'(?:ستون|مبلغ)?\s*(?:خرید|هزینه)\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if purchase_m and purchase_m.group(1):
                purchase_col = resolve_column(purchase_m.group(1), header_map)

            found_cols = extract_source_columns(text, header_map)

            if len(found_cols) >= 2:
                if not profit_col and not date_col:
                    profit_col = found_cols[0]
                    date_col = found_cols[1]
                elif profit_col and not date_col:
                    for fc in found_cols:
                        if fc != profit_col:
                            date_col = fc
                            break
                elif date_col and not profit_col:
                    for fc in found_cols:
                        if fc != date_col:
                            profit_col = fc
                            break
            elif len(found_cols) == 1:
                if not profit_col and not date_col:
                    profit_col = found_cols[0]

            val_method = "fifo"
            if any(w in text_lower for w in ["لایفو", "lifo"]):
                val_method = "lifo"
            elif any(w in text_lower for w in ["میانگین", "موزون", "میانگین موزون", "average"]):
                val_method = "average"
            elif any(w in text_lower for w in ["فایفو", "fifo"]):
                val_method = "fifo"

            method_fa = {
                "fifo": "فایفو (FIFO)",
                "lifo": "لایفو (LIFO)",
                "average": "میانگین موزون (Weighted Average)"
            }[val_method]

            period_fa = {
                "all": "روزانه، ماهانه، سالانه و کل",
                "daily": "روزانه",
                "monthly": "ماهانه",
                "yearly": "سالانه",
                "total": "کل",
                "by_item": "تفکیک کالاها (فروش منهای خرید)"
            }.get(period, "روزانه، ماهانه، سالانه و کل")

            return ProfitCalculationAction(
                sales_sheet=sales_s,
                purchase_sheet=purchase_s,
                date_col=date_col,
                profit_col=profit_col,
                sales_col=sales_col,
                purchase_col=purchase_col,
                valuation_method=val_method,
                period=period,
                create_report_sheet=True,
                description_persian=f"محاسبه سود {period_fa} به روش {method_fa}"
            )

        # 7.6 Sales Summary: محاسبه فروش روزانه، ماهانه، سالانه و کل
        if "فروش" in text and any(w in text for w in [
            "محاسبه", "حساب", "گزارش", "روزانه", "ماهانه", "سالانه", "کل", "چقدر", "چند", "آمار", "تحلیل", "جمع"
        ]):
            period = "all"
            if "روزانه" in text and not any(w in text for w in ["ماهانه", "سالانه", "کل"]):
                period = "daily"
            elif "ماهانه" in text and not any(w in text for w in ["روزانه", "سالانه", "کل"]):
                period = "monthly"
            elif "سالانه" in text and not any(w in text for w in ["روزانه", "ماهانه", "کل"]):
                period = "yearly"
            elif "کل" in text and not any(w in text for w in ["روزانه", "ماهانه", "سالانه"]):
                period = "total"

            date_col = None
            amount_col = None

            date_m = re.search(r'(?:ستون|بر اساس)?\s*تاریخ\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if date_m and date_m.group(1):
                date_col = resolve_column(date_m.group(1), header_map)

            amount_m = re.search(r'(?:ستون|مبلغ)?\s*فروش\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if amount_m and amount_m.group(1):
                amount_col = resolve_column(amount_m.group(1), header_map)

            found_cols = extract_source_columns(text, header_map)

            if len(found_cols) >= 2:
                if not amount_col and not date_col:
                    amount_col = found_cols[0]
                    date_col = found_cols[1]
                elif amount_col and not date_col:
                    for fc in found_cols:
                        if fc != amount_col:
                            date_col = fc
                            break
                elif date_col and not amount_col:
                    for fc in found_cols:
                        if fc != date_col:
                            amount_col = fc
                            break
            elif len(found_cols) == 1:
                if not amount_col and not date_col:
                    amount_col = found_cols[0]

            period_fa = {
                "all": "روزانه، ماهانه، سالانه و کل",
                "daily": "روزانه",
                "monthly": "ماهانه",
                "yearly": "سالانه",
                "total": "کل"
            }.get(period, "جامع")

            return SalesSummaryAction(
                date_col=date_col,
                amount_col=amount_col,
                period=period,
                create_report_sheet=True,
                description_persian=f"محاسبه فروش {period_fa}"
            )

        # 7.7 Insert Date: درج تاریخ در اکسل
        if "تاریخ" in text_lower and any(w in text_lower for w in ["بزن", "درج", "بنویس", "قرار", "بذار", "ثبت", "سطر", "ردیف", " تا "]):
            start_r, end_r = extract_row_range(text)
            start_r = start_r or 1
            end_r = end_r or 30

            date_m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', text)
            if date_m:
                y, m, d = int(date_m.group(1)), int(date_m.group(2)), int(date_m.group(3))
            else:
                y, m, d = 1403, 8, 29

            col_m = re.search(r'(?:در|توی|روی)?\s*ستون\s*([A-Za-z0-9\u0600-\u06FF]+)', text)
            col = resolve_column(col_m.group(1), header_map) if col_m else None
            if not col:
                cols = extract_source_columns(text, header_map)
                col = cols[0] if cols else 'A'

            mode = 'sequential' if any(w in text_lower for w in ['متوالی', 'روزشمار', 'پشت سر هم', 'افزایشی']) else 'fixed'
            return InsertDateAction(
                target_col=col,
                start_row=start_r,
                end_row=end_r,
                year=y,
                month=m,
                day=d,
                mode=mode,
                description_persian=f"درج تاریخ {y:04d}/{m:02d}/{d:02d} در ستون {col} از سطر {start_r} تا {end_r}"
            )

        # 8. Column Arithmetic: ضرب، جمع، تفریق، تقسیم
        # e.g.: ستون B را در C ضرب کن و در D قرار بده
        # e.g.: ستون A را در ستون B ضرب کن و جواب را در C بنویس
        # e.g.: ستون قیمت را در ستون تعداد ضرب کن
        # e.g.: ستون D رو ضربدر B کن و جواب رو E بذار
        # e.g.: حاصل‌ضرب B و D رو داخل E قرار بده
        # e.g.: B و D رو در هم ضرب کن و نتیجه رو در E بنویس
        # e.g.: توی E بنویس حاصلضرب D و B
        # e.g.: D رو تو B ضرب کن بریز تو E
        # e.g.: ستون C رو منهای B کن و توی A بنویس
        # e.g.: حاصل تفریق ستون A و B بره تو C
        # e.g.: مجموع D و B رو توی E بنویس
        detected_op = None
        for op_type, syn_list in ARITHMETIC_SYNONYMS.items():
            if any(syn in text_lower for syn in syn_list):
                detected_op = op_type
                break

        if detected_op:
            target_c, rem_text = extract_target_column(text, header_map)
            source_cols = extract_source_columns(rem_text, header_map)

            # If fewer than 2 source columns found in rem_text and no target_c, check entire text
            if len(source_cols) < 2 and not target_c:
                all_cols = extract_source_columns(text, header_map)
                if len(all_cols) >= 2:
                    source_cols = all_cols

            # Column arithmetic strictly requires at least 2 source columns
            if len(source_cols) >= 2:
                c1, c2 = source_cols[0], source_cols[1]
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

        # 9. Statistical calculations: جمع خودکار، میانگین، بیشترین، کمترین، میانه، تعداد سلول‌های پر
        # e.g.: میانگین ستون D را حساب کن
        # e.g.: مجموع ستون B را محاسبه کن
        # e.g.: بیشترین و کمترین ستون C
        # e.g.: ستون B رو میانگین بگیر
        detected_stat = None
        all_stat_syns = []
        for stat_code, syn_list in STATS_SYNONYMS.items():
            for syn in syn_list:
                all_stat_syns.append((syn, stat_code))
        # Prioritize longer phrases (e.g. 'سلول‌های پر' before 'شمارش')
        all_stat_syns.sort(key=lambda x: len(x[0]), reverse=True)

        for syn, stat_code in all_stat_syns:
            if syn in text_lower:
                if stat_code == "percentage" and "فرمت" in text_lower:
                    continue
                detected_stat = stat_code
                break

        if detected_stat:
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            source_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not source_col:
                source_cols = extract_source_columns(text, header_map)
                source_col = source_cols[0] if source_cols else None
            if not source_col:
                let_m = re.search(r'\b([A-Za-z]{1,2})\b', text)
                source_col = let_m.group(1).upper() if let_m else None

            if source_col:
                stat_fa = {
                    "average": "میانگین",
                    "sum": "جمع خودکار",
                    "max": "بیشترین مقدار (MAX)",
                    "min": "کمترین مقدار (MIN)",
                    "median": "میانه",
                    "count": "تعداد عددی",
                    "counta": "تعداد سلول‌های پر",
                    "percentage": "درصد",
                }.get(detected_stat, detected_stat)
                return ColumnStatsAction(
                    stat_type=detected_stat,
                    source_col=source_col,
                    description_persian=f"محاسبه {stat_fa} برای ستون {source_col}"
                )

        # 10. Merge Cells: ادغام سلول‌ها
        if any(w in text_lower for w in MERGE_WORDS):
            rng_m = re.search(r'([A-Za-z]+\d+)\s*(?::|تا|الی|to|-)\s*([A-Za-z]+\d+)', text)
            if rng_m:
                range_str = f"{rng_m.group(1).upper()}:{rng_m.group(2).upper()}"
            else:
                single_m = re.search(r'\b([A-Za-z]+\d+:[A-Za-z]+\d+)\b', text)
                range_str = single_m.group(1).upper() if single_m else "A1:B1"

            return MergeCellsAction(
                range_str=range_str,
                align_center=True,
                description_persian=f"ادغام سلول‌های محدوده {range_str} و تراز در مرکز"
            )

        # 11. Convert text to number: تبدیل متن به عدد و اصلاح داده‌های عددی
        if any(w in text_lower for w in TEXT_TO_NUMBER_WORDS) or ("متن به عدد" in text_lower or ("عددی" in text_lower and "متن" in text_lower)):
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not target_col:
                cols = extract_source_columns(text, header_map)
                target_col = cols[0] if cols else "A"
            return ConvertTextToNumberAction(
                target_col=target_col,
                description_persian=f"تبدیل متون عددی ستون {target_col} به عدد واقعی و اصلاح فرمت"
            )

        # 12. Split text / Text to columns: جدا کردن متن (تفکیک نام، نام خانوادگی و...)
        if any(w in text_lower for w in SPLIT_WORDS) or "تفکیک" in text_lower:
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            source_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not source_col:
                cols = extract_source_columns(text, header_map)
                source_col = cols[0] if cols else "A"

            delim = " "
            if any(w in text_lower for w in ["خط تیره", "دش", "-", "منها"]):
                delim = "-"
            elif any(w in text_lower for w in ["کاما", "ویرگول", ","]):
                delim = ","
            elif any(w in text_lower for w in ["اسلش", "/"]):
                delim = "/"

            h1 = "نام" if "نام" in text else "بخش ۱"
            h2 = "نام خانوادگی" if "خانوادگی" in text else "بخش ۲"

            return SplitTextAction(
                source_col=source_col,
                delimiter=delim,
                header1=h1,
                header2=h2,
                description_persian=f"تفکیک و جدا کردن متن ستون {source_col} به ستون‌های مجزا"
            )

        # 13. Smart Filter: فیلتر هوشمند اطلاعات بر اساس شرط یا اتوفیلتر
        if any(w in text_lower for w in CLEAR_FILTER_WORDS):
            return FilterDataAction(
                condition="clear",
                description_persian="پاک کردن تمام فیلترها و نمایش همه سطرها"
            )

        if any(w in text_lower for w in FILTER_WORDS):
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not target_col:
                cols = extract_source_columns(text, header_map)
                target_col = cols[0] if cols else None

            cond_type = None
            val = None
            if any(w in text_lower for w in ["بزرگتر", "بیشتر", "بالای", ">"]):
                cond_type = "gt"
                vm = re.search(r'(?:بزرگتر|بیشتر|بالای|>)\s*(?:از)?\s*([-+]?\d*\.?\d+)', text)
                val = float(vm.group(1)) if vm else None
            elif any(w in text_lower for w in ["کوچکتر", "کمتر", "پایین", "<"]):
                cond_type = "lt"
                vm = re.search(r'(?:کوچکتر|کمتر|پایین|<)\s*(?:از)?\s*([-+]?\d*\.?\d+)', text)
                val = float(vm.group(1)) if vm else None
            elif any(w in text_lower for w in ["مساوی", "برابر", "="]):
                cond_type = "eq"
                vm = re.search(r'(?:مساوی|برابر|=)\s*(?:با)?\s*([^\s]+)', text)
                val = vm.group(1) if vm else None
            elif any(w in text_lower for w in ["شامل", "حاوی"]):
                cond_type = "contains"
                vm = re.search(r'(?:شامل|حاوی)\s+([^\s]+)', text)
                val = vm.group(1) if vm else None

            if cond_type and target_col:
                return FilterDataAction(
                    target_col=target_col,
                    condition=cond_type,
                    value=val,
                    description_persian=f"فیلتر هوشمند ستون {target_col} بر اساس شرط"
                )
            else:
                return FilterDataAction(
                    condition="autofilter",
                    target_col=target_col,
                    description_persian="فعال‌سازی فیلتر هوشمند (AutoFilter) روی سرستون‌های جدول"
                )

        # 14. Set Cell Format: تغییر فرمت سلول (عدد، درصد، تاریخ، واحد پول)
        if "فرمت" in text_lower or any(w in text_lower for w in ["سه رقم", "هزارگان", "درصدی", "تومانی"]):
            fmt_type = "currency"
            if any(w in text_lower for w in CURRENCY_TOMAN_WORDS) or any(w in text_lower for w in ["تومان", "ریال", "واحد پول", "پول"]):
                fmt_type = "currency"
            elif any(w in text_lower for w in PERCENTAGE_WORDS) or any(w in text_lower for w in ["درصد", "%", "درصدی"]):
                fmt_type = "percentage"
            elif "تاریخ" in text_lower:
                fmt_type = "date"
            elif any(w in text_lower for w in ["سه رقم", "کاما", "هزارگان", "عدد", "عددی"]):
                fmt_type = "thousands"
            elif "متن" in text_lower:
                fmt_type = "text"

            range_match = re.search(r'\b([A-Za-z]+\d+:[A-Za-z]+\d+)\b', text)
            if range_match:
                target = range_match.group(1).upper()
            else:
                col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
                target = resolve_column(col_match.group(1), header_map) if col_match else None
                if not target:
                    cols = extract_source_columns(text, header_map)
                    target = cols[0] if cols else "A"

            fmt_fa_names = {
                "currency": "واحد پول (تومان)",
                "percentage": "درصد (%)",
                "date": "تاریخ (yyyy/mm/dd)",
                "thousands": "عدد با جداکننده سه‌رقم‌سه‌رقم",
                "text": "متن (Text)"
            }
            return SetCellFormatAction(
                target=target,
                format_type=fmt_type,
                description_persian=f"تغییر فرمت {target} به {fmt_fa_names.get(fmt_type, fmt_type)}"
            )

        # 15. Conditional Highlighting:
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
                val_match = re.search(rf'{cond_phrase}\s*([-+]?\d*\.?\d+)', text)
                if val_match:
                    threshold_val = float(val_match.group(1))
                    color = extract_color_hex(text, default="#90EE90")
                    col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
                    target_col = resolve_column(col_match.group(1), header_map) if col_match else None
                    if not target_col:
                        cols = extract_source_columns(text, header_map)
                        target_col = cols[0] if cols else None

                    cond_fa = {"gt": "بیشتر از", "lt": "کمتر از", "eq": "مساوی"}[cond_type]
                    return HighlightConditionAction(
                        target_col=target_col,
                        condition=cond_type,
                        value=threshold_val,
                        color_hex=color,
                        description_persian=f"هایلایت سلول‌های {cond_fa} {threshold_val} با رنگ {color}"
                    )

        # 16. Format cells (Font bold, size, color)
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
