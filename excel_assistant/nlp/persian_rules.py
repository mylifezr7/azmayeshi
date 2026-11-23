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
    InsertDateAction,
    PreventDuplicatesAction,
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
    TopSellingAction,
    CustomerLedgerAction,
    TrialBalanceAction,
    AccountBalanceAction,
    PayrollAction,
    DepreciationAction,
    FinancialStatementAction,
    DeepFileAuditAction,
    AutoChartAction,
    BackupAction,
    PdfExportAction,
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

        # 1.5 Deep File Audit / سلامت و خطاهای فایل اکسل
        if any(w in text_lower for w in ["سلامت فایل", "خطاهای فایل", "بررسی خطاهای فایل", "چکاپ فایل", "سلول های خالی", "فرمول های اشتباه", "تحلیل هوشمند فایل", "خطاهای محاسباتی", "داده های نامعتبر"]):
            return DeepFileAuditAction(description_persian="تحلیل هوشمند و عمیق سلامت فایل اکسل")

        # 1.6 Backup / نسخه پشتیبان
        if any(w in text_lower for w in ["پشتیبان", "بکاپ", "backup"]):
            return BackupAction(description_persian="ذخیره نسخه پشتیبان از فایل اکسل")

        # 1.7 PDF Export / خروجی PDF
        if any(w in text_lower for w in ["pdf", "پی دی اف", "پرینت"]) and any(w in text_lower for w in ["خروجی", "دانلود", "تبدیل", "صادر", "گرفتن", "بگیر"]):
            return PdfExportAction(description_persian="تولید خروجی PDF استاندارد از شیت فعال")

        # 1.8 Reconciliation / مغایرت‌گیری دو ستون یا دو فهرست
        if any(w in text_lower for w in ["مغایرت", "تطبیق دو ستون", "مقایسه دو ستون", "مغایرت گیری", "مغایرت‌گیری"]):
            cols_found = []
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text):
                res = resolve_column(c_str, header_map)
                if res and res not in cols_found:
                    cols_found.append(res)
            for c_let in re.findall(r'\b([A-Za-z]{1,2})\b', text):
                c_up = c_let.upper()
                if c_up not in cols_found:
                    cols_found.append(c_up)
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

        # 1.10 Payroll, Overtime, Insurance & Salary Tax / حقوق و دستمزد
        if any(w in text_lower for w in ["حقوق و دستمزد", "حقوق", "دستمزد", "اضافه کاری", "اضافه‌کاری", "بیمه کارگر", "بیمه تامین اجتماعی", "مالیات حقوق", "فیش حقوق"]):
            return PayrollAction(description_persian="محاسبه لیست حقوق، اضافه‌کاری، بیمه و مالیات")

        # 1.11 Inventory Alert & Valuation / موجودی انبار و هشدار کمبود موجودی
        if any(w in text_lower for w in ["کمبود موجودی", "هشدار موجودی", "کسری موجودی", "موجودی انبار", "کنترل انبار", "ارزش موجودی", "ارزش کل موجودی", "انبارگردانی"]):
            return InventoryAlertAction(description_persian="مدیریت موجودی انبار و هشدار کمبود موجودی")

        # 1.12 Top Selling Products / پرفروش‌ترین کالاها
        if any(w in text_lower for w in ["پرفروش ترین", "پرفروش‌ترین", "کالاهای پرفروش", "بیشترین فروش کالا", "پرفروش"]):
            return TopSellingAction(description_persian="گزارش پرفروش‌ترین کالاهای ثبت‌شده")

        # 1.13 Customer Ledger / بدهی و طلب مشتریان
        if any(w in text_lower for w in ["بدهی مشتری", "طلب مشتری", "بدهی و طلب", "مانده مشتری", "طرف حساب"]):
            return CustomerLedgerAction(description_persian="محاسبه مانده بدهی و طلب مشتریان")

        # 1.14 Trial Balance / تراز آزمایشی
        if any(w in text_lower for w in ["تراز آزمایشی", "تراز ازمایشی", "تراز ۴ ستونه", "تراز چهار ستونه", "تراز ۶ ستونه", "تراز شش ستونه"]):
            return TrialBalanceAction(description_persian="تهیه تراز آزمایشی چهار ستونه حسابداری")

        # 1.15 Account Balance / محاسبه مانده حساب (بدهکار و بستانکار)
        if any(w in text_lower for w in ["مانده حساب", "مانده بدهکار و بستانکار", "محاسبه مانده بدهکار", "مانده گیری", "ستون مانده"]):
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
        if any(w in text_lower for w in ["نمودار", "چارت"]) and any(w in text_lower for w in ["رسم", "بکش", "ایجاد", "نمایش", "تولید", "ساز"]):
            return AutoChartAction(description_persian="تولید نمودار گرافیکی خودکار اکسل")

        # 1.19 Percentage Price Adjustment / افزایش یا کاهش درصدی قیمت‌ها
        # e.g.: قیمت‌های ستون C را ۱۵ درصد افزایش بده
        # e.g.: ستون B را ۱۰ درصد کاهش بده
        if ("درصد" in text or "%" in text) and any(w in text for w in ["افزایش", "کاهش", "گران", "ارزان"]):
            pct_m = re.search(r'(\d+(?:\.\d+)?)\s*(?:درصد|%)', text)
            pct_val = float(pct_m.group(1)) if pct_m else 10.0
            is_decrease = any(w in text for w in ["کاهش", "ارزان", "کم کن"])
            op = "decrease" if is_decrease else "increase"

            col_m = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_m.group(1), header_map) if col_m else None
            if not target_col:
                let_m = re.search(r'\b([A-Za-z]{1,2})\b', text)
                target_col = let_m.group(1).upper() if let_m else "C"

            return PriceAdjustmentAction(
                target_col=target_col,
                percentage=pct_val,
                operation=op,
                description_persian=f"{'افزایش' if op == 'increase' else 'کاهش'} {pct_val:.0f} درصدی ستون {target_col}"
            )

        # 1.20 Row-by-Row Multi-Column Average / میانگین سطری بین ستون‌ها
        # e.g.: از ستون B و C میانگین بگیر و نتیجه را در D بنویس
        if "میانگین" in text and (" و " in text or "در " in text):
            target_match = re.search(r'(?:در|توی|داخل|به)\s+(?:ستون\s+)?([A-Za-z0-9\u0600-\u06FF]+)\s*(?:قرار\s*بده|بنویس|درج\s*کن|بریز|بگذار|ثبت\s*کن)?\s*[.]*$', text)
            target_c = resolve_column(target_match.group(1), header_map) if target_match else None

            all_cols = []
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text):
                res = resolve_column(c_str, header_map)
                if res and res not in all_cols and res != target_c:
                    all_cols.append(res)
            for c_let in re.findall(r'\b([A-Za-z]{1,2})\b', text):
                c_up = c_let.upper()
                if c_up not in all_cols and c_up != target_c:
                    all_cols.append(c_up)

            if len(all_cols) >= 2 and target_c:
                return ColumnArithmeticAction(
                    operation="average",
                    col1=all_cols[0],
                    col2=all_cols[1],
                    columns=all_cols,
                    target_col=target_c,
                    as_formula=True,
                    description_persian=f"میانگین سطری ستون‌های {' و '.join(all_cols)} و درج در ستون {target_c}"
                )

        # 2. Analyze sheet command
        if any(w in text_lower for w in ["تحلیل", "بررسی ساختار", "بررسی فایل", "انالیز"]):
            return AnalyzeSheetAction(description_persian="تحلیل هوشمند ساختار شیت و داده‌ها")

        # 3. Delete empty rows
        if ("ردیف" in text_lower or "سطر" in text_lower) and ("خالی" in text_lower) and any(w in text_lower for w in ["حذف", "پاک", "دلیت"]):
            return DeleteEmptyRowsAction(
                description_persian="حذف ردیف‌های خالی از شیت",
                is_destructive=True
            )

        # 4. Prevent duplicates / Delete duplicates / Highlight duplicates
        if "تکراری" in text_lower or any(w in text_lower for w in ["یکتا کن", "یکتاسازی"]):
            # Check target column if specified
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None

            # Check if user wants to prevent duplicates or validate uniqueness
            if any(w in text_lower for w in ["نباید", "جلوگیری", "جلوی", "مانع", "ارور", "خطا", "ثبت نشود", "وارد نشود", "قفل", "یکتا"]):
                return PreventDuplicatesAction(
                    target_col=target_col or "A",
                    description_persian=f"جلوگیری از ورود دادهٔ تکراری و قفل یکتاسازی ستون {target_col or 'A'}"
                )

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

        # 7.55 Profit Calculation: محاسبه سود (روزانه، ماهانه، سالانه، کل / فروش منهای خرید)
        # e.g.: محاسبه سود روزانه ماهانه و سالانه و کل
        # e.g.: محاسبه سود روزانه و ماهانه
        # e.g.: محاسبه سود روزانه
        # e.g.: سود ماهانه رو حساب کن
        # e.g.: محاسبه سود
        # e.g.: فروش رو منهای مبلغ خرید کن تا سود به دست بیاد
        # e.g.: تعداد فروش و قیمت فروش در شیت 1 و تعداد خرید و قیمت خرید در شیت 2
        if "سود" in text or ("فروش" in text and "خرید" in text and any(w in text for w in ["منهای", "منها", "تفریق", "-", "اختلاف", "به دست"])):
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

            found_cols = []
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text):
                res = resolve_column(c_str, header_map)
                if res and res not in found_cols:
                    found_cols.append(res)

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
                period=period,
                create_report_sheet=True,
                description_persian=f"محاسبه سود {period_fa}"
            )

        # 7.6 Sales Summary: محاسبه فروش روزانه، ماهانه، سالانه و کل
        # e.g.: محاسبه فروش روزانه و ماهانه و سالانه و کل
        # e.g.: فروش روزانه را حساب کن
        # e.g.: فروش ماهانه را حساب کن
        # e.g.: فروش سالانه رو محاسبه کن
        # e.g.: فروش کل چقدره؟
        # e.g.: گزارش فروش ستون D بر اساس تاریخ ستون A
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

            # Detect explicit column references
            date_m = re.search(r'(?:ستون|بر اساس)?\s*تاریخ\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if date_m and date_m.group(1):
                date_col = resolve_column(date_m.group(1), header_map)

            amount_m = re.search(r'(?:ستون|مبلغ)?\s*فروش\s*(?:ستون)?\s*([A-Za-z0-9\u0600-\u06FF]+)?', text)
            if amount_m and amount_m.group(1):
                amount_col = resolve_column(amount_m.group(1), header_map)

            # Check general column mentions if not yet matched
            found_cols = []
            for c_str in re.findall(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text):
                res = resolve_column(c_str, header_map)
                if res and res not in found_cols:
                    found_cols.append(res)

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
        # e.g.: از سطر 1 تا 30 در ستون A تاریخ بزن
        # e.g.: درج تاریخ 1403/08/29 از ردیف 1 تا 30 در ستون A
        # e.g.: از سطر 1 تا 30 در ستون 1 تاریخ بنویس
        if "تاریخ" in text and any(w in text for w in ["بزن", "درج", "بنویس", "قرار", "بذار", "ثبت", "سطر", "ردیف", " تا "]):
            row_m = re.search(r'از\s*(?:سطر|ردیف)\s*(\d+)\s*تا\s*(?:سطر|ردیف)?\s*(\d+)', text)
            start_r = int(row_m.group(1)) if row_m else 1
            end_r = int(row_m.group(2)) if row_m else 30

            date_m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', text)
            if date_m:
                y, m, d = int(date_m.group(1)), int(date_m.group(2)), int(date_m.group(3))
            else:
                y, m, d = 1403, 8, 29

            col_m = re.search(r'(?:در|توی|روی)?\s*ستون\s*([A-Za-z0-9\u0600-\u06FF]+)', text)
            col = resolve_column(col_m.group(1), header_map) if col_m else 'A'

            mode = 'sequential' if any(w in text for w in ['متوالی', 'روزشمار', 'پشت سر هم', 'افزایشی']) else 'fixed'
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

        # 9. Statistical calculations: جمع خودکار، میانگین، بیشترین، کمترین، میانه، تعداد سلول‌های پر
        # e.g.: میانگین ستون D را حساب کن
        # e.g.: جمع خودکار ستون B
        # e.g.: بیشترین و کمترین ستون C
        stats_map = {
            "میانگین": "average",
            "معدل": "average",
            "مجموع": "sum",
            "جمع خودکار": "sum",
            "اتوسام": "sum",
            "اتو سام": "sum",
            "جمع": "sum",
            "کل": "sum",
            "بیشترین": "max",
            "حداکثر": "max",
            "ماکزیمم": "max",
            "کمترین": "min",
            "حداقل": "min",
            "مینیمم": "min",
            "میانه": "median",
            "سلول های پر": "counta",
            "سلول‌های پر": "counta",
            "سلولهای پر": "counta",
            "غیر خالی": "counta",
            "غیرخالی": "counta",
            "تعداد": "count",
            "شمارش": "count",
            "درصد": "percentage",
        }

        for stat_word, stat_code in sorted(stats_map.items(), key=lambda x: len(x[0]), reverse=True):
            if stat_word in text:
                if stat_code == "percentage" and "فرمت" in text_lower:
                    continue
                # Find source column
                col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
                source_col = resolve_column(col_match.group(1), header_map) if col_match else None
                if not source_col:
                    # Try single letter token after word e.g. "میانگین D"
                    letter_match = re.search(r'\b([A-Za-z]{1,2})\b', text)
                    if letter_match:
                        source_col = letter_match.group(1).upper()

                if source_col:
                    stat_fa = {"average": "میانگین", "sum": "جمع خودکار", "max": "بیشترین مقدار (MAX)", "min": "کمترین مقدار (MIN)",
                               "median": "میانه", "count": "تعداد عددی", "counta": "تعداد سلول‌های پر", "percentage": "درصد"}[stat_code]
                    return ColumnStatsAction(
                        stat_type=stat_code,
                        source_col=source_col,
                        description_persian=f"محاسبه {stat_fa} برای ستون {source_col}"
                    )

        # 10. Merge Cells: ادغام سلول‌ها
        # e.g.: سلول‌های A1:D1 را ادغام کن
        # e.g.: ادغام A1 تا B1
        if any(w in text_lower for w in ["ادغام", "مرج", "merge", "یکپارچه سازی سلول"]):
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
        # e.g.: تبدیل متن به عدد در ستون B
        # e.g.: اصلاح داده‌های عددی ستون A
        # e.g.: اعداد متنی ستون C را به عدد واقعی تبدیل کن
        if any(w in text_lower for w in ["متن به عدد", "تبدیل به عدد", "اصلاح داده‌های عددی", "اصلاح اعداد", "اعداد متنی", "تبدیل اعداد", "اصلاح داده"]):
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not target_col:
                let_m = re.search(r'\b([A-Za-z]{1,2})\b', text)
                target_col = let_m.group(1).upper() if let_m else "A"
            return ConvertTextToNumberAction(
                target_col=target_col,
                description_persian=f"تبدیل متون عددی ستون {target_col} به عدد واقعی و اصلاح فرمت"
            )

        # 12. Split text / Text to columns: جدا کردن متن (تفکیک نام، نام خانوادگی و...)
        # e.g.: تفکیک نام و نام خانوادگی در ستون A
        # e.g.: جدا کردن متن ستون B بر اساس فاصله
        # e.g.: تفکیک ستون A با خط تیره
        if any(w in text_lower for w in ["جدا کردن متن", "تفکیک نام", "تفکیک ستون", "جداسازی متن", "جدا کردن ستون", "تفکیک متن", "جداسازی نام", "split"]):
            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            source_col = resolve_column(col_match.group(1), header_map) if col_match else None
            if not source_col:
                let_m = re.search(r'\b([A-Za-z]{1,2})\b', text)
                source_col = let_m.group(1).upper() if let_m else "A"

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
        # e.g.: فیلتر ستون B بزرگتر از 100
        # e.g.: فیلتر هوشمند جدول
        # e.g.: حذف فیلتر
        if "فیلتر" in text_lower or "filter" in text_lower:
            if any(w in text_lower for w in ["حذف فیلتر", "پاک کردن فیلتر", "بدون فیلتر", "پاکسازی فیلتر"]):
                return FilterDataAction(
                    condition="clear",
                    description_persian="پاک کردن تمام فیلترها و نمایش همه سطرها"
                )

            col_match = re.search(r'ستون\s+([A-Za-z0-9\u0600-\u06FF]+)', text)
            target_col = resolve_column(col_match.group(1), header_map) if col_match else None

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
        # e.g.: فرمت ستون B را به تومان تغییر بده
        # e.g.: فرمت درصدی برای ستون C
        # e.g.: فرمت ستون A تاریخ
        if "فرمت" in text_lower or any(w in text_lower for w in ["سه رقم", "هزارگان", "درصدی", "تومانی"]):
            fmt_type = "currency"
            if any(w in text_lower for w in ["تومان", "ریال", "واحد پول", "پول"]):
                fmt_type = "currency"
            elif any(w in text_lower for w in ["درصد", "%", "درصدی"]):
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
                    let_m = re.search(r'\b([A-Za-z]{1,2})\b', text)
                    target = let_m.group(1).upper() if let_m else "A"

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
