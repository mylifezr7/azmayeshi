"""
Persian Actionable Suggestion Engine.
Analyzes workbook structure and generates smart, natural suggestions for users.
"""
from typing import List
from excel_assistant.analyzer.inspector import SheetAnalysis, WorkbookAnalysis


def generate_smart_suggestions(analysis: SheetAnalysis) -> List[str]:
    """Generates intelligent Persian suggestions based on sheet analysis."""
    suggestions: List[str] = []

    # 1. Excel errors check
    if analysis.formula_report.error_cells:
        err_list = ", ".join(analysis.formula_report.error_cells[:4])
        suggestions.append(f"❌ در سلول‌های ({err_list}) خطای فرمولی وجود دارد. پیشنهاد می‌شود با بررسی فرمول یا جایگزینی با مقدار معتبر آن را اصلاح کنید.")

    # 2. Empty rows check
    if analysis.empty_row_indices:
        suggestions.append(f"🧹 تعداد {len(analysis.empty_row_indices)} ردیف خالی بین داده‌ها شناسایی شد. می‌توانید با گفتن «ردیف‌های خالی را حذف کن» شیت را مرتب کنید.")

    # 3. Duplicate rows check
    if analysis.duplicate_row_indices:
        suggestions.append(f"⚠️ تعداد {len(analysis.duplicate_row_indices)} ردیف تکراری کامل شناسایی شد. می‌توانید بگویید «داده‌های تکراری را حذف کن» یا «داده‌های تکراری را زرد کن».")

    # 4. Numeric column operations suggestion
    numeric_cols = [letter for letter, title, c_type in analysis.headers if c_type in ("numeric", "integer", "float", "currency")]
    if len(numeric_cols) >= 2:
        c1, c2 = numeric_cols[0], numeric_cols[1]
        t1 = analysis.columns[c1].header_name
        t2 = analysis.columns[c2].header_name
        suggestions.append(f"🔢 ستون‌های عددی '{t1}' ({c1}) و '{t2}' ({c2}) شناسایی شدند. برای مثال می‌توانید بگویید: «ستون {c1} را در ستون {c2} ضرب کن و در ستون بعدی بنویس».")
    elif len(numeric_cols) == 1:
        c = numeric_cols[0]
        t = analysis.columns[c].header_name
        suggestions.append(f"📈 ستون عددی '{t}' ({c}) وجود دارد. می‌توانید بگویید «میانگین ستون {c} را حساب کن» یا «مجموع ستون {c} را محاسبه کن».")

    # 5. Table creation suggestion
    if analysis.row_count > 2 and analysis.column_count >= 2:
        suggestions.append("📊 برای فیلترگذاری سریع و ظاهر حرفه‌ای‌تر، می‌توانید بگویید «یک جدول بساز».")

    # 6. Formatting suggestion
    suggestions.append("🎨 می‌توانید اندازه ستون‌ها را با گفتن «عرض ستون‌ها را خودکار تنظیم کن» مرتب نمایید.")

    return suggestions
