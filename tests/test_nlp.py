"""
Tests for Phase 3: Persian Natural Language Understanding and Intent Parsing.
"""
import pytest

from excel_assistant.nlp.intent_models import (
    ColumnArithmeticAction,
    ColumnStatsAction,
    HighlightDuplicatesAction,
    HighlightConditionAction,
    DeleteEmptyRowsAction,
    SortDataAction,
    AutoFitColumnsAction,
    CreateTableAction,
    UndoLastAction,
)
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.core.exceptions import AmbiguousCommandError


def test_persian_rule_parser_all_user_prompts():
    parser = PersianIntentParser()

    # 1. ستون A را در ستون B ضرب کن و جواب را در C بنویس
    cmd1 = "ستون A را در ستون B ضرب کن و جواب را در C بنویس."
    act1 = parser.parse(cmd1)
    assert isinstance(act1, ColumnArithmeticAction)
    assert act1.operation == "multiply"
    assert act1.col1 == "A"
    assert act1.col2 == "B"
    assert act1.target_col == "C"

    # 1.1 Colloquial user command from live screenshot: ستون B را در C ضرب کن و در D قرار بده
    cmd1_1 = "ستون B را در C ضرب کن و در D قرار بده"
    act1_1 = parser.parse(cmd1_1)
    assert isinstance(act1_1, ColumnArithmeticAction)
    assert act1_1.operation == "multiply"
    assert act1_1.col1 == "B"
    assert act1_1.col2 == "C"
    assert act1_1.target_col == "D"

    # 2. میانگین ستون D را حساب کن
    cmd2 = "میانگین ستون D را حساب کن."
    act2 = parser.parse(cmd2)
    assert isinstance(act2, ColumnStatsAction)
    assert act2.stat_type == "average"
    assert act2.source_col == "D"

    # 3. داده‌های تکراری را زرد کن
    cmd3 = "داده‌های تکراری را زرد کن."
    act3 = parser.parse(cmd3)
    assert isinstance(act3, HighlightDuplicatesAction)
    assert act3.color_hex == "#FFFF00"

    # 4. سلول‌های بیشتر از 100 را سبز کن
    cmd4 = "سلول‌های بیشتر از 100 را سبز کن."
    act4 = parser.parse(cmd4)
    assert isinstance(act4, HighlightConditionAction)
    assert act4.condition == "gt"
    assert act4.value == 100
    assert act4.color_hex == "#90EE90"

    # 5. ردیف‌های خالی را حذف کن
    cmd5 = "ردیف‌های خالی را حذف کن."
    act5 = parser.parse(cmd5)
    assert isinstance(act5, DeleteEmptyRowsAction)
    assert act5.is_destructive is True

    # 6. اطلاعات را از بیشترین به کمترین مرتب کن
    cmd6 = "اطلاعات را از بیشترین به کمترین مرتب کن."
    act6 = parser.parse(cmd6)
    assert isinstance(act6, SortDataAction)
    assert act6.ascending is False

    # 7. مجموع ستون B را محاسبه کن
    cmd7 = "مجموع ستون B را محاسبه کن."
    act7 = parser.parse(cmd7)
    assert isinstance(act7, ColumnStatsAction)
    assert act7.stat_type == "sum"
    assert act7.source_col == "B"

    # 8. ایجاد جدول
    cmd8 = "یک جدول بساز"
    act8 = parser.parse(cmd8)
    assert isinstance(act8, CreateTableAction)

    # 9. تنظیم عرض ستون‌ها
    cmd9 = "عرض ستون‌ها را خودکار تنظیم کن"
    act9 = parser.parse(cmd9)
    assert isinstance(act9, AutoFitColumnsAction)

    # 10. لغو آخرین عملیات
    cmd10 = "لغو آخرین عملیات"
    act10 = parser.parse(cmd10)
    assert isinstance(act10, UndoLastAction)

    # 11. Persian digits & conditional less than
    cmd11 = "سلول‌های کمتر از ۵۰ را قرمز کن"
    act11 = parser.parse(cmd11)
    assert isinstance(act11, HighlightConditionAction)
    assert act11.condition == "lt"
    assert act11.value == 50.0
    assert act11.color_hex == "#FF6B6B"

    # 12. Addition of columns
    cmd12 = "ستون A را با ستون B جمع کن و در ستون C بنویس"
    act12 = parser.parse(cmd12)
    assert isinstance(act12, ColumnArithmeticAction)
    assert act12.operation == "add"

    # 13. Subtraction & Division
    cmd13 = "ستون C منهای ستون B در ستون D"
    act13 = parser.parse(cmd13)
    assert isinstance(act13, ColumnArithmeticAction)
    assert act13.operation == "subtract"

    cmd14 = "ستون A تقسیم بر ستون B در ستون E"
    act14 = parser.parse(cmd14)
    assert isinstance(act14, ColumnArithmeticAction)
    assert act14.operation == "divide"


def test_header_resolution():
    parser = PersianIntentParser()

    class MockSheetAnalysis:
        header_to_col = {"قیمت": "B", "تعداد": "C", "فروش": "D"}

    cmd = "ستون قیمت را در ستون تعداد ضرب کن و در ستون فروش بنویس"
    act = parser.parse(cmd, sheet_analysis=MockSheetAnalysis())
    assert isinstance(act, ColumnArithmeticAction)
    assert act.col1 == "B"
    assert act.col2 == "C"
    assert act.target_col == "D"


def test_multi_column_arithmetic_parsing():
    parser = PersianIntentParser()

    cmd1 = "ستون های 2 و 4 و 5 رو با هم جمع کن و بذار در ستون 8"
    act1 = parser.parse(cmd1)
    assert isinstance(act1, ColumnArithmeticAction)
    assert act1.operation == "add"
    assert act1.columns == ["B", "D", "E"]
    assert act1.target_col == "H"

    cmd2 = "ستون های B و D و E را با هم ضرب کن و در ستون H قرار بده"
    act2 = parser.parse(cmd2)
    assert isinstance(act2, ColumnArithmeticAction)
    assert act2.operation == "multiply"
    assert act2.columns == ["B", "D", "E"]
    assert act2.target_col == "H"

    cmd3 = "جمع ستون 2 و 4 و 5 در ستون 8"
    act3 = parser.parse(cmd3)
    assert isinstance(act3, ColumnArithmeticAction)
    assert act3.operation == "add"
    assert act3.columns == ["B", "D", "E"]
    assert act3.target_col == "H"


def test_discount_parsing():
    from excel_assistant.nlp.intent_models import ApplyDiscountAction
    parser = PersianIntentParser()

    cmd1 = "تخفیف ستون B با درصد ستون C در ستون D بنویس"
    act1 = parser.parse(cmd1)
    assert isinstance(act1, ApplyDiscountAction)
    assert act1.price_col == "B"
    assert act1.discount_col == "C"
    assert act1.target_col == "D"

    cmd2 = "20 درصد تخفیف روی ستون B اعمال کن و در ستون E بذار"
    act2 = parser.parse(cmd2)
    assert isinstance(act2, ApplyDiscountAction)
    assert act2.price_col == "B"
    assert act2.discount_percent == 20.0
    assert act2.target_col == "E"


def test_insert_date_parsing():
    from excel_assistant.nlp.intent_models import InsertDateAction
    parser = PersianIntentParser()

    cmd1 = "از سطر 1 تا 30 در ستون A تاریخ بزن"
    act1 = parser.parse(cmd1)
    assert isinstance(act1, InsertDateAction)
    assert act1.target_col == "A"
    assert act1.start_row == 1
    assert act1.end_row == 30

    cmd2 = "درج تاریخ 1403/08/29 از ردیف 1 تا 30 در ستون B"
    act2 = parser.parse(cmd2)
    assert isinstance(act2, InsertDateAction)
    assert act2.target_col == "B"
    assert act2.year == 1403
    assert act2.month == 8
    assert act2.day == 29


def test_prevent_duplicates_parsing():
    from excel_assistant.nlp.intent_models import PreventDuplicatesAction
    parser = PersianIntentParser()

    cmd1 = "ستون 12 نباید داده تکراری داشته باشه"
    act1 = parser.parse(cmd1)
    assert isinstance(act1, PreventDuplicatesAction)
    assert act1.target_col == "L"

    cmd2 = "در ستون L جلوی ورود داده تکراری را بگیر"
    act2 = parser.parse(cmd2)
    assert isinstance(act2, PreventDuplicatesAction)
    assert act2.target_col == "L"


def test_ambiguous_command_error():
    parser = PersianIntentParser()
    with pytest.raises(AmbiguousCommandError) as exc_info:
        parser.parse("سلام چطوری هوا چطوره؟")
    assert "متوجه دستور" in str(exc_info.value)
    assert len(exc_info.value.suggested_options) > 0
