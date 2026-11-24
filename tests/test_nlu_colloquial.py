"""
Unit and regression tests for Persian NLU engine:
- Normalization (Persian/Arabic chars, digits, typos, colloquial phrases)
- Semantic slot extraction (target columns, source columns, percentages, ranges)
- Arbitrary sentence structures and colloquial phrasing
- Zero regressions across standard and colloquial inputs
- Ambiguity management and clarification prompts
"""
import pytest

from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.persian_normalizer import (
    normalize_persian_text,
    normalize_colloquial,
    extract_target_column,
    extract_source_columns,
    extract_percentage,
    extract_row_range,
    extract_column_identifier,
)
from excel_assistant.nlp.intent_models import (
    ColumnArithmeticAction,
    ColumnStatsAction,
    DeleteEmptyRowsAction,
    DeleteDuplicatesAction,
    PreventDuplicatesAction,
    HighlightDuplicatesAction,
    SortDataAction,
    FilterDataAction,
    ApplyDiscountAction,
    PriceAdjustmentAction,
    InsertDateAction,
    ReconciliationAction,
    InvoiceAction,
    PayrollAction,
    InventoryAlertAction,
    TrialBalanceAction,
    AccountBalanceAction,
    UndoLastAction,
    SaveWorkbookAction,
    BackupAction,
)
from excel_assistant.core.exceptions import AmbiguousCommandError


# =========================================================================
# 1. Normalization & Tokenization Tests
# =========================================================================

def test_character_and_digit_normalization():
    raw = "ستون B را در ١٢ ضرب کن و توی شیلد ذخیره کن"
    norm = normalize_persian_text(raw)
    assert "12" in norm
    assert "شیت" in norm


def test_typo_correction():
    typos = {
        "حمع ستون B": "جمع ستون B",
        "صرب B در C": "ضرب B در C",
        "تقریق A از B": "تفریق A از B",
        "تقصیم A بر B": "تقسیم A بر B",
        "فیلطر ستون B": "فیلتر ستون B",
        "ردیفهای خالی": "ردیف های خالی",
        "تکراریها را پاک کن": "تکراری ها را پاک کن",
    }
    for inp, expected_part in typos.items():
        norm = normalize_persian_text(inp)
        assert expected_part in norm


def test_colloquial_normalizations():
    assert "را" in normalize_colloquial("ستون D رو")
    assert "در" in normalize_colloquial("توی E")
    assert "کن" in normalize_colloquial("ضرب بکن")
    assert "پاسخ را" in normalize_colloquial("پاسخشو")
    assert "15 درصد" in normalize_colloquial("15% تخفیف")


def test_extract_column_identifier():
    assert extract_column_identifier("A") == "A"
    assert extract_column_identifier("b") == "B"
    assert extract_column_identifier("1") == "A"
    assert extract_column_identifier("2") == "B"
    assert extract_column_identifier("26") == "Z"
    assert extract_column_identifier("دی") == "D"
    assert extract_column_identifier("سی") == "C"
    assert extract_column_identifier("بی") == "B"


# =========================================================================
# 2. Semantic Slot Extraction Tests
# =========================================================================

def test_extract_target_column_all_positions():
    # Target at start
    tgt1, rem1 = extract_target_column("توی E بنویس حاصلضرب D و B")
    assert tgt1 == "E"
    assert "حاصلضرب D و B" in rem1

    tgt2, rem2 = extract_target_column("در ستون D قرار بده حاصل جمع A و B")
    assert tgt2 == "D"

    # Target at end with result noun
    tgt3, rem3 = extract_target_column("ستون D رو ضربدر B کن و جواب رو E بذار")
    assert tgt3 == "E"
    assert "E" not in rem3

    # Target at end with colloquial verb
    tgt4, rem4 = extract_target_column("D رو تو B ضرب کن بریز تو E")
    assert tgt4 == "E"

    # Target at end with preposition روی
    tgt5, rem5 = extract_target_column("A و B رو تقسیم کن روی C بنویس")
    assert tgt5 == "C"

    # No target
    tgt6, rem6 = extract_target_column("ستون D و B را ضرب کن")
    assert tgt6 is None


def test_extract_source_columns_various_formats():
    # Multi-number format
    cols1 = extract_source_columns("ستون های 2 و 4 و 5 رو با هم جمع کن")
    assert cols1 == ["B", "D", "E"]

    # Explicit Latin letters
    cols2 = extract_source_columns("ستون D رو ضربدر B کن")
    assert cols2 == ["D", "B"]

    # Phonetic Persian names
    cols3 = extract_source_columns("دی رو در بی ضرب کن")
    assert cols3 == ["D", "B"]

    # Subtraction direction: B را از C کم کن -> [C, B]
    cols4 = extract_source_columns("B رو از C کم کن")
    assert cols4 == ["C", "B"]


# =========================================================================
# 3. Arithmetic Operations Across Diverse Word Orders
# =========================================================================

@pytest.mark.parametrize("cmd, expected_op, expected_c1, expected_c2, expected_target", [
    ("ستون B را در C ضرب کن و در D قرار بده", "multiply", "B", "C", "D"),
    ("ستون D رو ضربدر B کن و جواب رو E بذار", "multiply", "D", "B", "E"),
    ("توی E بنویس حاصلضرب D و B", "multiply", "D", "B", "E"),
    ("حاصل‌ضرب B و D رو داخل E قرار بده", "multiply", "B", "D", "E"),
    ("B و D رو در هم ضرب کن و نتیجه رو در E بنویس", "multiply", "B", "D", "E"),
    ("D رو تو B ضرب کن بریز تو E", "multiply", "D", "B", "E"),
    ("B رو از C کم کن و در A بنویس", "subtract", "C", "B", "A"),
    ("ستون C رو منهای B کن و توی A بنویس", "subtract", "C", "B", "A"),
    ("مجموع D و B رو توی E بنویس", "add", "D", "B", "E"),
    ("A و B رو تقسیم کن روی C بنویس", "divide", "A", "B", "C"),
    ("ستون های 2 و 4 و 5 رو با هم جمع کن و توی ستون 6 بنویس", "add", "B", "D", "F"),
])
def test_colloquial_column_arithmetic_variations(cmd, expected_op, expected_c1, expected_c2, expected_target):
    parser = PersianIntentParser()
    act = parser.parse(cmd)
    assert isinstance(act, ColumnArithmeticAction)
    assert act.operation == expected_op
    assert act.col1 == expected_c1
    assert act.col2 == expected_c2
    assert act.target_col == expected_target


def test_arithmetic_with_header_map():
    parser = PersianIntentParser()
    header_map = {"قیمت": "B", "تعداد": "C", "جمع کل": "D"}
    cmd = "ستون قیمت را در ستون تعداد ضرب کن و در جمع کل بنویس"
    act = parser.parse(cmd, sheet_analysis=type("Analysis", (), {"header_to_col": header_map})())
    assert isinstance(act, ColumnArithmeticAction)
    assert act.operation == "multiply"
    assert act.col1 == "B"
    assert act.col2 == "C"
    assert act.target_col == "D"


# =========================================================================
# 4. Colloquial Data Cleaning Tests
# =========================================================================

@pytest.mark.parametrize("cmd", [
    "ردیف‌های خالی را حذف کن",
    "ردیف‌های خالی رو بنداز دور",
    "سطرهای بدون اطلاعات رو حذف کن",
    "خطوط خالی شیت رو پاک کن",
    "سطرهای خالی رو بردار",
])
def test_delete_empty_rows_colloquial(cmd):
    parser = PersianIntentParser()
    act = parser.parse(cmd)
    assert isinstance(act, DeleteEmptyRowsAction)
    assert act.is_destructive is True


@pytest.mark.parametrize("cmd", [
    "حذف داده‌های تکراری",
    "تکراری‌ها رو پاک کن",
    "سطرهای تکراری رو بنداز دور",
    "داده های تکراری ستون A را پاک کن",
])
def test_delete_duplicates_colloquial(cmd):
    parser = PersianIntentParser()
    act = parser.parse(cmd)
    assert isinstance(act, DeleteDuplicatesAction)
    assert act.is_destructive is True


def test_prevent_duplicates():
    parser = PersianIntentParser()
    act1 = parser.parse("ستون 12 نباید داده تکراری داشته باشه")
    assert isinstance(act1, PreventDuplicatesAction)
    assert act1.target_col == "L"

    act2 = parser.parse("در ستون B جلوی ورود داده تکراری را بگیر")
    assert isinstance(act2, PreventDuplicatesAction)
    assert act2.target_col == "B"


# =========================================================================
# 5. Colloquial Sorting & Filtering Tests
# =========================================================================

def test_colloquial_sorting():
    parser = PersianIntentParser()

    act1 = parser.parse("از بزرگ به کوچک ستون C رو مرتب کن")
    assert isinstance(act1, SortDataAction)
    assert act1.key_col == "C"
    assert act1.ascending is False

    act2 = parser.parse("ستون C رو سورت کن نزولی")
    assert isinstance(act2, SortDataAction)
    assert act2.key_col == "C"
    assert act2.ascending is False

    act3 = parser.parse("از کوچک به بزرگ سورت کن ستون B رو")
    assert isinstance(act3, SortDataAction)
    assert act3.key_col == "B"
    assert act3.ascending is True


def test_colloquial_filtering():
    parser = PersianIntentParser()

    act1 = parser.parse("فیلتر ستون B بزرگتر از 500")
    assert isinstance(act1, FilterDataAction)
    assert act1.target_col == "B"
    assert act1.condition == "gt"
    assert act1.value == 500.0

    act2 = parser.parse("فیلتر هوشمند جدول")
    assert isinstance(act2, FilterDataAction)
    assert act2.condition == "autofilter"

    act3 = parser.parse("حذف فیلتر")
    assert isinstance(act3, FilterDataAction)
    assert act3.condition == "clear"


# =========================================================================
# 6. Financial & Management Commands
# =========================================================================

def test_apply_discount_variations():
    parser = PersianIntentParser()

    act1 = parser.parse("20 درصد تخفیف روی ستون B اعمال کن و در ستون E بذار")
    assert isinstance(act1, ApplyDiscountAction)
    assert act1.price_col == "B"
    assert act1.discount_percent == 20.0
    assert act1.target_col == "E"

    act2 = parser.parse("تخفیف ستون B با ستون C در ستون D")
    assert isinstance(act2, ApplyDiscountAction)
    assert act2.price_col == "B"
    assert act2.discount_col == "C"
    assert act2.target_col == "D"


def test_price_adjustment_colloquial():
    parser = PersianIntentParser()

    act1 = parser.parse("قیمت‌های ستون C را ۱۵ درصد گرون کن")
    assert isinstance(act1, PriceAdjustmentAction)
    assert act1.target_col == "C"
    assert act1.percentage == 15.0
    assert act1.operation == "increase"

    act2 = parser.parse("ستون C رو ۱۵ درصد ارزون کن")
    assert isinstance(act2, PriceAdjustmentAction)
    assert act2.target_col == "C"
    assert act2.percentage == 15.0
    assert act2.operation == "decrease"


def test_insert_date_colloquial():
    parser = PersianIntentParser()

    act = parser.parse("از سطر 1 تا 30 در ستون A تاریخ بزن")
    assert isinstance(act, InsertDateAction)
    assert act.target_col == "A"
    assert act.start_row == 1
    assert act.end_row == 30


def test_accounting_actions():
    parser = PersianIntentParser()

    assert isinstance(parser.parse("مغایرت‌گیری ستون A و B"), ReconciliationAction)
    assert isinstance(parser.parse("ثبت فاکتور فروش رسمی"), InvoiceAction)
    assert isinstance(parser.parse("محاسبه حقوق و دستمزد و بیمه"), PayrollAction)
    assert isinstance(parser.parse("موجودی انبار و هشدار کمبود موجودی"), InventoryAlertAction)
    assert isinstance(parser.parse("تهیه تراز آزمایشی چهار ستونه"), TrialBalanceAction)
    assert isinstance(parser.parse("محاسبه مانده حساب بدهکار و بستانکار"), AccountBalanceAction)
    assert isinstance(parser.parse("لغو آخرین عملیات"), UndoLastAction)
    assert isinstance(parser.parse("فایل را ذخیره کن"), SaveWorkbookAction)
    assert isinstance(parser.parse("یک نسخه پشتیبان از فایل بگیر"), BackupAction)


# =========================================================================
# 7. Ambiguity Management & Clarification Prompts
# =========================================================================

def test_ambiguous_incomplete_commands():
    parser = PersianIntentParser()

    # Missing second column for arithmetic
    with pytest.raises(AmbiguousCommandError) as exc_info1:
        parser.parse("ستون B را ضرب کن")
    assert "متوجه دستور" in str(exc_info1.value)
    assert len(exc_info1.value.suggested_options) > 0

    # Unintelligible greeting / non-excel command
    with pytest.raises(AmbiguousCommandError) as exc_info2:
        parser.parse("سلام چطوری حالت خوبه")
    assert "متوجه دستور" in str(exc_info2.value)

    # Empty command
    with pytest.raises(AmbiguousCommandError) as exc_info3:
        parser.parse("   ")
    assert "نمی‌تواند خالی باشد" in str(exc_info3.value)
