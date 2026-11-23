"""
Tests for 12 Core Business & Productivity Features:
1. مرتب‌سازی اطلاعات (صعودی و نزولی)
2. فیلتر هوشمند اطلاعات (براساس شرط)
3. حذف داده‌های تکراری
4. حذف سطرهای خالی
5. جمع خودکار (AutoSum)
6. میانگین‌گیری داده‌ها
7. بیشترین و کمترین (MAX / MIN)
8. شمارش اطلاعات (سلول‌های پر یا عددی)
9. تبدیل متن به عدد (اصلاح داده‌های عددی)
10. جدا کردن متن (تفکیک نام، نام خانوادگی و...)
11. ادغام سلول‌ها
12. تغییر فرمت سلول (عدد، درصد، تاریخ و واحد پول)
"""
import pytest
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import (
    SortDataAction, FilterDataAction, DeleteDuplicatesAction, DeleteEmptyRowsAction,
    ColumnStatsAction, ConvertTextToNumberAction, SplitTextAction, MergeCellsAction,
    SetCellFormatAction
)
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.base import ActionContext


@pytest.fixture
def manager():
    mgr = ExcelConnectionManager(preferred_driver="openpyxl")
    mgr.connect()
    yield mgr
    mgr.close()


def test_sort_data(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "کد")
    sheet.set_value("B1", "نمره")
    sheet.set_value("A2", "A")
    sheet.set_value("B2", 12)
    sheet.set_value("A3", "B")
    sheet.set_value("B3", 20)
    sheet.set_value("A4", "C")
    sheet.set_value("B4", 15)

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    # Sort descending
    action_desc = SortDataAction(key_col="B", ascending=False)
    res = executor.execute(action_desc, ctx)
    assert res.success is True
    assert sheet.get_value("B2") == 20
    assert sheet.get_value("B3") == 15
    assert sheet.get_value("B4") == 12

    # Sort ascending
    action_asc = SortDataAction(key_col="B", ascending=True)
    res = executor.execute(action_asc, ctx)
    assert res.success is True
    assert sheet.get_value("B2") == 12
    assert sheet.get_value("B3") == 15
    assert sheet.get_value("B4") == 20


def test_smart_filter(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "نام")
    sheet.set_value("B1", "فروش")
    sheet.set_value("A2", "تهران")
    sheet.set_value("B2", 150)
    sheet.set_value("A3", "مشهد")
    sheet.set_value("B3", 80)
    sheet.set_value("A4", "شیراز")
    sheet.set_value("B4", 200)

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    # Filter B > 100
    action_filter = FilterDataAction(target_col="B", condition="gt", value=100)
    res = executor.execute(action_filter, ctx)
    assert res.success is True
    assert res.data["matched"] == 2
    assert res.data["hidden"] == 1

    # Clear filter
    action_clear = FilterDataAction(condition="clear")
    res_clear = executor.execute(action_clear, ctx)
    assert res_clear.success is True


def test_delete_duplicates(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "کد")
    sheet.set_value("A2", "101")
    sheet.set_value("A3", "102")
    sheet.set_value("A4", "101")  # Duplicate
    sheet.set_value("A5", "103")

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet, confirmed=True)

    action = DeleteDuplicatesAction(target_col="A")
    res = executor.execute(action, ctx)
    assert res.success is True
    assert res.data == 1
    assert sheet.get_value("A4") == "103"


def test_delete_empty_rows(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "نام")
    sheet.set_value("A2", "علی")
    sheet.set_value("A3", "")  # Empty row
    sheet.set_value("A4", "رضا")

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet, confirmed=True)

    action = DeleteEmptyRowsAction()
    res = executor.execute(action, ctx)
    assert res.success is True
    assert res.data == 1
    assert sheet.get_value("A3") == "رضا"


def test_autosum_and_stats(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "قیمت")
    sheet.set_value("A2", 1000)
    sheet.set_value("A3", 2000)
    sheet.set_value("A4", 3000)

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    # AutoSum
    res_sum = executor.execute(ColumnStatsAction(stat_type="sum", source_col="A"), ctx)
    assert res_sum.success is True
    assert res_sum.data == 6000
    assert sheet.get_formula("A5") == "=SUM(A2:A4)"

    # Average
    res_avg = executor.execute(ColumnStatsAction(stat_type="average", source_col="A", target_cell="C1"), ctx)
    assert res_avg.success is True
    assert res_avg.data == 2000
    assert sheet.get_formula("C1") == "=AVERAGE(A2:A4)"

    # Max and Min
    res_max = executor.execute(ColumnStatsAction(stat_type="max", source_col="A", target_cell="D1"), ctx)
    assert res_max.data == 3000
    res_min = executor.execute(ColumnStatsAction(stat_type="min", source_col="A", target_cell="D2"), ctx)
    assert res_min.data == 1000

    # Count of non-empty cells
    res_counta = executor.execute(ColumnStatsAction(stat_type="counta", source_col="A", target_cell="E1"), ctx)
    assert res_counta.data == 4  # Includes header or values


def test_convert_text_to_number(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "مبلغ")
    sheet.set_value("A2", "۱۲,۰۰۰ تومان")   # Persian digits with comma and currency
    sheet.set_value("A3", "۲۵,۵۰۰ ریال")
    sheet.set_value("A4", "10,000")

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    action = ConvertTextToNumberAction(target_col="A")
    res = executor.execute(action, ctx)
    assert res.success is True
    assert res.data == 3

    assert sheet.get_value("A2") == 12000
    assert isinstance(sheet.get_value("A2"), int)
    assert sheet.get_value("A3") == 25500
    assert sheet.get_value("A4") == 10000


def test_split_text(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "نام کامل")
    sheet.set_value("A2", "علی رضایی")
    sheet.set_value("A3", "محمد حسین احمدی")

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    action = SplitTextAction(source_col="A", delimiter=" ", header1="نام", header2="نام خانوادگی")
    res = executor.execute(action, ctx)
    assert res.success is True
    assert res.data == 2

    # Verify headers and split parts
    assert sheet.get_value("B1") == "نام"
    assert sheet.get_value("C1") == "نام خانوادگی"
    assert sheet.get_value("B2") == "علی"
    assert sheet.get_value("C2") == "رضایی"
    assert sheet.get_value("B3") == "محمد"
    assert sheet.get_value("C3") == "حسین احمدی"


def test_merge_cells(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "عنوان گزارش جامع")

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    action = MergeCellsAction(range_str="A1:D1", align_center=True)
    res = executor.execute(action, ctx)
    assert res.success is True
    assert "A1:D1" in res.affected_range


def test_set_cell_format(manager):
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", 1500000)

    executor = ActionExecutor()
    ctx = ActionContext(manager=manager, sheet=sheet)

    action = SetCellFormatAction(target="A1", format_type="currency")
    res = executor.execute(action, ctx)
    assert res.success is True


def test_nlp_parser_all_12_features():
    parser = PersianIntentParser()

    # 1. Sort
    a1 = parser.parse("مرتب‌سازی صعودی ستون B")
    assert isinstance(a1, SortDataAction) and a1.ascending is True
    a2 = parser.parse("مرتب‌سازی نزولی ستون C")
    assert isinstance(a2, SortDataAction) and a2.ascending is False

    # 2. Smart Filter
    a3 = parser.parse("فیلتر ستون B بزرگتر از 100")
    assert isinstance(a3, FilterDataAction) and a3.condition == "gt" and a3.value == 100
    a4 = parser.parse("فیلتر هوشمند جدول")
    assert isinstance(a4, FilterDataAction) and a4.condition == "autofilter"
    a5 = parser.parse("حذف فیلتر")
    assert isinstance(a5, FilterDataAction) and a5.condition == "clear"

    # 3. Delete Duplicates
    a6 = parser.parse("حذف داده‌های تکراری ستون A")
    assert isinstance(a6, DeleteDuplicatesAction)

    # 4. Delete Empty Rows
    a7 = parser.parse("سطرهای خالی را حذف کن")
    assert isinstance(a7, DeleteEmptyRowsAction)

    # 5. AutoSum
    a8 = parser.parse("جمع خودکار ستون B")
    assert isinstance(a8, ColumnStatsAction) and a8.stat_type == "sum"

    # 6. Average
    a9 = parser.parse("میانگین ستون C")
    assert isinstance(a9, ColumnStatsAction) and a9.stat_type == "average"

    # 7. Max & Min
    a10 = parser.parse("بیشترین مقدار ستون A")
    assert isinstance(a10, ColumnStatsAction) and a10.stat_type == "max"
    a11 = parser.parse("کمترین مقدار ستون A")
    assert isinstance(a11, ColumnStatsAction) and a11.stat_type == "min"

    # 8. Count
    a12 = parser.parse("شمارش سلول‌های پر ستون B")
    assert isinstance(a12, ColumnStatsAction) and a12.stat_type == "counta"

    # 9. Convert text to number
    a13 = parser.parse("تبدیل متن به عدد در ستون A")
    assert isinstance(a13, ConvertTextToNumberAction) and a13.target_col == "A"

    # 10. Split text
    a14 = parser.parse("تفکیک نام و نام خانوادگی در ستون A")
    assert isinstance(a14, SplitTextAction) and a14.source_col == "A"

    # 11. Merge cells
    a15 = parser.parse("سلول‌های A1:D1 را ادغام کن")
    assert isinstance(a15, MergeCellsAction) and a15.range_str == "A1:D1"

    # 12. Cell formatting
    a16 = parser.parse("فرمت ستون B را به تومان تغییر بده")
    assert isinstance(a16, SetCellFormatAction) and a16.format_type == "currency"
    a17 = parser.parse("فرمت ستون C را درصدی کن")
    assert isinstance(a17, SetCellFormatAction) and a17.format_type == "percentage"
