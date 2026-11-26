"""
Unit tests for Sales Calculation and Reporting (Daily, Monthly, Yearly, and Total).
"""
import pytest
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import SalesSummaryAction
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.base import ActionContext


def test_parse_sales_commands():
    parser = PersianIntentParser()

    # 1. Comprehensive all-in-one sales command
    action_all = parser.parse("محاسبه فروش روزانه و ماهانه و سالانه و کل")
    assert isinstance(action_all, SalesSummaryAction)
    assert action_all.period == "all"
    assert action_all.create_report_sheet is True

    # 2. Daily sales
    action_daily = parser.parse("فروش روزانه را حساب کن")
    assert isinstance(action_daily, SalesSummaryAction)
    assert action_daily.period == "daily"

    # 3. Monthly sales
    action_monthly = parser.parse("محاسبه فروش ماهانه")
    assert isinstance(action_monthly, SalesSummaryAction)
    assert action_monthly.period == "monthly"

    # 4. Yearly sales
    action_yearly = parser.parse("فروش سالانه رو حساب کن")
    assert isinstance(action_yearly, SalesSummaryAction)
    assert action_yearly.period == "yearly"

    # 5. Total sales
    action_total = parser.parse("فروش کل چقدره؟")
    assert isinstance(action_total, SalesSummaryAction)
    assert action_total.period == "total"

    # 6. Specific columns mentioned
    action_cols = parser.parse("محاسبه فروش ستون B بر اساس تاریخ ستون A")
    assert isinstance(action_cols, SalesSummaryAction)
    assert action_cols.amount_col == "B"
    assert action_cols.date_col == "A"


def test_execute_sales_summary_calculation():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()
    sheet = manager.get_active_sheet()

    # Set up sample sales data
    sheet.set_value("A1", "تاریخ")
    sheet.set_value("B1", "مبلغ فروش")

    # 2 transactions on 1403/08/01
    sheet.set_value("A2", "1403/08/01")
    sheet.set_value("B2", 100000)

    sheet.set_value("A3", "1403/08/01")
    sheet.set_value("B3", 150000)

    # 1 transaction on 1403/08/15
    sheet.set_value("A4", "1403/08/15")
    sheet.set_value("B4", 250000)

    # 1 transaction on 1403/09/01 (Different month)
    sheet.set_value("A5", "1403/09/01")
    sheet.set_value("B5", 500000)

    # 1 transaction in 1402 (Different year)
    sheet.set_value("A6", "1402/12/20")
    sheet.set_value("B6", 300000)

    executor = ActionExecutor()
    action = SalesSummaryAction(date_col="A", amount_col="B", period="all", create_report_sheet=True)
    ctx = ActionContext(manager=manager, sheet=sheet)

    res = executor.execute(action, ctx)

    assert res.success is True
    assert "گزارش_فروش" in manager.active_driver.get_sheet_names()
    assert res.data["total_sales"] == 1300000  # 100k + 150k + 250k + 500k + 300k
    assert res.data["total_count"] == 5
    assert res.data["avg_sale"] == 260000.0
    assert res.data["daily_count"] == 4   # 1403/08/01, 1403/08/15, 1403/09/01, 1402/12/20
    assert res.data["monthly_count"] == 3 # 1403/08, 1403/09, 1402/12
    assert res.data["yearly_count"] == 2  # 1403, 1402

    # Verify report sheet contents
    report_sheet = manager.active_driver.get_sheet("گزارش_فروش")
    assert "گزارش جامع" in str(report_sheet.get_value("A1"))
    assert "کل فروش" in str(report_sheet.get_value("A5"))
    assert "1,300,000" in str(report_sheet.get_value("B5"))

    manager.close()


def test_sales_auto_detection_of_columns():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()
    sheet = manager.get_active_sheet()

    # Column A: Customer name, Column B: Date, Column C: Product, Column D: Amount
    sheet.set_value("A1", "مشتری")
    sheet.set_value("B1", "تاریخ فاکتور")
    sheet.set_value("C1", "کالا")
    sheet.set_value("D1", "مبلغ نهایی")

    sheet.set_value("A2", "علی")
    sheet.set_value("B2", "1403/05/10")
    sheet.set_value("C2", "لپ‌تاپ")
    sheet.set_value("D2", 20000000)

    sheet.set_value("A3", "رضا")
    sheet.set_value("B3", "1403/05/11")
    sheet.set_value("C3", "موس")
    sheet.set_value("D3", 1000000)

    executor = ActionExecutor()
    # Notice: date_col and amount_col are None -> must auto-detect B and D!
    action = SalesSummaryAction(date_col=None, amount_col=None, period="all", create_report_sheet=True)
    ctx = ActionContext(manager=manager, sheet=sheet)

    res = executor.execute(action, ctx)

    assert res.success is True
    assert res.data["total_sales"] == 21000000
    assert res.data["total_count"] == 2

    manager.close()


def test_sales_auto_compute_when_unit_price_selected():
    """Verify sales amount is calculated as qty * unit_price * (1 - disc) when unit price column is selected."""
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()
    sheet = manager.get_active_sheet()

    sheet.set_value("A1", "تاریخ")
    sheet.set_value("B1", "تعداد")
    sheet.set_value("C1", "قیمت واحد (ریال)")
    sheet.set_value("D1", "تخفیف درصد")

    # Row 1: 10 * 100,000 = 1,000,000 (0% disc)
    sheet.set_value("A2", "1405/01/01")
    sheet.set_value("B2", 10)
    sheet.set_value("C2", 100000)
    sheet.set_value("D2", 0)

    # Row 2: 20 * 50,000 * 0.9 = 900,000 (10% disc)
    sheet.set_value("A3", "1405/01/02")
    sheet.set_value("B3", 20)
    sheet.set_value("C3", 50000)
    sheet.set_value("D3", 10)

    executor = ActionExecutor()
    # User or dialog picked amount_col="C" (which is unit price)
    action = SalesSummaryAction(date_col="A", amount_col="C", period="all", create_report_sheet=True)
    ctx = ActionContext(manager=manager, sheet=sheet)

    res = executor.execute(action, ctx)

    assert res.success is True
    # Instead of summing 100,000 + 50,000 = 150,000, it computes 1,000,000 + 900,000 = 1,900,000!
    assert res.data["total_sales"] == 1900000.0
    assert res.data["total_count"] == 2

    manager.close()


def test_sales_short_date_smart_month_and_year_grouping():
    """Verify MM/DD dates (e.g. 03/02, 03/05, 04/16, 05/20) correctly group into exactly 3 months."""
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()
    sheet = manager.get_active_sheet()

    sheet.set_value("A1", "سند")
    sheet.set_value("B1", "تاریخ")
    sheet.set_value("C1", "مبلغ")

    # 2 rows in Month 03
    sheet.set_value("A2", "SI01")
    sheet.set_value("B2", "03/02")
    sheet.set_value("C2", 1000)

    sheet.set_value("A3", "SI02")
    sheet.set_value("B3", "03/05")
    sheet.set_value("C3", 2000)

    # 1 row in Month 04
    sheet.set_value("A4", "SI03")
    sheet.set_value("B4", "04/16")
    sheet.set_value("C4", 3000)

    # 1 row in Month 05
    sheet.set_value("A5", "SI04")
    sheet.set_value("B5", "05/20")
    sheet.set_value("C5", 4000)

    executor = ActionExecutor()
    action = SalesSummaryAction(date_col="B", amount_col="C", period="all", create_report_sheet=True)
    ctx = ActionContext(manager=manager, sheet=sheet)
    res = executor.execute(action, ctx)

    assert res.success is True
    assert res.data["total_sales"] == 10000.0
    # Must group into exactly 3 months ('ماه 03', 'ماه 04', 'ماه 05'), NOT 4 separate days!
    assert res.data["monthly_count"] == 3

    manager.close()


