"""
Unit tests for Profit Calculation across multiple sheets (Sales - Purchase).
"""
import pytest
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import ProfitCalculationAction
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.base import ActionContext


def test_parse_profit_commands():
    parser = PersianIntentParser()

    # 1. User's exact prompt with 'شیلد 1' and 'شیلد 2'
    prompt = (
        "تعداد فروش و قیمت فروش رو در شیلد 1 نوشتم و تعداد خرید و قیمت خرید کالا هارو تو شیلد 2 نوشتم "
        "حالا فروش رو منهای مبلغ خرید کن تا سود به دست بیاد"
    )
    action = parser.parse(prompt)
    assert isinstance(action, ProfitCalculationAction)
    assert action.sales_sheet == "1"
    assert action.purchase_sheet == "2"

    # 2. User's new prompt: محاسبه سود روزانه ماهانه و سالانه و کل
    action_periods = parser.parse("محاسبه سود روزانه ماهانه و سالانه و کل")
    assert isinstance(action_periods, ProfitCalculationAction)
    assert action_periods.period == "all"

    # 3. Individual period prompts
    action_daily = parser.parse("محاسبه سود روزانه")
    assert isinstance(action_daily, ProfitCalculationAction)
    assert action_daily.period == "daily"

    action_monthly = parser.parse("محاسبه سود ماهانه")
    assert isinstance(action_monthly, ProfitCalculationAction)
    assert action_monthly.period == "monthly"

    action_yearly = parser.parse("سود سالانه رو حساب کن")
    assert isinstance(action_yearly, ProfitCalculationAction)
    assert action_yearly.period == "yearly"

    action_total = parser.parse("محاسبه سود کل")
    assert isinstance(action_total, ProfitCalculationAction)
    assert action_total.period == "total"

    # 4. Another natural phrasing
    action_phrase = parser.parse("فروش رو منهای مبلغ خرید کن تا سود به دست بیاد")
    assert isinstance(action_phrase, ProfitCalculationAction)


def test_execute_profit_calculation_with_dates_and_periods():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()

    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "تاریخ")
    sheet.set_value("B1", "کالا")
    sheet.set_value("C1", "تعداد فروش")
    sheet.set_value("D1", "قیمت فروش")

    sheet.set_value("A2", "1403/08/20")
    sheet.set_value("B2", "لپ‌تاپ")
    sheet.set_value("C2", 2)
    sheet.set_value("D2", 20000000)

    sheet.set_value("A3", "1403/08/21")
    sheet.set_value("B3", "موبایل")
    sheet.set_value("C3", 5)
    sheet.set_value("D3", 10000000)

    sheet.set_value("A4", "1403/09/01")
    sheet.set_value("B4", "موبایل")
    sheet.set_value("C4", 5)
    sheet.set_value("D4", 10000000)

    # Sheet 2: Purchases
    sheet2 = manager.active_driver.create_sheet("Sheet2")
    sheet2.set_value("A1", "نام کالا")
    sheet2.set_value("B1", "تعداد خرید")
    sheet2.set_value("C1", "قیمت خرید")

    sheet2.set_value("A2", "لپ‌تاپ")
    sheet2.set_value("B2", 2)
    sheet2.set_value("C2", 15000000)

    sheet2.set_value("A3", "موبایل")
    sheet2.set_value("B3", 10)
    sheet2.set_value("C3", 8000000)

    executor = ActionExecutor()
    action = ProfitCalculationAction(
        sales_sheet=manager.active_driver.get_sheet_names()[0],
        purchase_sheet="Sheet2",
        period="all",
        create_report_sheet=True
    )
    ctx = ActionContext(manager=manager, sheet=sheet)

    res = executor.execute(action, ctx)

    assert res.success is True
    assert "گزارش_سود" in manager.active_driver.get_sheet_names()
    assert res.data["daily_count"] >= 2
    assert res.data["monthly_count"] >= 2

    # Check report content
    report_sheet = manager.active_driver.get_sheet("گزارش_سود")
    report_text = ""
    for r in range(1, 40):
        for c in ["A", "B", "C", "D", "E", "F"]:
            v = report_sheet.get_value(f"{c}{r}")
            if v:
                report_text += f" {v}"

    assert "جدول تفکیک سود روزانه" in report_text
    assert "جدول تفکیک سود ماهانه" in report_text
    assert "سود خالص کل" in report_text

    manager.close()


def test_execute_profit_calculation_two_sheets():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()

    # Sheet 1: Sales
    sheet1 = manager.get_active_sheet()
    sheet1.set_value("A1", "نام کالا")
    sheet1.set_value("B1", "تعداد فروش")
    sheet1.set_value("C1", "قیمت فروش")

    sheet1.set_value("A2", "لپ‌تاپ")
    sheet1.set_value("B2", 5)
    sheet1.set_value("C2", 20000000)

    sheet1.set_value("A3", "موبایل")
    sheet1.set_value("B3", 10)
    sheet1.set_value("C3", 10000000)

    # Sheet 2: Purchases
    sheet2 = manager.active_driver.create_sheet("Sheet2")
    sheet2.set_value("A1", "نام کالا")
    sheet2.set_value("B1", "تعداد خرید")
    sheet2.set_value("C1", "قیمت خرید")

    sheet2.set_value("A2", "لپ‌تاپ")
    sheet2.set_value("B2", 5)
    sheet2.set_value("C2", 15000000)

    sheet2.set_value("A3", "موبایل")
    sheet2.set_value("B3", 10)
    sheet2.set_value("C3", 8000000)

    executor = ActionExecutor()
    action = ProfitCalculationAction(
        sales_sheet=manager.active_driver.get_sheet_names()[0],
        purchase_sheet="Sheet2",
        create_report_sheet=True
    )
    ctx = ActionContext(manager=manager, sheet=sheet1)

    res = executor.execute(action, ctx)

    assert res.success is True
    assert "گزارش_سود" in manager.active_driver.get_sheet_names()
    assert res.data["total_sales"] == 200000000    # (5*20M) + (10*10M) = 100M + 100M
    assert res.data["total_purchase"] == 155000000 # (5*15M) + (10*8M) = 75M + 80M
    assert res.data["total_profit"] == 45000000    # 200M - 155M = 45M
    assert res.data["item_count"] == 2

    # Verify report sheet
    report_sheet = manager.active_driver.get_sheet("گزارش_سود")
    assert "گزارش جامع سود و زیان" in str(report_sheet.get_value("A1"))
    assert "کل مبلغ فروش" in str(report_sheet.get_value("A5"))
    assert "کل مبلغ خرید" in str(report_sheet.get_value("A6"))
    assert "سود خالص کل" in str(report_sheet.get_value("A7"))
    assert "45,000,000" in str(report_sheet.get_value("B7"))

    manager.close()


def test_execute_profit_calculation_cogs_and_inventory_accounting():
    """
    Test user's specific real-world case:
    Bought 30 pens @ 10,000 Tomans (Total purchase batch = 300,000 Tomans).
    Sold 4 pens @ 12,000 Tomans (Total sales revenue = 48,000 Tomans).
    COGS must be 4 * 10,000 = 40,000 Tomans.
    Gross Profit must be 48,000 - 40,000 = +8,000 Tomans (NOT a loss of 48,000 - 300,000 = -252,000!).
    Remaining inventory must be 26 pens valued at 260,000 Tomans.
    """
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()

    # Sheet 1: Sales
    sheet1 = manager.get_active_sheet()
    sheet1.set_value("A1", "نام کالا")
    sheet1.set_value("B1", "تعداد فروش")
    sheet1.set_value("C1", "قیمت فروش")

    sheet1.set_value("A2", "خودکار")
    sheet1.set_value("B2", 4)
    sheet1.set_value("C2", 12000)

    # Sheet 2: Purchases
    sheet2 = manager.active_driver.create_sheet("Sheet2")
    sheet2.set_value("A1", "نام کالا")
    sheet2.set_value("B1", "تعداد خرید")
    sheet2.set_value("C1", "قیمت خرید")

    sheet2.set_value("A2", "خودکار")
    sheet2.set_value("B2", 30)
    sheet2.set_value("C2", 10000)

    executor = ActionExecutor()
    action = ProfitCalculationAction(
        sales_sheet=manager.active_driver.get_sheet_names()[0],
        purchase_sheet="Sheet2",
        create_report_sheet=True
    )
    ctx = ActionContext(manager=manager, sheet=sheet1)

    res = executor.execute(action, ctx)

    assert res.success is True
    assert res.data["total_sales"] == 48000
    assert res.data["total_purchase"] == 40000  # Cost of Goods Sold (4 * 10,000), not batch total 300,000
    assert res.data["total_profit"] == 8000    # +8,000 Tomans profit!
    assert res.data["remaining_stock"] == 26   # 30 - 4 = 26 pens remaining in stock
    assert res.data["stock_value"] == 260000   # 26 * 10,000 = 260,000 Tomans inventory value

    # Check report sheet formulas and values
    report_sheet = manager.active_driver.get_sheet("گزارش_سود")
    report_cells = {}
    for r in range(1, 25):
        for col_idx in range(1, 13):
            c_letter = chr(ord('A') + col_idx - 1)
            v = report_sheet.get_value(f"{c_letter}{r}")
            if v is not None:
                report_cells[f"{c_letter}{r}"] = v

    # Check that profit is 8,000 in KPI and table
    assert any("8,000" in str(v) for v in report_cells.values())
    assert any("260,000" in str(v) for v in report_cells.values())

    manager.close()

