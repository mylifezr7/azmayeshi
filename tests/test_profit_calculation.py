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

    # 2. General profit command
    action_simple = parser.parse("محاسبه سود")
    assert isinstance(action_simple, ProfitCalculationAction)

    # 3. Another natural phrasing
    action_phrase = parser.parse("فروش رو منهای مبلغ خرید کن تا سود به دست بیاد")
    assert isinstance(action_phrase, ProfitCalculationAction)


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
