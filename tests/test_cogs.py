"""
Tests for Cost of Goods Sold (COGS) / محاسبه بهای تمام شده کالای فروش رفته.
"""
import pytest
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import CostOfGoodsSoldAction
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.accounting_actions import AccountingActionHandler
from excel_assistant.actions.base import ActionContext
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.ui.command_search import search_commands


def test_parse_cogs_phrases():
    """Verify parsing various Persian phrases for Cost of Goods Sold."""
    parser = PersianIntentParser()

    # 1. Periodic / General
    action1 = parser.parse("محاسبه بهای تمام شده")
    assert isinstance(action1, CostOfGoodsSoldAction)
    assert action1.method == "periodic"

    action2 = parser.parse("بهای تمام شده کالای فروش رفته")
    assert isinstance(action2, CostOfGoodsSoldAction)
    assert action2.method == "periodic"

    action3 = parser.parse("cogs")
    assert isinstance(action3, CostOfGoodsSoldAction)
    assert action3.method == "periodic"

    action4 = parser.parse("قیمت تمام شده کالای فروش رفته")
    assert isinstance(action4, CostOfGoodsSoldAction)

    # 2. FIFO / LIFO / Average / Compare
    action_fifo = parser.parse("محاسبه بهای تمام شده به روش فایفو")
    assert isinstance(action_fifo, CostOfGoodsSoldAction)
    assert action_fifo.method == "fifo"

    action_lifo = parser.parse("محاسبه بهای تمام شده به روش لایفو")
    assert isinstance(action_lifo, CostOfGoodsSoldAction)
    assert action_lifo.method == "lifo"

    action_avg = parser.parse("بهای تمام شده با میانگین موزون")
    assert isinstance(action_avg, CostOfGoodsSoldAction)
    assert action_avg.method == "average"

    action_cmp = parser.parse("مقایسه بهای تمام شده فایفو و لایفو")
    assert isinstance(action_cmp, CostOfGoodsSoldAction)
    assert action_cmp.method == "compare"


def test_execute_cogs_periodic_standard():
    """Verify execution of standard periodic Statement of COGS."""
    mgr = ExcelConnectionManager()
    mgr.connect()
    sheet = mgr.get_active_sheet()
    ctx = ActionContext(sheet=sheet, manager=mgr)

    action = CostOfGoodsSoldAction(method="periodic", create_report_sheet=True)
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    assert "بهای تمام شده کالای فروش رفته" in res.message_persian
    assert "صورت بهای تمام شده (COGS)" in mgr.list_sheets()

    cogs_sheet = mgr.get_sheet("صورت بهای تمام شده (COGS)")
    # Check Header
    assert "صورت محاسبه بهای تمام شده" in str(cogs_sheet.get_value("A1"))
    # Check items
    assert cogs_sheet.get_value("B5") == "موجودی کالای اول دوره (Beginning Inventory)"
    assert cogs_sheet.get_value("D5") == 120_000_000.0
    assert cogs_sheet.get_value("C6") == 450_000_000.0
    assert cogs_sheet.get_value("C7") == 15_000_000.0
    assert cogs_sheet.get_value("C8") == 25_000_000.0
    # Check Net Purchases Formula
    assert cogs_sheet.get_value("D9") == "=C6+C7-C8"
    # Check Goods Available Formula
    assert cogs_sheet.get_value("D10") == "=D5+D9"
    # Check Ending Inventory
    assert cogs_sheet.get_value("D11") == 140_000_000.0
    # Check COGS Formula
    assert cogs_sheet.get_value("D12") == "=D10-D11"
    # Check Gross Profit Formula
    assert cogs_sheet.get_value("D15") == "=D14-D12"


def test_execute_cogs_comparison_methods():
    """Verify execution with FIFO vs LIFO vs Average comparison block."""
    mgr = ExcelConnectionManager()
    mgr.connect()
    sheet = mgr.get_active_sheet()
    ctx = ActionContext(sheet=sheet, manager=mgr)

    action = CostOfGoodsSoldAction(method="compare", create_report_sheet=True)
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    cogs_sheet = mgr.get_sheet("صورت بهای تمام شده (COGS)")
    # Row 19 should have the comparison title
    assert "مقایسه تحلیلی بهای تمام شده" in str(cogs_sheet.get_value("A19"))
    # Row 21 should be FIFO
    assert "فایفو" in str(cogs_sheet.get_value("B21"))
    # Row 22 should be LIFO
    assert "لایفو" in str(cogs_sheet.get_value("B22"))
    # Row 23 should be Weighted Average
    assert "میانگین" in str(cogs_sheet.get_value("B23"))


def test_execute_cogs_with_detected_sheet_data():
    """Verify COGS calculation extracts values if present in active sheet."""
    mgr = ExcelConnectionManager()
    mgr.connect()
    sheet = mgr.get_active_sheet()

    # Pre-populate sample trial balance with COGS accounts
    sheet.set_value("A1", "حساب موجودی کالای اول دوره")
    sheet.set_value("B1", 85_000_000)
    sheet.set_value("A2", "خرید طی دوره")
    sheet.set_value("B2", 300_000_000)
    sheet.set_value("A3", "کرایه حمل کالا")
    sheet.set_value("B3", 10_000_000)
    sheet.set_value("A4", "برگشت از خرید و تخفیفات")
    sheet.set_value("B4", 15_000_000)
    sheet.set_value("A5", "موجودی کالای پایان دوره")
    sheet.set_value("B5", 90_000_000)
    sheet.set_value("A6", "درآمد فروش کالا")
    sheet.set_value("B6", 500_000_000)

    ctx = ActionContext(sheet=sheet, manager=mgr)
    action = CostOfGoodsSoldAction(method="periodic")
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    cogs_sheet = mgr.get_sheet("صورت بهای تمام شده (COGS)")
    assert cogs_sheet.get_value("D5") == 85_000_000.0
    assert cogs_sheet.get_value("C6") == 300_000_000.0
    assert cogs_sheet.get_value("C7") == 10_000_000.0
    assert cogs_sheet.get_value("C8") == 15_000_000.0
    assert cogs_sheet.get_value("D11") == 90_000_000.0
    assert cogs_sheet.get_value("D14") == 500_000_000.0


def test_search_cogs():
    """Verify search engine returns COGS commands when searching for 'بهای تمام شده' or 'cogs'."""
    res_baha = search_commands("بهای تمام شده", limit=5)
    assert len(res_baha) > 0
    assert any("بهای تمام شده" in r["title"] for r in res_baha)
    assert res_baha[0]["cmd"] == "محاسبه بهای تمام شده کالای فروش رفته"

    res_cogs = search_commands("cogs", limit=5)
    assert len(res_cogs) > 0
    assert any("بهای تمام شده" in r["title"] or "cogs" in str(r["keywords"]).lower() for r in res_cogs)


def test_execute_cogs_with_transactional_multi_sheets():
    """Verify COGS dynamically scans multi-sheet workbooks (sales, purchases, inventory)."""
    mgr = ExcelConnectionManager()
    mgr.connect()
    sh_sales = mgr.create_sheet("فاکتور فروش")
    sh_purch = mgr.create_sheet("فاکتور خرید")
    sh_inv = mgr.create_sheet("حرکت انبار")

    # Purchase invoices
    sh_purch.set_value("A1", "سند")
    sh_purch.set_value("E1", "کد کالا")
    sh_purch.set_value("G1", "تعداد")
    sh_purch.set_value("H1", "قیمت واحد")

    sh_purch.set_value("A2", "PI501")
    sh_purch.set_value("E2", "IT01")
    sh_purch.set_value("G2", 5)
    sh_purch.set_value("H2", 690_000_000)

    sh_purch.set_value("A3", "PI502")
    sh_purch.set_value("E3", "IT02")
    sh_purch.set_value("G3", 20)
    sh_purch.set_value("H3", 92_000_000)

    sh_purch.set_value("A4", "PR601")  # purchase return
    sh_purch.set_value("E4", "IT02")
    sh_purch.set_value("G4", 2)
    sh_purch.set_value("H4", 92_000_000)

    # Sales invoices
    sh_sales.set_value("A1", "سند")
    sh_sales.set_value("E1", "کد کالا")
    sh_sales.set_value("G1", "تعداد")
    sh_sales.set_value("H1", "قیمت واحد")
    sh_sales.set_value("I1", "تخفیف%")

    sh_sales.set_value("A2", "SI301")
    sh_sales.set_value("E2", "IT01")
    sh_sales.set_value("G2", 2)
    sh_sales.set_value("H2", 980_000_000)
    sh_sales.set_value("I2", 3)

    sh_sales.set_value("A3", "SI302")
    sh_sales.set_value("E3", "IT02")
    sh_sales.set_value("G3", 8)
    sh_sales.set_value("H3", 145_000_000)
    sh_sales.set_value("I3", 0)

    # Inventory movements
    sh_inv.set_value("A1", "تاریخ")
    sh_inv.set_value("B1", "سند")
    sh_inv.set_value("C1", "کد کالا")
    sh_inv.set_value("E1", "نوع حرکت")
    sh_inv.set_value("F1", "تعداد +/-")

    sh_inv.set_value("A2", "03/01")
    sh_inv.set_value("B2", "OP01")
    sh_inv.set_value("C2", "IT01")
    sh_inv.set_value("E2", "اول دوره")
    sh_inv.set_value("F2", 3)

    sh_inv.set_value("A3", "03/01")
    sh_inv.set_value("B3", "OP02")
    sh_inv.set_value("C3", "IT02")
    sh_inv.set_value("E3", "اول دوره")
    sh_inv.set_value("F3", 12)

    ctx = ActionContext(sheet=sh_purch, manager=mgr)
    action = CostOfGoodsSoldAction(method="fifo")
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    cogs_sh = mgr.get_sheet("صورت بهای تمام شده (COGS)")
    
    # Beginning inventory: 3*690M + 12*92M = 2,070M + 1,104M = 3,174,000,000
    assert cogs_sh.get_value("D5") == 3_174_000_000.0
    # Gross purchases: 5*690M + 20*92M = 3,450M + 1,840M = 5,290,000,000
    assert cogs_sh.get_value("C6") == 5_290_000_000.0
    # Purchase returns: 2*92M = 184,000,000
    assert cogs_sh.get_value("C8") == 184_000_000.0
    # Net sales: (2*980M * 0.97) + (8*145M) = 1,901,200,000 + 1,160,000,000 = 3,061,200,000
    assert cogs_sh.get_value("D14") == 3_061_200_000.0

    # Ensure no notes start with '=' (which causes #NAME? in Excel)
    for r in range(5, 18):
        val = str(cogs_sh.get_value(f"E{r}") or "")
        assert not val.startswith("="), f"Row {r} note begins with '=': {val}"

