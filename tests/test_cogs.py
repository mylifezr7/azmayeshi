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

    # 1. On empty workbook without demo flag, must reject and warn user
    action_empty = CostOfGoodsSoldAction(method="periodic", create_report_sheet=True)
    res_empty = AccountingActionHandler.execute_cogs(action_empty, ctx)
    assert res_empty.success is False
    assert "اطلاعات کافی نیست" in res_empty.message_persian

    # 2. When allow_demo=True, verify benchmark demonstration report is generated
    action = CostOfGoodsSoldAction(method="periodic", create_report_sheet=True, allow_demo=True)
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

    action = CostOfGoodsSoldAction(method="compare", create_report_sheet=True, allow_demo=True)
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


def test_cogs_nine_sheets_comprehensive_case_study():
    """تست اعتبارسنجی صورت بهای تمام شده و اسکن همزمان ۹ شیت سناریوی پیچیده."""
    mgr = ExcelConnectionManager()
    mgr.connect()

    # 1. شیت موجودی اول دوره
    sh_beg = mgr.create_sheet("موجودی اول دوره")
    sh_beg.set_value("A1", "کد")
    sh_beg.set_value("B1", "شرح")
    sh_beg.set_value("C1", "تعداد")
    sh_beg.set_value("D1", "بهای واحد")
    sh_beg.set_value("A2", "A101"); sh_beg.set_value("C2", 35); sh_beg.set_value("D2", 29_228_571.43)
    sh_beg.set_value("A3", "B205"); sh_beg.set_value("C3", 50); sh_beg.set_value("D3", 18_780_000.0)
    sh_beg.set_value("A4", "C330"); sh_beg.set_value("C4", 12); sh_beg.set_value("D4", 68_500_000.0)
    sh_beg.set_value("A5", "D410"); sh_beg.set_value("C5", 25); sh_beg.set_value("D5", 12_700_000.0)

    # 2. شیت فاکتورهای خرید
    sh_p = mgr.create_sheet("فاکتورهای خرید")
    sh_p.set_value("A1", "فاکتور")
    sh_p.set_value("C1", "کد کالا")
    sh_p.set_value("E1", "تعداد")
    sh_p.set_value("F1", "قیمت واحد ناخالص")
    sh_p.set_value("G1", "تخفیف تجاری%")
    sh_p.set_value("A2", "PU001"); sh_p.set_value("C2", "A101"); sh_p.set_value("E2", 30); sh_p.set_value("F2", 31_000_000); sh_p.set_value("G2", 4)
    sh_p.set_value("A3", "PU002"); sh_p.set_value("C3", "B205"); sh_p.set_value("E3", 40); sh_p.set_value("F3", 19_800_000); sh_p.set_value("G3", 0)

    # 3. شیت برگشت از خرید
    sh_pr = mgr.create_sheet("برگشت از خرید")
    sh_pr.set_value("A1", "سند")
    sh_pr.set_value("D1", "کد کالا")
    sh_pr.set_value("E1", "تعداد برگشتی")
    sh_pr.set_value("F1", "بهای واحد پس از تخفیف")
    sh_pr.set_value("A2", "PR001"); sh_pr.set_value("D2", "B205"); sh_pr.set_value("E2", 5); sh_pr.set_value("F2", 19_800_000)

    # 4. شیت هزینه های حمل ورودی
    sh_fr = mgr.create_sheet("هزینه های حمل")
    sh_fr.set_value("A1", "سند")
    sh_fr.set_value("E1", "مبلغ")
    sh_fr.set_value("F1", "نوع هزینه")
    sh_fr.set_value("A2", "FR001"); sh_fr.set_value("E2", 18_000_000); sh_fr.set_value("F2", "حمل ورودی")
    sh_fr.set_value("A3", "FR010"); sh_fr.set_value("E3", 22_000_000); sh_fr.set_value("F3", "حمل فروش؛ هزینه دوره")

    # 5. شیت فاکتورهای فروش
    sh_s = mgr.create_sheet("فاکتورهای فروش")
    sh_s.set_value("A1", "فاکتور")
    sh_s.set_value("D1", "کد کالا")
    sh_s.set_value("E1", "تعداد")
    sh_s.set_value("F1", "قیمت واحد فروش")
    sh_s.set_value("G1", "تخفیف فروش%")
    sh_s.set_value("A2", "SA001"); sh_s.set_value("D2", "A101"); sh_s.set_value("E2", 18); sh_s.set_value("F2", 39_500_000); sh_s.set_value("G2", 5)

    # 6. شیت برگشت از فروش
    sh_sr = mgr.create_sheet("برگشت از فروش")
    sh_sr.set_value("A1", "سند")
    sh_sr.set_value("C1", "کد کالا")
    sh_sr.set_value("D1", "فاکتور مرجع")
    sh_sr.set_value("E1", "تعداد")
    sh_sr.set_value("F1", "وضعیت کالا")
    sh_sr.set_value("A2", "SR001"); sh_sr.set_value("C2", "A101"); sh_sr.set_value("D2", "SA001"); sh_sr.set_value("E2", 2); sh_sr.set_value("F2", "سالم")

    ctx = ActionContext(sheet=sh_beg, manager=mgr)
    action = CostOfGoodsSoldAction(method="fifo")
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    cogs_sh = mgr.get_sheet("صورت بهای تمام شده (COGS)")

    # 1. موجودی اول دوره باید بزرگتر از صفر و دقیق باشد
    assert round(cogs_sh.get_value("D5"), 0) == 3_101_500_000.0
    # 2. خرید ناخالص (30*31M + 40*19.8M = 930M + 792M = 1,722M)
    assert cogs_sh.get_value("C6") == 1_722_000_000.0
    # 3. هزینه حمل فقط ورودی لحاظ شود (FR001 = 18M) و FR010 (22M فروش) حذف شود
    assert cogs_sh.get_value("C7") == 18_000_000.0
    # 4. تخفیفات تجاری خرید (37.2M) + برگشت از خرید (99M) = 136.2M
    assert cogs_sh.get_value("C8") == 136_200_000.0
    # 5. درآمد فروش خالص پس از کسر تخفیف فروش و برگشت از فروش SA001
    # 18 * 39.5M = 711M; Disc 5% = 35.55M -> Net unit = 37.525M. Return 2 units = 75.05M -> Net = 600.4M
    assert cogs_sh.get_value("D14") == 600_400_000.0


def test_cogs_challenge_4_comprehensive():
    """Verify Challenge 4 (Mehr 1405) with all 10 sheets, freight allocation by weight, and discounts."""
    import os
    if not os.path.exists("Challenge_4_Accounting.xlsx"):
        pytest.skip("Challenge_4_Accounting.xlsx not found")

    mgr = ExcelConnectionManager()
    mgr.connect("Challenge_4_Accounting.xlsx")
    sheet = mgr.get_active_sheet()
    ctx = ActionContext(sheet=sheet, manager=mgr)

    action = CostOfGoodsSoldAction(method="fifo")
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    assert "تطبیق‌های حسابرسی و انبارگردانی اعمال‌شده در چالش ۴" in res.message_persian

    cogs_sh = mgr.get_sheet("صورت بهای تمام شده (COGS)")
    assert cogs_sh.get_value("D5") == 1_654_350_000.0   # موجودی اول دوره
    assert cogs_sh.get_value("C6") == 2_087_690_000.0   # خرید ناخالص
    assert cogs_sh.get_value("C7") == 40_500_000.0      # کرایه حمل ورودی
    assert cogs_sh.get_value("C8") == 99_690_000.0      # برگشت و تخفیفات خرید
    assert cogs_sh.get_value("D11") == 1_615_427_740.0  # موجودی پایان دوره (فایفو)
    assert cogs_sh.get_value("D14") == 2_598_410_000.0  # فروش خالص


def test_cogs_challenge_5_comprehensive():
    """Verify Challenge 5 (Aban 1405) with multi-item invoices, physical counts, consignments, and per-invoice freight."""
    import os
    if not os.path.exists("Challenge_5_Accounting.xlsx"):
        pytest.skip("Challenge_5_Accounting.xlsx not found")

    mgr = ExcelConnectionManager()
    mgr.connect("Challenge_5_Accounting.xlsx")
    sheet = mgr.get_active_sheet()
    ctx = ActionContext(sheet=sheet, manager=mgr)

    action = CostOfGoodsSoldAction(method="fifo")
    res = AccountingActionHandler.execute_cogs(action, ctx)

    assert res.success is True
    assert "تطبیق‌های حسابرسی و انبارگردانی اعمال‌شده در چالش ۵" in res.message_persian

    cogs_sh = mgr.get_sheet("صورت بهای تمام شده (COGS)")
    assert cogs_sh.get_value("D5") == 1_297_750_000.0   # موجودی اول دوره
    assert cogs_sh.get_value("C6") == 2_042_740_000.0   # خرید ناخالص
    assert cogs_sh.get_value("C7") == 42_500_000.0      # کرایه حمل ورودی
    assert cogs_sh.get_value("C8") == 67_136_792.0      # برگشت و تخفیفات خرید
    assert cogs_sh.get_value("D11") == 1_599_190_572.0  # موجودی پایان دوره (فایفو)
    assert cogs_sh.get_value("D14") == 2_185_250_000.0  # فروش خالص



